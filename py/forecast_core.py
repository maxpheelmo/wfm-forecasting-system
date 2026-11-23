# -*- coding: utf-8 -*-
"""
Ядро прогноза нагрузки: чистые функции на дневных рядах (без БД и Excel).

Дневные ряды (DataFrame с индексом-датой):
  voice_entry      звонки, начавшиеся в основной IVR / IVR второй линии
  voice_entry_queue  из них дошли до рабочей очереди
  voice_direct     в рабочей очереди, минуя основной IVR (партнёрская линия, идентификация, переводы)
  chat_entry       чат-сессии на входе КЦ
  chat_entry_queue   из них дошли до рабочей очереди
  mail             письма в «Входящие Email»

Модель (одна и та же для любого горизонта):
  объём дня   = уровень × индекс дня недели × запас на внеплановые всплески × события
  доля до очереди = по текущему режиму (робот/бот), с будущими сменами режима из календаря
  очереди         = вход × доля (+ прямые для голоса)
  30 мин      = день × профиль дня (будни / Сб / Вс; праздник = Вс)

Прошлые аномалии — автоматически: день отличается от медианы того же дня недели (±3 недели) больше чем на
SPIKE_UP / SPIKE_DOWN → исключается из уровня и индексов; его избыток учитывается средним «запасом».
Смена режима доли — автоматический поиск излома (бинарная сегментация по сглаженной доле) + календарь.
"""
import numpy as np
import pandas as pd

WD = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]

DEFAULTS = dict(
    level_weeks=12,         # окно уровня (бэктест 22.10: 12 нед лучше 8 и 6 на горизонте 1–3 мес.)
    weekday_weeks=12,       # окно индексов дня недели
    share_weeks=6,          # окно доли до очереди (внутри текущего режима)
    reserve_weeks=12,       # окно оценки запаса на всплески
    spike_up=0.5,           # +50% к норме — всплеск
    spike_down=-0.5,        # −50% — провал
    use_reserve=True,       # запас на внеплановые всплески для входов (голос, чат вход, почта)
    reserve_queue={"voice": True, "chat": False},   # чат: всплески входа поглощает бот — в очереди запас не нужен
    clean=True,
    startup_days=7,
    trend="none",           # none | damped
    cp_min_days=14,         # минимальный режим для поиска излома
    cp_min_jump=0.04,       # минимальный скачок доли (4 п.п.)
    volume_breaks=("voice_direct",),  # ряды, где ищем изломы уровня объёма (уровень — после последнего излома)
    vol_min_jump=0.25,      # минимальный скачок объёма (±25%) для излома
    half_life=None,         # квартал: обычная медиана; недельный пересчёт: 3 (свежие недели весят больше)
    promo_segments=False,   # акция: один множитель на весь период (бэктест 22.10 лучше, чем 1-я неделя / остальные)
    promo_shrink=True,      # мало прошлых запусков -> эффект осторожнее: 1 + (м − 1)·n/(n+1)
    dom_effect=False,       # день месяца (1-е, 2-е, предпоследний, последний): коэффициент по истории
    dom_weeks=52,           # окно оценки коэффициентов дня месяца
    dom_min_n=5,            # минимум случаев для коэффициента
    new_year=True,          # новогодний коэффициент 1–10.01 (если ≥ 2 лет истории); False — только «как воскресенье»
    incident_days=None,     # дни инцидентов Service Desk (incidents.py): в такие дни порог всплеска ниже
    inc_spike_up=0.2,       # порог всплеска в день инцидента (+20% к норме вместо spike_up)
    promo_reserve=False,    # квартал (№16): равномерный запас на крупные акции = избыток прошлых запусков «акция» / весь объём
    promo_reserve_weeks=52, # окно оценки запаса на акции (крупные акции — раз в 3–4 месяца, нужно ≥ года)
)

VOLUME_SERIES = ["voice_entry", "voice_direct", "chat_entry", "mail"]
SHARES = {"voice": ("voice_entry", "voice_entry_queue"), "chat": ("chat_entry", "chat_entry_queue")}


# ---------------------------------------------------------------------------
# Очистка и режимы
# ---------------------------------------------------------------------------
def weekday_norm(s):
    """Медиана того же дня недели за ±3 недели (без самого дня)."""
    return pd.concat([s.shift(7 * k) for k in (-3, -2, -1, 1, 2, 3)], axis=1).median(axis=1, skipna=True)


def weekday_norm_past(s):
    """Только прошлое: медиана того же дня недели за 1–4 недели назад (для дней у края истории)."""
    return pd.concat([s.shift(7 * k) for k in (1, 2, 3, 4)], axis=1).median(axis=1, skipna=True)


def series_start(s):
    tail = s.dropna().iloc[-60:]
    if tail.empty or tail.median() <= 0:
        return None
    ok = s[s >= 0.3 * tail.median()]
    return ok.index.min() if len(ok) else None


def detect_spikes(s, cfg, exclude=()):
    """-> (маска аномальных дней, норма). Норма — центрированная медиана; у правого края — прошлая."""
    st = series_start(s)
    if st is None:
        return pd.Series(False, index=s.index), s * np.nan
    s = s.where(s.index >= st)
    norm = weekday_norm(s)
    past = weekday_norm_past(s)
    edge = s.index > s.index.max() - pd.Timedelta(days=21)
    norm = norm.where(~edge, past.combine_first(norm))
    dev = s / norm - 1
    mask = ((dev > cfg["spike_up"]) | (dev < cfg["spike_down"])).fillna(False)
    inc = cfg.get("incident_days")
    if inc is not None and len(inc):
        # день инцидента (сбой источник / клиентского онлайна): всплеск засчитывается при меньшем отклонении
        mask |= (s.index.isin(pd.DatetimeIndex(inc)) & (dev > cfg.get("inc_spike_up", 0.2))).fillna(False)
    mask &= s.index >= st + pd.Timedelta(days=cfg["startup_days"])
    if len(exclude):
        mask |= s.index.isin(pd.DatetimeIndex(exclude))
    return mask, norm


def find_breaks(share, cfg, min_days=None, min_jump=None):
    """Изломы уровня доли: бинарная сегментация по 7-дневной медиане. -> список дат начала нового режима."""
    min_days = min_days or cfg["cp_min_days"]
    min_jump = min_jump or cfg["cp_min_jump"]
    s = share.dropna()
    if len(s) < 2 * min_days:
        return []
    sm = s.rolling(7, center=True, min_periods=4).median()
    out = []

    def split(lo, hi):
        x = sm.iloc[lo:hi].values
        n = len(x)
        if n < 2 * min_days:
            return
        best, best_i = None, None
        total = ((x - x.mean()) ** 2).sum()
        for i in range(min_days, n - min_days + 1):
            a, b = x[:i], x[i:]
            sse = ((a - a.mean()) ** 2).sum() + ((b - b.mean()) ** 2).sum()
            if best is None or sse < best:
                best, best_i = sse, i
        a, b = x[:best_i], x[best_i:]
        if abs(np.median(b) - np.median(a)) >= min_jump and best < 0.6 * total:
            out.append(sm.index[lo + best_i])
            split(lo, lo + best_i)
            split(lo + best_i, hi)

    split(0, len(sm))
    return sorted(out)


def merge_breaks(auto, manual, min_days=14):
    """Объединить изломы: ближе min_days друг к другу — один излом (календарный в приоритете)."""
    out = sorted(set(pd.Timestamp(x) for x in manual))
    for b in sorted(auto):
        if all(abs((b - m).days) >= min_days for m in out):
            out.append(b)
    out = sorted(out)
    merged = []
    for b in out:
        if merged and (b - merged[-1]).days < min_days and b not in manual:
            continue
        merged.append(b)
    return merged


# ---------------------------------------------------------------------------
# Компоненты
# ---------------------------------------------------------------------------
def weekday_index(s, mask, end, weeks):
    w = weeks
    while True:
        lo = end - pd.Timedelta(weeks=w)
        x = s[(s.index > lo) & (s.index <= end) & ~mask].dropna()
        if len(x) >= 21 or w >= max(26, weeks):
            break
        w += 2
    if x.empty:
        return pd.Series(1.0, index=range(7))
    m = x.groupby(x.index.weekday).mean().reindex(range(7))
    m = m.fillna(m.mean())
    return m / m.mean()


def level_at(s, mask, end, weeks, wd_idx, trend="none", min_days=10, half_life=None):
    # если в окне мало чистых дней (длинная акция/турнир исключены) — расширяем окно назад, до 26 недель
    w = weeks
    while True:
        lo = end - pd.Timedelta(weeks=w)
        x = s[(s.index > lo) & (s.index <= end) & ~mask].dropna()
        if len(x) >= min_days or w >= max(26, weeks):
            break
        w += 2
    if x.empty:
        return 0.0, 0.0, {"from": end, "to": end, "days": 0, "weeks": float(w)}
    des = x / wd_idx.reindex(x.index.weekday).values
    if half_life:
        # взвешенная медиана: свежие недели весят больше (вес 0,5 каждые half_life недель)
        age = (end - des.index).days.values / 7.0
        wts = 0.5 ** (age / half_life)
        order = np.argsort(des.values)
        cw = np.cumsum(wts[order]) / wts.sum()
        level = float(des.values[order][np.searchsorted(cw, 0.5)])
    else:
        level = float(des.median())
    slope = 0.0
    if trend == "damped" and len(des) >= 28:
        wk = des.resample("W").mean().dropna()
        if len(wk) >= 4:
            t = np.arange(len(wk))
            slopes = [(wk.iloc[j] - wk.iloc[i]) / (j - i) for i in range(len(wk)) for j in range(i + 1, len(wk))]
            slope = float(np.median(slopes)) / 7.0          # в день
            level = float(wk.iloc[-2:].mean()) if len(wk) >= 2 else level
    return level, slope, {"from": x.index.min(), "to": x.index.max(), "days": int(len(x)), "weeks": float(w)}


def reserve_rate(s, mask, norm, end, weeks, planned=()):
    """Доля объёма, которую дают внеплановые всплески: Σ(факт − норма) по всплескам / Σ факт."""
    lo = end - pd.Timedelta(weeks=weeks)
    win = (s.index > lo) & (s.index <= end)
    spikes = win & mask & ~s.index.isin(pd.DatetimeIndex(planned)) & (s > norm)
    tot = float(s[win].sum())
    if not tot:
        return 0.0, 0.0, 0.0
    excess = float((s[spikes] - norm[spikes]).sum())
    return float(min(0.3, excess / tot)), excess, tot


def share_at(entry, reached, mask, end, weeks, breaks):
    start = max([b for b in breaks if b <= end], default=reached.dropna().index.min())
    lo = max(start, end - pd.Timedelta(weeks=weeks))
    win = (entry.index >= lo) & (entry.index <= end) & ~mask & entry.notna().values & reached.notna().values
    e, r = float(entry[win].sum()), float(reached[win].sum())
    days = int(win.sum())
    return (float(r / e) if e else np.nan), start, days, e, r, lo


def dom_key(d):
    """Позиция дня в месяце, для которой проверяем особый эффект (смена статусов VIP, выплаты)."""
    last = (d + pd.offsets.MonthEnd(0)).day
    return {1: "1-е", 2: "2-е", last: "последний", last - 1: "предпоследний"}.get(d.day)


def dom_factors(s, excl_days, origin, weeks=52, min_n=5):
    """Коэффициенты дня месяца: медиана (факт / медиана того же дня недели ±3 нед) по прошлым случаям."""
    lo = origin - pd.Timedelta(weeks=weeks)
    x = s[(s.index > lo) & (s.index <= origin - pd.Timedelta(days=21))].dropna()   # нужна норма с двух сторон
    norm = weekday_norm(s).reindex(x.index)
    r = (x / norm).replace([np.inf, -np.inf], np.nan)
    r = r[~r.index.isin(pd.DatetimeIndex(list(excl_days)))].dropna()
    keys = pd.Series([dom_key(d) for d in r.index], index=r.index)
    out = {}
    for k, idx in keys.dropna().groupby(keys.dropna()).groups.items():
        v = r.loc[idx]
        if len(v) >= min_n:
            out[k] = (float(v.median()), int(len(v)))
    return out


def month_factors(s, origin, holidays=(), promo_days=(), min_years=2, min_days=10):
    """Индекс календарного месяца к обычному дню того же года. Только для года (№138).

    Внутри каждого года день делится на медиану своего дня недели, затем на медиану года.
    Так рост уровня между годами не попадает в индекс, а провал на целый месяц — попадает
    (локальная норма ±3 недели его прячет). Не входят праздники, 1–10 января и дни акций.
    Месяц берётся, если таких лет ≥ min_years. Иначе множитель 1. Среднее 12 множителей = 1.
    -> (dict 1..12 -> множитель, dict 1..12 -> сколько лет) или None.
    """
    st = series_start(s)
    s = s.dropna().astype(float)
    if st is not None:
        s = s[s.index >= st]
    s = s[s.index <= origin]
    if len(s) < 60:
        return None
    ban = set(pd.DatetimeIndex(holidays).normalize()) | set(pd.DatetimeIndex(list(promo_days)).normalize())
    ratios = {m: [] for m in range(1, 13)}
    years = {m: 0 for m in range(1, 13)}
    for year, ys in s.groupby(s.index.year):
        wd = ys.groupby(ys.index.weekday).median()
        if wd.median() <= 0 or wd.isna().all():
            continue
        wd = wd / wd.mean()
        des = ys / wd.reindex(ys.index.weekday).to_numpy()
        clean = []
        for d, v in des.items():
            d = pd.Timestamp(d).normalize()
            if d in ban or (d.month == 1 and d.day <= 10) or not np.isfinite(v) or v <= 0:
                continue
            clean.append((d, float(v)))
        if len(clean) < 60:
            continue
        base = float(np.median([v for _, v in clean]))
        if not np.isfinite(base) or base <= 0:
            continue
        by_m = {m: [] for m in range(1, 13)}
        for d, v in clean:
            by_m[d.month].append(v / base)
        for m, vals in by_m.items():
            if len(vals) >= min_days:
                ratios[m].append(float(np.median(vals)))
                years[m] += 1
    if sum(years.values()) < min_years:
        return None
    factors = {}
    for m in range(1, 13):
        factors[m] = float(np.median(ratios[m])) if years[m] >= min_years else 1.0
    mean = float(np.mean(list(factors.values())))
    if not np.isfinite(mean) or mean <= 0:
        return None
    factors = {m: float(v / mean) for m, v in factors.items()}
    return factors, years


def robot_regimes(robot, entry, on_thr=0.05, off_thr=0.02, persist=5):
    """Голосовой робот: включён/выключен по доле входящих звонков, прошедших через робота.
    Гистерезис: включается при доле > on_thr, выключается при доле < off_thr; новое состояние должно держаться
    ≥ persist дней подряд (иначе — дребезг около порога, режим не меняем). -> (даты смен, включён ли сейчас, периоды)."""
    rs = (robot / entry).where(entry > 0).dropna()
    if rs.empty:
        return [], None, []
    sm = rs.rolling(3, min_periods=2).median()
    state = bool(sm.iloc[0] > on_thr)
    start, run_start, run_len = sm.index[0], None, 0
    changes, periods = [], []
    for d, v in sm.items():
        flip = (not state and v > on_thr) or (state and v < off_thr)
        if flip:
            run_start = run_start or d
            run_len += 1
            if run_len >= persist:
                periods.append((start, run_start - pd.Timedelta(days=1), state))
                changes.append(run_start)
                state, start, run_start, run_len = not state, run_start, None, 0
        else:
            run_start, run_len = None, 0
    periods.append((start, sm.index[-1], state))
    return changes, state, periods


def promo_runs(ev, action="акция"):
    """Строки календаря с действием «акция» -> список dict(name, channel, a, b, m)."""
    out = []
    for r in ev.itertuples():
        if str(getattr(r, "action", "")) != action:
            continue
        a = pd.Timestamp(r.date_from)
        b = pd.Timestamp(r.date_to) if pd.notna(r.date_to) else a
        name = str(getattr(r, "name", "") or "").strip().upper() or str(getattr(r, "comment", "") or "").strip().upper()
        out.append(dict(name=name, channel=str(r.channel), a=a, b=b, m=r.multiplier))
    # одинаковые акции (название + канал), пересекающиеся по датам, — один запуск (защита от дублей в календаре)
    out.sort(key=lambda x: (x["name"], x["channel"], x["a"]))
    merged = []
    for x in out:
        y = merged[-1] if merged else None
        if y and y["name"] == x["name"] and y["channel"] == x["channel"] and x["a"] <= y["b"] + pd.Timedelta(days=1):
            y["b"] = max(y["b"], x["b"])
            if pd.isna(y["m"]):
                y["m"] = x["m"]
        else:
            merged.append(dict(x))
    return merged


def learn_promo(s, runs, promo_days, origin, seg=7):
    """Эффект прошлых запусков акции для ряда s: {название: [(м1, м2), ...]}.
    м1 — первые seg дней, м2 — остальные; база — медиана тех же дней недели за 1–12 недель до старта (без акций)."""
    res = {}
    for r in runs:
        if r["a"] > origin:
            continue
        days = [d for d in pd.date_range(r["a"], min(r["b"], origin)) if pd.notna(s.get(d))]
        if not days:
            continue
        f, b = [], []
        for d in days:
            vals = []
            for k in range(1, 13):
                x = d - pd.Timedelta(weeks=k)
                if x < r["a"] and x not in promo_days and pd.notna(s.get(x)):
                    vals.append(s[x])
                if len(vals) == 4:
                    break
            f.append(s[d])
            b.append(np.median(vals) if vals else np.nan)
        f, b = np.array(f, float), np.array(b, float)
        ok = ~np.isnan(b)
        if ok.sum() < 3:
            continue
        rel = np.array([(d - r["a"]).days for d in days])
        m1 = f[ok & (rel < seg)].sum() / b[ok & (rel < seg)].sum() if (ok & (rel < seg)).any() else np.nan
        m2 = f[ok & (rel >= seg)].sum() / b[ok & (rel >= seg)].sum() if (ok & (rel >= seg)).sum() >= 3 else np.nan
        res.setdefault(r["name"], []).append((m1, m2))
    return res


def promo_excess_rate(s, runs, promo_days, origin, weeks=52):
    """Доля объёма, которую дали прошлые запуски акций за окно: Σ(факт − база) по дням акций / Σ факт.
    База дня — медиана тех же дней недели за 1–12 недель до старта без акций (как в learn_promo)."""
    lo = origin - pd.Timedelta(weeks=weeks)
    win = s[(s.index > lo) & (s.index <= origin)].dropna()
    if win.empty or win.sum() <= 0:
        return 0.0, 0, 0.0, 0.0
    exc, n = 0.0, 0
    for r in runs:
        if r["b"] < lo or r["a"] > origin:
            continue
        days = [d for d in pd.date_range(max(r["a"], lo), min(r["b"], origin)) if pd.notna(s.get(d))]
        got = False
        for d in days:
            vals = []
            for k in range(1, 13):
                x = d - pd.Timedelta(weeks=k)
                if x < r["a"] and x not in promo_days and pd.notna(s.get(x)):
                    vals.append(s[x])
                if len(vals) == 4:
                    break
            if vals:
                exc += s[d] - float(np.median(vals))
                got = True
        n += got
    return float(min(0.3, max(0.0, exc / win.sum()))), n, float(exc), float(win.sum())


def new_year_factor(s, origin, min_years=2):
    """Новогодние 1–10.01 относительно воскресений 11.01–07.02 того же года; медиана по годам (нужно ≥ min_years)."""
    f = []
    s = s.dropna()
    if s.empty:
        return None, 0
    for y in range(s.index.min().year, origin.year + 1):
        if pd.Timestamp(y, 2, 7) > origin:
            continue
        a = s[pd.Timestamp(y, 1, 1):pd.Timestamp(y, 1, 10)].dropna()
        sun = s[pd.Timestamp(y, 1, 11):pd.Timestamp(y, 2, 7)]
        sun = sun[sun.index.weekday == 6].dropna()
        if len(a) >= 8 and len(sun) >= 3 and sun.median() > 0:
            f.append(a.mean() / sun.median())
    return (float(np.median(f)), len(f)) if len(f) >= min_years else (None, len(f))


# ---------------------------------------------------------------------------
# Прогноз
# ---------------------------------------------------------------------------
def day_type(d, holidays=()):
    if pd.Timestamp(d).normalize() in holidays:
        return "Вс"
    return {5: "Сб", 6: "Вс"}.get(pd.Timestamp(d).weekday(), "бдн")


def forecast_daily(hist, origin, start, end, cfg=None, events=None, holidays=(), manual_breaks=None, log=None):
    """
    hist     — дневные ряды (см. модуль), используются только даты ≤ origin
    start..end — период прогноза (включительно)
    events   — DataFrame календаря: date_from, date_to, channel (Все/Голос/Чат/Почта), action, multiplier
    holidays — даты, которые считаются воскресеньем
    -> DataFrame по дням: все ряды + voice_queue, chat_queue; log (dict) — параметры для отчёта
    """
    cfg = {**DEFAULTS, **(cfg or {})}
    log = log if log is not None else {}
    log["rules"] = dict(spike_up=cfg["spike_up"], share_weeks=cfg["share_weeks"],
                        reserve_weeks=cfg["reserve_weeks"], reserve_cap=0.3)
    h = hist[hist.index <= origin].copy()
    days = pd.date_range(start, end, freq="D")
    hol = pd.DatetimeIndex(holidays)
    ev = events if events is not None else pd.DataFrame(columns=["date_from", "date_to", "channel", "action", "multiplier"])
    chan_of = {"voice_entry": "Голос", "voice_direct": "Голос", "chat_entry": "Чат", "mail": "Почта"}

    def ev_days(action, chan, future=None):
        out = {}
        for r in ev.itertuples():
            if str(r.action) != action or str(r.channel) not in (chan, "Все"):
                continue
            a = pd.Timestamp(r.date_from)
            b = pd.Timestamp(r.date_to) if pd.notna(r.date_to) else a
            for d in pd.date_range(a, b):
                out[d] = r.multiplier
        return out

    runs = promo_runs(ev)

    def promo_days_for(chan):
        dd = set()
        for r in runs:
            if r["channel"] in (chan, "Все"):
                dd |= set(pd.date_range(r["a"], r["b"]))
        return dd

    out = pd.DataFrame(index=days)
    masks, spike_ratios = {}, []
    log["promo"], log["new_year"] = [], {}
    for col in VOLUME_SERIES:
        s = h[col].astype(float)
        pdays = promo_days_for(chan_of[col])
        excl = list(ev_days("исключить из истории", chan_of[col])) + [d for d in pdays if d <= origin] \
            + [d for d in hol if d <= origin]
        if cfg["clean"]:
            mask, norm = detect_spikes(s, cfg, excl)
        else:
            mask, norm = pd.Series(False, index=s.index), weekday_norm(s)
        masks[col] = mask
        st = series_start(s)
        if st is not None:
            mask = mask | (s.index < st + pd.Timedelta(days=cfg["startup_days"]))
        wdi = weekday_index(s, mask, origin, cfg["weekday_weeks"])
        vbreaks = []
        lw = cfg["level_weeks"]
        if col in (cfg["volume_breaks"] or ()) and st is not None:
            des = np.log((s / wdi.reindex(s.index.weekday).values).where(~mask & (s.index >= st)).clip(lower=1))
            vbreaks = find_breaks(des, cfg, min_jump=np.log(1 + cfg["vol_min_jump"]))
            last_b = max(vbreaks, default=None)
            if last_b is not None:
                lw = min(lw, max(2, (origin - last_b).days / 7))
        level, slope, level_meta = level_at(s, mask, origin, lw, wdi, cfg["trend"], half_life=cfg.get("half_life"))
        if cfg["use_reserve"] and cfg["clean"]:
            rr, rr_ex, rr_tot = reserve_rate(s, masks[col], norm, origin, cfg["reserve_weeks"], excl)
        else:
            rr, rr_ex, rr_tot = 0.0, 0.0, 0.0
        ratios = (s / norm)[masks[col] & (s > norm)]
        spike_ratios += list(ratios.dropna().values) if col in ("voice_entry", "chat_entry") else []
        steps = (days - origin).days.values
        damp = np.array([sum(0.97 ** k for k in range(1, int(n) + 1)) for n in steps]) if slope else 0
        lvl = np.maximum(level + slope * damp, 0) if slope else np.full(len(days), level)
        idx = np.array([wdi[6] if d in hol else wdi[d.weekday()] for d in days])
        val = lvl * idx / (1 - rr) if rr < 1 else lvl * idx
        if col in (cfg.get("month_series") or ()):
            # Год (№138): множитель месяца. Квартал этот ключ не передаёт. 1–10.01 и праздники
            # остаются на новогоднем коэффициенте и «как воскресенье», индекс на них не кладётся.
            got = month_factors(s, origin, hol, pdays)
            if got:
                factors, years = got
                hol_set = set(pd.DatetimeIndex(hol).normalize())
                mult = np.array([
                    1.0 if ((d.month == 1 and d.day <= 10) or d in hol_set) else factors[int(d.month)]
                    for d in days])
                val = val * mult
                log.setdefault("month_season", {})[col] = {
                    "index": {str(m): round(factors[m], 3) for m in range(1, 13)},
                    "years": {str(m): years[m] for m in range(1, 13)}}
        if cfg.get("dom_effect") and col in ("voice_entry", "chat_entry", "mail"):
            # день месяца: 1-е / 2-е / предпоследний / последний — коэффициент по прошлым месяцам
            dfs = dom_factors(s, set(excl), origin, cfg["dom_weeks"], cfg["dom_min_n"])
            mult = np.array([dfs.get(dom_key(d), (1.0, 0))[0] if d not in hol else 1.0 for d in days])
            val = val * mult
            log.setdefault("day_of_month", {})[col] = {k: dict(factor=round(v[0], 3), n=v[1]) for k, v in dfs.items()}
        if col in ("voice_entry", "chat_entry"):
            # новогодние каникулы: коэффициент к «как воскресенье» по прошлым годам (если есть ≥ 2 лет)
            fny, nyears = new_year_factor(s, origin) if cfg.get("new_year", True) else (None, 0)
            ny = np.array([d in hol and d.month == 1 and d.day <= 10 for d in days])
            if fny and ny.any():
                val = np.where(ny, val * fny, val)
            log["new_year"][col] = dict(factor=round(fny, 3) if fny else None, years=nyears)
            # акции: будущие дни × множитель (указанный в календаре или по прошлым запускам с тем же названием)
            learned = learn_promo(s, [r for r in runs if r["channel"] in (chan_of[col], "Все")], pdays, origin)
            for r in runs:
                if r["channel"] not in (chan_of[col], "Все") or r["b"] <= origin:
                    continue
                hist_m = learned.get(r["name"], [])
                m1 = np.nanmean([x[0] for x in hist_m]) if hist_m else np.nan
                m2 = np.nanmean([x[1] for x in hist_m]) if hist_m and not all(np.isnan(x[1]) for x in hist_m) else m1
                m1 = m1 if np.isfinite(m1) else m2
                if hist_m and not cfg.get("promo_segments", True):
                    # один множитель: среднее по запускам (у запуска — среднее его доступных частей)
                    per_run = [np.mean([v for v in x if np.isfinite(v)]) for x in hist_m if any(np.isfinite(v) for v in x)]
                    m1 = m2 = float(np.mean(per_run)) if per_run else np.nan
                if hist_m and cfg.get("promo_shrink", False):
                    k_ = len(hist_m) / (len(hist_m) + 1)      # мало прошлых запусков — эффект осторожнее
                    m1, m2 = 1 + (m1 - 1) * k_, 1 + (m2 - 1) * k_
                explicit = pd.notna(r["m"]) and r["m"] not in ("", None)
                src = "календарь" if explicit else (f"прошлые запуски: {len(hist_m)}" if hist_m else "нет истории — множитель 1, укажите вручную")
                for i, d in enumerate(days):
                    if r["a"] <= d <= r["b"]:
                        rel = (d - r["a"]).days
                        m = float(r["m"]) if explicit else (m1 if rel < 7 else m2) if hist_m else 1.0
                        val[i] *= m if np.isfinite(m) else 1.0
                log["promo"].append(dict(series=col, name=r["name"], period=f"{r['a'].date()} – {r['b'].date()}",
                                         source=src, m_first_week=None if explicit or not hist_m else round(float(m1), 3),
                                         m_rest=None if explicit or not hist_m else round(float(m2), 3),
                                         m_set=float(r["m"]) if explicit else None))
        rp = 0.0
        if cfg.get("promo_reserve") and col in ("voice_entry", "chat_entry"):
            # равномерный запас на крупные акции: будущие дни без известной акции × 1/(1 − доля акций за год)
            crun = [r for r in runs if r["channel"] in (chan_of[col], "Все")]
            rp, n_runs, rp_ex, rp_tot = promo_excess_rate(s, crun, pdays, origin, cfg.get("promo_reserve_weeks", 52))
            known = np.array([any(r["a"] <= d <= r["b"] for r in crun if r["b"] > origin) for d in days])
            if rp > 0:
                val = np.where(known, val, val / (1 - rp))
            log.setdefault("promo_reserve", {})[col] = dict(
                rate=round(rp, 4), runs=n_runs, excess=round(rp_ex, 1), total=round(rp_tot, 1),
                weeks=cfg.get("promo_reserve_weeks", 52))
        out[col] = val
        log[col] = dict(level=round(level, 1), slope_per_day=round(slope, 2), reserve=round(rr, 4),
                        reserve_excess=round(rr_ex, 1), reserve_total=round(rr_tot, 1),
                        level_weeks=round(lw, 1),
                        level_from=str(pd.Timestamp(level_meta["from"]).date()),
                        level_to=str(pd.Timestamp(level_meta["to"]).date()),
                        level_days=level_meta["days"], level_window_weeks=round(level_meta["weeks"], 1),
                        volume_breaks=[str(b.date()) for b in vbreaks],
                        weekday_index={WD[i]: round(float(wdi[i]), 3) for i in range(7)},
                        excluded_days=[d.strftime("%d.%m.%Y") for d in s.index[masks[col]]][-60:])
    typical_spike = float(np.median(spike_ratios)) if spike_ratios else 1.5
    log["typical_spike"] = round(typical_spike, 2)

    # будущие события-множители
    for col in VOLUME_SERIES:
        for d, m in ev_days("множитель", chan_of[col]).items():
            if d in out.index:
                out.loc[d, col] *= float(m) if pd.notna(m) and m not in ("", None) else typical_spike

    # доли до очереди по режимам
    for key, (ecol, rcol) in SHARES.items():
        chan = "Голос" if key == "voice" else "Чат"
        e, r = h[ecol].astype(float), h[rcol].astype(float)
        e = e.where(r.notna())          # вход до начала ряда «дошли до очереди» (внешняя история) в доле не участвует
        sh = (r / e).where(e > 0)
        st = series_start(r)
        sh_hist = sh[sh.index >= st + pd.Timedelta(days=cfg["startup_days"])] if st is not None else sh
        auto = find_breaks(sh_hist, cfg)
        man = [pd.Timestamp(x) for x in (manual_breaks or {}).get(chan, []) if pd.Timestamp(x) <= origin]
        if key == "voice" and "voice_robot" in h.columns and h["voice_robot"].notna().any():
            # робот виден в данных: включение/выключение — смена режима сразу, без календаря и без ожидания 14 дней
            rb, on_now, periods = robot_regimes(h["voice_robot"].astype(float), h[ecol].astype(float))
            man += [d for d in rb if d <= origin]
            shr = {}
            for a_, b_, st_ in periods:
                w_ = (e.index >= a_) & (e.index <= b_)
                if e[w_].sum() > 0 and (b_ - a_).days >= 6:
                    shr[st_] = float(r[w_].sum() / e[w_].sum())          # последний период каждого состояния
            log["robot"] = dict(on_now=on_now, changes=[str(x.date()) for x in rb],
                                share_robot_on=round(shr[True], 4) if True in shr else None,
                                share_robot_off=round(shr[False], 4) if False in shr else None)
        man += [d for d, _ in ev_days("смена режима", chan).items() if d <= origin]
        breaks = merge_breaks(auto, man, cfg["cp_min_days"])
        mask = masks[ecol] | (sh - sh.rolling(15, center=True, min_periods=5).median()).abs().gt(0.08).fillna(False)
        share, reg_start, ndays, e_sum, r_sum, share_from = share_at(e, r, mask, origin, cfg["share_weeks"], breaks)
        prev = [b for b in breaks if b < pd.Timestamp(reg_start)]
        p_lo = max(prev) if prev else r.dropna().index.min()
        pw = (e.index >= p_lo) & (e.index < pd.Timestamp(reg_start)) & ~mask
        prev_share = float(r[pw].sum() / e[pw].sum()) if e[pw].sum() else None
        share_note = ""
        if not np.isfinite(share) or ndays < 3:
            # режим начался только что: чистых дней почти нет — берём все дни режима, иначе прошлый режим
            w2 = (e.index >= pd.Timestamp(reg_start)) & (e.index <= origin)
            if e[w2].sum() and int(w2.sum()) >= 2:
                e_sum, r_sum = float(e[w2].sum()), float(r[w2].sum())
                share, ndays = (r_sum / e_sum), int(w2.sum())
                share_from = pd.Timestamp(reg_start)
                share_note = f"новый режим: всего {ndays} дн., доля неустойчива"
            elif prev_share:
                e_sum, r_sum = float(e[pw].sum()), float(r[pw].sum())
                share, ndays = prev_share, 0
                share_from = p_lo
                share_note = "новый режим без данных: взята доля прошлого режима"
        # запас на всплески на уровне очередей: избыток «дошли до очереди» в дни всплесков входа / весь «дошли до очереди»
        rr_e = log[ecol]["reserve"]
        rr_g, g_ex, g_tot = 0.0, 0.0, 0.0
        if cfg["use_reserve"] and cfg["clean"]:
            m_g, n_g = detect_spikes(r, cfg, list(ev_days("исключить из истории", chan)))
            rr_g, g_ex, g_tot = reserve_rate(r, masks[ecol] | m_g, n_g, origin, cfg["reserve_weeks"], list(ev_days("исключить из истории", chan)))
        if cfg["reserve_queue"].get(key, True):
            queue_factor = (1 - rr_e) / (1 - rr_g) if rr_g < 1 else 1.0  # вход уже умножен на 1/(1−rr_e)
        else:
            queue_factor, rr_g, g_ex, g_tot = (1 - rr_e), 0.0, 0.0, 0.0   # очереди = очищенный вход × доля
        series = pd.Series(share, index=days)
        fut = {d: m for d, m in ev_days("смена режима", chan).items() if d > origin}
        for d0 in sorted(fut):
            m = fut[d0]
            if pd.notna(m) and m not in ("", None):
                series[series.index >= d0] = float(m)
        out[f"{key}_share"] = series
        out[f"{key}_queue_factor"] = queue_factor
        log[f"{key}_share"] = dict(share=round(share, 4), regime_start=str(pd.Timestamp(reg_start).date()), days_used=ndays,
                                   entry_sum=round(e_sum, 1), reached_sum=round(r_sum, 1),
                                   share_from=None if pd.isna(pd.Timestamp(share_from)) else str(pd.Timestamp(share_from).date()),
                                   auto_breaks=[str(b.date()) for b in auto], manual_breaks=[str(b.date()) for b in man],
                                   future_changes={str(k.date()): v for k, v in fut.items()},
                                   reserve_queue=round(rr_g, 4),
                                   reserve_queue_excess=round(g_ex, 1), reserve_queue_total=round(g_tot, 1),
                                   note=share_note,
                                   previous_regime_share=round(prev_share, 4) if prev_share else None,
                                   previous_regime=f"{pd.Timestamp(p_lo).date()} – {(pd.Timestamp(reg_start) - pd.Timedelta(days=1)).date()}")
    out["voice_entry_queue"] = out["voice_entry"] * out["voice_share"] * out.pop("voice_queue_factor")
    out["voice_queue"] = out["voice_entry_queue"] + out["voice_direct"]
    out["chat_entry_queue"] = out["chat_entry"] * out["chat_share"] * out.pop("chat_queue_factor")
    out["chat_queue"] = out["chat_entry_queue"]
    return out, log


def profile_30(interval_df, value_col, channel_col, exclude_days=(), holidays=()):
    """Профиль 30 мин: доля получаса в сутках по типу дня (бдн/Сб/Вс) — сумма получаса / сумма суток."""
    it = interval_df.copy()
    it["interval_start"] = pd.to_datetime(it["interval_start"])
    it["day"] = it["interval_start"].dt.normalize()
    it = it[~it["day"].isin(pd.DatetimeIndex(exclude_days))]
    it["dtype"] = [day_type(d, holidays) for d in it["day"]]
    it["slot"] = it["interval_start"].dt.hour * 2 + (it["interval_start"].dt.minute >= 30)
    num = it.groupby([channel_col, "dtype", "slot"])[value_col].sum()
    den = it.groupby([channel_col, "dtype"])[value_col].sum()
    prof = (num / den).rename("share").reset_index()
    full = pd.MultiIndex.from_product([prof[channel_col].unique(), ["бдн", "Сб", "Вс"], range(48)],
                                      names=[channel_col, "dtype", "slot"])
    prof = prof.set_index([channel_col, "dtype", "slot"])["share"].reindex(full).fillna(0.0)
    s = prof.groupby(level=[0, 1]).transform("sum")
    prof = (prof / s.where(s > 0)).fillna(1 / 48)
    return prof.sort_index()


def to_intervals(daily, profiles, mapping, holidays=()):
    """daily[channel_col] × профиль -> long DataFrame (interval_start, channel, value)."""
    rows = []
    for col, (chan_name, prof_key) in mapping.items():
        for d, v in daily[col].items():
            dt_ = day_type(d, holidays)
            p = profiles.loc[(prof_key, dt_)].values
            for k in range(48):
                rows.append((d + pd.Timedelta(minutes=30 * k), chan_name, float(v) * float(p[k])))
    return pd.DataFrame(rows, columns=["interval_start", "channel", "value"])


# ---------------------------------------------------------------------------
# Бэктест
# ---------------------------------------------------------------------------
TARGETS = ["voice_entry", "voice_queue", "chat_entry", "chat_queue"]


def naive_forecasts(h, origin, days):
    last = h[h.index <= origin]
    out = {}
    for name, k in (("наивный: неделю назад", 1), ("среднее 4 таких же дней", 4)):
        df = pd.DataFrame(index=days)
        for col in TARGETS:
            s = last[col]
            vals = []
            for d in days:
                wd = d.weekday()
                same = s[s.index.weekday == wd].iloc[-k:]
                vals.append(float(same.mean()) if len(same) else np.nan)
            df[col] = vals
        out[name] = df
    return out


def backtest(hist, variants, first_origin, last_origin, step_days=7, horizon_weeks=12, events=None, holidays=(),
             manual_breaks=None):
    """-> DataFrame: variant, target, origin, horizon_week, fact, fc (по дням)."""
    hist = hist.copy()
    hist["voice_queue"] = hist["voice_entry_queue"] + hist["voice_direct"]
    hist["chat_queue"] = hist["chat_entry_queue"]
    end_data = hist.dropna(subset=["voice_entry", "chat_entry"]).index.max()
    rows = []
    o = pd.Timestamp(first_origin)
    while o <= pd.Timestamp(last_origin):
        days = pd.date_range(o + pd.Timedelta(days=1), min(o + pd.Timedelta(weeks=horizon_weeks), end_data))
        if len(days) == 0:
            break
        fcs = naive_forecasts(hist, o, days)
        for name, cfg in variants.items():
            fcs[name], _ = forecast_daily(hist, o, days[0], days[-1], cfg, events, holidays, manual_breaks)
        for name, fc in fcs.items():
            for col in TARGETS:
                for d in days:
                    rows.append((name, col, o, (d - o).days // 7 + 1, d, hist.at[d, col], fc.at[d, col]))
        o += pd.Timedelta(days=step_days)
    return pd.DataFrame(rows, columns=["variant", "target", "origin", "hweek", "day", "fact", "fc"])


def bt_metrics(bt, exclude_days=()):
    bt = bt[~bt["day"].isin(pd.DatetimeIndex(exclude_days))].dropna(subset=["fact", "fc"])
    bt = bt.assign(h=pd.cut(bt["hweek"], [0, 2, 4, 8, 13], labels=["1–2 нед", "3–4 нед", "5–8 нед", "9–12 нед"]))
    g = bt.groupby(["target", "variant", "h"], observed=True)
    day = g.apply(lambda x: pd.Series({
        "WAPE день": (x.fact - x.fc).abs().sum() / x.fact.sum(),
        "Смещение": (x.fc - x.fact).sum() / x.fact.sum(),
        "дней": len(x)}), include_groups=False)
    # 4-недельные блоки от каждой точки прогноза: |Σпрогноз − Σфакт| / Σфакт
    blk = bt.assign(block=(bt["hweek"] - 1) // 4).groupby(["target", "variant", "origin", "block"]).agg(
        fact=("fact", "sum"), fc=("fc", "sum"), n=("fact", "size"))
    blk = blk[blk["n"] >= 25].reset_index()
    blk["err"] = (blk.fc - blk.fact).abs() / blk.fact
    b = blk.groupby(["target", "variant"])["err"].mean().rename("Ошибка 4 недель")
    return day.reset_index(), b.reset_index()
