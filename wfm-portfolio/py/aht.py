# -*- coding: utf-8 -*-
"""
Прогноз AHT по 30 минутам: голос = разговор + поствызывная, чат = от распределения на оператора до закрытия (дашборд/6, №72).

Данные: представление v_aht_interval (views/60_v_aht_interval.sql).
Модель — «профиль × уровень»:
  уровень   — AHT канала за последние level_weeks недель (сумма времени / сумма обработанных);
  профиль   — AHT часа × типа дня (бдн / Сб / Вс; праздник = Вс) к уровню, за profile_weeks недель,
              со сжатием к 1 при малом числе обращений в ячейке (shrink);
  тренд     — по желанию: наклон недельного AHT (медиана попарных наклонов) за trend_weeks недель,
              затухающий (damp в неделю) — №3 «с трендом»; принимается только по бэктесту.
AHT 30-минутки = AHT её часа. Голос берётся только за дни, где есть статусы операторов (поствызывная).
"""
import numpy as np
import pandas as pd

from forecast_core import day_type

AHT_DEFAULTS = dict(level_weeks=8, profile_weeks=8, shrink=200, trend=False, trend_weeks=12, damp=0.9)
CHANNELS = {"Голос": "Голос очереди", "Чат": "Чат очереди"}
SQL = """select channel, interval_start, handled, work_sec, wrap_sec, wrap_found, wrap_covered
         from v_aht_interval"""


def load(con):
    import wfm_store
    d = wfm_store.query(SQL, con)
    d["interval_start"] = pd.to_datetime(d["interval_start"])
    d = d[(d["channel"] != "Голос") | (d["wrap_covered"] == 1)].copy()     # голос — только с поствызывной
    d["handle_sec"] = d["work_sec"] + d["wrap_sec"]
    return d


def _prep(d, holidays):
    x = d.copy()
    x["day"] = x["interval_start"].dt.normalize()
    x["hour"] = x["interval_start"].dt.hour
    dts = {dd: day_type(dd, holidays) for dd in x["day"].unique()}
    x["dtype"] = x["day"].map(dts)
    return x


def fit(d, origin, cfg=None, holidays=()):
    """d — история (load), только до origin включительно. -> модель (dict по каналам)."""
    cfg = {**AHT_DEFAULTS, **(cfg or {})}
    x = _prep(d[d["interval_start"] < pd.Timestamp(origin) + pd.Timedelta(days=1)], holidays)
    model = {}
    for ch, g in x.groupby("channel"):
        lo_l = pd.Timestamp(origin) - pd.Timedelta(weeks=cfg["level_weeks"])
        lw = g[g["day"] > lo_l]
        if lw["handled"].sum() == 0:
            continue
        if cfg.get("level_stat") == "median":                              # медиана дневного AHT за N недель
            dd = lw.groupby("day")[["handle_sec", "handled"]].sum()
            level = float((dd["handle_sec"] / dd["handled"]).median())
        else:
            level = lw["handle_sec"].sum() / lw["handled"].sum()
        lo_p = pd.Timestamp(origin) - pd.Timedelta(weeks=cfg["profile_weeks"])
        pw = g[g["day"] > lo_p]
        base = pw["handle_sec"].sum() / pw["handled"].sum()
        cell = pw.groupby(["dtype", "hour"])[["handle_sec", "handled"]].sum()
        ratio = (cell["handle_sec"] / cell["handled"]) / base
        w = cell["handled"] / (cell["handled"] + cfg["shrink"])           # мало обращений -> ближе к 1
        prof = (1 + (ratio - 1) * w).reindex(pd.MultiIndex.from_product([["бдн", "Сб", "Вс"], range(24)]))
        prof = prof.fillna(prof.groupby(level=1).transform("mean")).fillna(1.0)
        slope = 0.0
        if cfg["trend"]:
            lo_t = pd.Timestamp(origin) - pd.Timedelta(weeks=cfg["trend_weeks"])
            wk = g[g["day"] > lo_t].set_index("day")[["handle_sec", "handled"]].resample("W").sum()
            wk = (wk["handle_sec"] / wk["handled"]).dropna()
            if len(wk) >= 6:
                sl = [(wk.iloc[j] - wk.iloc[i]) / (j - i) for i in range(len(wk)) for j in range(i + 1, len(wk))]
                slope = float(np.median(sl)) / level                       # доля уровня в неделю
        model[ch] = dict(level=level, profile=prof, slope=slope, origin=pd.Timestamp(origin), cfg=cfg)
    return model


def predict(model, slots, holidays=()):
    """slots — DatetimeIndex 30-минуток. -> DataFrame (interval_start, channel, aht_sec)."""
    slots = pd.DatetimeIndex(slots)
    rows = []
    for ch, m in model.items():
        cfg = m["cfg"]
        dts = [day_type(s.normalize(), holidays) for s in slots]
        prof = m["profile"].reindex(list(zip(dts, slots.hour))).values
        weeks = np.maximum((slots - m["origin"]).days.values / 7.0, 0)
        # затухающий тренд: сумма damp^k за прошедшие недели (не уходит в бесконечность)
        cum = (1 - cfg["damp"] ** weeks) / (1 - cfg["damp"]) if m["slope"] else 0
        aht = m["level"] * prof * (1 + m["slope"] * cum)
        rows.append(pd.DataFrame({"interval_start": slots, "channel": ch, "aht_sec": aht}))
    return pd.concat(rows, ignore_index=True) if rows else pd.DataFrame(columns=["interval_start", "channel", "aht_sec"])


# ---------------------------------------------------------------------------
# Бэктест AHT: прогноз от точки на 1–13 недель вперёд; вес — фактическое число обработанных
# (ошибка именно AHT, а не объёма). Метрики: отклонение AHT месяца и ошибка по часам (Σ|Δ handle| / Σ handle).
# ---------------------------------------------------------------------------
VARIANTS = {
    "простая: медиана 4 нед": dict(level_weeks=4, level_stat="median", shrink=10 ** 9),
    "простая: медиана 8 нед": dict(level_weeks=8, level_stat="median", shrink=10 ** 9),
    "простая: медиана 12 нед": dict(level_weeks=12, level_stat="median", shrink=10 ** 9),
    "профиль час×тип дня, 8 нед": dict(),
    "профиль, уровень 4 нед": dict(level_weeks=4),
    "профиль, уровень 12 нед": dict(level_weeks=12, profile_weeks=12),
    "профиль + тренд": dict(trend=True),
}


def backtest(d, holidays=(), horizon_weeks=13, step_days=7, min_hist_weeks=8):
    x = _prep(d, holidays)
    first = x.groupby("channel")["day"].min().max() + pd.Timedelta(weeks=min_hist_weeks)
    last_day = x["day"].max() - pd.Timedelta(days=1)
    out = []
    o = first
    while o < last_day - pd.Timedelta(days=7):
        fut = x[(x["day"] > o) & (x["day"] <= o + pd.Timedelta(weeks=horizon_weeks)) & (x["day"] <= last_day)]
        for vn, cfg in VARIANTS.items():
            m = fit(d, o, cfg, holidays)
            p = predict(m, fut["interval_start"].unique(), holidays)
            f = fut.merge(p, on=["interval_start", "channel"], how="inner")
            f = f.assign(variant=vn, origin=o, hweek=((f["day"] - o).dt.days - 1) // 7 + 1,
                         fc_handle=f["aht_sec"] * f["handled"])
            out.append(f[["variant", "channel", "origin", "hweek", "day", "hour", "handled", "handle_sec", "fc_handle"]])
        o += pd.Timedelta(days=step_days)
    return pd.concat(out, ignore_index=True) if out else pd.DataFrame()


def bt_summary(bt):
    bt = bt.assign(h=pd.cut(bt["hweek"], [0, 4, 8, 13], labels=["1–4 нед", "5–8 нед", "9–13 нед"]),
                   month=bt["day"].dt.to_period("M").astype(str))
    # месяц: AHT факт / AHT прогноз по каждой точке × месяцу × горизонту
    mon = bt.groupby(["variant", "channel", "h", "origin", "month"], observed=True)[["handle_sec", "fc_handle", "handled"]].sum()
    mon = mon[mon["handled"] > 0].reset_index()
    mon["dev"] = (mon["handle_sec"] / mon["fc_handle"] - 1).abs() * 100
    mon["bias"] = (mon["handle_sec"] / mon["fc_handle"] - 1) * 100
    hr = bt.groupby(["variant", "channel", "h", "origin", "day", "hour"], observed=True)[["handle_sec", "fc_handle"]].sum().reset_index()
    hr["abs"] = (hr["handle_sec"] - hr["fc_handle"]).abs()
    hs = hr.groupby(["variant", "channel", "h"], observed=True).apply(
        lambda g: g["abs"].sum() / g["handle_sec"].sum() * 100, include_groups=False).rename("Ошибка по часам, %")
    ms = mon.groupby(["variant", "channel", "h"], observed=True).agg(
        **{"Отклонение AHT месяца, %": ("dev", "mean"), "Смещение (факт/прогноз − 100), %": ("bias", "mean"),
           "Точек": ("dev", "size")})
    return ms.join(hs).round(2).reset_index().rename(columns={"variant": "Вариант", "channel": "Канал", "h": "Горизонт"})
