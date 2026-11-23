# -*- coding: utf-8 -*-
"""
Прогноз нагрузки по 30 минутам (голос, чат, почта) из хранилища + журнал версий.

Запуск: wfm.bat — «План на квартал», «Недельный пересчёт»; tools\\backtest.bat — бэктест.
Или: py wfm_forecast.py [--to 2027-03-31] [--from 2026-11-01] [--weekly] [--year] [--backtest] [--publish]

Что делает (подробно — forecast_core.py):
  * история: источник + внешняя история входа КЦ (импорт: tools\\import_history.bat), там где источника ещё не было;
  * прошлые аномалии и смены режима робота/бота находит сам;
  * календарь events.xlsx: «акция» (повторяющиеся крупные; эффект — по прошлым запускам с тем же Названием),
    «смена режима» (Множитель = новая доля до очереди), «исключить из истории», «множитель», «фон» (не используется);
  * праздники (wfm_settings.HOLIDAYS) — как воскресенье, новогодние 1–10.01 — с коэффициентом по прошлым годам;
  * сохраняет версию в DB\\forecast_run и DB\\forecast_interval (только дописывание);
  * отчёт: Сводка, По_дням, 30 мин, По_неделям, Бот_и_робот, Параметры.
"""
import argparse
import datetime as dt
import json
import os
import sys

import numpy as np
import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import forecast_core as F
import wfm_settings as S
import wfm_store
from wfm_common import EVENTS_FILE, error_line, is_db_timeout, report_path, stamp, store_root

CHANNELS = {   # колонка ядра -> (название канала в хранилище, профиль)
    "voice_entry": ("Голос вход", "Голос вход"),
    "voice_direct": ("Голос прямой", "Голос прямой"),
    "voice_queue": ("Голос очереди", "Голос"),
    "chat_entry": ("Чат вход", "Чат вход"),
    "chat_queue": ("Чат очереди", "Чат"),
    "mail": ("Почта", "Почта"),
}
WD = F.WD
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")

SQL_DAILY = {
    "queue": "select day, channel, sum(offered) as v from v_interval_volume group by 1, 2",
    "chat": "select day, sum(sessions) as entry, sum(reached_queue) as entry_queue from v_chat_sessions_interval group by 1",
    "voice": """select day, sum(entry_sessions) as entry, sum(entry_reached_queue) as entry_queue, sum(direct_queue) as direct,
                       sum(entry_via_robot) as robot
                from v_voice_sessions_interval group by 1""",
}
SQL_PROFILE = """
select channel, interval_start, offered as v from v_interval_volume
where day >= {lo} and day <= {hi}
union all select 'Чат вход', interval_start, sessions from v_chat_sessions_interval where day >= {lo} and day <= {hi}
union all select 'Голос вход', interval_start, entry_sessions from v_voice_sessions_interval where day >= {lo} and day <= {hi}
union all select 'Голос прямой', interval_start, direct_queue from v_voice_sessions_interval where day >= {lo} and day <= {hi}
"""


# ---------------------------------------------------------------------------
def load_ext_history():
    """Внешняя история входа КЦ (tools\\import_history.bat): DataFrame day × series, уже в нашем определении."""
    p = os.path.join(store_root(), "ext_history", "ext_daily.parquet")
    if not os.path.exists(p):
        return None
    x = pd.read_parquet(p)
    x["day"] = pd.to_datetime(x["day"])
    return x.pivot_table(index="day", columns="series", values="value")


def load_history(con, ext=True):
    h = _load_source(con)
    x = load_ext_history() if ext else None
    if x is None:
        return h
    idx = pd.date_range(min(h.index.min(), x.index.min()), h.index.max(), freq="D")
    h = h.reindex(idx)
    for col in x.columns:
        if col not in h.columns:
            continue
        st = F.series_start(h[col])
        # источник — с момента, когда ряд «встал» (+7 дней на разгон); раньше — внешняя история
        source_from = st + pd.Timedelta(days=7) if st is not None else h.index.max() + pd.Timedelta(days=1)
        ext_col = x[col].reindex(idx)
        h[col] = h[col].where((h.index >= source_from) | ext_col.isna()).combine_first(ext_col)
    return h


def _load_source(con):
    q = lambda sql: wfm_store.query(sql, con)
    g = q(SQL_DAILY["queue"])
    g["day"] = pd.to_datetime(g["day"])
    g = g.pivot_table(index="day", columns="channel", values="v", aggfunc="sum")
    c = q(SQL_DAILY["chat"]).assign(day=lambda x: pd.to_datetime(x["day"])).set_index("day")
    v = q(SQL_DAILY["voice"]).assign(day=lambda x: pd.to_datetime(x["day"])).set_index("day")
    yesterday = pd.Timestamp(dt.date.today()) - pd.Timedelta(days=1)
    idx = pd.date_range(min(g.index.min(), c.index.min(), v.index.min()), yesterday, freq="D")
    h = pd.DataFrame(index=idx)
    h["voice_entry"], h["voice_entry_queue"], h["voice_direct"] = v["entry"], v["entry_queue"], v["direct"]
    h["voice_robot"] = v["robot"]
    h["chat_entry"], h["chat_entry_queue"] = c["entry"], c["entry_queue"]
    h["mail"] = g.get("Почта")
    h["fact_voice_queue"], h["fact_chat_queue"] = g.get("Голос"), g.get("Чат")
    return h.astype(float)


def load_events():
    cols = ["date_from", "date_to", "channel", "type", "action", "name", "group", "multiplier", "comment"]
    if not os.path.exists(EVENTS_FILE):
        return pd.DataFrame(columns=cols)
    ev = pd.read_excel(EVENTS_FILE, sheet_name=0)
    ev = ev.rename(columns={"Дата с": "date_from", "Дата по": "date_to", "Канал": "channel", "Тип": "type",
                            "Действие": "action", "Название": "name", "Группа акции": "group",
                            "Множитель": "multiplier", "Комментарий": "comment"})
    for c in ("name", "comment", "group"):
        if c not in ev.columns:
            ev[c] = ""
    ev[["name", "comment", "group"]] = ev[["name", "comment", "group"]].fillna("").astype(str)
    ev = ev[[c for c in cols if c in ev.columns]].dropna(subset=["date_from"])
    ev["date_from"] = pd.to_datetime(ev["date_from"])
    ev["date_to"] = pd.to_datetime(ev["date_to"], errors="coerce")
    ev["action"] = ev["action"].astype(str).str.strip().str.lower()
    ev["channel"] = ev["channel"].astype(str).str.strip()
    ev["multiplier"] = pd.to_numeric(ev["multiplier"].astype(str).str.replace(",", "."), errors="coerce")
    # акции: модель учится по ГРУППЕ, а не по свободному названию. Группа пустая — определяем по ключевым словам.
    groups = getattr(S, "PROMO_GROUPS", {})
    is_promo = ev["action"] == "акция"
    for i in ev.index[is_promo]:
        g = ev.at[i, "group"].strip()
        if not g:
            t = (ev.at[i, "name"] + " " + ev.at[i, "comment"]).lower()
            g = next((k for k, v in groups.items() if any(w in t for w in v["keywords"])), "Другое")
        ev.at[i, "group"] = g
        ev.at[i, "name"] = g                       # по этому полю ядро ищет прошлые запуски
        if g in groups and not groups[g]["use"] and pd.isna(ev.at[i, "multiplier"]):
            ev.at[i, "action"] = "фон"             # группа не участвует (если множитель не задан вручную)
    return ev.reset_index(drop=True)


def default_period():
    t = pd.Timestamp(dt.date.today())
    q_end = pd.Timestamp(year=t.year, month=((t.month - 1) // 3 + 1) * 3, day=1) + pd.offsets.MonthEnd(0)
    nxt = q_end + pd.offsets.QuarterEnd(1)
    return t + pd.Timedelta(days=1), nxt


def year_period(today=None):
    """Календарный год для пункта «План на год»: с июля — следующий год, раньше — текущий."""
    t = pd.Timestamp(today or dt.date.today())
    y = t.year + 1 if t.month >= 7 else t.year
    return pd.Timestamp(year=y, month=1, day=1), pd.Timestamp(year=y, month=12, day=31)


def status_mode_name(run_id, mode=None):
    """«квартал 29.10.2026 22-32» из run_id вида 2026-10-29_223256 (режим берётся из _REPORT_MODE)."""
    t = dt.datetime.strptime(run_id, "%Y-%m-%d_%H%M%S")
    return f"{mode or _REPORT_MODE.get('mode', 'квартал')} {stamp(t)}"


_REPORT_MODE = {}


def to_pydate(x):
    return None if pd.isna(x) else pd.Timestamp(x).date()


# ---------------------------------------------------------------------------
# Паспорт версии (№34): какая модель и какие решения журнала в ней действуют. Менять при изменении модели.
MODEL_VERSION = ("v2 (уровень 12 нед, запас на всплески) · квартал: запас на акции №44 · AHT: медиана 8 нед №45 · "
                 "без робота №30 · без инцидентов №43 · партнёрская линия в очереди №36")

WEEKLY_CFG = dict(half_life=3, reserve_queue={"voice": True, "chat": True})   # бэктест 22.10: лучше для 1–2 недель


def season_gate(h, ev, holidays, origin):
    """Принять месячный индекс ряда, только если на горизонте 6–12 мес он лучше плоского v2 (№138).

    Ошибка — среднее |факт / прогноз − 1| по месяцам. Будущие акции скрыты, как в квартальном бэктесте.
    Мало точек или выигрыш меньше 1 п.п. — ряд не берём. Голос и чат решаются отдельно.
    """
    names = {"chat_entry": "Чат вход", "voice_entry": "Голос вход"}
    starts = {c: F.series_start(h[c].dropna()) for c in names}
    rows = []
    months = pd.period_range((origin - pd.DateOffset(months=18)).to_period("M"), origin.to_period("M"), freq="M")
    flat_cfg = dict(promo_reserve=True)
    sea_cfg = dict(promo_reserve=True, month_series=tuple(names))
    for m in months:
        s, e = m.start_time, min(m.end_time.normalize(), origin)
        if (e - s).days < 20:
            continue
        for k in range(6, 13):
            o = (m - k + 1).start_time - pd.Timedelta(days=1)
            ev_o = _bt_events(ev, o, {})
            flat, _ = F.forecast_daily(h, o, s, e, flat_cfg, ev_o, holidays)
            sea, _ = F.forecast_daily(h, o, s, e, sea_cfg, ev_o, holidays)
            for col, nm in names.items():
                fa = h.loc[s:e, col]
                if starts[col] is None or o < starts[col] + pd.Timedelta(weeks=8) or fa.isna().any() or len(fa) == 0:
                    continue
                if fa.sum() <= 0 or flat[col].sum() <= 0 or sea[col].sum() <= 0:
                    continue
                if not np.isfinite(flat[col].sum()) or not np.isfinite(sea[col].sum()):
                    continue
                rows.append({"series": nm, "col": col,
                             "flat_err": abs(fa.sum() / flat[col].sum() - 1) * 100,
                             "sea_err": abs(fa.sum() / sea[col].sum() - 1) * 100,
                             "flat_bias": (fa.sum() / flat[col].sum() - 1) * 100,
                             "sea_bias": (fa.sum() / sea[col].sum() - 1) * 100})
        print(f"   индекс {m}")
    det = pd.DataFrame(rows)
    summary, accepted = [], []
    for col, nm in names.items():
        d = det[det["col"] == col] if len(det) else det
        n = int(len(d))
        if n == 0:
            summary.append(dict(series=nm, n=0, flat_err=None, sea_err=None, flat_bias=None, sea_bias=None,
                                accepted=False, why="нет точек на горизонте 6–12 мес"))
            print(f"   {nm}: нет точек")
            continue
        flat_err = round(float(d["flat_err"].mean()), 2)
        sea_err = round(float(d["sea_err"].mean()), 2)
        flat_bias = round(float(d["flat_bias"].mean()), 2)
        sea_bias = round(float(d["sea_bias"].mean()), 2)
        gain = round(flat_err - sea_err, 2)
        ok = n >= 8 and gain >= 1.0
        why = (f"ошибка {flat_err}% → {sea_err}% (выигрыш {gain} п.п., точек {n})" if ok else
               f"не принят: ошибка {flat_err}% → {sea_err}% (выигрыш {gain} п.п.), точек {n}; нужно ≥ 8 и выигрыш ≥ 1 п.п.")
        summary.append(dict(series=nm, n=n, flat_err=flat_err, sea_err=sea_err, flat_bias=flat_bias,
                            sea_bias=sea_bias, accepted=bool(ok), why=why))
        if ok:
            accepted.append(col)
        print(f"   {nm}: {why}; смещение {flat_bias}% → {sea_bias}%")
    indexes = {}
    for col, nm in names.items():
        chan = "Голос" if col.startswith("voice") else "Чат"
        pdays = set()
        for r in F.promo_runs(ev):
            if r["channel"] in (chan, "Все"):
                pdays |= set(pd.date_range(r["a"], r["b"]))
        got = F.month_factors(h[col], origin, holidays, pdays)
        if not got:
            indexes[col] = None
            print(f"   индекс {nm} на {origin:%d.%m.%Y}: нет (мало полных лет)")
        else:
            factors, years = got
            indexes[col] = {"index": {str(m): round(factors[m], 3) for m in range(1, 13)},
                            "years": {str(m): int(years[m]) for m in range(1, 13)}}
            print("   индекс " + nm + ": " + ", ".join(f"{m}={factors[m]:.3f}" for m in range(1, 13)))
    def profile(col, year):
        s = h[col].dropna()
        st = F.series_start(s)
        if st is not None:
            s = s[s.index >= st]
        per = {}
        for mth in range(1, 13):
            part = s[(s.index.year == year) & (s.index.month == mth)]
            if len(part) >= 20:
                per[mth] = float(part.mean())
        if len(per) < 3:
            return {}
        base = float(np.mean(list(per.values())))
        return {str(mth): round(v / base, 3) for mth, v in per.items()}
    public = {"accepted": list(accepted),
              "rule": "горизонт 6–12 мес; ошибка месяца ниже плоского v2 хотя бы на 1 п.п.; точек не меньше 8",
              "rows": summary, "index": indexes,
              "voice_2026": profile("voice_entry", 2026), "chat_2026": profile("chat_entry", 2026)}
    if public["voice_2026"]:
        print("   голос 2026, день к среднему полного месяца: " +
              ", ".join(f"{m}={public['voice_2026'][m]}" for m in sorted(public["voice_2026"], key=int)))
    return {"accepted": accepted, "public": public}


def run_forecast(con, start, end, publish=None, comment="", mode="квартал"):
    holidays = pd.DatetimeIndex(pd.to_datetime(getattr(S, "HOLIDAYS", [])))
    h = load_history(con)
    origin = h.dropna(subset=["voice_entry", "chat_entry"]).index.max()
    ev = load_events()
    cfg = dict(WEEKLY_CFG if mode == "неделя" else {})
    if mode != "неделя" and getattr(S, "QUARTER_PROMO_RESERVE", False):
        cfg["promo_reserve"] = True          # №44: будущие акции на квартал неизвестны — равномерный запас по истории
    season_info = None
    if mode == "год":
        # №138: индекс только здесь. Квартал и неделя cfg["month_series"] не получают.
        print("Бэктест месячного индекса против плоского v2, горизонт 6–12 мес...")
        season_info = season_gate(h, ev, holidays, origin)
        cfg["month_series"] = tuple(season_info["accepted"])
        if not comment:
            comment = ("черновик года №138, индекс принят, не публиковать" if season_info["accepted"]
                       else "черновик года №138, индекс не принят, не публиковать")
    if getattr(S, "USE_INCIDENTS", False):
        import incidents
        incidents.load()
        cfg["incident_days"] = incidents.incident_days()

    log = {}
    daily, log = F.forecast_daily(h, origin, start, end, cfg, ev, holidays, None, log)
    if season_info is not None:
        log["month_gate"] = season_info["public"]
    for p in log.get("promo", []):
        if p["source"].startswith("нет истории"):
            print(f"   ВНИМАНИЕ: акция «{p['name']}» ({p['period']}) — нет прошлых запусков, множитель 1. "
                  f"Укажите множитель в events.xlsx.")
    # сценарий без будущих смен режима (если они есть) — для сравнения
    fut_reg = ev[(ev["action"] == "смена режима") & (ev["date_from"] > origin)]
    alt = None
    if len(fut_reg):
        alt, _ = F.forecast_daily(h, origin, start, end, cfg, ev.drop(fut_reg.index), holidays)
    # сценарий по роботу — автоматически: если робот сейчас выключен — «если заработает» (доля как в прошлый период
    # с роботом), если включён — «если отключат»
    robot_alt, robot_lbl = None, None
    rlog = log.get("robot") or {}
    target = rlog.get("share_robot_on") if rlog.get("on_now") is False else rlog.get("share_robot_off") if rlog.get("on_now") else None
    if target:
        row = pd.DataFrame([dict(date_from=start, date_to=pd.NaT, channel="Голос", type="бот", action="смена режима",
                                 name="робот (сценарий)", multiplier=target, comment="")])
        robot_alt, _ = F.forecast_daily(h, origin, start, end, cfg, pd.concat([ev, row], ignore_index=True), holidays)
        robot_lbl = (f"Голос очереди: если робот заработает (доля {target:.2f})" if rlog.get("on_now") is False
                     else f"Голос очереди: если робота отключат (доля {target:.2f})")
    short = None
    if mode != "неделя":
        short, _ = F.forecast_daily(h, origin, start, end, WEEKLY_CFG, ev, holidays)   # для сравнения: свежие недели

    # профили 30 мин по последним 8 неделям без аномальных дней
    lo, hi = origin - pd.Timedelta(weeks=8), origin
    it = wfm_store.query(SQL_PROFILE.format(lo=f"date '{lo:%Y-%m-%d}'", hi=f"date '{hi:%Y-%m-%d}'"), con)
    bad = set()
    for col in F.VOLUME_SERIES:
        m, _ = F.detect_spikes(h[col], {**F.DEFAULTS, **cfg})
        bad |= set(m[m].index)
    prof = F.profile_30(it, "v", "channel", sorted(bad), holidays)
    mapping = {k: v for k, v in CHANNELS.items()}
    intervals = F.to_intervals(daily, prof, mapping, holidays)
    aht_iv = None
    try:
        import aht as A
        ad = A.load(con)
        if len(ad):
            am = A.fit(ad, origin, getattr(S, "AHT_CFG", {}), holidays)
            slots = intervals["interval_start"].drop_duplicates().sort_values()
            aht_iv = A.predict(am, slots, holidays)
            aht_iv["channel"] = "AHT " + aht_iv["channel"].map(A.CHANNELS) + ", с"
            intervals = pd.concat([intervals, aht_iv.rename(columns={"aht_sec": "value"})], ignore_index=True)
            aht_weeks = int(getattr(S, "AHT_CFG", {}).get("level_weeks", 8))
            bits = [f"{k.lower()} {v['level']:.0f} с" for k, v in am.items()]
            print(f"AHT — медиана факта за {aht_weeks} недель до {origin:%d.%m.%Y}: " + ", ".join(bits) + ".")
    except Exception as e:                                   # AHT не должен ломать прогноз объёма
        print(f"   ВНИМАНИЕ: прогноз AHT не построен: {e}")

    # версия
    run_id = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    eff = {**F.DEFAULTS, **cfg}
    params = {"origin": str(origin.date()), "mode": mode,
              "settings": {k: (len(v) if k == "incident_days" and v is not None else v) for k, v in eff.items()
                           if k != "reserve_queue"},
              "reserve_queue": eff["reserve_queue"], "aht_cfg": getattr(S, "AHT_CFG", {}), "log": log,
              "events": ev.assign(date_from=ev["date_from"].astype(str), date_to=ev["date_to"].astype(str)).to_dict("records")}
    status = "опубликован" if publish else "черновик"
    now = pd.Timestamp.now().floor("s")
    model_label = MODEL_VERSION
    if mode == "год":
        acc = ", ".join(season_info["accepted"]) if season_info and season_info["accepted"] else "не принят"
        model_label = MODEL_VERSION + f" · год: месячный индекс {acc} (№138)"
    run = pd.DataFrame([dict(run_id=run_id, created_at=now, kind=f"модель: {mode}",
                             period_from=start, period_to=end, scenario="база (календарь)", status=status,
                             params=json.dumps(params, ensure_ascii=False, default=str), comment=comment,
                             data_to=origin, model=model_label, published_at=now if publish else pd.NaT)])
    root = store_root()
    for name, df in (("forecast_run", run), ("forecast_interval", intervals.assign(run_id=run_id))):
        d = os.path.join(root, name)
        os.makedirs(d, exist_ok=True)
        tb = pa.Table.from_pandas(df, preserve_index=False)
        tmp = os.path.join(d, f"run_{run_id}.parquet.tmp")
        pq.write_table(tb, tmp, compression="zstd")
        os.replace(tmp, os.path.join(d, f"run_{run_id}.parquet"))

    _REPORT_MODE["mode"] = mode
    out = write_report(h, daily, alt, log, ev, origin, start, end, run_id, status, holidays, con, short, intervals,
                       (robot_lbl, robot_alt))
    return run_id, out


def list_runs():
    """Все версии прогноза (паспорта) из DB\\forecast_run — без подключения к хранилищу источник."""
    d = os.path.join(store_root(), "forecast_run")
    files = sorted(f for f in os.listdir(d) if f.endswith(".parquet")) if os.path.isdir(d) else []
    if not files:
        return pd.DataFrame()
    r = pd.concat([pd.read_parquet(os.path.join(d, f)) for f in files], ignore_index=True)
    for c in ("data_to", "model", "published_at", "comment"):
        if c not in r.columns:
            r[c] = None
    return r.sort_values("created_at", ascending=False).reset_index(drop=True)


def publish_run(run_id, comment=""):
    """Черновик -> «опубликован» (№34): по этой версии строятся потребность и графики. Опубликованную не трогаем."""
    path = os.path.join(store_root(), "forecast_run", f"run_{run_id}.parquet")
    if not os.path.exists(path):
        raise ValueError(f"Версия {run_id} не найдена")
    r = pd.read_parquet(path)
    if (r["status"] == "опубликован").any():
        raise ValueError(f"Версия {run_id} уже опубликована — опубликованные версии не меняются (№34)")
    r["status"] = "опубликован"
    r["published_at"] = pd.Timestamp.now().floor("s")
    if comment:
        r["comment"] = comment
    tmp = path + ".tmp"
    pq.write_table(pa.Table.from_pandas(r, preserve_index=False), tmp, compression="zstd")
    os.replace(tmp, path)
    return r.iloc[0]


def previous_published(con, start, end, run_id):
    try:
        runs = wfm_store.query("select run_id, created_at, status from forecast_run", wfm_store.connect())
    except Exception:
        return None, None
    runs = runs[(runs["status"] == "опубликован") & (runs["run_id"] != run_id)].sort_values("created_at")
    if runs.empty:
        return None, None
    rid = runs.iloc[-1]["run_id"]
    con2 = wfm_store.connect()
    df = wfm_store.query(f"select cast(interval_start as date) as day, channel, sum(value) as v from forecast_interval "
                         f"where run_id = '{rid}' group by 1, 2", con2)
    con2.close()
    df["day"] = pd.to_datetime(df["day"])
    return rid, df.pivot_table(index="day", columns="channel", values="v")


def write_report(h, daily, alt, log, ev, origin, start, end, run_id, status, holidays, con, short=None, intervals=None,
                 robot=(None, None)):
    ev = ev[ev["action"] != "фон"]
    names = {k: v[0] for k, v in CHANNELS.items()}
    d = daily[list(CHANNELS)].rename(columns=names).copy()
    d["Доля голоса до очереди"] = daily["voice_share"]
    d["Доля чата до очереди"] = daily["chat_share"]
    d.insert(0, "День недели", [WD[x.weekday()] for x in d.index])
    d.insert(1, "Праздник", ["да" if x in holidays else "" for x in d.index])
    evtxt = []
    for x in d.index:
        t = [f"{r.channel}: {r.action}{(' ' + str(r.name)) if str(getattr(r, 'name', '') or '').strip() else ''}"
             f"{'' if pd.isna(r.multiplier) else ' ×' + str(r.multiplier)}"
             for r in ev.itertuples() if r.date_from <= x <= (r.date_to if pd.notna(r.date_to) else r.date_from)
             or (r.action == "смена режима" and r.date_from == x)]
        evtxt.append("; ".join(t))
    d["События"] = evtxt
    # Вход в КЦ — линия очереди. Чужие линии (смежный отдел, отдел претензий и их чаты), не дошедшие до очереди, из входа вынуты.
    # Прямые звонки во вход входят: они сразу в рабочей очереди. На оператора не пересчитывается.
    side = voice_funnel_side(con, origin)
    chat_side = chat_funnel_side(con, origin)
    d["Голос вход"] = d["Голос вход"] * (1 - side["share"]) + d["Голос прямой"]
    d["Чат вход"] = d["Чат вход"] * (1 - chat_side["share"])
    d = d.drop(columns=["Голос прямой"])
    volume_cols = ["Голос вход", "Голос очереди", "Чат вход", "Чат очереди", "Почта"]

    # сводка по месяцам + прошлая опубликованная версия
    m = d[volume_cols].resample("MS").sum().round(0)
    m.index = m.index.strftime("%Y-%m")
    _, prev = previous_published(con, start, end, run_id)
    if prev is not None:
        pm = prev.reindex(columns=["Голос очереди", "Чат очереди"]).resample("MS").sum()
        pm.index = pm.index.strftime("%Y-%m")
        for ch in ("Голос очереди", "Чат очереди"):
            m[f"{ch}: прошлая опубл."] = pm[ch].reindex(m.index).round(0)
            m[f"{ch}: разница"] = (m[ch] / m[f"{ch}: прошлая опубл."] - 1).round(3)
    if alt is not None:
        am = alt[["voice_queue", "chat_queue"]].resample("MS").sum().round(0)
        am.index = am.index.strftime("%Y-%m")
        m["Голос очереди: без будущих смен режима"] = am["voice_queue"]
        m["Чат очереди: без будущих смен режима"] = am["chat_queue"]
    if robot[1] is not None:
        rm = robot[1]["voice_queue"].resample("MS").sum().round(0)
        rm.index = rm.index.strftime("%Y-%m")
        m[robot[0]] = rm
    if short is not None:
        sm = short[["voice_queue", "chat_queue"]].resample("MS").sum().round(0)
        sm.index = sm.index.strftime("%Y-%m")
        m["Голос очереди: по свежим неделям"] = sm["voice_queue"]
        m["Чат очереди: по свежим неделям"] = sm["chat_queue"]
    if intervals is not None and len(intervals):
        ivm = intervals.pivot_table(index="interval_start", columns="channel", values="value", aggfunc="sum")
        for ac, vc in (("AHT Голос очереди, с", "Голос очереди"), ("AHT Чат очереди, с", "Чат очереди")):
            if ac in ivm.columns and vc in ivm.columns:          # AHT месяца — взвешенный по прогнозу объёма
                a_ = ((ivm[ac] * ivm[vc]).resample("MS").sum() / ivm[vc].resample("MS").sum()).round(0)
                a_.index = a_.index.strftime("%Y-%m")
                m[ac] = a_.reindex(m.index)
    days_in = d.resample("MS").size()
    days_in.index = days_in.index.strftime("%Y-%m")
    m.insert(0, "Дней", days_in)
    m["Без оператора, голос"] = _without_share(m["Голос вход"], m["Голос очереди"])
    m["Без оператора, чат"] = _without_share(m["Чат вход"], m["Чат очереди"])
    if start.day > 1:                                   # первый месяц неполный — подписать
        m = m.rename(index={f"{start:%Y-%m}": f"{start:%Y-%m} (с {start:%d.%m})"})
    m.index.name = "Месяц"
    lab = m.index[0]
    E, G = float(m.loc[lab, "Голос вход"]), float(m.loc[lab, "Голос очереди"])
    Ce, Cg = float(m.loc[lab, "Чат вход"]), float(m.loc[lab, "Чат очереди"])
    wo_example = (lab, E, G, Ce, Cg)

    # бот и робот: история долей по неделям + прогноз + чувствительность
    hh = h[h.index >= origin - pd.Timedelta(weeks=16)]
    wk = pd.DataFrame({
        "Голос вход": hh["voice_entry"], "Голос дошло": hh["voice_entry_queue"],
        "Чат вход": hh["chat_entry"], "Чат дошло": hh["chat_entry_queue"]}).resample("W-SUN", label="left", closed="left").sum()
    wk["Доля голоса до очереди (факт)"] = (wk["Голос дошло"] / wk["Голос вход"]).round(3)
    wk["Доля чата до очереди (факт)"] = (wk["Чат дошло"] / wk["Чат вход"]).round(3)
    wk.index = wk.index.date
    wk.index.name = "Неделя с"
    sens = pd.DataFrame([
        {"Канал": "Голос", "Вход в сутки (прогноз)": round(daily["voice_entry"].mean()),
         "Доля до очереди (прогноз)": round(log["voice_share"]["share"], 3),
         "Режим с": log["voice_share"]["regime_start"],
         "Доля в прошлом режиме": log["voice_share"]["previous_regime_share"],
         "Прошлый режим": log["voice_share"]["previous_regime"],
         "+5 п.п. бота = обращений в сутки": -round(daily["voice_entry"].mean() * 0.05)},
        {"Канал": "Чат", "Вход в сутки (прогноз)": round(daily["chat_entry"].mean()),
         "Доля до очереди (прогноз)": round(log["chat_share"]["share"], 3),
         "Режим с": log["chat_share"]["regime_start"],
         "Доля в прошлом режиме": log["chat_share"]["previous_regime_share"],
         "Прошлый режим": log["chat_share"]["previous_regime"],
         "+5 п.п. бота = обращений в сутки": -round(daily["chat_entry"].mean() * 0.05)},
    ])

    # параметры: без «Голос прямой» — эти звонки уже внутри голоса очереди
    level_rows = []
    for col in F.VOLUME_SERIES:
        if col == "voice_direct":
            continue
        L = log[col]
        wdi = L["weekday_index"]
        level_rows.append({
            "Ряд": names[col],
            "Уровень в сутки": round(float(L["level"]), 1),
            "Окно, недель": L["level_weeks"],
            "Запас на всплески": L["reserve"],
            "Излом уровня": ", ".join(L["volume_breaks"]) or "нет",
            **{f"Индекс {k}": wdi[k] for k in WD},
            "Исключённые дни": ", ".join(L["excluded_days"][-15:]) or "нет",
        })
    level = pd.DataFrame(level_rows)

    def _txt(v):
        if isinstance(v, list):
            return ", ".join(str(x) for x in v) or "нет"
        if isinstance(v, dict):
            return ", ".join(f"{k}: {val}" for k, val in v.items()) or "нет"
        if v is None or (isinstance(v, str) and v.strip() in ("", "{}", "None")):
            return "нет"
        return v

    share_rows = []
    for title, key in (("Голос", "voice_share"), ("Чат", "chat_share")):
        info = log[key]
        share_rows.append({
            "Канал": title,
            "Доля до оператора": info.get("share"),
            "Режим с": info.get("regime_start"),
            "Дней в расчёте": info.get("days_used"),
            "Изломы по данным": _txt(info.get("auto_breaks")),
            "Изломы из календаря": _txt(info.get("manual_breaks")),
            "Будущая смена": _txt(info.get("future_changes")),
            "Запас в очереди": info.get("reserve_queue"),
            "Прошлая доля": info.get("previous_regime_share"),
            "Прошлый режим": info.get("previous_regime") or "",
            "Заметка": info.get("note") or "",
        })
    shares = pd.DataFrame(share_rows)
    if shares["Заметка"].astype(str).str.strip().eq("").all():
        shares = shares.drop(columns=["Заметка"])

    wkly = d[volume_cols].resample("W-SUN", label="left", closed="left").sum().round(0)
    d["Без оператора, голос"] = _without_share(d["Голос вход"], d["Голос очереди"])
    d["Без оператора, чат"] = _without_share(d["Чат вход"], d["Чат очереди"])
    wkly["Без оператора, голос"] = _without_share(wkly["Голос вход"], wkly["Голос очереди"])
    wkly["Без оператора, чат"] = _without_share(wkly["Чат вход"], wkly["Чат очереди"])
    wkly.index = [f"{x:%d.%m}–{x + pd.Timedelta(days=6):%d.%m}" for x in wkly.index]
    wkly.index.name = "Неделя (пн–вс)"
    promo = pd.DataFrame(log.get("promo", []))
    if len(promo):
        promo["series"] = promo["series"].map(lambda x: names.get(x, x))
        promo = promo.rename(columns={"series": "Ряд", "name": "Акция", "period": "Период", "source": "Откуда множитель",
                                      "m_first_week": "Множитель", "m_set": "Множитель из календаря"})
        promo = promo.drop(columns=[c for c in ("m_rest",) if c in promo.columns])
    ny = pd.DataFrame([{"Ряд": names.get(k, k),
                        "Коэффициент к воскресенью": v["factor"] if v.get("factor") else "мало лет",
                        "Лет истории": v.get("years") or 0}
                       for k, v in log.get("new_year", {}).items()])
    pr = pd.DataFrame([{"Ряд": names.get(k, k), "Запас на акции": v["rate"], "Запусков за год": v["runs"]}
                       for k, v in log.get("promo_reserve", {}).items()])

    # 30 минут — один лист. Часы и матрицы по каналам те же числа.
    it_long = None
    if intervals is not None and len(intervals):
        order = ["Голос вход", "Голос очереди", "Чат вход", "Чат очереди", "Почта"]
        aht_cols = ["AHT Голос очереди, с", "AHT Чат очереди, с"]
        ivall = intervals.pivot_table(index="interval_start", columns="channel", values="value", aggfunc="sum")
        if "Голос вход" in ivall.columns:
            direct_iv = ivall["Голос прямой"] if "Голос прямой" in ivall.columns else 0
            ivall["Голос вход"] = ivall["Голос вход"] * (1 - side["share"]) + direct_iv
        if "Чат вход" in ivall.columns:
            ivall["Чат вход"] = ivall["Чат вход"] * (1 - chat_side["share"])
        iv = ivall[[c for c in order if c in ivall.columns]]
        av = ivall[[c for c in aht_cols if c in ivall.columns]]
        ts = iv.index
        lab = [f"{t:%H:%M}–{t + pd.Timedelta(minutes=30):%H:%M}" for t in ts]
        it_long = pd.concat([iv, av.round(0)], axis=1)
        it_long.insert(0, "Тип дня", [F.day_type(t.normalize(), holidays) for t in ts])
        it_long.insert(0, "День недели", [WD[t.weekday()] for t in ts])
        it_long.insert(0, "Интервал", lab)
        it_long.insert(0, "Дата", ts.normalize().date)
        it_long = it_long.reset_index(drop=True)

    lead = ["Вход в КЦ, голос", "На оператора, голос", "Без оператора, голос",
            "Вход в КЦ, чат", "На оператора, чат", "Без оператора, чат", "Почта"]
    m = _order(_relabel(m), ["Дней"] + lead)
    d = _order(_relabel(d), ["День недели", "Праздник"] + lead)
    wkly = _order(_relabel(wkly), lead)
    if it_long is not None:
        it_long = _relabel(it_long)

    out = report_path(f"Прогноз {status_mode_name(run_id)}.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        m.to_excel(w, sheet_name="Сводка")
        dd = d.copy()
        dd.index = dd.index.date
        dd.index.name = "Дата"
        dd.to_excel(w, sheet_name="По_дням")
        if it_long is not None:
            it_long.to_excel(w, sheet_name="30 мин", index=False)
        wkly.to_excel(w, sheet_name="По_неделям")
        sens.to_excel(w, sheet_name="Бот_и_робот", index=False)
        wk.to_excel(w, sheet_name="Бот_и_робот", startrow=len(sens) + 3)
        pws = w.book.create_sheet("Параметры")
        lines = _param_lines(log, side, chat_side, wo_example, origin, start, end, run_id, status,
                             intervals, fresh=short is not None)
        param_headers, explain_n = write_parameters(
            pws, level, shares, ny, pr, promo if len(promo) else None, lines)
        for wsn in w.book.sheetnames:
            ws = w.book[wsn]
            for c in ws[1]:
                if c.value is not None:
                    c.fill, c.font = HDR_FILL, HDR_FONT
            for col in ws.columns:
                letter = get_column_letter(col[0].column)
                width = max((len(str(c.value)) for c in col[:60] if c.value is not None), default=0)
                ws.column_dimensions[letter].width = min(max(10, width + 2), 45)
            if wsn not in ("Параметры", "30 мин"):
                ws.freeze_panes = "B2"
            fmt_sheet(ws)
        if "30 мин" in w.book.sheetnames:
            w.book["30 мин"].freeze_panes = "E2"
        for hr in param_headers:
            for c in pws[hr]:
                if c.value is not None:
                    c.fill, c.font = HDR_FILL, HDR_FONT
        if "Исключённые дни" in level.columns:
            pws.column_dimensions[get_column_letter(list(level.columns).index("Исключённые дни") + 1)].width = 55
        for letter, width in (("A", 16), ("B", 24), ("C", 62), ("D", 36)):
            pws.column_dimensions[letter].width = width
        pws.freeze_panes = "A2"
    return out


SIDE_VOICE = ("Входящая линия смежного отдела", "Входящая линия (отдел претензий)")
SIDE_CHAT = ("чат VIP и смежный отдел", "чат отдел претензий", "чат Завершение и оценка (смежный отдел)", "чат Завершение и оценка отдел претензий")


def _relabel(df):
    """В книге: вход в КЦ и на оператора. Имена каналов в хранилище не меняются."""
    ren = {}
    for col in df.columns:
        name = str(col)
        if name.startswith("AHT"):
            continue
        new = name
        for old, nxt in (("Голос вход", "Вход в КЦ, голос"), ("Чат вход", "Вход в КЦ, чат"),
                         ("Голос очереди", "На оператора, голос"), ("Чат очереди", "На оператора, чат")):
            if new == old or new.startswith(old + ":"):
                new = nxt + new[len(old):]
                break
        if new != name:
            ren[col] = new
    return df.rename(columns=ren)


def _order(df, lead):
    front = [c for c in lead if c in df.columns]
    return df[front + [c for c in df.columns if c not in front]]


def _without_share(entry, on_operator):
    """Без оператора = (вход в КЦ − на оператора) / вход в КЦ."""
    if isinstance(entry, pd.Series):
        den = entry.replace(0, pd.NA)
        return ((entry - on_operator) / den).astype(float)
    if not entry:
        return None
    return (entry - on_operator) / entry


def voice_funnel_side(con, origin, weeks=6):
    """Доля входа меню, ушедшая на смежный отдел или отдел претензий и не вставшая в рабочую очередь. Журнал №106."""
    origin = pd.Timestamp(origin).normalize()
    lo = origin - pd.Timedelta(weeks=weeks)
    hi = origin + pd.Timedelta(days=1)
    titles = ", ".join("'" + t.replace("'", "''") + "'" for t in SIDE_VOICE)
    sql = f"""
    with ids as (
        select uuid as project_id from dim_projects where title in ({titles})
    ),
    q as (
        select session_id, project_id,
               row_number() over (partition by session_id order by enqueued_time, id) as rn
        from voice_queued_calls
        where enqueued_time >= timestamp '{lo:%Y-%m-%d}'
          and enqueued_time < timestamp '{hi:%Y-%m-%d}'
    ),
    entry_s as (
        select q.session_id
        from q join voice_entry_projects e on e.project_id = q.project_id
        where q.rn = 1
    ),
    one as (
        select q.session_id,
               max(case when g.project_id is not null then 1 else 0 end) as reached,
               max(case when i.project_id is not null then 1 else 0 end) as side
        from q
        join entry_s s on s.session_id = q.session_id
        left join support_queues g on g.project_id = q.project_id and g.channel = 'Голос'
        left join ids i on i.project_id = q.project_id
        group by q.session_id
    )
    select count(*) as menu,
           sum(case when reached = 0 and side = 1 then 1 else 0 end) as side
    from one
    """
    df = wfm_store.query(sql, con)
    menu = float(df["menu"].iloc[0] or 0)
    side = float(df["side"].iloc[0] or 0)
    return {"menu": menu, "side": side, "share": (side / menu) if menu else 0.0,
            "lo": lo, "hi": origin, "weeks": weeks}


def chat_funnel_side(con, origin, weeks=6):
    """Доля чатов, ушедших на VIP/смежный отдел или отдел претензий и не вставших в рабочую очередь. Журнал №107."""
    origin = pd.Timestamp(origin).normalize()
    lo = origin - pd.Timedelta(weeks=weeks)
    hi = origin + pd.Timedelta(days=1)
    titles = ", ".join("'" + t.replace("'", "''") + "'" for t in SIDE_CHAT)
    sql = f"""
    with s as (
        select i.session_id,
               max(case when g.project_id is not null then 1 else 0 end) as reached,
               max(case when p.title in ({titles}) then 1 else 0 end) as side
        from chat_session_iterations i
        left join dim_projects p on p.uuid = i.project_id
        left join support_queues g on g.project_id = i.project_id and g.channel = 'Чат'
        where coalesce(p.datachannel, 'im') <> 'email'
          and i.start_time >= timestamp '{lo:%Y-%m-%d}'
          and i.start_time < timestamp '{hi:%Y-%m-%d}'
        group by i.session_id
    )
    select count(*) as menu,
           sum(case when reached = 0 and side = 1 then 1 else 0 end) as side
    from s
    """
    df = wfm_store.query(sql, con)
    menu = float(df["menu"].iloc[0] or 0)
    side = float(df["side"].iloc[0] or 0)
    return {"menu": menu, "side": side, "share": (side / menu) if menu else 0.0,
            "lo": lo, "hi": origin, "weeks": weeks}


def _num(x, digits=0):
    return f"{float(x):,.{digits}f}".replace(",", " ").replace(".", ",")


def _pct(x):
    return f"{float(x):.1%}".replace(".", ",")


def _when(value):
    t = pd.Timestamp(value)
    return f"{t:%d.%m.%Y}" if pd.notna(t) else str(value)


def _weeks(w):
    w = float(w)
    return _num(w, 0 if abs(w - round(w)) < 0.05 else 1)


def _row(block, step, how, num=""):
    return (block, step, how, num)


def _wrap_lines(text, width):
    text = str(text or "")
    per = max(int(width * 0.85), 8)
    return max(1, -(-len(text) // per))


def _level_rows(block, L, rules):
    """Обычный день, запас и индекс — отдельными строками."""
    level = float(L["level"])
    reserve = float(L["reserve"])
    weeks = L.get("level_window_weeks") or L.get("level_weeks") or 0
    fresh = _REPORT_MODE.get("mode") == "неделя"
    how = "Взвешенная медиана" if fresh else "Медиана"
    tail = " Свежие недели весят больше: вес падает вдвое каждые 3 недели." if fresh else ""
    if L.get("level_from") and L.get("level_days"):
        src = (f"{how} {_num(L['level_days'])} чистых дней с {_when(L['level_from'])} по {_when(L['level_to'])}, "
               f"окно {_weeks(weeks)} недель. Перед медианой каждые сутки делим на индекс их дня недели.{tail}")
    else:
        src = (f"{how} чистых дней за {_weeks(weeks)} недель. "
               f"Перед медианой каждые сутки делим на индекс их дня недели.{tail}")
    rows = [_row(block, "Обычный день", src, _num(level, 1))]
    spike = float(rules.get("spike_up") or 0.5)
    rows.append(_row(
        block, "Какие дни не берём",
        f"День выше или ниже соседних таких же дней недели больше чем на {_pct(spike)}. "
        f"Соседние — медиана того же дня недели за 3 недели до и 3 после. "
        f"Список — колонка «Исключённые дни» в таблице ниже.",
    ))
    ex, tot = L.get("reserve_excess"), L.get("reserve_total")
    rw = rules.get("reserve_weeks") or 12
    if reserve > 0 and ex is not None and tot:
        cap = ""
        if float(tot) and float(ex) / float(tot) > reserve + 1e-6:
            cap = f" Больше {_pct(rules.get('reserve_cap') or 0.3)} не берём."
        rows.append(_row(
            block, "Запас на всплески",
            f"(факт − обычный день) по дням всплесков / все обращения за {_weeks(rw)} недель.{cap}",
            f"{_num(ex, 1)} / {_num(tot, 1)} = {_pct(reserve)}",
        ))
        rows.append(_row(block, "Делитель", "1 − запас на всплески.", _num(1 - reserve, 3)))
        rows.append(_row(
            block, "Сутки с запасом", "Обычный день делим на делитель.",
            f"{_num(level, 1)} / {_num(1 - reserve, 3)} = {_num(level / (1 - reserve), 1)}",
        ))
        index_how = "Сутки с запасом умножаем на индекс дня недели. "
    elif reserve > 0:
        rows.append(_row(block, "Делитель", f"1 − запас {_pct(reserve)}.", _num(1 - reserve, 3)))
        rows.append(_row(
            block, "Сутки с запасом", "Обычный день делим на делитель.",
            f"{_num(level, 1)} / {_num(1 - reserve, 3)} = {_num(level / (1 - reserve), 1)}",
        ))
        index_how = "Сутки с запасом умножаем на индекс дня недели. "
    else:
        rows.append(_row(block, "Запас на всплески", "Всплесков в окне нет, обычный день не делим.", "0"))
        index_how = "Обычный день умножаем на индекс дня недели. "
    rows.append(_row(
        block, "Индекс дня",
        index_how + "Индекс = среднее этого дня недели / среднее всех дней в том же окне. "
        "Праздник берёт индекс воскресенья.",
        "таблица «Уровень»",
    ))
    return rows


def _calendar_rows(block, col, log):
    """Новый год и акции — отдельные строки, только если в прогоне они есть."""
    rows = []
    ny = (log.get("new_year") or {}).get(col) or {}
    factor = ny.get("factor")
    if factor:
        rows.append(_row(
            block, "1–10 января",
            "Эти дни делим на воскресенья 11.01–07.02 того же года и берём медиану по годам.",
            f"× {_num(factor, 3)}, лет: {int(ny.get('years') or 0)}",
        ))
    elif col in ("voice_entry", "chat_entry") and "new_year" in log:
        rows.append(_row(
            block, "1–10 января",
            "Отдельного множителя нет: таких лет меньше двух. День остаётся как воскресенье.",
            "× 1",
        ))
    if any(p.get("series") == col for p in (log.get("promo") or [])):
        rows.append(_row(block, "День акции", "Умножаем на множитель из таблицы «Акции».", "таблица «Акции»"))
    reserve = (log.get("promo_reserve") or {}).get(col) or {}
    rate = float(reserve.get("rate") or 0)
    if rate > 0:
        ex, tot = reserve.get("excess"), reserve.get("total")
        weeks = reserve.get("weeks") or 52
        num = f"делим на {_num(1 - rate, 3)}"
        if ex is not None and tot:
            num = f"{_num(ex, 1)} / {_num(tot, 1)} = {_pct(rate)}; делим на {_num(1 - rate, 3)}"
        rows.append(_row(
            block, "День без акции",
            f"Насколько прошлые акции были выше обычного дня, делим на все обращения за {_weeks(weeks)} недель. "
            f"Запусков: {reserve.get('runs')}.",
            num,
        ))
    return rows


def _queue_rows(block, info, entry, origin, rules):
    share = float(info["share"])
    rr_e = float(entry["reserve"])
    rr_g = float(info.get("reserve_queue") or 0)
    e_sum, r_sum = info.get("entry_sum"), info.get("reached_sum")
    frac = f"{_num(r_sum, 1)} / {_num(e_sum, 1)} = {_pct(share)}" if e_sum else _pct(share)
    frm = _when(info["share_from"]) if info.get("share_from") else _when(info.get("regime_start"))
    sw = int(rules.get("share_weeks") or 6)
    rows = [_row(
        block, "До оператора",
        f"Дошедшие до оператора / все обращения с {frm} по {_when(origin)}. "
        f"Режим с {_when(info.get('regime_start'))}, дней {info.get('days_used')}. Окно не длиннее {sw} недель.",
        frac,
    )]
    if info.get("note"):
        rows.append(_row(block, "До оператора, заметка", str(info["note"]).rstrip("."), ""))
    if rr_g > 0:
        rows.append(_row(
            block, "На оператора",
            "Сутки × доля до оператора × (1 − запас входа) / (1 − запас по дошедшим). "
            "Первый множитель снимает запас, уже заложенный в сутки.",
            f"× {_pct(share)} × {_num(1 - rr_e, 3)} / {_num(1 - rr_g, 3)}",
        ))
        tail = _reserve_tail(info.get("reserve_queue_excess"), info.get("reserve_queue_total"), info.get("reserve_queue"), rules)
        if tail:
            rows.append(_row(block, "Запас по дошедшим", "Считается так же, как запас на всплески, но по дошедшим до оператора.", tail[2:]))
    elif rr_e > 0:
        rows.append(_row(
            block, "На оператора",
            "Сутки × доля до оператора × (1 − запас). Запас всплеска в очередь не входит.",
            f"× {_pct(share)} × {_num(1 - rr_e, 3)}",
        ))
    else:
        rows.append(_row(block, "На оператора", "Сутки × доля до оператора.", f"× {_pct(share)}"))
    fut = info.get("future_changes") or {}
    if isinstance(fut, dict) and fut:
        nxt = ", ".join(f"{_when(d)} → {_pct(v)}" for d, v in fut.items())
        rows.append(_row(block, "Дальше по календарю", "Доля с этих дат берётся из календаря.", nxt))
    return rows


def _reserve_tail(excess, total, rate, rules):
    if excess is None or not total or not rate:
        return ""
    rw = rules.get("reserve_weeks") or 12
    return f": {_num(excess, 1)} / {_num(total, 1)} за {_weeks(rw)} недель"


def _aht_row(intervals, origin):
    if intervals is None or not len(intervals) or "channel" not in intervals.columns:
        return None
    weeks = int(getattr(S, "AHT_CFG", {}).get("level_weeks", 8))
    bits = []
    for ch, label in (("AHT Голос очереди, с", "голос"), ("AHT Чат очереди, с", "чат")):
        part = intervals.loc[intervals["channel"] == ch, "value"]
        if len(part):
            bits.append(f"{label} {_num(float(part.median()), 1)} с")
    if not bits:
        return None
    return _row(
        "Общее", "AHT",
        f"По каждому дню делим время обработки на число обработанных и берём медиану этих дней "
        f"за {weeks} недель до {_when(origin)}. На все получасы одно число.",
        ", ".join(bits),
    )


def _param_lines(log, side, chat_side, wo_example, origin, start, end, run_id, status, intervals, fresh=False):
    """Строки листа: блок, шаг, откуда цифра, результат."""
    rules = log.get("rules") or {}
    stamp = dt.datetime.strptime(run_id, "%Y-%m-%d_%H%M%S")
    rows = [
        _row("Прогон", "Когда", "Черновик или опубликованная версия.", f"{stamp:%d.%m.%Y %H:%M}, {status}"),
        _row("Прогон", "Данные по", "Последний день факта, который вошёл в расчёт.", f"{origin:%d.%m.%Y}"),
        _row("Прогон", "Период", "Какие даты прогнозируем.", f"{start:%d.%m.%Y}–{end:%d.%m.%Y}"),
    ]
    rows += _level_rows("Голос", log["voice_entry"], rules)
    rows += _calendar_rows("Голос", "voice_entry", log)
    rows.append(_row(
        "Голос", "Другие линии",
        f"Ушли на другие линии и до оператора не дошли / все обращения "
        f"за {side['lo']:%d.%m.%Y}–{side['hi']:%d.%m.%Y}. "
        f"Вход = сутки × (1 − эта доля) + звонки сразу в очередь.",
        f"{_num(side['side'])} / {_num(side['menu'])} = {_pct(side['share'])}",
    ))
    rows += _queue_rows("Голос", log["voice_share"], log["voice_entry"], origin, rules)
    rows.append(_row(
        "Голос", "Сразу в очередь",
        "Эти звонки прибавляем и ко входу, и к тем, кто дошёл до оператора.",
        "",
    ))
    rows += _level_rows("Чат", log["chat_entry"], rules)
    rows += _calendar_rows("Чат", "chat_entry", log)
    rows.append(_row(
        "Чат", "Другие линии",
        f"Ушли на другие линии и до оператора не дошли / все чаты "
        f"за {chat_side['lo']:%d.%m.%Y}–{chat_side['hi']:%d.%m.%Y}. "
        f"Вход = сутки × (1 − эта доля).",
        f"{_num(chat_side['side'])} / {_num(chat_side['menu'])} = {_pct(chat_side['share'])}",
    ))
    rows += _queue_rows("Чат", log["chat_share"], log["chat_entry"], origin, rules)
    if "mail" in log:
        rows += _level_rows("Почта", log["mail"], rules)
        rows.append(_row("Почта", "Доля", "Доли нет: вход равен числу на оператора.", "без оператора 0%"))
    rows.append(_row(
        "Общее", "30 минут",
        "Сутки умножаем на долю получаса. Доля = обращения этого получаса / обращения суток, "
        "по факту 8 недель, отдельно будни, суббота и воскресенье. Праздник считаем воскресеньем.",
    ))
    aht = _aht_row(intervals, origin)
    if aht:
        rows.append(aht)
    robot = log.get("robot") or {}
    if robot.get("on_now") is False and robot.get("share_robot_on"):
        rows.append(_row("Общее", "Робот", "Сейчас выключен. Колонка сценария — если заработает.", _pct(robot["share_robot_on"])))
    elif robot.get("on_now") is True and robot.get("share_robot_off"):
        rows.append(_row("Общее", "Робот", "Сейчас включён. Колонка сценария — если отключат.", _pct(robot["share_robot_off"])))
    lab, E, G, Ce, Cg = wo_example
    rows.append(_row(
        "Общее", "Без оператора",
        "(вход − на оператора) / вход.",
        f"{lab}. Голос: {_num(E - G)} / {_num(E)} = {_pct((E - G) / E)}. "
        f"Чат: {_num(Ce - Cg)} / {_num(Ce)} = {_pct((Ce - Cg) / Ce)}.",
    ))
    if fresh:
        rows.append(_row("Общее", "По свежим неделям", "Колонка на сводке: тот же расчёт, свежие недели весят больше.", ""))
    return rows


def write_parameters(ws, level, shares, ny, promo_reserve, promo=None, explain=None):
    """Лист «Параметры»: подписи по-русски, каждый блок со своей шапкой."""
    blocks = [("Уровень", level), ("Доля до оператора", shares)]
    if ny is not None and len(ny):
        blocks.append(("Новый год, 1–10 января", ny))
    if promo_reserve is not None and len(promo_reserve):
        blocks.append(("Запас на крупные акции", promo_reserve))
    if promo is not None and len(promo):
        blocks.append(("Акции в периоде", promo))
    headers = []
    row = 1
    explain_n = 0
    edge = Side(style="thin", color="FFD9D9D9")
    thin = Border(left=edge, right=edge, top=edge, bottom=edge, diagonal=Side())
    if explain:
        headers.append(row)
        for c, name in enumerate(("Блок", "Шаг", "Как получили", "Число"), 1):
            cell = ws.cell(row, c, name)
            cell.fill, cell.font = HDR_FILL, HDR_FONT
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            cell.border = thin
        row += 1
        for rec in explain:
            vals = rec if isinstance(rec, tuple) else ("", "", rec, "")
            for c, v in enumerate(vals[:4], 1):
                cell = ws.cell(row, c, v or None)
                cell.alignment = Alignment(wrap_text=True, vertical="center")
                cell.font = Font(color="1F3864")
                cell.border = thin
            how_n = _wrap_lines(vals[2] if len(vals) > 2 else "", 62)
            num_n = _wrap_lines(vals[3] if len(vals) > 3 else "", 36)
            ws.row_dimensions[row].height = max(18, 15 * max(how_n, num_n))
            explain_n = row
            row += 1
        row += 1
    for title, frame in blocks:
        ws.cell(row, 1, title).font = Font(bold=True, color="1F3864")
        row += 1
        headers.append(row)
        for c, name in enumerate(frame.columns, 1):
            cell = ws.cell(row, c, name)
            cell.fill, cell.font = HDR_FILL, HDR_FONT
        for rec in frame.itertuples(index=False):
            row += 1
            for c, v in enumerate(rec, 1):
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    continue
                cell = ws.cell(row, c, v.item() if hasattr(v, "item") and not isinstance(v, str) else v)
                _fmt_param(cell, frame.columns[c - 1])
        row += 2
    ws.column_dimensions["A"].width = 28
    if "Исключённые дни" in level.columns:
        ws.column_dimensions[get_column_letter(list(level.columns).index("Исключённые дни") + 1)].width = 55
    return headers, explain_n


def _fmt_param(cell, header):
    if not isinstance(cell.value, (int, float)) or isinstance(cell.value, bool):
        return
    h = str(header)
    if h.startswith("Запас") or h.startswith("Доля") or h.startswith("Прошлая"):
        cell.number_format = "0.0%"
    elif h.startswith("Индекс") or h.startswith("Коэффициент"):
        cell.number_format = "0.000"
    elif h.startswith("Уровень"):
        cell.number_format = "#,##0.0"


def fmt_sheet(ws):
    """Форматы чисел: обращения — целые с разделителем (за 30 мин и час — с десятыми), доли — %, даты — ДД.ММ.ГГГГ."""
    title = ws.title
    small = title in ("30 мин", "Часы") or title.endswith("30 мин")
    hdr = [c.value for c in ws[1]]
    for row in ws.iter_rows(min_row=2):
        for c in row:
            v, h_ = c.value, hdr[c.column - 1] if c.column - 1 < len(hdr) else None
            if isinstance(v, (dt.date, dt.datetime, pd.Timestamp)):
                c.number_format = "DD.MM.YYYY"
            elif isinstance(v, float) or isinstance(v, int) and not isinstance(v, bool):
                h_ = str(h_ or "")
                if "Доля" in h_ or "разница" in h_ or "Запас" in h_ or "Без оператора" in h_ or title == "Бот_и_робот" and "Доля" in h_:
                    c.number_format = "0.0%"
                elif title == "Параметры":
                    continue
                else:
                    c.number_format = "#,##0.0" if small else "#,##0"


# ---------------------------------------------------------------------------
# Варианты бэктеста горизонта квартала. Проверены 29.10 (№41) и отклонены: уровень 8 / 16 нед, без запаса,
# запас чата в очереди, без новогоднего коэффициента — чтобы прогнать их снова, верните в словарь:
#   "уровень 8 нед": dict(level_weeks=8, weekday_weeks=8), "уровень 16 нед": dict(level_weeks=16, weekday_weeks=16),
#   "без запаса на всплески": dict(use_reserve=False), "чат: запас и в очереди": dict(reserve_queue={"voice": True, "chat": True}),
#   "без новогоднего коэфф.": dict(new_year=False)
# Варианты с инцидентами добавляются сами, если в inputs есть выгрузка Service Desk.
# Квартал без знания будущих акций (как в реальности: Jira — за 2–3 недели): акции «акция» с датой начала после точки
# прогноза из календаря убираются. Вариант с ключом _future_promos знает их заранее (только для справки).
BT_VARIANTS = {
    "v2 (текущая)": {},
    "v2 + запас на акции": dict(promo_reserve=True),
    "v2 + запас на акции + всплеск 30%": dict(promo_reserve=True, spike_up=0.3),
    "v2, акции известны заранее (справка)": dict(_future_promos=True),
}
BT_HORIZONS = (1, 2, 3, 4, 5)


def _bt_cfg(cfg):
    return {k: v for k, v in cfg.items() if not k.startswith("_")}


def _bt_events(ev, origin, cfg):
    """Квартальный бэктест: будущие «акции» неизвестны (кроме варианта _future_promos)."""
    if cfg.get("_future_promos") or ev is None or ev.empty:
        return ev
    return ev[(ev["action"] != "акция") | (ev["date_from"] <= origin)]          # месяцев вперёд; Q1 от начала ноября = 3 (январь), 4 (февраль), 5 (март)


def run_backtest(con):
    """Бэктест «как в трекере» на горизонте квартала: факт / прогноз за месяц, прогноз сделан за 1–5 месяцев
    (точка прогноза — последний день месяца). Варианты BT_VARIANTS. Отдельно — аналоги нашего случая:
    прогноз от 31.10 на январь–март следующего года (где хватает истории; голос — только с VOICE_EXT_FROM).
    Плюс неделя: недельный пересчёт против квартальной модели (последние 26 недель).
    Календарь events.xlsx и праздники считаются известными заранее (как при реальном планировании)."""
    h = load_history(con)
    ev = load_events()
    holidays = pd.DatetimeIndex(pd.to_datetime(getattr(S, "HOLIDAYS", [])))
    h["voice_queue"] = h["voice_entry_queue"] + h["voice_direct"]
    h["chat_queue"] = h["chat_entry_queue"]
    end = h.dropna(subset=["voice_entry", "chat_entry"]).index.max()
    names = {"voice_entry": "Голос вход", "voice_queue": "Голос очереди", "chat_entry": "Чат вход", "chat_queue": "Чат очереди",
             "mail": "Почта"}
    starts = {c: F.series_start(h[c].dropna()) for c in names}
    rows = []
    variants = dict(BT_VARIANTS)
    weekly = {"недельный пересчёт": WEEKLY_CFG, "v2 (текущая)": {}}
    import incidents
    if getattr(S, "BT_WITH_INCIDENTS", False) and incidents.load() is not None:   # №43: по умолчанию не гоняем
        inc_all = incidents.incident_days(["КЦ", "клиентский онлайн"])
        inc_kc = incidents.incident_days(["КЦ"])
        print(f"Инциденты: дней КЦ + клиентский онлайн {len(inc_all)}, только КЦ {len(inc_kc)}")
        variants.update({
            "+ инциденты (КЦ+клиент, 20%)": dict(incident_days=inc_all, inc_spike_up=0.2),
            "+ инциденты (КЦ+клиент, 30%)": dict(incident_days=inc_all, inc_spike_up=0.3),
            "+ инциденты (только КЦ, 20%)": dict(incident_days=inc_kc, inc_spike_up=0.2)})
        weekly["недельный + инциденты (КЦ+клиент, 20%)"] = {**WEEKLY_CFG, "incident_days": inc_all, "inc_spike_up": 0.2}

    def rec(block, variant, kind, per, o, s, e, f):
        for c, nm in names.items():
            fa = h.loc[s:e, c]
            if starts[c] is None or o < starts[c] + pd.Timedelta(weeks=8) or fa.isna().any() or len(fa) == 0 \
                    or not np.isfinite(f[c].sum()) or f[c].sum() <= 0:
                continue
            rows.append({"Блок": block, "Вариант": variant, "Ряд": nm, "Горизонт": kind, "Период": per,
                         "Прогноз_от": o.date(), "Факт": round(fa.sum()), "Прогноз": round(f[c].sum()),
                         "Факт/прогноз, %": round(fa.sum() / f[c].sum() * 100, 1),
                         "Ошибка по дням, %": round((fa - f[c]).abs().sum() / fa.sum() * 100, 1)})

    # 1) месяцы последних 12 месяцев × горизонт 1–5 мес × варианты
    months = pd.period_range((end - pd.DateOffset(months=12)).to_period("M"), end.to_period("M"), freq="M")
    total = sum(1 for m in months if (min(m.end_time.normalize(), end) - m.start_time).days >= 13) \
        * len(BT_HORIZONS) * len(variants)
    print(f"Бэктест: месяцы {months[0]} – {months[-1]}, горизонт {BT_HORIZONS[0]}–{BT_HORIZONS[-1]} мес, "
          f"вариантов {len(variants)} ({total} прогнозов); история с {h.index.min():%d.%m.%Y}. Это займёт минуты...")
    done = 0
    for m in months:
        s, e = m.start_time, min(m.end_time.normalize(), end)
        if (e - s).days < 13:
            continue
        for k in BT_HORIZONS:
            o = (m - k + 1).start_time - pd.Timedelta(days=1)
            for vn, cfg in variants.items():
                f, _ = F.forecast_daily(h, o, s, e, _bt_cfg(cfg), _bt_events(ev, o, cfg), holidays)
                rec("месяцы", vn, f"за {k} мес", str(m), o, s, e, f)
                done += 1
        print(f"   {m}: готово ({done}/{total})")

    # 2) аналоги Q1: прогноз от 31.10 на январь–март следующего года
    for y in range(h.index.min().year + 1, end.year + 1):
        o = pd.Timestamp(y - 1, 10, 31)
        for vn, cfg in variants.items():
            f, _ = F.forecast_daily(h, o, pd.Timestamp(y, 1, 1), pd.Timestamp(y, 3, 31), _bt_cfg(cfg), _bt_events(ev, o, cfg),
                                    holidays)
            for mi, mname in ((1, "январь"), (2, "февраль"), (3, "март")):
                s = pd.Timestamp(y, mi, 1)
                e = min(s + pd.offsets.MonthEnd(0), end)
                if e >= s:
                    rec("Q1 от 31.10", vn, mname, f"{y}-{mi:02d}", o, s, e, f.loc[s:e])
            s, e = pd.Timestamp(y, 1, 1), min(pd.Timestamp(y, 1, 10), end)
            if e >= s:
                rec("Q1 от 31.10", vn, "1–10 января", f"{y}-01-01..10", o, s, e, f.loc[s:e])

    # 3) неделя: недельный пересчёт против квартальной модели
    for o in pd.date_range(end - pd.Timedelta(weeks=26), end - pd.Timedelta(days=7), freq="W-SUN"):
        s, e = o + pd.Timedelta(days=1), o + pd.Timedelta(days=7)
        for vn, cfg in weekly.items():
            f, _ = F.forecast_daily(h, o, s, e, cfg, ev, holidays)
            rec("недели", vn, "неделя", f"{s:%d.%m}–{e:%d.%m}", o, s, e, f)

    det = pd.DataFrame(rows)
    det["dev"] = (det["Факт/прогноз, %"] - 100).abs()
    agg = lambda d, by: d.groupby(by).agg(**{
        "Среднее отклонение, %": ("dev", "mean"),
        "Смещение (факт/прогноз − 100), %": ("Факт/прогноз, %", lambda x: (x - 100).mean()),
        "В пределах ±10%": ("dev", lambda x: (x <= 10).mean()),
        "Ошибка по дням, %": ("Ошибка по дням, %", "mean"),
        "Точек": ("dev", "size")}).round(2).reset_index()
    mon = det[det["Блок"] == "месяцы"]
    tbl = agg(mon, ["Ряд", "Горизонт", "Вариант"])
    q1h = agg(mon[mon["Горизонт"].isin(["за 3 мес", "за 4 мес", "за 5 мес"])], ["Ряд", "Вариант"])   # горизонт Q1
    q1h = q1h.sort_values(["Ряд", "Среднее отклонение, %"])
    ana = det[det["Блок"] == "Q1 от 31.10"].drop(columns=["dev", "Блок"])
    wk = agg(det[det["Блок"] == "недели"], ["Ряд", "Вариант"])

    out = report_path(f"Бэктест объёма {stamp()}.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        q1h.to_excel(w, sheet_name="Итог_горизонт_Q1", index=False)
        ws = w.book["Итог_горизонт_Q1"]
        notes = ["Горизонт Q1 = месяцы, спрогнозированные за 3–5 месяцев (как прогноз от начала ноября на январь–март).",
                 "Отклонение = |факт / прогноз − 100%| по сумме месяца. Смещение > 0 — прогноз занижен.",
                 "Прогноз на каждую точку строится только по данным до даты «Прогноз_от» (без заглядывания в будущее).",
                 "Квартал: будущие акции из календаря НЕ известны (как в реальности), кроме варианта «акции известны заранее».",
                 "Вариант лучше текущего, если ниже отклонение по очереди-рядам и смещение не хуже; разница < 1 п.п. — шум."]
        for i, n_ in enumerate(notes, start=len(q1h) + 3):
            ws.cell(i, 1, n_).font = Font(italic=True, color="595959")
        tbl.to_excel(w, sheet_name="По_горизонтам", index=False)
        ana.to_excel(w, sheet_name="Аналоги_Q1", index=False)
        wk.to_excel(w, sheet_name="Недели", index=False)
        det.drop(columns="dev").to_excel(w, sheet_name="Детали", index=False)
        for wsn in w.book.sheetnames:
            ws = w.book[wsn]
            for c in ws[1]:
                c.fill, c.font = HDR_FILL, HDR_FONT
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 22
            ws.freeze_panes = "A2"
    pd.set_option("display.width", 200)
    print("\nГоризонт Q1 (прогноз за 3–5 мес), среднее отклонение месяца:")
    print(q1h[["Ряд", "Вариант", "Среднее отклонение, %", "Смещение (факт/прогноз − 100), %", "Ошибка по дням, %", "Точек"]].to_string(index=False))
    h1 = agg(mon[mon["Горизонт"] == "за 1 мес"], ["Ряд", "Вариант"]).sort_values(["Ряд", "Среднее отклонение, %"])
    print("\nГоризонт 1 мес, среднее отклонение месяца / ошибка по дням:")
    print(h1[["Ряд", "Вариант", "Среднее отклонение, %", "Ошибка по дням, %", "Точек"]].to_string(index=False))
    print("\nНедели (последние 26), среднее отклонение недели / ошибка по дням:")
    print(wk.sort_values(["Ряд", "Среднее отклонение, %"])[["Ряд", "Вариант", "Среднее отклонение, %", "Ошибка по дням, %", "Точек"]].to_string(index=False))
    if len(ana):
        print("\nАналоги Q1 (прогноз от 31.10), факт/прогноз, %, вариант v2:")
        print(ana[ana["Вариант"] == "v2 (текущая)"][["Ряд", "Период", "Горизонт", "Факт/прогноз, %"]].to_string(index=False))
    return out


def run_backtest_aht(con):
    """Бэктест AHT (№3): варианты aht.VARIANTS против простой медианы последних N недель, горизонт 1–13 недель."""
    import aht as A
    holidays = pd.DatetimeIndex(pd.to_datetime(getattr(S, "HOLIDAYS", [])))
    d = A.load(con)
    span = d.groupby("channel")["interval_start"].agg(["min", "max"])
    print("История AHT: " + "; ".join(f"{c} {r['min']:%d.%m.%Y} – {r['max']:%d.%m.%Y}" for c, r in span.iterrows()))
    print("Считаю (точки раз в неделю, горизонт до 13 недель)...")
    bt = A.backtest(d, holidays)
    tbl = A.bt_summary(bt)
    # контроль: AHT по месяцам из v_aht_interval (сверка с базовой линией 21.10)
    x = d.assign(month=d["interval_start"].dt.to_period("M").astype(str)).groupby(["channel", "month"])[
        ["work_sec", "wrap_sec", "handled", "wrap_found"]].sum()
    x["AHT, с"] = ((x["work_sec"] + x["wrap_sec"]) / x["handled"]).round(0)
    x["разговор, с"] = (x["work_sec"] / x["handled"]).round(0)
    x["поствызывная, с"] = (x["wrap_sec"] / x["handled"]).round(0)
    x["доля звонков с найденной поствызывной"] = (x["wrap_found"] / x["handled"]).round(3)
    x = x.reset_index()
    out = report_path(f"Бэктест AHT {stamp()}.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        tbl.sort_values(["Канал", "Горизонт", "Отклонение AHT месяца, %"]).to_excel(w, sheet_name="Итог", index=False)
        x.to_excel(w, sheet_name="AHT_по_месяцам", index=False)
        for wsn in w.book.sheetnames:
            ws = w.book[wsn]
            for c in ws[1]:
                c.fill, c.font = HDR_FILL, HDR_FONT
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 24
    pd.set_option("display.width", 220)
    print("\nAHT по месяцам (контроль; базовая линия 21.10: голос 214–243 с, чат 8,5–9,2 мин):")
    print(x[["channel", "month", "handled", "AHT, с", "разговор, с", "поствызывная, с", "доля звонков с найденной поствызывной"]].to_string(index=False))
    print("\nБэктест AHT:")
    print(tbl.sort_values(["Канал", "Горизонт", "Отклонение AHT месяца, %"]).to_string(index=False))
    return out


def main():
    ap = argparse.ArgumentParser(description="Прогноз нагрузки")
    ap.add_argument("--from", dest="start", default=None)
    ap.add_argument("--to", dest="end", default=None)
    ap.add_argument("--backtest", action="store_true")
    ap.add_argument("--backtest-aht", action="store_true", help="бэктест прогноза AHT")
    ap.add_argument("--publish-run", default=None, help="опубликовать готовую версию (run_id)")
    ap.add_argument("--publish", action="store_true", help="сразу пометить версию опубликованной")
    ap.add_argument("--weekly", action="store_true", help="недельный пересчёт на 2 недели вперёд")
    ap.add_argument("--year", action="store_true", help="план на год: v2 + месячный индекс, если бэктест 6–12 мес его принял (№138)")
    a = ap.parse_args()
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    con = wfm_store.connect()
    if a.publish_run:
        x = publish_run(a.publish_run)
        print(f"Версия {x['run_id']} опубликована {x['published_at']:%d.%m.%Y %H:%M}.")
        return
    if a.backtest_aht:
        out = run_backtest_aht(con)
        print(f"\nГотово: {out}")
        return
    if a.backtest:
        out = run_backtest(con)
        print(f"\nГотово: {out}")
        return
    mode = "год" if a.year else ("неделя" if a.weekly else "квартал")
    if a.year:
        d0, d1 = year_period()
    else:
        d0, d1 = default_period()
        if a.weekly:
            d1 = d0 + pd.Timedelta(days=13)
    start = pd.Timestamp(a.start) if a.start else d0
    end = pd.Timestamp(a.end) if a.end else d1
    if not (a.start or a.end or a.weekly or a.year) and sys.stdin.isatty():
        ans = input(f"Период прогноза {start:%d.%m.%Y} – {end:%d.%m.%Y}. Enter — согласиться, или введите дату конца (ДД.ММ.ГГГГ): ").strip()
        if ans:
            end = pd.Timestamp(dt.datetime.strptime(ans, "%d.%m.%Y"))
    publish = a.publish
    comment = ""
    # Год не спрашивает «опубликовать»: №138, публикуется отдельно вручную.
    if not publish and sys.stdin.isatty() and not a.weekly and not a.year:
        publish = input("Опубликовать эту версию (по ней строятся графики)? (д/н): ").strip().lower() in ("д", "да", "y", "l")
        if publish:
            comment = input("Комментарий к версии (Enter — пропустить): ").strip()
    print()
    run_id, out = run_forecast(con, start, end, publish, comment, mode)
    con.close()
    names = {"квартал": "Квартал", "неделя": "Неделя", "год": "Год"}
    state = "Опубликована." if publish else "Не опубликована."
    print(f"{names[mode]} {start:%d.%m.%Y}–{end:%d.%m.%Y}. {state}")
    print(f"Версия {run_id}.")
    print(f"\nФайл: {out}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        if not is_db_timeout(e):
            import traceback
            traceback.print_exc()
        print("\n" + error_line(e))
        sys.exit(1)
