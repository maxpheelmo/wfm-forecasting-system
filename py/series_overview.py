# -*- coding: utf-8 -*-
"""
Обзор рядов нагрузки из хранилища (к источник не подключается) + черновик календаря событий.

Голос и чат — цепочкой «вход × доля до очереди» (под бота/робота):
  Голос вход        — звонки, начавшиеся в основной IVR / IVR второй линии (wfm_settings.VOICE_ENTRY_PROJECTS)
  Голос прямой      — звонки в рабочей очереди, начавшиеся не в IVR КЦ (партнёрская линия, идентификация, переводы)
  Доля голоса до очереди — из «Голос вход» дошли до рабочей очереди
  Робот, звонков    — из «Голос вход» прошли через голосового робота (справочно)
  Голос очереди         — уникальные звонки в рабочих очередях (как в отчётах)
  Чат вход          — все чат-сессии (входы КЦ, включая MAX)
  Доля чата до очереди  — чат-сессии, дошедшие до рабочей очереди
  Чат очереди           — уникальные чаты в рабочих очередях
  Почта             — уникальные письма в «Входящие Email»

Подозрительный день: объём отличается от «нормы» больше чем на VOL_THRESHOLD (норма — медиана того же
дня недели за 3 недели до и 3 недели после, сам день не входит); доли — больше чем на SHARE_THRESHOLD п.п.

Результат: reports\\checks\\series_overview_ГГГГММДД_ЧЧММ.xlsx — Сводка, По_дням (+ график), По_неделям, Дни_недели,
           Профиль_30мин, Доли_по_часам, Подозрительные_дни.
Границы смены режима (робот/бот) берутся из events.xlsx (действие «смена режима»), иначе — по умолчанию.
Запуск: tools\\series_overview.bat
"""
import datetime as dt
import os
import sys

import numpy as np
import pandas as pd
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Font, PatternFill

import wfm_store
from wfm_common import EVENTS_FILE, error_line, is_db_timeout, report_path, stamp

VOL_THRESHOLD = 0.50     # отклонение объёма от нормы дня недели (дневной шум у очереди ±15–17%)
SHARE_THRESHOLD = 0.08   # отклонение доли до очереди, 0.08 = 8 п.п.
PROFILE_WEEKS = 8        # окно для профилей
WD = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
SERIES = ["Голос вход", "Голос прямой", "Голос очереди", "Чат вход", "Чат очереди", "Почта"]
SHARES = {"Доля голоса до очереди": "Голос вход", "Доля чата до очереди": "Чат вход"}   # доля -> её вход
FLAG_SERIES = ["Голос вход", "Голос очереди", "Чат вход", "Чат очереди"]   # «Голос прямой» и почта малы и шумны — не размечаем
STARTUP_DAYS = 7         # первая неделя после начала ряда — разгон, не размечаем
# Смена режима доли до очереди (робот/бот): соседей через границу при расчёте нормы не берём.
# Если рядом лежит events.xlsx — даты берутся из строк с действием «смена режима».
DEFAULT_REGIME_BREAKS = {"Голос": ["2026-09-20"]}
CHANNEL = {"Голос вход": "Голос", "Голос прямой": "Голос", "Голос очереди": "Голос", "Доля голоса до очереди": "Голос",
           "Чат вход": "Чат", "Чат очереди": "Чат", "Доля чата до очереди": "Чат", "Почта": "Почта"}

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
FLAG_FILL = PatternFill("solid", fgColor="FCE4D6")

SQL_DAY = "select day, channel, sum(offered) as offered from v_interval_volume group by 1, 2"
SQL_CHAT = "select day, sum(sessions) as sessions, sum(reached_queue) as reached_queue from v_chat_sessions_interval group by 1"
SQL_VOICE = """
select day, sum(entry_sessions) as entry, sum(entry_reached_queue) as entry_queue,
       sum(entry_via_robot) as robot, sum(direct_queue) as direct
from v_voice_sessions_interval group by 1
"""
SQL_INTERVAL = """
select channel, interval_start, offered from v_interval_volume
where day >= current_date - {days} * interval '1 day' and day < current_date
"""
SQL_HOUR = """
select 'Чат' as channel, extract(hour from interval_start) as hour, sum(sessions) as entry, sum(reached_queue) as reached
from v_chat_sessions_interval
where day >= current_date - {days} * interval '1 day' and day < current_date group by 1, 2
union all
select 'Голос', extract(hour from interval_start), sum(entry_sessions), sum(entry_reached_queue)
from v_voice_sessions_interval
where day >= current_date - {days} * interval '1 day' and day < current_date group by 1, 2
"""


def series_start(s):
    """Первый день, с которого ряд «рабочий»: объём ≥ 30% медианы последних 60 дней."""
    tail = s.dropna().iloc[-60:]
    if tail.empty or tail.median() <= 0:
        return None
    ok = s[s >= 0.3 * tail.median()]
    return ok.index.min() if len(ok) else None


def weekday_norm(s, breaks=()):
    """Медиана того же дня недели за ±3 недели, без самого дня и без соседей по другую сторону смены режима."""
    cols = []
    for k in (-3, -2, -1, 1, 2, 3):
        sh = s.shift(7 * k)
        src = s.index - pd.Timedelta(days=7 * k)
        for b in breaks:
            b = pd.Timestamp(b)
            sh = sh.where(~(((s.index >= b) & (src < b)) | ((s.index < b) & (src >= b))))
        cols.append(sh)
    return pd.concat(cols, axis=1).median(axis=1, skipna=True)


def regime_breaks():
    """Даты смены режима по каналам: из events.xlsx (действие «смена режима»), иначе по умолчанию."""
    try:
        ev = pd.read_excel(EVENTS_FILE, sheet_name=0)
        ev = ev[ev["Действие"].astype(str).str.strip().str.lower() == "смена режима"]
        out = {}
        for _, r in ev.iterrows():
            chans = ["Голос", "Чат", "Почта"] if str(r["Канал"]).strip() == "Все" else [str(r["Канал"]).strip()]
            for ch in chans:
                out.setdefault(ch, []).append(pd.Timestamp(r["Дата с"]))
        return out, "events.xlsx"
    except Exception:
        return {k: [pd.Timestamp(x) for x in v] for k, v in DEFAULT_REGIME_BREAKS.items()}, "по умолчанию"


def day_type(d):
    return {5: "Сб", 6: "Вс"}.get(pd.Timestamp(d).weekday(), "бдн")


def style(ws, df, pct_cols=(), flag_col=None):
    for c in ws[1]:
        c.fill, c.font = HDR_FILL, HDR_FONT
    for i, col in enumerate(df.columns, 1):
        letter = ws.cell(1, i).column_letter
        ws.column_dimensions[letter].width = max(10, min(40, len(str(col)) + 3))
        if col in pct_cols:
            for cell in ws[letter][1:]:
                cell.number_format = "0.0%"
    if flag_col is not None and flag_col in df.columns:
        j = list(df.columns).index(flag_col) + 1
        for r in range(2, len(df) + 2):
            if ws.cell(r, j).value:
                for cc in range(1, len(df.columns) + 1):
                    ws.cell(r, cc).fill = FLAG_FILL
    ws.freeze_panes = "B2"


def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    started = dt.datetime.now()
    today = pd.Timestamp(dt.date.today())
    con = wfm_store.connect()
    q = lambda sql: wfm_store.query(sql, con)

    # ---------------- дневные ряды ----------------
    d = q(SQL_DAY)
    d["day"] = pd.to_datetime(d["day"])
    p = d.pivot_table(index="day", columns="channel", values="offered", aggfunc="sum")
    c = q(SQL_CHAT)
    c["day"] = pd.to_datetime(c["day"])
    c = c.set_index("day")
    v = q(SQL_VOICE)
    v["day"] = pd.to_datetime(v["day"])
    v = v.set_index("day")
    idx = pd.date_range(min(p.index.min(), c.index.min(), v.index.min()), today - pd.Timedelta(days=1), freq="D")
    days = pd.DataFrame(index=idx)
    days["Голос вход"] = v["entry"]
    days["Голос прямой"] = v["direct"]
    days["Доля голоса до очереди"] = v["entry_queue"] / v["entry"]
    days["Робот, звонков"] = v["robot"]
    days["Голос очереди"] = p.get("Голос")
    days["Чат вход"] = c["sessions"]
    days["Доля чата до очереди"] = c["reached_queue"] / c["sessions"]
    days["Чат очереди"] = p.get("Чат")
    days["Почта"] = p.get("Почта")
    days = days.astype(float)

    starts = {col: series_start(days[col]) for col in SERIES}
    for sh, base in SHARES.items():
        starts[sh] = starts[base]
    first = min(x for x in starts.values() if x is not None)
    days = days[days.index >= first]                  # до начала рабочей истории — тестовые единичные строки

    breaks, breaks_src = regime_breaks()
    flags = []
    for col in SERIES:
        st = starts[col]
        s = days[col].where(days.index >= st) if st is not None else days[col] * np.nan
        norm = weekday_norm(s, breaks.get(CHANNEL[col], []))
        days[f"{col}: норма"] = norm.round(0)
        days[f"{col}: откл."] = (s / norm - 1).round(3)
        if col in FLAG_SERIES and st is not None:
            ok = days.index >= st + pd.Timedelta(days=STARTUP_DAYS)
            for day in days.index[((days[f"{col}: откл."].abs() > VOL_THRESHOLD) & ok).fillna(False)]:
                flags.append((day, col, s[day], norm[day], s[day] / norm[day] - 1))
    for sh in SHARES:
        st = starts[sh]
        s = days[sh].where(days.index >= st) if st is not None else days[sh] * np.nan
        norm = weekday_norm(s, breaks.get(CHANNEL[sh], []))
        days[f"{sh}: норма"] = norm.round(3)
        days[f"{sh}: откл. п.п."] = (s - norm).round(3)
        if st is not None:
            ok = days.index >= st + pd.Timedelta(days=STARTUP_DAYS)
            for day in days.index[((days[f"{sh}: откл. п.п."].abs() > SHARE_THRESHOLD) & ok).fillna(False)]:
                flags.append((day, sh, s[day], norm[day], s[day] - norm[day]))

    fl = pd.DataFrame(flags, columns=["Дата", "Ряд", "Факт", "Норма", "Отклонение"]).sort_values(["Дата", "Ряд"])
    flagged_days = set(fl["Дата"])
    days.insert(0, "День недели", [WD[x.weekday()] for x in days.index])
    days["Подозрительный"] = ["да" if x in flagged_days else "" for x in days.index]

    # ---------------- недели ----------------
    wk = days[SERIES + ["Робот, звонков"]].copy()
    wk["_vg"] = days["Доля голоса до очереди"] * days["Голос вход"]
    wk["_cg"] = days["Доля чата до очереди"] * days["Чат вход"]
    wk = wk.resample("W-SUN", label="left", closed="left").sum(min_count=1)
    full = days.resample("W-SUN", label="left", closed="left").size()
    wk = wk[full == 7]
    wk.insert(2, "Доля голоса до очереди", (wk.pop("_vg") / wk["Голос вход"]).round(3))
    wk.insert(6, "Доля чата до очереди", (wk.pop("_cg") / wk["Чат вход"]).round(3))
    for col in SERIES:
        wk[f"{col}: к пред. неделе"] = (wk[col] / wk[col].shift(1) - 1).round(3)
    wk.index.name = "Неделя с"

    # ---------------- профиль дней недели (чистые дни последних N недель) ----------------
    lo = today - pd.Timedelta(weeks=PROFILE_WEEKS)
    clean = days[(days.index >= lo) & (days["Подозрительный"] == "")]
    wdp = clean.groupby("День недели")[SERIES].mean().reindex(WD)
    wdp_idx = (wdp / wdp.mean()).round(3)
    wdp_idx.columns = [f"{c_}: индекс" for c_ in wdp.columns]
    wd_tab = pd.concat([wdp.round(0), wdp_idx], axis=1)
    wd_tab.insert(0, "Чистых дней", clean.groupby("День недели").size().reindex(WD))

    # ---------------- профиль 30 мин рабочих очередей (чистые дни) ----------------
    it = q(SQL_INTERVAL.format(days=PROFILE_WEEKS * 7))
    it["interval_start"] = pd.to_datetime(it["interval_start"])
    it["day"] = it["interval_start"].dt.normalize()
    it = it[~it["day"].isin(flagged_days)]
    it["Тип дня"] = it["day"].map(day_type)
    it["Время"] = it["interval_start"].dt.strftime("%H:%M")
    prof = (it.groupby(["channel", "Тип дня", "Время"])["offered"].sum()
            / it.groupby(["channel", "Тип дня"])["offered"].sum()).unstack(["channel", "Тип дня"])
    prof = prof.reindex([f"{h:02d}:{m:02d}" for h in range(24) for m in (0, 30)]).fillna(0.0)
    prof.columns = [f"{a} {b}" for a, b in prof.columns]
    prof = prof[sorted(prof.columns)].round(4)
    prof.index.name = "Время"

    hr = q(SQL_HOUR.format(days=PROFILE_WEEKS * 7))
    hr["hour"] = hr["hour"].astype(int)
    hr["share"] = hr["reached"] / hr["entry"]
    hours = hr.pivot_table(index="hour", columns="channel", values=["entry", "share"])
    hours.columns = [("Вход " if a == "entry" else "Доля до очереди ") + b.lower() for a, b in hours.columns]
    hours = hours.round(3)
    hours.index.name = "Час"
    con.close()

    # ---------------- сводка ----------------
    last = days[days.index >= lo]
    rows = []
    for col in SERIES + list(SHARES):
        if col in SHARES:
            base = SHARES[col]
            avg = round((last[col] * last[base]).sum() / last[base].sum(), 3) if last[base].sum() else None
        else:
            avg = round(last[col].mean(), 0)
        rows.append({"Ряд": col, "История с": starts[col].date() if starts[col] is not None else None,
                     f"Среднее, {PROFILE_WEEKS} нед.": avg,
                     f"Чистых дней, {PROFILE_WEEKS} нед.": int(((last["Подозрительный"] == "") & last[col].notna()).sum()),
                     "Подозрительных дней всего": int((fl["Ряд"] == col).sum())})
    summ = pd.DataFrame(rows)

    # ---------------- xlsx ----------------
    out = report_path(f"Обзор рядов {stamp(started)}.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        summ.to_excel(w, sheet_name="Сводка", index=False)
        ws = w.book["Сводка"]
        style(ws, summ)
        notes = [f"Факт по {(today - pd.Timedelta(days=1)):%d.%m.%Y} включительно (сегодня не берётся).",
                 f"Подозрительный день: объём отличается от медианы того же дня недели за ±3 недели больше чем на "
                 f"{VOL_THRESHOLD:.0%}; доли до очереди — больше чем на {SHARE_THRESHOLD * 100:.0f} п.п.",
                 "Голос очереди ≈ Голос вход × Доля голоса до очереди + Голос прямой; Чат очереди ≈ Чат вход × Доля чата до очереди.",
                 "Смена режима робота/бота видна как ступенька в доле до очереди (лист «По_неделям»). "
                 f"Границы режима ({breaks_src}): " + "; ".join(f"{k} с {', '.join(f'{x:%d.%m.%Y}' for x in v)}" for k, v in breaks.items()),
                 f"Размечаются: {', '.join(FLAG_SERIES)} и доли до очереди; первая неделя каждого ряда (разгон) не размечается.",
                 "Профили — по чистым дням последних 8 недель. Лист «По_дням» — с графиком.",
                 "Прошлые аномалии размечаются автоматически; в events.xlsx — только будущие события (акции, запуск робота/бота)."]
        for i, n_ in enumerate(notes, start=len(summ) + 3):
            ws.cell(i, 1, n_).font = Font(italic=True, color="595959")

        out_days = days.copy()
        out_days.index = out_days.index.date
        out_days.index.name = "Дата"
        out_days.to_excel(w, sheet_name="По_дням")
        ws = w.book["По_дням"]
        style(ws, out_days.reset_index(), pct_cols=[x for x in out_days.columns if "откл." in x or "Доля" in x],
              flag_col="Подозрительный")
        n = len(out_days) + 1
        chart = LineChart()
        chart.title, chart.height, chart.width = "Нагрузка по дням", 9, 30
        for col in ("Голос вход", "Голос очереди", "Чат вход", "Чат очереди"):
            j = list(out_days.columns).index(col) + 2
            chart.add_data(Reference(ws, min_col=j, min_row=1, max_row=n), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=n))
        chart.y_axis.title = "обращений/сут"
        ws.add_chart(chart, ws.cell(2, len(out_days.columns) + 3).coordinate)

        wk_out = wk.copy()
        wk_out.index = wk_out.index.date
        wk_out.index.name = "Неделя с"
        wk_out.to_excel(w, sheet_name="По_неделям")
        style(w.book["По_неделям"], wk_out.reset_index(),
              pct_cols=[x for x in wk_out.columns if "к пред." in x or "Доля" in x])

        wd_tab.to_excel(w, sheet_name="Дни_недели")
        style(w.book["Дни_недели"], wd_tab.reset_index())

        prof.to_excel(w, sheet_name="Профиль_30мин")
        style(w.book["Профиль_30мин"], prof.reset_index(), pct_cols=list(prof.columns))

        hours.to_excel(w, sheet_name="Доли_по_часам")
        style(w.book["Доли_по_часам"], hours.reset_index(), pct_cols=[x for x in hours.columns if "Доля" in x])

        fl_out = fl.copy()
        fl_out["Дата"] = fl_out["Дата"].dt.date
        fl_out["День недели"] = [WD[pd.Timestamp(x).weekday()] for x in fl_out["Дата"]]
        fl_out = fl_out[["Дата", "День недели", "Ряд", "Факт", "Норма", "Отклонение"]]
        fl_out.to_excel(w, sheet_name="Подозрительные_дни", index=False)
        style(w.book["Подозрительные_дни"], fl_out, pct_cols=["Отклонение"])

    print(f"Готово: {os.path.abspath(out)}")
    print(f"Подозрительных дней: {len(flagged_days)} (по рядам: "
          + ", ".join(f"{k} {v_}" for k, v_ in fl.groupby('Ряд').size().items()) + ")")
    for col in SERIES:
        print(f"  {col}: история с {starts[col]:%d.%m.%Y}" if starts[col] is not None else f"  {col}: нет данных")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        if not is_db_timeout(e):
            import traceback
            traceback.print_exc()
        print("\n" + error_line(e))
        sys.exit(1)
