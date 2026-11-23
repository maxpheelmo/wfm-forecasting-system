# -*- coding: utf-8 -*-
"""
Точность наших прогнозов против факта (пункт «Точность» в wfm.bat): все версии из DB\\forecast_*.

Метрики по каждой версии и ряду:
  * месяц — факт / прогноз, % (как в трекере);
  * неделя (пн–вс) — факт / прогноз, %;
  * день — ошибка по дням (Σ|факт − прогноз| / Σ факт);
  * 30 мин (рабочей очереди) — ошибка по интервалам.
Сравниваются только дни после даты составления версии, где уже есть полный факт.
Ряды: «Голос очереди» / «Чат очереди» = поступило в рабочей очереди; «Голос вход» / «Чат вход» = вход КЦ.
"""
import datetime as dt
import os
import sys

import pandas as pd
from openpyxl.styles import Font, PatternFill

import wfm_store
import wfm_forecast as WF
from wfm_common import error_line, is_db_timeout, report_path, stamp, store_root

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
SQL_INT = "select interval_start, channel, offered as v from v_interval_volume where channel in ('Голос', 'Чат')"


def facts(con):
    h = WF.load_history(con, ext=False)
    f = pd.DataFrame({"Голос очереди": h["fact_voice_queue"], "Чат очереди": h["fact_chat_queue"],
                      "Голос вход": h["voice_entry"], "Чат вход": h["chat_entry"]})
    last = h.dropna(subset=["voice_entry", "chat_entry"]).index.max()
    return f[f.index <= last]


def our_forecasts(con):
    try:
        runs = wfm_store.query("select run_id, created_at, kind, status, comment from forecast_run", con)
        di = wfm_store.query("select run_id, cast(interval_start as date) as day, channel, sum(value) as v "
                             "from forecast_interval group by 1, 2, 3", con)
    except Exception:
        return pd.DataFrame(), pd.DataFrame()
    di["day"] = pd.to_datetime(di["day"])
    runs["Прогноз"] = runs.apply(lambda r: f"наш {r['run_id']} ({r['kind']}, {r['status']})", axis=1)
    d = di.merge(runs[["run_id", "Прогноз", "created_at"]], on="run_id")
    d = d[d["channel"].isin(["Голос очереди", "Чат очереди", "Голос вход", "Чат вход"])]
    return d.rename(columns={"channel": "Ряд"})[["Прогноз", "created_at", "day", "Ряд", "v"]], runs


def score(fc, fact):
    """fc: Прогноз, day, Ряд, v  -> детали по дням с фактом."""
    fl = fact.stack().rename("fact").reset_index().rename(columns={"level_0": "day", "level_1": "Ряд"})
    fl.columns = ["day", "Ряд", "fact"]
    d = fc.merge(fl, on=["day", "Ряд"], how="inner").dropna(subset=["fact", "v"])
    d = d[d["fact"] > 0]
    # прогноз нашей версии, сделанный внутри периода, оцениваем только по дням после даты составления
    d = d[d["created_at"].isna() | (d["day"] > pd.to_datetime(d["created_at"]).dt.normalize())]
    return d


def agg(d, key):
    g = d.groupby(["Прогноз", "Ряд", key])
    t = g.agg(Дней=("fact", "size"), Факт=("fact", "sum"), Прогноз_сумма=("v", "sum"),
              Ош=("fact", lambda x: 0)).reset_index()
    err = g.apply(lambda x: (x["fact"] - x["v"]).abs().sum() / x["fact"].sum() * 100 if x["fact"].sum() > 0 else float("nan"),
                  include_groups=False).values
    t["Факт/прогноз, %"] = (t["Факт"] / t["Прогноз_сумма"].where(t["Прогноз_сумма"] > 0) * 100).round(1)
    t["Ошибка по дням, %"] = pd.Series(err).round(1).values
    t["Факт"] = t["Факт"].round(0)
    t["Прогноз_сумма"] = t["Прогноз_сумма"].round(0)
    return t.drop(columns="Ош").rename(columns={"Прогноз_сумма": "Прогноз, сумма"})


def interval_score(con, runs):
    try:
        it = wfm_store.query(SQL_INT, con)
        fi = wfm_store.query("select run_id, interval_start, channel, value from forecast_interval "
                             "where channel in ('Голос очереди', 'Чат очереди')", con)
    except Exception:
        return pd.DataFrame()
    it["interval_start"] = pd.to_datetime(it["interval_start"])
    it["channel"] = it["channel"] + " очереди"
    fi["interval_start"] = pd.to_datetime(fi["interval_start"])
    last = it["interval_start"].max().normalize()            # последний день неполный — не берём
    m = fi.merge(it, on=["interval_start", "channel"], how="left").fillna({"v": 0})
    m = m[m["interval_start"] < last]
    m = m.merge(runs[["run_id", "created_at"]], on="run_id")
    m = m[m["interval_start"] > pd.to_datetime(m["created_at"]).dt.normalize() + pd.Timedelta(days=1)]
    if m.empty:
        return pd.DataFrame()
    m["month"] = m["interval_start"].dt.strftime("%Y-%m")
    out = m.groupby(["run_id", "channel", "month"]).apply(
        lambda x: pd.Series({"Интервалов": len(x), "Ошибка по 30 мин, %": round((x["v"] - x["value"]).abs().sum() / x["v"].sum() * 100, 1)}),
        include_groups=False).reset_index()
    return out.rename(columns={"channel": "Ряд", "month": "Месяц"})


def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    con = wfm_store.connect()
    fact = facts(con)
    ours, runs = our_forecasts(con)
    allf = ours
    if allf.empty:
        print("Нет прогнозов для сравнения: сначала сделайте прогноз («План на квартал»), а после наступления периода — этот пункт.")
        return
    d = score(allf, fact)
    if d.empty:
        print("Пока нет дней, где есть и прогноз, и факт: сравнение появится, когда наступит период прогноза "
              "и данные за него будут загружены.")
        return
    d["Месяц"] = d["day"].dt.strftime("%Y-%m")
    wk = d["day"] - pd.to_timedelta(d["day"].dt.weekday, unit="D")
    d["Неделя"] = wk.dt.strftime("%Y-%m-%d")
    months = agg(d, "Месяц")
    weeks = agg(d, "Неделя")
    weeks = weeks[weeks["Дней"] == 7]
    summ = months.assign(dev=(months["Факт/прогноз, %"] - 100).abs()).groupby(["Прогноз", "Ряд"]).agg(
        Месяцев=("Месяц", "size"), **{"Среднее отклонение месяца, %": ("dev", "mean"),
                                       "Средняя ошибка по дням, %": ("Ошибка по дням, %", "mean")}).round(1)
    wsum = weeks.assign(dev=(weeks["Факт/прогноз, %"] - 100).abs()).groupby(["Прогноз", "Ряд"]).agg(
        Недель=("Неделя", "size"), **{"Среднее отклонение недели, %": ("dev", "mean")}).round(1)
    summ = summ.join(wsum, how="outer").reset_index()
    ints = interval_score(con, runs) if len(runs) else pd.DataFrame()
    con.close()
    out = report_path(f"Точность {stamp()}.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        summ.to_excel(w, sheet_name="Итог", index=False)
        ws = w.book["Итог"]
        notes = ["Отклонение = |факт / прогноз − 100%|. Ошибка по дням = Σ|факт − прогноз| / Σ факт.",
                 "«очереди» — поступило в рабочей очереди (на нём строятся графики); «вход» — вход КЦ.",
                 "Версии оцениваются только по дням после даты составления."]
        for i, n_ in enumerate(notes, start=len(summ) + 3):
            ws.cell(i, 1, n_).font = Font(italic=True, color="595959")
        months.to_excel(w, sheet_name="Месяцы", index=False)
        weeks.to_excel(w, sheet_name="Недели", index=False)
        if len(ints):
            ints.to_excel(w, sheet_name="30 мин", index=False)
        d.assign(day=d["day"].dt.date).drop(columns=["created_at"]).rename(
            columns={"day": "Дата", "v": "Прогноз", "fact": "Факт"}).to_excel(w, sheet_name="Дни", index=False)
        for wsn in w.book.sheetnames:
            ws = w.book[wsn]
            for c in ws[1]:
                c.fill, c.font = HDR_FILL, HDR_FONT
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col[:50])
                ws.column_dimensions[col[0].column_letter].width = min(max(10, width + 2), 60)
            ws.freeze_panes = "A2"
    print(f"Версий в сравнении: {summ['Прогноз'].nunique()}. Месяцы, недели и дни — в файле.")
    print(f"Отчёт: {out}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        if not is_db_timeout(e):
            import traceback
            traceback.print_exc()
        print("\n" + error_line(e))
        sys.exit(1)
