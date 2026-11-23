# -*- coding: utf-8 -*-
"""
Проверка хранилища через представления (без подключения к источник): то, на чём будет строиться прогноз.

Результат: reports\\checks\\store_check_ГГГГММДД_ЧЧММ.xlsx (только агрегаты, без логинов и данных клиентов):
  Таблицы          строк и диапазон дат по каждой таблице хранилища
  Каналы_по_дням   голос / чат / почта очереди по дням: попаданий, уникальных, принято, потеряно, SL
  SL_неделя_42     12.10–18.10: сверка с недельным отчётом (голос 30,9%, чат 49,0%)
  Чат_бот_по_дням  чат-сессии: всего, через бота, дошли до очереди, без оператора, доля без оператора
  Профиль_30мин    среднее «поступило уникальных» по дню недели × 30 мин, последние 8 недель
  Очереди_проверка где отвечает отдел поддержки; «НЕТ В СПИСКЕ» — кандидаты в wfm_settings.SUPPORT_QUEUES
  Входы_проверка   откуда начинались сессии, дошедшие до рабочих очередей; «НЕИЗВЕСТНЫЙ ВХОД» — проверить

Запуск: tools\\check_store.bat
"""
import datetime as dt
import os
import sys

import pandas as pd

import wfm_store
from wfm_common import error_line, report_path, stamp

SQL = {
    "Таблицы": """
        select 'agent_status_changes' as tbl, count(*) as rows, min(entered) as min_dt, max(leaved) as max_dt from agent_status_changes
        union all select 'voice_queued_calls', count(*), min(enqueued_time), max(dequeued_time) from voice_queued_calls
        union all select 'voice_call_legs', count(*), min(ended), max(ended) from voice_call_legs
        union all select 'chat_session_iterations', count(*), min(start_time), max(finish_time) from chat_session_iterations
        union all select 'dim_projects', count(*), null, null from dim_projects
        union all select 'dim_employees', count(*), null, null from dim_employees""",
    "Каналы_по_дням": """
        select day, channel,
               sum(entries) as entries, sum(offered) as offered, sum(handled) as handled,
               sum(abandoned) as abandoned, sum(short_abandoned) as short_abandoned,
               sum(handled_in_sl) as handled_in_sl,
               case when channel = 'Голос'
                    then 1.0 * sum(handled_in_sl) / nullif(sum(handled) + sum(abandoned) - sum(short_abandoned), 0)
                    else 1.0 * sum(handled_in_sl) / nullif(sum(entries), 0) end as sl,
               case when channel = 'Голос' then sum(talk_sec) / nullif(sum(handled), 0) end as avg_talk_sec
        from v_interval_volume
        where day >= current_date - interval '120 days'
        group by day, channel
        order by day, channel""",
    "SL_неделя_42": """
        select channel,
               sum(entries) as entries, sum(offered) as offered, sum(handled) as handled,
               sum(abandoned) as abandoned, sum(short_abandoned) as short_abandoned, sum(handled_in_sl) as handled_in_sl,
               case when channel = 'Голос'
                    then 1.0 * sum(handled_in_sl) / nullif(sum(handled) + sum(abandoned) - sum(short_abandoned), 0)
                    else 1.0 * sum(handled_in_sl) / nullif(sum(entries), 0) end as sl,
               case channel when 'Голос' then 0.309 when 'Чат' then 0.490 end as sl_weekly_report
        from v_interval_volume
        where interval_start >= timestamp '2026-10-12 00:00:00' and interval_start < timestamp '2026-10-19 00:00:00'
        group by channel
        order by channel""",
    "Чат_бот_по_дням": """
        select day, sum(sessions) as sessions, sum(sessions_via_bot) as via_bot,
               sum(reached_queue) as reached_queue, sum(answered_queue) as answered_queue,
               sum(no_operator) as no_operator, sum(bot_no_operator) as bot_no_operator,
               1.0 * sum(no_operator) / nullif(sum(sessions), 0) as share_no_operator,
               1.0 * sum(reached_queue) / nullif(sum(sessions), 0) as share_reached_queue
        from v_chat_sessions_interval
        where day >= current_date - interval '120 days'
        group by day
        order by day""",
    "Профиль_30мин": """
        with x as (
            select channel, day, interval_start, offered
            from v_interval_volume
            where day >= current_date - interval '56 days' and day < current_date)
        select channel,
               extract(isodow from interval_start) as weekday,
               extract(hour from interval_start) * 100 + extract(minute from interval_start) as hhmm,
               round(sum(offered) / 8.0, 2) as avg_offered
        from x
        group by 1, 2, 3
        order by 1, 2, 3""",
    "Очереди_проверка": """
        select * from v_queue_check order by channel, answered_by_support desc""",
    "Входы_проверка": """
        select * from v_entry_check order by channel, reached_queue desc""",
}


def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    started = dt.datetime.now()
    con = wfm_store.connect()
    out_tables = {}
    for name, sql in SQL.items():
        try:
            out_tables[name] = wfm_store.query(sql, con)
            print(f"  {name}: {len(out_tables[name])} строк")
        except Exception as e:
            out_tables[name] = pd.DataFrame({"ОШИБКА": [str(e)[:500]]})
            print(f"  {name}: ОШИБКА {e}")
    con.close()

    w42 = out_tables.get("SL_неделя_42")
    if w42 is not None and "sl" in w42:
        for r in w42.itertuples():
            ref = f" (недельный отчёт: {r.sl_weekly_report:.1%})" if pd.notna(r.sl_weekly_report) else ""
            print(f"  SL 12.10–18.10, {r.channel}: {r.sl:.1%}{ref}" if pd.notna(r.sl) else f"  SL {r.channel}: нет данных")
    miss = out_tables.get("Очереди_проверка")
    if miss is not None and "status" in miss:
        m = miss[miss["status"] == "НЕТ В СПИСКЕ"]
        print(f"  рабочих очередей не из списка: {len(m)}")
        for r in miss[miss["status"] == "исключена (другой отдел)"].itertuples():
            print(f"  исключена из очереди: {r.title or r.project_id} ({r.channel})")
    ent = out_tables.get("Входы_проверка")
    if ent is not None and "status" in ent:
        for r in ent.itertuples():
            print(f"  вход {r.channel}: {r.title or r.project_id} — {r.status}, {r.share_of_queue:.1%} сессий рабочих очередей")

    out = report_path(f"Проверка хранилища {stamp(started)}.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        for name, df in out_tables.items():
            df.to_excel(w, sheet_name=name[:31], index=False)
            ws = w.book[name[:31]]
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(max(10, width + 2), 60)
            for c in ws[1]:
                if c.value in ("sl", "sl_weekly_report", "share_no_operator", "share_reached_queue", "share_support", "share_of_queue"):
                    for cell in ws[c.column_letter][1:]:
                        cell.number_format = "0.0%"
            ws.freeze_panes = "A2"
    print(f"\nГотово: {os.path.abspath(out)}  ({(dt.datetime.now() - started).seconds} с)")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n" + error_line(e))
        sys.exit(1)
