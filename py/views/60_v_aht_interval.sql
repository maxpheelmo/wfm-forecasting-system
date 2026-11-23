-- Время обработки очереди по 30-минутным интервалам постановки в очередь (журнал решений №3). Основа прогноза AHT.
-- Голос: разговор (плечо оператора: connected → ended) + поствызывная обработка.
--   Поствызывная = статус оператора wrapup* (agent_status_changes), начавшийся в окне [конец разговора − 5 с;
--   конец разговора + 60 с] у того же оператора; если подходят несколько звонков — к ближайшему по времени.
--   Статусы пишутся по времени рабочей станции оператора — отсюда допуск −5 с.
--   Слои статусов (wrapup и wrapup#канал) пишутся строками с одинаковым началом — берём одну, самую длинную.
--   Источник хранит статусы 180 дней: wrap_covered = 1 только там, где статусы у нас есть (по дню).
-- Чат: от распределения на оператора (oper_distribute_time) до закрытия / закрытые с ответом.
--   Ожидание в очереди не входит. Так же дашборд/6 (№72). Отсечения по длительности нет.
--   Параллельность чатов здесь не учитывается (решение №4 — отдельный замер).
-- Выбросы голоса (разговор дольше 4 ч, поствызывная дольше 30 мин) отбрасываются.
create or replace view v_aht_interval as
with calls as (
    select queue_row_id, operator_login, wait_start, talk_end_time,
           talk_sec
    from v_voice_contacts
    where final_stage = 'operator' and operator_login is not null and talk_end_time is not null
      and talk_sec >= 0 and talk_sec <= 4 * 3600
),
wr_raw as (
    select login, entered,
           extract(epoch from leaved) - extract(epoch from entered) as sec
    from agent_status_changes
    where lower(status) like 'wrapup%' and leaved is not null
),
wr as (
    select login, min(entered) as entered, max(sec) as sec
    from wr_raw
    where sec >= 0 and sec <= 1800
    group by login, date_trunc('second', entered)
),
cand as (
    select w.login, w.entered, w.sec, c.queue_row_id,
           row_number() over (partition by w.login, w.entered
                              order by abs(extract(epoch from w.entered) - extract(epoch from c.talk_end_time))) as rn
    from wr w
    join calls c on c.operator_login = w.login
                and w.entered >= c.talk_end_time - interval '5 seconds'
                and w.entered <= c.talk_end_time + interval '60 seconds'
),
wrap_per_call as (
    select queue_row_id, sum(sec) as wrap_sec
    from cand where rn = 1
    group by queue_row_id
),
st_days as (
    select distinct cast(entered as date) as day from agent_status_changes
),
v as (
    select 'Голос' as channel, c.wait_start as t, c.talk_sec as work_sec,
           coalesce(w.wrap_sec, 0) as wrap_sec,
           case when w.wrap_sec is not null then 1 else 0 end as wrap_found,
           case when d.day is not null then 1 else 0 end as wrap_covered
    from calls c
    left join wrap_per_call w on w.queue_row_id = c.queue_row_id
    left join st_days d on d.day = cast(c.talk_end_time as date)
    union all
    select 'Чат', t.wait_start,
           extract(epoch from t.finish_time) - extract(epoch from i.oper_distribute_time),
           0, 0, 1
    from v_text_contacts t
    join chat_session_iterations i on i.id = t.iteration_id
    where t.channel = 'Чат' and t.answered_time is not null
      and i.oper_distribute_time is not null and t.finish_time is not null
      and t.finish_time >= i.oper_distribute_time
)
select channel,
       cast(t as date) as day,
       date_trunc('hour', t)
         + case when extract(minute from t) >= 30 then interval '30 minutes' else interval '0 minutes' end
         as interval_start,
       count(*) as handled,
       sum(work_sec) as work_sec,
       sum(wrap_sec) as wrap_sec,
       sum(wrap_found) as wrap_found,
       min(wrap_covered) as wrap_covered
from v
where t is not null
group by 1, 2, 3;
