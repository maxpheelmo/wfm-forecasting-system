-- Чат-сессии (все входы, не почта) по 30-минутным интервалам начала сессии — для доли бота.
--   sessions          всего чат-сессий
--   sessions_via_bot  прошли через проект бота (bot_projects)
--   reached_queue       дошли до рабочей очереди (support_queues, канал «Чат»)
--   answered_queue      получили ответ оператора в рабочей очереди
--   no_operator       ни в одной итерации не было оператора
--   bot_no_operator   через бота и без оператора (бот закрыл ИЛИ клиент ушёл — не различаем)
create or replace view v_chat_sessions_interval as
with s as (
    select i.session_id,
           min(i.start_time) as session_start,
           max(case when b.project_id is not null then 1 else 0 end) as via_bot,
           max(case when g.project_id is not null then 1 else 0 end) as reached_queue,
           max(case when i.oper_id is not null then 1 else 0 end) as reached_operator,
           max(case when g.project_id is not null and i.first_oper_answer_time is not null then 1 else 0 end) as answered_queue
    from chat_session_iterations i
    left join dim_projects p on p.uuid = i.project_id
    left join bot_projects b on b.project_id = i.project_id
    left join support_queues g on g.project_id = i.project_id and g.channel = 'Чат'
    where coalesce(p.datachannel, 'im') <> 'email'
    group by i.session_id
)
select cast(session_start as date) as day,
       date_trunc('hour', session_start)
         + case when extract(minute from session_start) >= 30 then interval '30 minutes' else interval '0 minutes' end
         as interval_start,
       count(*) as sessions,
       sum(via_bot) as sessions_via_bot,
       sum(reached_queue) as reached_queue,
       sum(answered_queue) as answered_queue,
       sum(case when reached_operator = 0 then 1 else 0 end) as no_operator,
       sum(case when via_bot = 1 and reached_operator = 0 then 1 else 0 end) as bot_no_operator
from s
where session_start is not null
group by 1, 2;
