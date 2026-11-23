-- Проверка входов за 30 дней: откуда начинались сессии, дошедшие до рабочих очередей (голос и чат; почта — без входа).
-- Первый проект сессии сверяется со списками wfm_settings:
--   'вход голоса'           — VOICE_ENTRY_PROJECTS (основной IVR, IVR второй линии): учитывается как «Голос вход»
--   'голосовой робот'       — VOICE_ROBOT_PROJECTS
--   'чат-бот'               — BOT_PROJECTS
--   'вход / служебный чата' — CHAT_SERVICE_PROJECTS (вход: сайт, приложение, мессенджеры)
--   'сразу в рабочую очередь'   — сессия началась прямо в рабочей очереди (партнёрская линия, идентификация и т.п.): «Голос прямой»
--   'исключена (другой отдел)'    — EXCLUDED_QUEUES
--   'НЕИЗВЕСТНЫЙ ВХОД'      — ни в одном списке: новый канал входа или перевод из чужой очереди — проверить
-- Чат-вход в модели = все чат-сессии без почты, поэтому новый чат-вход не теряется, но его надо внести в список.
-- Голос: звонки с неизвестного входа попадают в «Голос прямой» — сумма верна, разложение «вход × доля» — нет.
create or replace view v_entry_check as
with vq as (
    select q.session_id, q.project_id,
           row_number() over (partition by q.session_id order by q.enqueued_time, q.id) as rn
    from voice_queued_calls q
    where q.enqueued_time >= current_date - interval '30 days'
),
vs as (
    select vq.session_id,
           max(case when vq.rn = 1 then vq.project_id end) as first_project,
           max(case when g.project_id is not null then 1 else 0 end) as reached
    from vq
    left join support_queues g on g.project_id = vq.project_id and g.channel = 'Голос'
    group by vq.session_id
),
ci as (
    select i.session_id, i.project_id,
           row_number() over (partition by i.session_id order by i.start_time, i.number_in_session, i.id) as rn
    from chat_session_iterations i
    left join dim_projects p on p.uuid = i.project_id
    where i.start_time >= current_date - interval '30 days'
      and coalesce(p.datachannel, 'im') <> 'email'
),
cs as (
    select ci.session_id,
           max(case when ci.rn = 1 then ci.project_id end) as first_project,
           max(case when g.project_id is not null then 1 else 0 end) as reached
    from ci
    left join support_queues g on g.project_id = ci.project_id and g.channel = 'Чат'
    group by ci.session_id
),
x as (
    select 'Голос' as channel, first_project, count(*) as sessions, sum(reached) as reached_queue from vs group by first_project
    union all
    select 'Чат', first_project, count(*), sum(reached) from cs group by first_project
)
select x.channel,
       x.first_project as project_id,
       coalesce(p.title, '') as title,
       x.sessions,
       x.reached_queue,
       round(1.0 * x.reached_queue / sum(x.reached_queue) over (partition by x.channel), 4) as share_of_queue,
       case when e.project_id is not null then 'вход голоса'
            when r.project_id is not null then 'голосовой робот'
            when b.project_id is not null then 'чат-бот'
            when sp.project_id is not null then 'вход / служебный чата'
            when g.project_id is not null then 'сразу в рабочую очередь'
            when ng.project_id is not null then 'исключена (другой отдел)'
            else 'НЕИЗВЕСТНЫЙ ВХОД' end as status
from x
left join dim_projects p on p.uuid = x.first_project
left join voice_entry_projects e on e.project_id = x.first_project
left join voice_robot_projects r on r.project_id = x.first_project
left join bot_projects b on b.project_id = x.first_project
left join service_projects sp on sp.project_id = x.first_project
left join support_queues g on g.project_id = x.first_project
left join excluded_queues ng on ng.project_id = x.first_project
where x.reached_queue > 0;
