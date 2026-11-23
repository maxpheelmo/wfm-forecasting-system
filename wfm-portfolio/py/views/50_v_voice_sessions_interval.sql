-- Голосовые сессии по 30-минутным интервалам начала сессии — для цепочки «вход × доля до очереди».
--   entry_sessions     сессии, начавшиеся во входной точке (voice_entry_projects: основной IVR, IVR второй линии)
--   entry_reached_queue  из них дошли до рабочей очереди
--   entry_via_robot    из них прошли через голосового робота (voice_robot_projects)
--   robot_reached_queue  прошли через робота и дошли до рабочей очереди
--   direct_queue         дошли до рабочей очереди, начавшись НЕ во входной точке (партнёрская линия, идентификация, переводы)
-- Голос в рабочих очередях (сессии) = entry_reached_queue + direct_queue.
create or replace view v_voice_sessions_interval as
with q as (
    select qc.session_id, qc.project_id, qc.enqueued_time,
           row_number() over (partition by qc.session_id order by qc.enqueued_time, qc.id) as rn
    from voice_queued_calls qc
),
s as (
    select q.session_id,
           min(q.enqueued_time) as s_start,
           max(case when q.rn = 1 and e.project_id is not null then 1 else 0 end) as from_entry,
           max(case when g.project_id is not null then 1 else 0 end) as reached_queue,
           max(case when r.project_id is not null then 1 else 0 end) as via_robot
    from q
    left join voice_entry_projects e on e.project_id = q.project_id
    left join support_queues g on g.project_id = q.project_id and g.channel = 'Голос'
    left join voice_robot_projects r on r.project_id = q.project_id
    group by q.session_id
)
select cast(s_start as date) as day,
       date_trunc('hour', s_start)
         + case when extract(minute from s_start) >= 30 then interval '30 minutes' else interval '0 minutes' end
         as interval_start,
       sum(from_entry) as entry_sessions,
       sum(case when from_entry = 1 and reached_queue = 1 then 1 else 0 end) as entry_reached_queue,
       sum(case when from_entry = 1 and via_robot = 1 then 1 else 0 end) as entry_via_robot,
       sum(case when via_robot = 1 and reached_queue = 1 then 1 else 0 end) as robot_reached_queue,
       sum(case when from_entry = 0 and reached_queue = 1 then 1 else 0 end) as direct_queue
from s
where s_start is not null
group by 1, 2;
