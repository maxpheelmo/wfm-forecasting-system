-- Голос: одна строка = одно попадание звонка в рабочую очередь.
-- wait_start = начало ожидания (unblocked_time, иначе enqueued_time).
-- wait_sec   = до ответа оператора (принятые) или до выхода из очереди (остальные).
-- is_first_entry = 1 у первого попадания сессии в рабочей очереди (для «поступило уникальных»).
create or replace view v_voice_contacts as
select q.id as queue_row_id,
       q.session_id,
       q.project_id,
       g.queue_name,
       coalesce(q.unblocked_time, q.enqueued_time) as wait_start,
       q.enqueued_time,
       q.dequeued_time,
       q.final_stage,
       case when q.final_stage = 'operator' then l.dst_abonent end as operator_login,
       case when q.final_stage = 'operator' then l.connected end as answered_time,
       case when q.final_stage = 'operator' then l.ended end as talk_end_time,
       extract(epoch from coalesce(case when q.final_stage = 'operator' then l.connected end, q.dequeued_time))
         - extract(epoch from coalesce(q.unblocked_time, q.enqueued_time)) as wait_sec,
       case when q.final_stage = 'operator'
            then extract(epoch from l.ended) - extract(epoch from l.connected) end as talk_sec,
       case when row_number() over (partition by q.session_id order by q.enqueued_time, q.id) = 1
            then 1 else 0 end as is_first_entry
from voice_queued_calls q
join support_queues g on g.project_id = q.project_id and g.channel = 'Голос'
left join voice_call_legs l on l.session_id = q.session_id and l.leg_id = q.next_leg_id;
