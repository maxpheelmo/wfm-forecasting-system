-- Чаты и почта: одна строка = одна итерация (попадание в рабочую очередь).
-- is_answered = был первый ответ оператора; is_closed_no_answer = закрыта без ответа оператора.
-- is_first_entry = 1 у первого попадания сессии в рабочей очереди этого канала.
create or replace view v_text_contacts as
select i.id as iteration_id,
       i.session_id,
       i.project_id,
       g.queue_name,
       g.channel,
       coalesce(i.unblock_time, i.start_time) as wait_start,
       i.start_time,
       i.first_oper_answer_time as answered_time,
       i.finish_time,
       i.oper_id as operator_login,
       i.result,
       i.script_name,
       extract(epoch from i.first_oper_answer_time)
         - extract(epoch from coalesce(i.unblock_time, i.start_time)) as wait_answer_sec,
       case when i.first_oper_answer_time is not null then 1 else 0 end as is_answered,
       case when i.first_oper_answer_time is null and i.finish_time is not null then 1 else 0 end as is_closed_no_answer,
       case when row_number() over (partition by g.channel, i.session_id
                                    order by i.start_time, i.number_in_session, i.id) = 1
            then 1 else 0 end as is_first_entry
from chat_session_iterations i
join support_queues g on g.project_id = i.project_id and g.channel in ('Чат', 'Почта');
