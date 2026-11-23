-- Проверка состава рабочих очередей за 30 дней: где отвечают операторы отдела поддержки.
-- status = 'НЕТ В СПИСКЕ' — отдел поддержки отвечает в очереди, которой нет в wfm_settings.SUPPORT_QUEUES.
-- status = 'в списке, отдел не отвечает' — очередь в списке, но ответов отдела поддержки за 30 дней нет.
-- status = 'исключена (другой отдел)' — очередь другой группы (excluded_queues, напр. партнёрская линия): не предлагаем.
-- status = 'служебный' — бот, вход, оценка (bot_projects, service_projects): не очередь, не предлагаем.
create or replace view v_queue_check as
with voice as (
    select q.project_id,
           'Голос' as channel,
           count(*) as answered,
           sum(case when d.department is not null then 1 else 0 end) as answered_by_support
    from voice_queued_calls q
    join voice_call_legs l on l.session_id = q.session_id and l.leg_id = q.next_leg_id
    left join dim_employees e on e.login = l.dst_abonent
    left join support_departments d on d.department = e.department
    where q.final_stage = 'operator'
      and q.enqueued_time >= current_date - interval '30 days'
    group by q.project_id
),
txt as (
    select i.project_id,
           case when p.datachannel = 'email' then 'Почта' else 'Чат' end as channel,
           count(*) as answered,
           sum(case when d.department is not null then 1 else 0 end) as answered_by_support
    from chat_session_iterations i
    left join dim_projects p on p.uuid = i.project_id
    left join dim_employees e on e.login = i.oper_id
    left join support_departments d on d.department = e.department
    where i.first_oper_answer_time is not null
      and i.start_time >= current_date - interval '30 days'
    group by i.project_id, case when p.datachannel = 'email' then 'Почта' else 'Чат' end
),
a as (
    select * from voice
    union all
    select * from txt
),
allq as (
    select coalesce(a.project_id, g.project_id) as project_id,
           coalesce(a.channel, g.channel) as channel,
           coalesce(a.answered, 0) as answered,
           coalesce(a.answered_by_support, 0) as answered_by_support,
           case when g.project_id is not null then 1 else 0 end as in_list
    from a
    full outer join support_queues g on g.project_id = a.project_id
)
select x.project_id,
       coalesce(p.title, '') as title,
       x.channel,
       x.answered,
       x.answered_by_support,
       case when x.answered > 0 then round(1.0 * x.answered_by_support / x.answered, 3) end as share_support,
       case when x.in_list = 0 and ng.project_id is not null then 'исключена (другой отдел)'
            when x.in_list = 0 and (b.project_id is not null or sp.project_id is not null) then 'служебный'
            when x.in_list = 0 and x.answered_by_support > 0 then 'НЕТ В СПИСКЕ'
            when x.in_list = 1 and x.answered_by_support = 0 then 'в списке, отдел не отвечает'
            when x.in_list = 1 then 'в списке'
            else '' end as status
from allq x
left join dim_projects p on p.uuid = x.project_id
left join bot_projects b on b.project_id = x.project_id
left join service_projects sp on sp.project_id = x.project_id
left join excluded_queues ng on ng.project_id = x.project_id
where x.in_list = 1 or x.answered_by_support > 0;
