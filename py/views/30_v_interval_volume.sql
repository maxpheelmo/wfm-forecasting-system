-- Нагрузка очереди по каналу и 30-минутным интервалам (основа прогноза).
--   entries         попаданий в рабочей очереди (с переводами)
--   offered         уникальных обращений (первое попадание сессии в рабочей очереди канала)
--   handled         голос: принято оператором; чат/почта: был ответ оператора
--   abandoned       голос: потеряно в очереди; чат/почта: закрыто без ответа оператора
--   short_abandoned голос: потеряно быстрее порога «короткий потерянный»
--   handled_in_sl   принято / отвечено в пределах порога SL канала
--   talk_sec        голос: сумма разговора принятых, сек
-- SL голоса = handled_in_sl / (handled + abandoned - short_abandoned); чата = handled_in_sl / entries.
create or replace view v_interval_volume as
with p as (
    select max(case when name = 'sl_voice_sec' then value end) as sl_voice,
           max(case when name = 'sl_chat_sec' then value end) as sl_chat,
           max(case when name = 'sl_mail_sec' then value end) as sl_mail,
           max(case when name = 'short_abandon_sec' then value end) as short_ab
    from wfm_params
),
v as (
    select 'Голос' as channel,
           c.wait_start as t,
           c.is_first_entry,
           case when c.final_stage = 'operator' then 1 else 0 end as handled,
           case when c.final_stage = 'queue' then 1 else 0 end as abandoned,
           case when c.final_stage = 'queue' and c.wait_sec < p.short_ab then 1 else 0 end as short_abandoned,
           case when c.final_stage = 'operator' and c.wait_sec <= p.sl_voice then 1 else 0 end as handled_in_sl,
           coalesce(c.talk_sec, 0) as talk_sec
    from v_voice_contacts c cross join p
    union all
    select c.channel,
           c.wait_start,
           c.is_first_entry,
           c.is_answered,
           c.is_closed_no_answer,
           0,
           case when c.is_answered = 1
                 and c.wait_answer_sec <= case when c.channel = 'Чат' then p.sl_chat else p.sl_mail end
                then 1 else 0 end,
           0
    from v_text_contacts c cross join p
)
select channel,
       cast(t as date) as day,
       date_trunc('hour', t)
         + case when extract(minute from t) >= 30 then interval '30 minutes' else interval '0 minutes' end
         as interval_start,
       count(*) as entries,
       sum(is_first_entry) as offered,
       sum(handled) as handled,
       sum(abandoned) as abandoned,
       sum(short_abandoned) as short_abandoned,
       sum(handled_in_sl) as handled_in_sl,
       sum(talk_sec) as talk_sec
from v
where t is not null
group by 1, 2, 3;
