# -*- coding: utf-8 -*-
"""
Пример конфигурации. Живой модуль, который импортирует код, — `py/wfm_settings.py`.

Как пользоваться:
  1. Подключение к реплике задайте в `.env` (см. `.env.example`) либо правьте SOURCE_DB
     в `py/wfm_settings.py`. Пароль в файлы не кладите: используйте `py/save_password.py`.
  2. Идентификаторы очередей, входов, бота и робота в `SUPPORT_QUEUES` и соседних словарях —
     нейтральные заглушки. Замените их на id своей платформы контакт-центра.
  3. Имена сырых таблиц в загрузчике и SQL-витринах тоже обобщены
     (`voice_queued_calls`, `chat_session_iterations`, `dim_projects`, …).
     Подставьте имена своей схемы или заведите представления-синонимы в реплике.

Этот файл сам по себе не подключается. Он нужен как короткая памятка рядом с README.
"""

SOURCE_DB = dict(
    host="postgres.example.local",
    port=5432,
    dbname="report_db",
    user="analytics_user",
)

# Пример заглушек очередей (полный список — в py/wfm_settings.py):
SUPPORT_QUEUES = {
    "q_voice_sla0": ("SLA-0", "Голос"),
    "q_chat_sla1": ("чат SLA1", "Чат"),
    "q_mail_inbox": ("Email", "Почта"),
}

STORE_DIR = "data"
BACKEND = "parquet"
