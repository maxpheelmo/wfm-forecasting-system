# -*- coding: utf-8 -*-
"""
Единая точка чтения хранилища WFM. Прогноз и отчёты читают ТОЛЬКО через query().

    import wfm_store
    df = wfm_store.query("select * from v_interval_volume where channel = 'Голос'")

Что видно в SQL (одинаково для обоих backend):
  сырые таблицы источника   agent_status_changes, voice_queued_calls, voice_call_legs, chat_session_iterations
                      (у чатов — последняя версия каждой итерации);
  справочники         dim_projects, dim_employees (последняя выгрузка),
                      *_history — все выгрузки с колонкой _loaded_at;
  настройки           support_queues, bot_projects, support_departments, wfm_params (из wfm_settings.py);
  представления       файлы views/*.sql по порядку имён (v_voice_contacts, v_interval_volume, ...).

backend = "parquet": DuckDB поверх папки DB, представления создаются при каждом подключении.
backend = "postgres": своя PostgreSQL; таблицы с теми же именами, init_postgres() создаёт
                      настройки и представления из тех же SQL-файлов.
SQL в views/ пишем на общем подмножестве DuckDB и PostgreSQL.
"""
import glob
import os

import pandas as pd

import wfm_settings as S
from wfm_common import HERE, store_root

VIEWS_DIR = os.path.join(HERE, "views")

# сырые таблицы: имя -> (ключ, строки могут меняться)
RAW_TABLES = {
    "agent_status_changes": ("sequence_id", False),
    "voice_queued_calls": ("id", False),
    "voice_call_legs": ("id", False),
    "chat_session_iterations": ("id", True),
}
DIM_TABLES = ["dim_projects", "dim_employees"]
OPTIONAL_TABLES = ["forecast_run", "forecast_interval"]   # появляются после первого прогноза


# ---------------------------------------------------------------------------
# Настройки как таблицы
# ---------------------------------------------------------------------------
def settings_tables():
    """(имя, [(колонка, тип)], строки) — одинаково для DuckDB и PostgreSQL."""
    return [
        ("support_queues", [("project_id", "varchar"), ("queue_name", "varchar"), ("channel", "varchar")],
         [(k, v[0], v[1]) for k, v in S.SUPPORT_QUEUES.items()]),
        ("bot_projects", [("project_id", "varchar"), ("bot_name", "varchar")],
         list(S.BOT_PROJECTS.items())),
        ("voice_entry_projects", [("project_id", "varchar"), ("project_name", "varchar")],
         list(getattr(S, "VOICE_ENTRY_PROJECTS", {}).items())),
        ("voice_robot_projects", [("project_id", "varchar"), ("project_name", "varchar")],
         list(getattr(S, "VOICE_ROBOT_PROJECTS", {}).items())),
        ("service_projects", [("project_id", "varchar"), ("project_name", "varchar")],
         list(getattr(S, "CHAT_SERVICE_PROJECTS", {}).items())),
        ("excluded_queues", [("project_id", "varchar"), ("project_name", "varchar")],
         list(getattr(S, "EXCLUDED_QUEUES", {}).items())),
        ("support_departments", [("department", "varchar")], [(d,) for d in S.SUPPORT_DEPARTMENTS]),
        ("wfm_params", [("name", "varchar"), ("value", "double precision")],
         [("sl_voice_sec", float(S.SL_THRESHOLD_SEC["Голос"])),
          ("sl_chat_sec", float(S.SL_THRESHOLD_SEC["Чат"])),
          ("sl_mail_sec", float(S.SL_THRESHOLD_SEC["Почта"])),
          ("short_abandon_sec", float(S.SHORT_ABANDON_SEC))]),
    ]


def view_files():
    return sorted(glob.glob(os.path.join(VIEWS_DIR, "*.sql")))


def _create_settings(execute, cascade=False):
    # в PostgreSQL от таблиц настроек зависят представления: cascade их удаляет, _create_views создаёт заново
    for name, cols, rows in settings_tables():
        execute(f"drop table if exists {name}{' cascade' if cascade else ''}")
        execute(f"create table {name} ({', '.join(f'{c} {t}' for c, t in cols)})")
        for r in rows:
            vals = ", ".join("null" if v is None else (repr(v) if isinstance(v, float) else "'" + str(v).replace("'", "''") + "'")
                             for v in r)
            execute(f"insert into {name} values ({vals})")


def _create_views(execute):
    for path in view_files():
        with open(path, encoding="utf-8") as f:
            sql = f.read().strip().rstrip(";")
        try:
            execute(sql)
        except Exception as e:
            raise RuntimeError(f"Ошибка в представлении {os.path.basename(path)}: {e}") from e


# ---------------------------------------------------------------------------
# Parquet + DuckDB
# ---------------------------------------------------------------------------
def _glob(name):
    return os.path.join(store_root(), name, "*.parquet").replace("\\", "/")


def _has_files(name):
    return bool(glob.glob(os.path.join(store_root(), name, "*.parquet")))


def _duck_columns(con, src):
    return [r[0] for r in con.execute(f"describe select * from {src}").fetchall()]


def raw_select_sql(con, name, key, mutable):
    """SQL сырой таблицы из файлов: у изменяемых — последняя версия строки по _loaded_at."""
    src = f"read_parquet('{_glob(name)}', union_by_name=true)"
    if not mutable:
        return f"select * from {src}"
    cols = _duck_columns(con, src)
    order = "_loaded_at desc nulls last" if "_loaded_at" in cols else key
    return (f"select * exclude (_rn) from (select *, row_number() over "
            f"(partition by {key} order by {order}) as _rn from {src}) where _rn = 1")


def connect_duckdb():
    import duckdb
    con = duckdb.connect()
    for name, (key, mutable) in RAW_TABLES.items():
        if _has_files(name):
            con.execute(f"create or replace view {name} as {raw_select_sql(con, name, key, mutable)}")
    for name in DIM_TABLES:
        if _has_files(name):
            src = f"read_parquet('{_glob(name)}', union_by_name=true)"
            con.execute(f"create or replace view {name}_history as select * from {src}")
            con.execute(f"create or replace view {name} as select * exclude (_loaded_at) from {src} "
                        f"where _loaded_at = (select max(_loaded_at) from {src})")
    for name in OPTIONAL_TABLES:
        if _has_files(name):
            con.execute(f"create or replace view {name} as select * from read_parquet('{_glob(name)}', union_by_name=true)")
    _create_settings(con.execute)
    missing = [n for n in list(RAW_TABLES) + DIM_TABLES if not _has_files(n)]
    if missing:
        raise RuntimeError(f"В хранилище {store_root()} нет данных таблиц: {', '.join(missing)}. "
                           "Сначала загрузите данные (wfm.bat → Настройки → обновить данные источника).")
    if not view_files():
        raise RuntimeError(f"Не найдены SQL-файлы представлений в папке {VIEWS_DIR}. "
                           "Папка views из архива должна лежать рядом с wfm_store.py.")
    _create_views(con.execute)
    return con


# ---------------------------------------------------------------------------
# Своя PostgreSQL (на будущее)
# ---------------------------------------------------------------------------
def connect_postgres():
    import psycopg
    from wfm_common import get_password
    db = dict(S.TARGET_DB)
    return psycopg.connect(**db, password=get_password(db["user"]), autocommit=True)


def init_postgres(conn=None):
    """Создать в своей PostgreSQL таблицы настроек и представления из views/*.sql."""
    conn = conn or connect_postgres()
    with conn.cursor() as cur:
        _create_settings(cur.execute, cascade=True)
        _create_views(cur.execute)
    return conn


# ---------------------------------------------------------------------------
def connect():
    if S.BACKEND == "parquet":
        return connect_duckdb()
    if S.BACKEND == "postgres":
        return connect_postgres()
    raise ValueError(f"Неизвестный BACKEND в wfm_settings.py: {S.BACKEND}")


def query(sql, con=None):
    """SQL -> pandas.DataFrame. Если подключение не передано — открывает и закрывает своё."""
    own = con is None
    con = con or connect()
    try:
        if S.BACKEND == "parquet":
            return con.execute(sql).df()
        with con.cursor() as cur:
            cur.execute(sql)
            cols = [c.name for c in cur.description]
            return pd.DataFrame(cur.fetchall(), columns=cols)
    finally:
        if own:
            con.close()
