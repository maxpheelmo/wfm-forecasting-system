# -*- coding: utf-8 -*-
"""
Общее для загрузчика и проверок: подключение к реплике источника и чтение в pyarrow с повтором.
"""
import datetime as dt
import getpass
import os
import time

import pyarrow as pa

import wfm_settings as S

HERE = os.path.dirname(os.path.abspath(__file__))       # папка py (скрипты, views)
ROOT = os.path.dirname(HERE)                             # корень: wfm.bat, py, tools, inputs, DB, reports
INPUTS_DIR = os.path.join(ROOT, "inputs")                # входные данные помимо источника
REPORTS_DIR = os.path.join(ROOT, "reports")
EVENTS_FILE = os.path.join(INPUTS_DIR, "events.xlsx")

# Отчёт -> подпапка reports по началу имени файла (новые виды отчётов — дописать сюда)
REPORT_KINDS = {
    "forecast": ("Прогноз ", "Бэктест ", "Точность ", "forecast_", "backtest_", "accuracy_"),
    "need": ("Потребность ",),
    "checks": ("Проверка хранилища ", "Обзор рядов ", "Сверка ", "store_check_", "series_overview_", "razvedka_",
               "check_profile30_", "sl_threshold_", "bot_outcomes_", "voice_routes_"),
}


def error_line(exc):
    """Одна строка. У таймаута реплики список адресов не печатаем."""
    first = next((line.strip() for line in str(exc).splitlines() if line.strip()), type(exc).__name__)
    if not first.endswith("."):
        first += "."
    return f"ОШИБКА: {type(exc).__name__}: {first}"


def is_db_timeout(exc):
    head = f"{type(exc).__name__} {str(exc).splitlines()[0] if str(exc) else ''}".lower()
    return "timeout" in head


def report_kind(name):
    return next((k for k, pref in REPORT_KINDS.items() if name.startswith(pref)), "other")


def report_path(name):
    """Путь для отчёта: reports\\<вид>\\name (папка создаётся при необходимости).
    Если файл с таким именем уже есть (два запуска в одну минуту) — добавляется « (2)», « (3)»…"""
    d = os.path.join(REPORTS_DIR, report_kind(name))
    os.makedirs(d, exist_ok=True)
    path, (base, ext), k = os.path.join(d, name), os.path.splitext(name), 2
    while os.path.exists(path):
        path = os.path.join(d, f"{base} ({k}){ext}")
        k += 1
    return path


def stamp(t=None):
    """Метка времени для имени файла: 29.10.2026 22-32 (двоеточие в именах файлов Windows запрещено)."""
    return f"{(t or dt.datetime.now()):%d.%m.%Y %H-%M}"

# Типы PostgreSQL (oid) -> типы Parquet/Arrow. Всё неизвестное сохраняем строкой.
PG_TO_ARROW = {
    16: pa.bool_(),                      # bool
    20: pa.int64(),                      # bigint
    21: pa.int16(),                      # smallint
    23: pa.int32(),                      # integer
    700: pa.float32(),                   # real
    701: pa.float64(),                   # double
    1700: pa.float64(),                  # numeric
    25: pa.string(),                     # text
    1043: pa.string(),                   # varchar
    1042: pa.string(),                   # char
    1082: pa.date32(),                   # date
    1114: pa.timestamp("us"),            # timestamp without time zone
    1184: pa.timestamp("us", tz="UTC"),  # timestamptz
}


def source_user():
    return os.environ.get("SOURCE_PG_USER", S.SOURCE_DB["user"])


def get_password(user):
    pw = os.environ.get("SOURCE_PG_PASSWORD")
    if pw:
        return pw
    try:
        import keyring
        pw = keyring.get_password(S.KEYRING_SERVICE, user)
        if pw:
            return pw
    except Exception:
        pass
    return getpass.getpass(f"Пароль PostgreSQL для {user} (чтобы не вводить каждый раз — wfm.bat → Настройки → пароль): ")


def connect_source():
    import psycopg
    user = source_user()
    db = dict(S.SOURCE_DB, user=user)
    return psycopg.connect(
        **db, password=get_password(user), autocommit=True, connect_timeout=15,
        options="-c default_transaction_read_only=on -c statement_timeout=900000",
    )


def _retry(fn, retries=3):
    """Реплика иногда обрывает запрос «conflict with recovery» — повторяем."""
    for attempt in range(1, retries + 1):
        try:
            return fn()
        except Exception as e:
            if attempt < retries and "conflict with recovery" in str(e):
                print(f"      реплика прервала запрос (conflict with recovery), повтор {attempt}...")
                time.sleep(5 * attempt)
                continue
            raise


def fetch_arrow(conn, sql, params=None):
    """Запрос -> pyarrow.Table с типами по описанию колонок PostgreSQL."""
    def run():
        with conn.cursor() as cur:
            cur.execute(sql, params or {})
            return cur.description, cur.fetchall()
    desc, rows = _retry(run)
    arrays, names = [], []
    for i, col in enumerate(desc):
        t = PG_TO_ARROW.get(col.type_code, pa.string())
        vals = [r[i] for r in rows]
        if pa.types.is_string(t):
            vals = [None if v is None else (v if isinstance(v, str) else str(v)) for v in vals]
        elif pa.types.is_floating(t):
            vals = [None if v is None else float(v) for v in vals]
        elif pa.types.is_timestamp(t) and t.tz is not None:
            vals = [None if v is None else v.astimezone(dt.timezone.utc) for v in vals]
        arrays.append(pa.array(vals, type=t))
        names.append(col.name)
    return pa.Table.from_arrays(arrays, names=names)


def fetch_rows(conn, sql, params=None):
    def run():
        with conn.cursor() as cur:
            cur.execute(sql, params or {})
            return cur.fetchall()
    return _retry(run)


def store_root():
    p = S.STORE_DIR
    return p if os.path.isabs(p) else os.path.join(ROOT, p)
