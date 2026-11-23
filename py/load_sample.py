# -*- coding: utf-8 -*-
"""Записать синтетические CSV из sample_data/ в хранилище Parquet (папка data/).

Данные выдуманы и нужны только для того, чтобы поднять DuckDB и посмотреть витрины.
К живой реплике скрипт не подключается.
"""
import os
import sys

import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import wfm_settings as S  # noqa: E402
from wfm_common import store_root  # noqa: E402

SAMPLE = os.path.join(ROOT, "sample_data")

TABLES = {
    "voice_queued_calls": "voice_queued_calls.csv",
    "voice_call_legs": "voice_call_legs.csv",
    "chat_session_iterations": "chat_session_iterations.csv",
    "agent_status_changes": "agent_status_changes.csv",
    "dim_projects": "dim_projects.csv",
    "dim_employees": "dim_employees.csv",
}

TIME_COLS = {
    "voice_queued_calls": ["enqueued_time", "unblocked_time", "dequeued_time"],
    "voice_call_legs": ["created", "connected", "ended"],
    "chat_session_iterations": [
        "start_time", "unblock_time", "distribute_time", "oper_distribute_time",
        "first_oper_answer_time", "first_answer_time", "finish_time",
    ],
    "agent_status_changes": ["entered", "leaved"],
}


def main():
    root = store_root()
    os.makedirs(root, exist_ok=True)
    loaded_at = pd.Timestamp.now()
    for name, fname in TABLES.items():
        path = os.path.join(SAMPLE, fname)
        if not os.path.exists(path):
            print(f"нет файла {path}, пропускаю")
            continue
        df = pd.read_csv(path)
        for col in TIME_COLS.get(name, []):
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce")
        if "removed" in df.columns:
            df["removed"] = df["removed"].astype(bool)
        df["_loaded_at"] = loaded_at
        dest_dir = os.path.join(root, name)
        os.makedirs(dest_dir, exist_ok=True)
        dest = os.path.join(dest_dir, "sample.parquet")
        table = pa.Table.from_pandas(df, preserve_index=False)
        tmp = dest + ".tmp"
        pq.write_table(table, tmp, compression="zstd")
        os.replace(tmp, dest)
        print(f"{name}: {len(df)} строк → {dest}")
    print("Готово. Хранилище:", root)
    print("Проверка витрин: python -c \"import sys; sys.path.insert(0, 'py'); import wfm_store; "
          "print(wfm_store.query('select channel, count(*) n from v_interval_volume group by 1'))\"")


if __name__ == "__main__":
    main()
