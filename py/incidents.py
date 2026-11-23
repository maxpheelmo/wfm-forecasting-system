# -*- coding: utf-8 -*-
"""
Инциденты из Service Desk (журнал решений №35, №42).

Вход: самая свежая выгрузка в inputs\\ — exportSD*.xlsx (или .csv) с колонками
      Номер, Дата создания, Дата решения, Тема, Услуга, Приоритет, Описание (лишние колонки не мешают).
Массовый инцидент распознаётся по шаблону описания:
      «Время возникновения: ДД.ММ.ГГГГ ЧЧ:ММ … Время фиксации: … [теги влияния]».
Группа (wfm_settings.INCIDENT_*):
      КЦ                 — сбой источник, чат-бота, голосового робота;
      клиентский онлайн  — линия, линия обслуживания, платежи, акции, регистрация/идентификация и т.п.;
      не клиентский      — внутренние сервисы, автоматизация, реклама, VPN (в прогноз не идёт);
      прочее             — остальное (в прогноз не идёт).
Выход: DB\\incidents\\incidents.parquet — ТОЛЬКО разобранные поля (номер, услуга, группа, приоритет, начало, решение,
       теги). Описание и тема не сохраняются: в обращениях КЦ бывают данные клиентов.
Пересобирается сам, если в inputs появилась выгрузка новее.
"""
import glob
import os
import re

import pandas as pd

import wfm_settings as S
from wfm_common import INPUTS_DIR, store_root

COLS = {"Номер": "number", "Дата создания": "created", "Дата решения": "resolved", "Услуга": "service",
        "Приоритет": "priority", "Описание": "description"}
RE_START = re.compile(r"Время возникновения:\s*(\d{2}\.\d{2}\.\d{4}\s+\d{1,2}:\d{2})")
RE_TAGS = re.compile(r"Время фиксации:[^\[]*\[([^\]]*)\]")
OUT = ("incidents", "incidents.parquet")


def latest_export():
    files = [f for f in glob.glob(os.path.join(INPUTS_DIR, "exportSD*")) if f.lower().endswith((".xlsx", ".csv"))
             and not os.path.basename(f).startswith("~$")]
    return max(files, key=os.path.getmtime) if files else None


def classify(service, tags):
    u, t = str(service), str(tags).lower()
    if any(k in u for k in S.INCIDENT_SERVICES_KC):
        return "КЦ"
    if any(k in u for k in S.INCIDENT_SERVICES_NOT_CLIENT):
        return "не клиентский"
    if any(k in t for k in S.INCIDENT_CLIENT_TAGS) or any(u.startswith(k) for k in S.INCIDENT_SERVICES_CLIENT):
        return "клиентский онлайн"
    return "прочее"


def parse_export(path):
    raw = pd.read_csv(path, sep=None, engine="python") if path.lower().endswith(".csv") else pd.read_excel(path)
    miss = [c for c in COLS if c not in raw.columns]
    if miss:
        raise ValueError(f"В выгрузке {os.path.basename(path)} нет колонок: {', '.join(miss)}")
    d = raw[list(COLS)].rename(columns=COLS)
    desc = d.pop("description").astype(str)
    d["start"] = pd.to_datetime(desc.str.extract(RE_START)[0], format="%d.%m.%Y %H:%M", errors="coerce")
    d["tags"] = desc.str.extract(RE_TAGS)[0].fillna("").str.strip()
    del desc
    d = d[d["start"].notna()].copy()
    d["created"] = pd.to_datetime(d["created"], errors="coerce")
    d["resolved"] = pd.to_datetime(d["resolved"], errors="coerce")
    d["number"] = d["number"].astype(str).str.replace(r"\.0$", "", regex=True)
    d["group"] = [classify(u, t) for u, t in zip(d["service"], d["tags"])]
    d["source"] = os.path.basename(path)
    return d[["number", "service", "group", "priority", "start", "resolved", "created", "tags", "source"]]


def load(verbose=True):
    """-> DataFrame инцидентов (из кэша DB\\incidents или свежей выгрузки) или None, если выгрузки нет."""
    out = os.path.join(store_root(), *OUT)
    src = latest_export()
    if src and (not os.path.exists(out) or os.path.getmtime(src) > os.path.getmtime(out)):
        d = parse_export(src)
        if os.path.exists(out):                       # старые инциденты, которых нет в новой выгрузке, сохраняем
            old = pd.read_parquet(out)
            d = pd.concat([old[~old["number"].isin(d["number"])], d], ignore_index=True)
        os.makedirs(os.path.dirname(out), exist_ok=True)
        d.to_parquet(out + ".tmp", index=False)
        os.replace(out + ".tmp", out)
        if verbose:
            g = d["group"].value_counts().to_dict()
            print(f"Инциденты Service Desk: {os.path.basename(src)} → {len(d)} массовых, "
                  f"{d['start'].min():%d.%m.%Y} – {d['start'].max():%d.%m.%Y}; по группам: {g}")
    if not os.path.exists(out):
        return None
    return pd.read_parquet(out)


def incident_days(groups=None):
    """Дни инцидентов выбранных групп (по умолчанию wfm_settings.INCIDENT_GROUPS_USED): от начала до решения,
    но не дольше INCIDENT_MAX_DAYS суток. -> DatetimeIndex или None."""
    d = load(verbose=False)
    if d is None or d.empty:
        return None
    groups = groups or S.INCIDENT_GROUPS_USED
    d = d[d["group"].isin(groups)]
    days = set()
    for a, b in zip(d["start"], d["resolved"]):
        b = b if pd.notna(b) and b >= a else a
        b = min(b, a + pd.Timedelta(days=S.INCIDENT_MAX_DAYS - 1))
        days |= set(pd.date_range(a.normalize(), b.normalize()))
    return pd.DatetimeIndex(sorted(days))


if __name__ == "__main__":
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    x = load()
    print("Выгрузки нет: положите exportSD*.xlsx в inputs." if x is None else
          x.groupby(["group", "service"]).size().sort_values(ascending=False).to_string())
