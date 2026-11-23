# -*- coding: utf-8 -*-
"""
Общий загрузчик хранилища WFM из реплики источника. Запускается сам перед каждым пунктом консоли
(wfm.bat), вручную — «Настройки → обновить данные источника».

Что грузит (все проекты источника, фильтр рабочих очередей — в представлениях):
  agent_status_changes          статусы операторов (источник хранит только 180 дней!)
  voice_queued_calls            голос: попадания в очередь
  voice_call_legs                  голос: плечи звонков (БЕЗ номеров: src_id/dst_id/IP не берём,
                             src/dst_abonent — только логин оператора и имя IVR, номера и IP не забираем)
  chat_session_iterations  чаты и почта (БЕЗ from_addr и remote_name)
  dim_projects, dim_employees — справочники, копия с датой при каждом запуске

Как устроено (разведки 20.10.2026): строки в источник пишутся по ЗАВЕРШЕНИИ (выход из статуса,
выход из очереди, конец плеча, конец итерации).
  * Первый запуск таблицы — полный слепок по суткам (от старых дней к новым), файлы
    snap_ГГГГ-ММ-ДД.parquet. Прерванный слепок продолжается со следующего запуска.
  * Дальше — догрузка: строки с временем окончания >= (наше последнее − OVERLAP_DAYS),
    только новые ключи -> inc_ГГГГММДД_ЧЧММСС.parquet.
  * Чаты/почта: незакрытые итерации перечитываются; изменившиеся строки дописываются
    новой версией, при чтении (wfm_store) берётся последняя.
  * Статусы, голос, плечи считаются неизменяемыми: если строка в источнике изменилась —
    предупреждение (такого быть не должно).

Сверки после каждого запуска (папка DB\\_log):
  load_runs.csv              по таблицам: сколько записано, всего, версии, расхождения по дням
  check_days_<таблица>.csv   строк по дням у нас и в источнике. Если у нас меньше (источник дописал строки за
                             прошлые дни задним числом) — такие дни перечитываются целиком и дописываются
                             автоматически (файлы fix_*.parquet). Если у нас больше — источник удалил часть строк,
                             у нас они сохраняются (не ошибка).
  depth.csv                  самые ранние/поздние даты таблиц источника (если min сдвигается — таблица чистится)
  queue_check.csv        очереди, где отвечает отдел поддержки, но их нет в wfm_settings.SUPPORT_QUEUES
  entry_check.csv            откуда начинались сессии, дошедшие до рабочих очередей (30 дней); неизвестный вход —
                             предупреждение (новый канал или перевод из чужой очереди)

Запуск вручную: py wfm_load.py [--only agent_status_changes,voice_call_legs] [--no-checks]
"""
import argparse
import csv
import datetime as dt
import glob
import json
import os
import sys
import time

import duckdb
import pyarrow as pa
import pyarrow.compute as pc
import pyarrow.parquet as pq

import wfm_settings as S
from wfm_common import connect_source, error_line, fetch_arrow, fetch_rows, is_db_timeout, store_root

SNAPSHOT_MARK = "_snapshot_done.json"

# Колонки voice_call_legs без данных клиентов
LEGS_SELECT = """
    id, leg_id, session_id, base_id, protocol,
    case when src_abonent_type in ('SP', 'IVR') then src_abonent end as src_abonent,
    case when dst_abonent_type in ('SP', 'IVR') then dst_abonent end as dst_abonent,
    src_abonent_type, dst_abonent_type, incoming, intrusion,
    created, connected, ended, initiated, pickedup,
    voip_reason, internal_reason, cpd_result, text_reason, reason_treat, ended_by, is_bind"""

MBSTAT_SELECT = """
    id, session_id, project_id, attempt_num, number_in_session, oper_id, script_name,
    start_time, unblock_time, distribute_time, oper_distribute_time,
    first_oper_answer_time, first_answer_time, finish_time, result,
    unblocks_before, session_restored"""

# name, key, day_col (слепок и сверка по дням), inc_col (догрузка), mutable, select, open_cond
TABLES = [
    dict(name="agent_status_changes", key="sequence_id", day_col="entered", inc_col="leaved",
         mutable=False, select="*", open_cond=None, purged=True),
    dict(name="voice_queued_calls", key="id", day_col="enqueued_time", inc_col="dequeued_time",
         mutable=False, select="*", open_cond=None, purged=False),
    dict(name="voice_call_legs", key="id", day_col="ended", inc_col="ended",
         mutable=False, select=LEGS_SELECT, open_cond=None, purged=False),
    dict(name="chat_session_iterations", key="id", day_col="start_time", inc_col="finish_time",
         mutable=True, select=MBSTAT_SELECT, open_cond="finish_time is null", purged=False),
]

DIMS = [
    dict(name="dim_projects", select="""
        uuid, title, creationdate, removed, removaldate, formtemplate,
        description, datachannel, stateuuid, state, calllimit,
        servicelevelparameter, scripttype, ivrscript, callstopcriteriauuid, callstopcriteriatitle,
        ivrpolicyuuid, ivrpolicytitle"""),
    dict(name="dim_employees", select="login, title, department, post, removed"),
]

DEPTH_TABLES = [
    ("agent_status_changes", "entered"), ("voice_queued_calls", "enqueued_time"), ("voice_call_legs", "ended"),
    ("chat_session_iterations", "start_time"), ("src_phone_calls", "creationdate"),
    ("src_custom_forms", "creationdate"), ("call_status", "entered"),
]


# ---------------------------------------------------------------------------
# Хранилище (Parquet)
# ---------------------------------------------------------------------------
class Store:
    def __init__(self, root):
        self.root = os.path.abspath(root)
        self.log = os.path.join(self.root, "_log")
        os.makedirs(self.log, exist_ok=True)
        self.con = duckdb.connect()

    def dir(self, name):
        d = os.path.join(self.root, name)
        os.makedirs(d, exist_ok=True)
        return d

    def files(self, name):
        return sorted(glob.glob(os.path.join(self.dir(name), "*.parquet")))

    def src(self, name):
        g = os.path.join(self.dir(name), "*.parquet").replace("\\", "/")
        return f"read_parquet('{g}', union_by_name=true)"

    def columns(self, name):
        if not self.files(name):
            return []
        return [r[0] for r in self.con.execute(f"describe select * from {self.src(name)}").fetchall()]

    def latest(self, name, key, mutable):
        """SQL: строки хранилища, у изменяемых таблиц — последняя версия."""
        if not mutable:
            return f"select * from {self.src(name)}"
        order = "_loaded_at desc nulls last" if "_loaded_at" in self.columns(name) else key
        return (f"select * exclude (_rn) from (select *, row_number() over (partition by {key} "
                f"order by {order}) as _rn from {self.src(name)}) where _rn = 1")

    def write(self, name, table, fname, loaded_at):
        table = table.append_column("_loaded_at", pa.array([loaded_at] * table.num_rows, type=pa.timestamp("us")))
        path = os.path.join(self.dir(name), fname)
        tmp = path + ".tmp"
        pq.write_table(table, tmp, compression="zstd")
        os.replace(tmp, path)   # атомарно: обрыв посередине не оставит битый файл

    def append_csv(self, fname, row):
        path = os.path.join(self.log, fname)
        new = not os.path.exists(path)
        with open(path, "a", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=list(row), delimiter=";")
            if new:
                w.writeheader()
            w.writerow(row)

    def write_csv(self, fname, rows, fields):
        with open(os.path.join(self.log, fname), "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=fields, delimiter=";")
            w.writeheader()
            w.writerows(rows)


# ---------------------------------------------------------------------------
# Загрузка таблицы
# ---------------------------------------------------------------------------
def snapshot(conn, store, t, loaded_at):
    name, day = t["name"], t["day_col"]
    lo, hi = fetch_rows(conn, f"select min({day})::date, localtimestamp::date from {name}")[0]
    have = {os.path.basename(p) for p in store.files(name)}
    total, done_days = 0, 0
    if lo is not None:
        days = [lo + dt.timedelta(days=i) for i in range((hi - lo).days + 1)]
        todo = [d for d in days if f"snap_{d:%Y-%m-%d}.parquet" not in have]
        print(f"    слепок: дней в источнике {len(days)} ({lo:%d.%m.%Y} – {hi:%d.%m.%Y}), осталось {len(todo)}")
        t0 = time.time()
        for i, d in enumerate(todo, 1):
            tb = fetch_arrow(conn, f"select {t['select']} from {name} where {day} >= %(a)s and {day} < %(b)s",
                             dict(a=d, b=d + dt.timedelta(days=1)))
            if tb.num_rows:
                store.write(name, tb, f"snap_{d:%Y-%m-%d}.parquet", loaded_at)
            total += tb.num_rows
            done_days += 1
            if i == 1 or i % 25 == 0 or i == len(todo):
                print(f"      {d:%d.%m.%Y}: {tb.num_rows} строк  [{i}/{len(todo)}, {time.time() - t0:.0f} с]")
    # строки без даты (у чатов start_time может быть пустым)
    if "snap_nodate.parquet" not in have:
        tb = fetch_arrow(conn, f"select {t['select']} from {name} where {day} is null")
        if tb.num_rows:
            store.write(name, tb, "snap_nodate.parquet", loaded_at)
            print(f"      без даты: {tb.num_rows} строк")
        total += tb.num_rows
    with open(os.path.join(store.dir(name), SNAPSHOT_MARK), "w", encoding="utf-8") as f:
        json.dump(dict(done_at=dt.datetime.now().isoformat(timespec="seconds"),
                       first_day=str(lo), last_day=str(hi)), f, ensure_ascii=False)
    return total, done_days


def new_and_changed(store, t, batch):
    """Разделить пачку из источника на новые ключи и изменённые строки (против последней версии)."""
    name, key = t["name"], t["key"]
    if not store.files(name) or batch.num_rows == 0:
        return batch, batch.slice(0, 0)
    store_cols = set(store.columns(name))
    cols = [c for c in batch.column_names if c != key and c in store_cols]
    cond = " or ".join(f"b.{c} is distinct from s.{c}" for c in cols) or "false"
    store.con.register("batch", batch)
    try:
        latest = store.latest(name, key, t["mutable"])
        new_ids = [r[0] for r in store.con.execute(
            f"select b.{key} from batch b anti join ({latest}) s using ({key})").fetchall()]
        changed_ids = [r[0] for r in store.con.execute(
            f"select b.{key} from batch b join (select * from ({latest}) where {key} in (select {key} from batch)) s "
            f"using ({key}) where {cond}").fetchall()]
    finally:
        store.con.unregister("batch")
    ktype = batch.schema.field(key).type
    new = batch.filter(pc.is_in(batch[key], value_set=pa.array(new_ids, type=ktype)))
    changed = batch.filter(pc.is_in(batch[key], value_set=pa.array(changed_ids, type=ktype)))
    return new, changed


def increment(conn, store, t, loaded_at, overlap_days):
    name, key, inc = t["name"], t["key"], t["inc_col"]
    mx = store.con.execute(f"select max({inc}) from {store.src(name)}").fetchone()[0] if store.files(name) else None
    if mx is None:
        print("    догрузка: в хранилище нет строк с датой окончания — пропуск")
        return 0, 0
    since = mx - dt.timedelta(days=overlap_days)
    where = f"{inc} >= %(s)s" + (f" or {t['open_cond']}" if t["open_cond"] else "")
    batch = fetch_arrow(conn, f"select {t['select']} from {name} where {where}", dict(s=since))
    new, changed = new_and_changed(store, t, batch)
    to_write = pa.concat_tables([new, changed]) if t["mutable"] else new
    if to_write.num_rows:
        store.write(name, to_write, f"inc_{loaded_at:%Y%m%d_%H%M%S}.parquet", loaded_at)
    print(f"    догрузка с {since:%d.%m.%Y %H:%M}: получено {batch.num_rows}, новых {new.num_rows}, "
          f"изменённых {changed.num_rows}{' (дописаны новой версией)' if t['mutable'] and changed.num_rows else ''}")
    return new.num_rows, changed.num_rows


def check_days(conn, store, t):
    """Сверка числа строк по дням. -> (дни, где у нас меньше, чем в источник; дни, где больше; первый день источник)."""
    name, key, day = t["name"], t["key"], t["day_col"]
    src = {d: n for d, n in fetch_rows(conn, f"select {day}::date, count(*) from {name} group by 1")}
    ours = {d: n for d, n in store.con.execute(
        f"select cast({day} as date), count(*) from ({store.latest(name, key, t['mutable'])}) group by 1").fetchall()}
    real = [d for d in src if d is not None]
    first = min(real) if real else None
    today = dt.date.today()
    rows, short, extra = [], [], []
    for d in sorted(set(src) | set(ours), key=lambda x: (x is None, x)):
        s, o = src.get(d), ours.get(d, 0)
        if t["purged"] and d is not None and first is not None and d < first:
            note = "источник уже удалил"
        elif t["purged"] and d == first:
            note = "первый день источника (частично удалён)"
        elif d == today:
            note = "сегодня (идут новые строки)" if s != o else ""
        elif (s or 0) > o:
            note = "НЕДОСТАЧА у нас"
            short.append(d)
        elif o > (s or 0):
            note = "в источнике удалена часть строк (у нас сохранены)"
            extra.append(d)
        else:
            note = ""
        rows.append(dict(day=d if d is not None else "без даты", store=o, source=s if s is not None else "",
                         diff=o - (s or 0), note=note))
    store.write_csv(f"check_days_{name}.csv", rows, ["day", "store", "source", "diff", "note"])
    return short, extra, first


def repair_days(conn, store, t, days, loaded_at):
    """Дни, где у нас строк меньше: источник дописал их задним числом (после окна догрузки).
    Перечитываем такие дни целиком и дописываем недостающие строки. -> {день: добавлено строк}."""
    name, day = t["name"], t["day_col"]
    added = {}
    for d in days:
        if d is None:
            continue
        batch = fetch_arrow(conn, f"select {t['select']} from {name} where {day} >= %(a)s and {day} < %(b)s",
                            dict(a=d, b=d + dt.timedelta(days=1)))
        new, changed = new_and_changed(store, t, batch)
        to_write = pa.concat_tables([new, changed]) if t["mutable"] else new
        if to_write.num_rows:
            store.write(name, to_write, f"fix_{loaded_at:%Y%m%d_%H%M%S}_{d:%Y%m%d}.parquet", loaded_at)
        added[d] = new.num_rows
    return added


def load_table(conn, store, t, loaded_at, overlap_days, checks):
    t0 = time.time()
    name = t["name"]
    print(f"\n[{name}]")
    snap_rows = snap_days = 0
    if not os.path.exists(os.path.join(store.dir(name), SNAPSHOT_MARK)):
        snap_rows, snap_days = snapshot(conn, store, t, loaded_at)
        mode = "слепок+догрузка"
    else:
        mode = "догрузка"
    new_rows, changed = increment(conn, store, t, loaded_at, overlap_days)
    r = store.con.execute(f"select count(*), count(distinct {t['key']}), min({t['day_col']}), max({t['inc_col']}) "
                          f"from {store.src(name)}").fetchone()
    short, extra, src_first = (check_days(conn, store, t) if checks else ([], [], None))
    fixed = {}
    if short:
        fixed = repair_days(conn, store, t, short, loaded_at)
        print("    дописано задним числом (источник добавил строки за прошлые дни): "
              + ", ".join(f"{d:%d.%m} +{n}" for d, n in fixed.items()))
        short, extra, src_first = check_days(conn, store, t)
        new_rows += sum(fixed.values())
        r = store.con.execute(f"select count(*), count(distinct {t['key']}), min({t['day_col']}), max({t['inc_col']}) "
                              f"from {store.src(name)}").fetchone()
    bad = short
    versions = r[0] - r[1]
    flags = []
    if versions and not t["mutable"]:
        flags.append(f"ДУБЛИКАТЫ ключа: {versions}")
    if changed and not t["mutable"]:
        flags.append(f"ИЗМЕНЁННЫЕ строки в источнике: {changed} (не записаны; сообщить)")
    if bad:
        flags.append(f"НЕДОСТАЧА по дням и после дозагрузки: {len(bad)} ({', '.join(f'{d:%d.%m}' for d in bad[:8])}"
                     f"{' …' if len(bad) > 8 else ''}) — _log/check_days_{name}.csv")
    info = []
    if extra:
        info.append(f"в источник удалена часть строк за {len(extra)} дн. ({', '.join(f'{d:%d.%m}' for d in extra[:6])}"
                    f"{' …' if len(extra) > 6 else ''}) — у нас сохранены, это нормально")
    size = sum(os.path.getsize(p) for p in store.files(name)) / 2 ** 20
    print(f"    итого у нас: {r[0]} строк ({r[1]} ключей), {size:.0f} МБ; "
          f"даты {r[2]:%d.%m.%Y} – {r[3]:%d.%m.%Y %H:%M}" if r[2] and r[3] else f"    итого у нас: {r[0]} строк")
    for i_ in info:
        print("    к сведению: " + i_)
    for f_ in flags:
        print("    ВНИМАНИЕ: " + f_)
    store.append_csv("load_runs.csv", dict(
        run=loaded_at, table=name, mode=mode, snapshot_days=snap_days, snapshot_rows=snap_rows,
        new_rows=new_rows, changed_rows=changed, store_rows=r[0], store_keys=r[1], versions=versions,
        store_min_day=r[2], store_max_end=r[3], source_first_day=src_first if src_first else "",
        days_mismatch=len(bad), size_mb=round(size, 1), seconds=round(time.time() - t0, 1)))
    return flags


def load_dims(conn, store, loaded_at):
    print("\n[справочники]")
    for d in DIMS:
        tb = fetch_arrow(conn, f"select {d['select']} from {d['name']}")
        store.write(d["name"], tb, f"dim_{loaded_at:%Y%m%d_%H%M%S}.parquet", loaded_at)
        print(f"    {d['name']}: {tb.num_rows} строк")


def depth(conn, store, loaded_at):
    for tbl, col in DEPTH_TABLES:
        try:
            lo, hi = fetch_rows(conn, f"select min({col}), max({col}) from {tbl}")[0]
            store.append_csv("depth.csv", dict(run=loaded_at, table=tbl, column=col, min_dt=lo, max_dt=hi, error=""))
        except Exception as e:
            store.append_csv("depth.csv", dict(run=loaded_at, table=tbl, column=col, min_dt="", max_dt="",
                                               error=str(e)[:300]))


def queue_check(store):
    """Очереди, где отвечает отдел поддержки, но их нет в списке (через представления wfm_store)."""
    import wfm_store
    df = wfm_store.query("select * from v_queue_check order by channel, answered_by_support desc")
    store.write_csv("queue_check.csv", df.to_dict("records"), list(df.columns))
    return df[df["status"] == "НЕТ В СПИСКЕ"]


ENTRY_WARN_SHARE = 0.005   # неизвестный вход даёт ≥ 0,5% сессий рабочих очередей канала — предупреждать


def entry_check(store):
    """Входы: первый проект сессий, дошедших до рабочих очередей (через представления wfm_store)."""
    import wfm_store
    df = wfm_store.query("select * from v_entry_check order by channel, reached_queue desc")
    store.write_csv("entry_check.csv", df.to_dict("records"), list(df.columns))
    return df[(df["status"] == "НЕИЗВЕСТНЫЙ ВХОД") & (df["share_of_queue"] >= ENTRY_WARN_SHARE)]


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Загрузка хранилища WFM из реплики источника")
    ap.add_argument("--only", default="", help="через запятую: какие таблицы грузить (по умолчанию все)")
    ap.add_argument("--no-checks", action="store_true", help="без сверки по дням (быстрее)")
    ap.add_argument("--overlap-days", type=float, default=S.OVERLAP_DAYS)
    a = ap.parse_args()

    t0 = time.time()
    loaded_at = dt.datetime.now().replace(microsecond=0)
    store = Store(store_root())
    only = {x.strip() for x in a.only.split(",") if x.strip()}
    tables = [t for t in TABLES if not only or t["name"] in only]
    print(f"Хранилище: {store.root}")
    print(f"Подключение к {S.SOURCE_DB['host']}:{S.SOURCE_DB['port']}/{S.SOURCE_DB['dbname']} ...")
    conn = connect_source()

    flags = {}
    load_dims(conn, store, loaded_at)
    for t in tables:
        f_ = load_table(conn, store, t, loaded_at, a.overlap_days, not a.no_checks)
        if f_:
            flags[t["name"]] = f_
    depth(conn, store, loaded_at)
    conn.close()

    print("\n[проверка рабочих очередей]")
    try:
        missing = queue_check(store)
        if len(missing):
            flags["рабочей очереди"] = [f"отдел поддержки отвечает в очередях не из списка: " + "; ".join(
                f"{r.title or r.project_id} ({r.channel}, {int(r.answered_by_support)})" for r in missing.itertuples())
                + " — добавить в wfm_settings.SUPPORT_QUEUES?"]
            print("    " + flags["рабочей очереди"][0])
        else:
            print("    все очереди, где отвечает отдел поддержки, есть в списке")
    except Exception as e:
        flags["рабочей очереди"] = [f"проверка не выполнена: {e}"]
        print(f"    НЕ УДАЛОСЬ: {e}")

    print("\n[проверка входов]")
    try:
        unknown = entry_check(store)
        if len(unknown):
            flags["входы"] = ["сессии очереди начинаются в проектах не из списков wfm_settings: " + "; ".join(
                f"{r.title or r.project_id} ({r.channel}, {r.share_of_queue:.1%} сессий рабочих очередей)" for r in unknown.itertuples())
                + " — новый вход? добавить в VOICE_ENTRY_PROJECTS / CHAT_SERVICE_PROJECTS"]
            print("    " + flags["входы"][0])
        else:
            print("    все сессии рабочих очередей приходят через известные входы")
    except Exception as e:
        flags["входы"] = [f"проверка не выполнена: {e}"]
        print(f"    НЕ УДАЛОСЬ: {e}")

    print(f"\nГотово за {time.time() - t0:.0f} с. Журнал: {store.log}")
    if flags:
        print("ВНИМАНИЕ — есть замечания (выше). Пришлите вывод консоли.")
    else:
        print("Проверки: замечаний нет.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nПрервано. Уже записанные дни сохранены; следующий запуск продолжит.")
        sys.exit(1)
    except Exception as e:
        print("\n" + error_line(e))
        if not is_db_timeout(e):
            print("Уже записанные дни сохранены; после исправления просто запустите ещё раз.")
        sys.exit(1)
