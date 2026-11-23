# -*- coding: utf-8 -*-
"""Потребность: люди в линии по 30 мин, голос и чат. Почта — часы месяца, без Erlang.

Запуск из wfm.bat, пункт «Потребность»: спросит квартал или год.
Файл: «Потребность квартал …» или «Потребность год …».
Без флага — показатели и период из inputs/need.txt, пустой run_id = опубликованный квартал на эти даты.
Объём и AHT — из forecast_interval выбранной версии. В расчёт идёт очередь «на оператора» (канал «Голос очереди» / «Чат очереди»).
Erlang A — тот же расчёт, что в report/week_report/erlang_calc.py.
Объём в прогнозе за 30 мин; калькулятор ждёт объём за час, поэтому в формулу идёт объём получаса × 2.
"""
import importlib.util
import os
import sys
from collections import defaultdict
from datetime import datetime, timedelta

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import wfm_store  # noqa: E402
from wfm_common import INPUTS_DIR, report_path, stamp  # noqa: E402

CFG_PATH = os.path.join(INPUTS_DIR, "need.txt")
KEYS = (
    "run_id", "period_from", "period_to", "sl",
    "voice_sl_sec", "voice_patience_sec", "voice_short_sec", "voice_k", "voice_aht_sec",
    "chat_sl_sec", "chat_patience_sec", "chat_k", "chat_concurrency", "chat_aht_sec",
    "planned", "absence", "breaks", "utz",
)
HOURS_WEEK = 42.0   # часов графика на человека в неделю, 2/2 (№122). В квоту, не в Erlang.
ABSENCE_NAME = "Отсутствия_2025-2026.xlsx"
ABSENCE_FROM = datetime(2025, 1, 1).date()
ABSENCE_TO = datetime(2026, 10, 31).date()
ABSENCE_DEPT = "Отдел поддержки"
# АД = отпуск без содержания (№142). Прогул — та же неявка, в файле 4 дня.
UNPLANNED_KINDS = {
    "Болезнь",
    "Отсутствие по невыясненным причинам",
    "Отпуск неоплачиваемый по разрешению работодателя",
    "Прогул",
}
MATERNITY_KINDS = {
    "Отпуск по беременности и родам",
    "Отпуск по уходу за ребенком",
}
# Янв–окт: 2025 и 2026. Ноя–дек: только 2025, в файле 2026 этих месяцев нет.
ABSENCE_YEARS = {m: (2025, 2026) if m <= 10 else (2025,) for m in range(1, 13)}
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
BAND = PatternFill("solid", fgColor="F3F6FB")
SECTION_FILL = PatternFill("solid", fgColor="1F3864")
SECTION_FONT = Font(bold=True, color="FFFFFF")
THIN = Border(
    left=Side(style="thin", color="D0D7E2"),
    right=Side(style="thin", color="D0D7E2"),
    top=Side(style="thin", color="D0D7E2"),
    bottom=Side(style="thin", color="D0D7E2"),
)


def _erlang():
    # Калькулятор Erlang A в исходной сборке жил отдельным модулем и в этот
    # репозиторий не входил. Сначала ищем локальную копию рядом со скриптами.
    candidates = [
        os.path.join(HERE, "erlang_calc.py"),
        os.path.abspath(os.path.join(HERE, "..", "..", "report", "week_report", "erlang_calc.py")),
    ]
    path = next((p for p in candidates if os.path.isfile(p)), None)
    if not path:
        sys.exit(
            "Не найден erlang_calc.py (Erlang A). Положите модуль в py/erlang_calc.py "
            "или укажите свой путь. В публичный репозиторий исходный калькулятор не входил."
        )
    spec = importlib.util.spec_from_file_location("erlang_calc", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    cache = {}
    orig = mod._sl_int

    def cached(n, lam, aht, theta, t, short):
        key = (n, round(lam, 8), round(aht, 4), round(theta, 10), t, short)
        hit = cache.get(key)
        if hit is None:
            hit = orig(n, lam, aht, theta, t, short)
            cache[key] = hit
        return hit

    mod._sl_int = cached
    return mod


def _num(text, name):
    if text is None or str(text).strip() == "":
        return None
    try:
        return float(str(text).strip().replace(",", "."))
    except ValueError:
        sys.exit(f"need.txt: «{name}» не число: {text}")


def _date_or_empty(text, name):
    text = (text or "").strip()
    if not text:
        return pd.NaT
    try:
        return pd.Timestamp(text)
    except Exception:
        sys.exit(f"need.txt: {name} не дата: {text}")


def load_cfg(path):
    if not os.path.exists(path):
        sys.exit(f"Нет файла показателей: {path}")
    raw = {}
    with open(path, encoding="utf-8-sig") as f:
        for line in f:
            body = line.split("#", 1)[0].strip()
            if not body or "=" not in body:
                continue
            k, v = body.split("=", 1)
            raw[k.strip()] = v.strip()
    missing = [k for k in KEYS if k not in raw]
    if missing:
        sys.exit("need.txt: нет строк: " + ", ".join(missing))
    cfg = {k: raw[k] for k in KEYS}
    for k in KEYS:
        if k in ("run_id", "period_from", "period_to"):
            continue
        cfg[k] = _num(raw[k], k)
    cfg["run_id"] = (cfg["run_id"] or "").strip()
    cfg["period_from"] = _date_or_empty(cfg["period_from"], "period_from")
    cfg["period_to"] = _date_or_empty(cfg["period_to"], "period_to")
    if pd.isna(cfg["period_from"]) != pd.isna(cfg["period_to"]):
        sys.exit("need.txt: period_from и period_to либо оба пустые, либо оба заполнены")
    if pd.notna(cfg["period_from"]) and cfg["period_to"] < cfg["period_from"]:
        sys.exit("need.txt: period_to раньше period_from")
    sl = cfg["sl"]
    if sl > 1:
        sl = sl / 100.0
    if not 0 < sl < 1:
        sys.exit("need.txt: sl должен быть между 0 и 1 (или 1…99, если писать проценты)")
    cfg["sl"] = sl
    for k in ("voice_k", "chat_k", "chat_concurrency", "utz"):
        if not cfg[k] or cfg[k] <= 0:
            sys.exit(f"need.txt: {k} должен быть больше 0")
    if not 0 <= cfg["breaks"] < 1:
        sys.exit("need.txt: breaks должен быть от 0 до 1, не включая 1")
    for k in ("planned", "absence"):
        if cfg[k] > 1:
            cfg[k] = cfg[k] / 100.0
        if not 0 <= cfg[k] < 1:
            sys.exit(f"need.txt: {k} должен быть от 0 до 1, не включая 1")
    if cfg["absence"] < cfg["planned"]:
        sys.exit("need.txt: таргет абсентеизма меньше планового")
    cfg["unplanned"] = cfg["absence"] - cfg["planned"]
    if cfg["utz"] > 1:
        sys.exit("need.txt: utz должен быть не больше 1")
    maternity = _num(raw.get("maternity", "0"), "maternity") or 0
    if maternity not in (0, 1):
        sys.exit("need.txt: maternity — 0 или 1")
    cfg["maternity"] = bool(maternity)
    return cfg


def pick_run(kind, published_only):
    """Последняя версия нужного вида. published_only — только статус «опубликован»."""
    runs = wfm_store.query(
        "select run_id, created_at, period_from, period_to, status from forecast_run "
        f"where kind = '{kind}'")
    if runs.empty:
        return None
    runs["created_at"] = pd.to_datetime(runs["created_at"])
    runs["period_from"] = pd.to_datetime(runs["period_from"])
    runs["period_to"] = pd.to_datetime(runs["period_to"])
    if published_only:
        runs = runs[runs["status"] == "опубликован"]
        if runs.empty:
            return None
    return runs.sort_values("created_at").iloc[-1]


def apply_horizon(cfg, horizon):
    """Меню: квартал или год. Без горизонта — прежнее правило need.txt."""
    if horizon is None:
        cfg["run_id"] = resolve_run(cfg)
        return
    if cfg["run_id"]:
        print("Номер в need.txt для этого запуска не используется: горизонт выбран в меню.")
    if horizon == "quarter":
        row = pick_run("модель: квартал", True)
        if row is None:
            sys.exit("Нет опубликованного квартального прогноза")
        print(f"Квартал: последний опубликованный {row['run_id']}, "
              f"{pd.Timestamp(row['period_from']):%d.%m.%Y}–{pd.Timestamp(row['period_to']):%d.%m.%Y}")
    elif horizon == "year":
        row = pick_run("модель: год", True)
        if row is None:
            row = pick_run("модель: год", False)
            if row is None:
                sys.exit("Нет прогноза на год")
            print(f"Опубликованного года нет. Беру последний черновик {row['run_id']}")
        else:
            print(f"Год: последний опубликованный {row['run_id']}")
    else:
        sys.exit(f"Неизвестный горизонт: {horizon}")
    cfg["run_id"] = str(row["run_id"])
    cfg["period_from"] = pd.Timestamp(row["period_from"])
    cfg["period_to"] = pd.Timestamp(row["period_to"])


def resolve_run(cfg):
    """Пустой run_id — последний опубликованный квартал, который закрывает период."""
    pinned = cfg["run_id"].replace("'", "")
    if pinned:
        print(f"Прогноз указан в файле: {pinned}")
        return pinned
    if pd.isna(cfg["period_from"]) or pd.isna(cfg["period_to"]):
        row = pick_run("модель: квартал", True)
        if row is None:
            sys.exit("Нет опубликованного квартального прогноза")
        cfg["period_from"] = pd.Timestamp(row["period_from"])
        cfg["period_to"] = pd.Timestamp(row["period_to"])
        print(f"Квартал: последний опубликованный {row['run_id']}, "
              f"{cfg['period_from']:%d.%m.%Y}–{cfg['period_to']:%d.%m.%Y}")
        return str(row["run_id"])
    runs = wfm_store.query(
        "select run_id, created_at, period_from, period_to from forecast_run "
        "where status = 'опубликован' and kind = 'модель: квартал'")
    if runs.empty:
        sys.exit("Нет опубликованного квартального прогноза")
    runs["created_at"] = pd.to_datetime(runs["created_at"])
    runs["period_from"] = pd.to_datetime(runs["period_from"])
    runs["period_to"] = pd.to_datetime(runs["period_to"])
    cover = runs[(runs["period_from"] <= cfg["period_from"]) & (runs["period_to"] >= cfg["period_to"])]
    if cover.empty:
        sys.exit("Нет опубликованного квартала на даты из need.txt")
    rid = str(cover.sort_values("created_at").iloc[-1]["run_id"])
    print(f"Прогноз: последний опубликованный квартал {rid}")
    return rid


def load_forecast(cfg):
    rid = cfg["run_id"].replace("'", "")
    df = wfm_store.query(
        "select interval_start, channel, value from forecast_interval "
        f"where run_id = '{rid}'")
    if df.empty:
        sys.exit(f"В хранилище нет прогноза {cfg['run_id']}")
    df["interval_start"] = pd.to_datetime(df["interval_start"])
    end = cfg["period_to"] + pd.Timedelta(days=1)
    df = df[(df["interval_start"] >= cfg["period_from"]) & (df["interval_start"] < end)]
    wide = df.pivot_table(index="interval_start", columns="channel", values="value", aggfunc="sum")
    need_cols = ["Голос очереди", "Чат очереди", "AHT Голос очереди, с", "AHT Чат очереди, с", "Почта"]
    missing = [c for c in need_cols if c not in wide.columns]
    if missing:
        sys.exit("В прогнозе нет каналов: " + ", ".join(missing))
    out = wide[need_cols].dropna()
    if out.empty:
        sys.exit("В выбранном периоде нет интервалов прогноза")
    return out


def _aht(series, override):
    if override:
        return float(override)
    med = float(series.median())
    if abs(float(series.max()) - float(series.min())) > 0.05:
        print(f"  AHT в прогнозе не постоянный: {series.min():.1f}…{series.max():.1f} с, по интервалам берётся своё значение")
        return None
    return med


def _as_date(value):
    if isinstance(value, datetime):
        return value.date()
    return datetime.strptime(str(value)[:10], "%d.%m.%Y").date()


def absence_profile(level):
    """12 долей. Среднее = уровень года. Форма — календарные дни очереди, не сам процент дней.

    Декрет и уход за ребёнком в долю не входят: это головы, их потом вычитают из пула.
    """
    path = os.path.join(INPUTS_DIR, ABSENCE_NAME)
    if not os.path.exists(path):
        sys.exit(f"Нет файла отсутствий: {path}")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sh = wb.active
    missed = defaultdict(set)
    maternity = defaultdict(set)
    for i, row in enumerate(sh.iter_rows(values_only=True), 1):
        if i < 6:
            continue
        dept, name, kind, _cal, _work, d1, d2 = (list(row) + [None] * 7)[:7]
        if dept != ABSENCE_DEPT or not kind or not name or not d1 or not d2:
            continue
        kind = str(kind)
        if kind not in UNPLANNED_KINDS and kind not in MATERNITY_KINDS:
            continue
        a = max(_as_date(d1), ABSENCE_FROM)
        b = min(_as_date(d2), ABSENCE_TO)
        if b < a:
            continue
        cur = a
        while cur <= b:
            if kind in UNPLANNED_KINDS:
                missed[(cur.year, cur.month)].add((str(name), cur))
            else:
                maternity[(cur.year, cur.month)].add(str(name))
            cur += timedelta(days=1)
    wb.close()

    raw = []
    for month in range(1, 13):
        years = ABSENCE_YEARS[month]
        counts = [len(missed[(year, month)]) for year in years]
        raw.append(sum(counts) / len(years))
    mean_days = sum(raw) / 12.0
    if mean_days <= 0:
        sys.exit("В файле отсутствий нет календарных дней внеплана")
    rows = []
    for month, days in enumerate(raw, 1):
        years = ABSENCE_YEARS[month]
        index = days / mean_days
        share = level * index
        if share >= 1:
            sys.exit(f"Внеплан месяца {month} получился {share:.0%}. Уровень года слишком большой.")
        heads = [len(maternity[(year, month)]) for year in years]
        rows.append({
            "Месяц": month,
            "Лет в индексе": len(years),
            "Кален. дни, среднее": round(days, 2),
            "Индекс": index,
            "Внеплан": share,
            "Декрет: голов, факт": round(sum(heads) / len(heads), 1),
        })
    return pd.DataFrame(rows).set_index("Месяц")


def mail_aht(run_id):
    """Медиана дневных средних: от назначения до ответа, 8 недель до data_to прогноза."""
    rid = str(run_id).replace("'", "")
    meta = wfm_store.query(f"select data_to from forecast_run where run_id = '{rid}'")
    if meta.empty or pd.isna(meta.iloc[0]["data_to"]):
        sys.exit(f"У прогноза {rid} нет даты «данные по», AHT почты не к чему мерить")
    hi = pd.Timestamp(meta.iloc[0]["data_to"]).normalize()
    lo = hi - pd.Timedelta(weeks=8)
    days = wfm_store.query(f"""
        select cast(coalesce(i.unblock_time, i.start_time) as date) as day,
               count(*) as n,
               sum(extract(epoch from (i.first_oper_answer_time - i.oper_distribute_time))) as sec
        from chat_session_iterations i
        join support_queues g on g.project_id = i.project_id and g.channel = 'Почта'
        where i.first_oper_answer_time is not null
          and i.oper_distribute_time is not null
          and i.first_oper_answer_time >= i.oper_distribute_time
          and cast(coalesce(i.unblock_time, i.start_time) as date) > date '{lo:%Y-%m-%d}'
          and cast(coalesce(i.unblock_time, i.start_time) as date) <= date '{hi:%Y-%m-%d}'
        group by 1
    """)
    if days.empty or float(days["n"].sum()) <= 0:
        sys.exit(f"Нет отвеченных писем за {lo:%d.%m.%Y}–{hi:%d.%m.%Y}")
    aht = float((days["sec"] / days["n"]).median())
    if not aht or aht <= 0:
        sys.exit("AHT почты получился нулевым")
    return {"aht": aht, "lo": lo, "hi": hi, "letters": int(days["n"].sum()), "days": int(len(days))}


def monthly_people(out, cfg, mail_volume, mail_aht_sec):
    """Люди для квоты. Чел.-ч линии не меняются. Отпуск сюда не входит. Почта — часы месяца."""
    profile = absence_profile(cfg["unplanned"])
    df = out.copy()
    df["Интервал"] = pd.to_datetime(df["Интервал"])
    mail_volume = mail_volume.copy()
    mail_volume.index = pd.to_datetime(mail_volume.index)
    hpw = HOURS_WEEK * (1 - cfg["breaks"]) * cfg["utz"]
    rows = []
    for month, g in df.groupby(df["Интервал"].dt.to_period("M")):
        days = int(g["Интервал"].dt.normalize().nunique())
        weeks = days / 7.0
        vh = float(g["Голос: в линии"].sum()) * 0.5
        ch = float(g["Чат: в линии"].sum()) * 0.5
        letters = float(mail_volume[mail_volume.index.to_period("M") == month].sum())

        def people(hours):
            return hours / (hpw * weeks)

        pv, pc = people(vh), people(ch)
        pm = people(letters * mail_aht_sec / 3600.0)
        share = float(profile.loc[month.month, "Внеплан"])
        total = cfg["planned"] + share
        if total >= 1:
            sys.exit(f"Абсентеизм месяца {month} получился {total:.0%}. Таргет слишком большой.")
        k = 1 - total
        rows.append({
            "Месяц": str(month),
            "Дней": days,
            "Голос: чел.-ч": round(vh, 1),
            "Чат: чел.-ч": round(ch, 1),
            "Почта: писем": round(letters, 1),
            "Почта: чел.-ч": round(letters * mail_aht_sec / 3600.0, 1),
            "Голос: люди на линию": round(pv, 1),
            "Чат: люди на линию": round(pc, 1),
            "Почта: люди": round(pm, 1),
            "Всего: люди на линию": round(pv + pc + pm, 1),
            "Внеплан": share,
            "Абсентеизм": total,
            "Голос: люди с абсентеизмом": round(pv / k, 1),
            "Чат: люди с абсентеизмом": round(pc / k, 1),
            "Почта: люди с абсентеизмом": round(pm / k, 1),
            "Всего: люди с абсентеизмом": round((pv + pc + pm) / k, 1),
        })
        if cfg["maternity"]:
            rows[-1]["Декрет: голов, факт"] = float(profile.loc[month.month, "Декрет: голов, факт"])
    return pd.DataFrame(rows), hpw, profile


def _month_label(value):
    """Как лист «Месяцы» в книге года: 01.2027."""
    period = pd.Period(str(value), freq="M")
    return f"{period.month:02d}.{period.year}"


def _half_hours(out):
    """Как лист «30 мин» прогноза: дата и слот отдельно."""
    df = out.copy()
    ts = pd.to_datetime(df["Интервал"])
    df.insert(0, "Дата", [t.date() for t in ts])
    end = ts + pd.Timedelta(minutes=30)
    df["Интервал"] = [f"{a:%H:%M}–{b:%H:%M}" for a, b in zip(ts, end)]
    return df


def _present(months, profile, show_maternity):
    """Месяц как 01.2027. Числа те же. Декрет — только если включён в need.txt."""
    mon = months.copy()
    mon["Месяц"] = mon["Месяц"].map(_month_label)
    years = [pd.Period(str(v), freq="M").year for v in months["Месяц"]]
    year = max(set(years), key=years.count)
    prof = profile.reset_index()
    prof["Месяц"] = prof["Месяц"].map(lambda n: f"{int(n):02d}.{year}")
    prof = prof.rename(columns={"Кален. дни, среднее": "Календарные дни"})
    if show_maternity:
        mon = mon.rename(columns={"Декрет: голов, факт": "Декрет, человек"})
        prof = prof.rename(columns={"Декрет: голов, факт": "Декрет, человек"})
    else:
        mon = mon.drop(columns=["Декрет: голов, факт"], errors="ignore")
        prof = prof.drop(columns=["Декрет: голов, факт"], errors="ignore")
    mon["Ставки в графике"] = mon["Всего: люди с абсентеизмом"]
    return mon, prof


def _paint_header(ws, band=False):
    widths = {
        "Месяц": 12, "Дата": 14, "Дней": 10, "Лет в индексе": 16, "Календарные дни": 18,
        "Индекс": 12, "Внеплан": 12, "Абсентеизм": 14, "Декрет, человек": 18,
        "Голос: чел.-ч": 16, "Чат: чел.-ч": 16, "Почта: писем": 16, "Почта: чел.-ч": 16,
        "Голос: люди на линию": 24, "Чат: люди на линию": 22, "Почта: люди": 16,
        "Всего: люди на линию": 24,
        "Голос: люди с абсентеизмом": 28, "Чат: люди с абсентеизмом": 26,
        "Почта: люди с абсентеизмом": 28, "Всего: люди с абсентеизмом": 28,
        "Ставки в графике": 20,
        "Интервал": 14, "Голос: обращений": 18, "Голос: AHT, с": 16, "Голос: в линии": 16,
        "Чат: обращений": 16, "Чат: AHT, с": 14, "Чат: в линии": 14,
    }
    formats = {
        "Внеплан": "0.00%",
        "Абсентеизм": "0.00%",
        "Индекс": "0.000",
        "Календарные дни": "#,##0.0",
        "Декрет, человек": "0.0",
        "Дней": "0",
        "Лет в индексе": "0",
        "Дата": "DD.MM.YYYY",
        "Голос: обращений": "#,##0.000",
        "Чат: обращений": "#,##0.000",
        "Голос: AHT, с": "#,##0.0",
        "Чат: AHT, с": "#,##0.0",
        "Голос: в линии": "#,##0.00",
        "Чат: в линии": "#,##0.00",
    }
    one_dec = {
        "Голос: чел.-ч", "Чат: чел.-ч", "Почта: писем", "Почта: чел.-ч",
        "Голос: люди на линию", "Чат: люди на линию", "Почта: люди", "Всего: люди на линию",
        "Голос: люди с абсентеизмом", "Чат: люди с абсентеизмом", "Почта: люди с абсентеизмом",
        "Всего: люди с абсентеизмом",
        "Ставки в графике",
    }
    headers = [c.value for c in ws[1]]
    for cell in ws[1]:
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = THIN
        ws.column_dimensions[get_column_letter(cell.column)].width = widths.get(cell.value, 14)
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        stripe = band and row[0].row % 2 == 0
        for cell in row:
            header = headers[cell.column - 1]
            fmt = formats.get(header) or ("#,##0.0" if header in one_dec else None)
            if fmt and isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                cell.number_format = fmt
            elif header == "Дата" and cell.value is not None:
                cell.number_format = "DD.MM.YYYY"
            if stripe:
                cell.fill = BAND
            if band:
                cell.border = THIN
                cell.alignment = Alignment(vertical="center")


def _fmt(x, digits=0):
    return f"{float(x):,.{digits}f}".replace(",", " ").replace(".", ",")


def _pct(x):
    return f"{float(x):.1%}".replace(".", ",")


def _how(cfg, voice_aht, chat_aht, voice_n, chat_n, voice_h, chat_h, months, mail):
    """Расчёт этого прогона словами, с его числами."""
    quota = HOURS_WEEK * (1 - cfg["breaks"]) * cfg["utz"]
    aht_v = _fmt(voice_aht, 1) + " с" if voice_aht else "AHT своего получаса"
    aht_c = _fmt(chat_aht, 1) + " с" if chat_aht else "AHT своего получаса"
    lines = [
        "Как считали этот прогон.",
        (f"Голос. Берём обращения за полчаса и умножаем на 2 — калькулятору нужен час. "
         f"Дальше цель {_pct(cfg['sl'])} за {cfg['voice_sl_sec']:.0f} с и AHT {aht_v}. "
         f"Калькулятор возвращает людей в линии. Человеко-часы = сумма этих людей × 0,5 = {_fmt(voice_h, 1)} "
         f"на {_fmt(voice_n)} обращений."),
        (f"Чат. То же: полчаса × 2, цель {_pct(cfg['sl'])} за {cfg['chat_sl_sec']:.0f} с, AHT {aht_c}, "
         f"один оператор ведёт {cfg['chat_concurrency']:.0f} чата. "
         f"Человеко-часы = {_fmt(chat_h, 1)} на {_fmt(chat_n)} обращений."),
        (f"Люди на линию. Берём человеко-часы и делим на {_fmt(quota, 2)} ч × (дни месяца / 7). "
         f"{_fmt(HOURS_WEEK)} ч графика × (1 − {_pct(cfg['breaks'])}) × {_pct(cfg['utz'])} = {_fmt(quota, 2)} ч линии на человека в неделю."),
    ]
    if len(months):
        r = months.iloc[0]
        lines.append(
            f"Пример, {_month_label(r['Месяц'])}, голос. {_fmt(r['Голос: чел.-ч'], 1)} чел.-ч делим на "
            f"({_fmt(quota, 2)} × {_fmt(r['Дней'])} / 7) = {_fmt(r['Голос: люди на линию'], 1)} человек на линию. "
            f"С абсентеизмом {_pct(r['Абсентеизм'])} ({_pct(cfg['planned'])} план + {_pct(r['Внеплан'])} внеплан): "
            f"{_fmt(r['Голос: люди на линию'], 1)} / (1 − {_pct(r['Абсентеизм'])}) = "
            f"{_fmt(r['Голос: люди с абсентеизмом'], 1)}."
        )
        lines.append(
            f"Тот же месяц, чат. {_fmt(r['Чат: чел.-ч'], 1)} чел.-ч делим на "
            f"({_fmt(quota, 2)} × {_fmt(r['Дней'])} / 7) = {_fmt(r['Чат: люди на линию'], 1)} на линию. "
            f"С абсентеизмом: {_fmt(r['Чат: люди на линию'], 1)} / (1 − {_pct(r['Абсентеизм'])}) = "
            f"{_fmt(r['Чат: люди с абсентеизмом'], 1)}."
        )
    lines.append(
        f"Почта. Erlang нет, на лист «30 мин» не пишется. Письма месяца × AHT {_fmt(mail['aht'], 1)} с = часы. "
        f"AHT — медиана дневных средних с {mail['lo']:%d.%m.%Y} по {mail['hi']:%d.%m.%Y}: "
        f"{_fmt(mail['letters'])} писем за {mail['days']} дней. Делитель тот же, {_fmt(quota, 2)} ч."
    )
    if len(months):
        r = months.iloc[0]
        lines.append(
            f"Пример, {_month_label(r['Месяц'])}, почта. {_fmt(r['Почта: писем'], 1)} писем дают "
            f"{_fmt(r['Почта: чел.-ч'], 1)} ч. Делим на ({_fmt(quota, 2)} × {_fmt(r['Дней'])} / 7) = "
            f"{_fmt(r['Почта: люди'], 1)}. С абсентеизмом: {_fmt(r['Почта: люди'], 1)} / "
            f"(1 − {_pct(r['Абсентеизм'])}) = {_fmt(r['Почта: люди с абсентеизмом'], 1)}."
        )
    lines.append(
        "«Всего» и «Ставки в графике» — голос, чат и почта. "
        "«Ставки в графике» равны «Всего: люди с абсентеизмом»: столько ставок должно быть на работе в этом месяце."
    )
    lines.append(
        f"Плановые {_pct(cfg['planned'])} входят в каждый месяц, включая февраль. "
        f"Внеплан года {_pct(cfg['unplanned'])} — это таргет {_pct(cfg['absence'])} минус план. "
        "Календарь отпусков этот процент не меняет. Декрет не вычитается."
    )
    return lines


def _parameters(ws, cfg, voice_aht, chat_aht, voice_n, chat_n, voice_h, chat_h, hpw, status, months, mail, mail_n, mail_h):
    """Лист «Параметры»: подписи по-русски, проценты как проценты."""
    aht_v = round(voice_aht, 1) if voice_aht else "своё значение в каждом получасе"
    aht_c = round(chat_aht, 1) if chat_aht else "своё значение в каждом получасе"
    blocks = [
        ("Прогноз", [
            ("Номер", cfg["run_id"], None),
            ("Что это", status, None),
            ("Период", f"{cfg['period_from']:%d.%m.%Y}–{cfg['period_to']:%d.%m.%Y}", None),
        ]),
        ("Очередь", [
            ("SL", cfg["sl"], "0%"),
            ("Голос: ответ за, с", cfg["voice_sl_sec"], "0"),
            ("Голос: терпение, с", cfg["voice_patience_sec"], "#,##0"),
            ("Голос: короткие потерянные, с", cfg["voice_short_sec"], "0"),
            ("Голос: K", cfg["voice_k"], "0.00"),
            ("Голос: AHT, с", aht_v, "0.0" if isinstance(aht_v, float) else None),
            ("Чат: ответ за, с", cfg["chat_sl_sec"], "0"),
            ("Чат: терпение, с", cfg["chat_patience_sec"], "#,##0"),
            ("Чат: K", cfg["chat_k"], "0.00"),
            ("Чат: параллельность", cfg["chat_concurrency"], "0"),
            ("Чат: AHT, с", aht_c, "0.0" if isinstance(aht_c, float) else None),
            ("Почта: AHT, с", round(mail["aht"], 1), "0.0"),
            ("Почта: окно AHT", f"{mail['lo']:%d.%m.%Y}–{mail['hi']:%d.%m.%Y}, медиана дневных средних, {mail['letters']} писем", None),
        ]),
        ("Люди", [
            ("Часов графика на человека в неделю", HOURS_WEEK, "0"),
            ("Перерывы", cfg["breaks"], "0.0%"),
            ("UTZ", cfg["utz"], "0.0%"),
            ("Часов линии на человека в неделю", round(hpw, 2), "0.00"),
            ("Плановый абсентеизм", cfg["planned"], "0%"),
            ("Таргет абсентеизма на год", cfg["absence"], "0%"),
            ("Внеплан, уровень года", cfg["unplanned"], "0%"),
            ("Доля месяца", "план + внеплан месяца. Внеплан месяца = уровень года × индекс. Среднее 12 месяцев = таргет", None),
            ("Откуда индекс", f"{ABSENCE_NAME}, календарные дни очереди {ABSENCE_FROM:%d.%m.%Y}–{ABSENCE_TO:%d.%m.%Y}. Болезнь, неявка, отпуск без содержания", None),
            ("Отпуск в календаре", f"{cfg['planned']:.0%} уже в численности каждого месяца. Календарь 7+14+7 этот процент не меняет", None),
        ] + ([("Декрет", "колонка факта на листах. Из людей не вычтен", None)] if cfg.get("maternity") else []),
        ),
        ("Итог", [
            ("Голос: обращений", round(voice_n, 1), "#,##0.0"),
            ("Чат: обращений", round(chat_n, 1), "#,##0.0"),
            ("Голос: в линии, чел.-ч", round(voice_h, 1), "#,##0.0"),
            ("Чат: в линии, чел.-ч", round(chat_h, 1), "#,##0.0"),
            ("Почта: писем", round(mail_n, 1), "#,##0.0"),
            ("Почта: чел.-ч", round(mail_h, 1), "#,##0.0"),
        ]),
    ]
    notes = _how(cfg, voice_aht, chat_aht, voice_n, chat_n, voice_h, chat_h, months, mail)
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 88
    row = 1
    label_font = Font(color="1F3864")
    for title, items in blocks:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        head = ws.cell(row, 1, title)
        head.fill = SECTION_FILL
        head.font = SECTION_FONT
        head.alignment = Alignment(vertical="center", indent=1)
        ws.cell(row, 2).fill = SECTION_FILL
        ws.row_dimensions[row].height = 22
        row += 1
        for name, value, fmt in items:
            left = ws.cell(row, 1, name)
            left.font = label_font
            left.alignment = Alignment(vertical="center", indent=1)
            left.border = THIN
            right = ws.cell(row, 2, value)
            right.alignment = Alignment(vertical="center", wrap_text=True)
            right.border = THIN
            if fmt and isinstance(value, (int, float)) and not isinstance(value, bool):
                right.number_format = fmt
            if isinstance(value, str) and len(value) > 60:
                ws.row_dimensions[row].height = 32
            row += 1
        row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    head = ws.cell(row, 1, "Как считали")
    head.fill = SECTION_FILL
    head.font = SECTION_FONT
    head.alignment = Alignment(vertical="center", indent=1)
    ws.cell(row, 2).fill = SECTION_FILL
    ws.row_dimensions[row].height = 22
    for line in notes:
        row += 1
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        cell = ws.cell(row, 1, line)
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[row].height = 48 if len(line) > 120 else 22
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False


def write_book(path, out, months, profile, cfg, voice_aht, chat_aht, voice_n, chat_n, voice_h, chat_h, hpw, status, mail, mail_n, mail_h):
    mon, prof = _present(months, profile, cfg.get("maternity"))
    slots = _half_hours(out)
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        mon.to_excel(xw, sheet_name="Месяцы", index=False)
        prof.to_excel(xw, sheet_name="Внеплан", index=False)
        slots.to_excel(xw, sheet_name="30 мин", index=False)
        xw.book.create_sheet("Параметры")
        _paint_header(xw.book["Месяцы"], band=True)
        _paint_header(xw.book["Внеплан"], band=True)
        _paint_header(xw.book["30 мин"], band=False)
        _parameters(
            xw.book["Параметры"], cfg, voice_aht, chat_aht,
            voice_n, chat_n, voice_h, chat_h, hpw, status, months, mail, mail_n, mail_h,
        )


def line_people(ec, ch, calls_30, aht, target):
    """Люди в линии на этот получас. calls_30 — обращения за 30 мин; в Erlang идёт за час."""
    if calls_30 <= 0 or not aht or aht <= 0:
        return 0.0
    return float(ch.need(calls_30 * 2.0, aht, target))


def main(horizon=None):
    cfg = load_cfg(CFG_PATH)
    apply_horizon(cfg, horizon)
    print(f"Показатели: {CFG_PATH}")
    print(f"Прогноз {cfg['run_id']}, {cfg['period_from']:%d.%m.%Y}–{cfg['period_to']:%d.%m.%Y}, SL {cfg['sl']:.0%}")
    fc = load_forecast(cfg)
    mail = mail_aht(cfg["run_id"])
    print(
        f"AHT почты {mail['aht']:.1f} с, окно {mail['lo']:%d.%m.%Y}–{mail['hi']:%d.%m.%Y}, "
        f"писем {mail['letters']}, дней {mail['days']}"
    )
    voice_aht = _aht(fc["AHT Голос очереди, с"], cfg["voice_aht_sec"])
    chat_aht = _aht(fc["AHT Чат очереди, с"], cfg["chat_aht_sec"])
    if voice_aht:
        print(f"AHT голос {voice_aht:.1f} с")
    if chat_aht:
        print(f"AHT чат {chat_aht:.1f} с")
    ec = _erlang()
    voice = ec.Channel("Голос", cfg["voice_sl_sec"], cfg["voice_patience_sec"], cfg["voice_short_sec"], cfg["voice_k"], 1)
    chat = ec.Channel("Чат", cfg["chat_sl_sec"], cfg["chat_patience_sec"], 0, cfg["chat_k"], cfg["chat_concurrency"])
    rows = []
    capped = 0
    for ts, r in fc.iterrows():
        va = voice_aht if voice_aht else float(r["AHT Голос очереди, с"])
        ca = chat_aht if chat_aht else float(r["AHT Чат очереди, с"])
        nv = line_people(ec, voice, float(r["Голос очереди"]), va, cfg["sl"])
        nc = line_people(ec, chat, float(r["Чат очереди"]), ca, cfg["sl"])
        hi_v = max(2.0, float(r["Голос очереди"]) * 2 * va / 3600 / 1 * 3 + 3) if va else 0
        hi_c = max(2.0, float(r["Чат очереди"]) * 2 * ca / 3600 / cfg["chat_concurrency"] * 3 + 3) if ca else 0
        if (float(r["Голос очереди"]) > 0 and nv >= hi_v - 1e-6) or (float(r["Чат очереди"]) > 0 and nc >= hi_c - 1e-6):
            capped += 1
        rows.append({
            "Интервал": ts,
            "Голос: обращений": round(float(r["Голос очереди"]), 3),
            "Голос: AHT, с": round(va, 1),
            "Голос: в линии": round(nv, 2),
            "Чат: обращений": round(float(r["Чат очереди"]), 3),
            "Чат: AHT, с": round(ca, 1),
            "Чат: в линии": round(nc, 2),
        })
    out = pd.DataFrame(rows)
    voice_h = float(out["Голос: в линии"].sum()) * 0.5
    chat_h = float(out["Чат: в линии"].sum()) * 0.5
    voice_n = float(out["Голос: обращений"].sum())
    chat_n = float(out["Чат: обращений"].sum())
    print(f"Обращений: голос {voice_n:.1f}, чат {chat_n:.1f}")
    print(f"В линии, чел.-ч: голос {voice_h:.1f}, чат {chat_h:.1f}, всего {voice_h + chat_h:.1f}")
    mail_n = float(fc["Почта"].sum())
    mail_h = mail_n * mail["aht"] / 3600.0
    print(f"Почта: писем {mail_n:.1f}, чел.-ч {mail_h:.1f}")
    months, hpw, profile = monthly_people(out, cfg, fc["Почта"], mail["aht"])
    lo, hi = float(profile["Внеплан"].min()), float(profile["Внеплан"].max())
    print(f"Квота: {hpw:.2f} ч линии на человека в неделю.")
    tail = " Колонка декрета включена, из людей не вычтена." if cfg["maternity"] else ""
    print(
        f"Абсентеизм: план {cfg['planned']:.0%} каждый месяц, таргет года {cfg['absence']:.0%}, "
        f"внеплан {cfg['unplanned']:.0%}. Доля месяца {cfg['planned'] + lo:.2%}…{cfg['planned'] + hi:.2%}. "
        f"Делит людей, не часы.{tail}"
    )
    if capped:
        print(f"ВНИМАНИЕ: в {capped} интервалах поиск упёрся в потолок формулы. Пришлите этот вывод.")
    if horizon == "year":
        status = "черновик года, не опубликован"
    elif horizon == "quarter":
        status = "опубликованный квартал"
    else:
        status = "номер из need.txt"
    kind = "год" if horizon == "year" else "квартал"
    path = report_path(f"Потребность {kind} {stamp()}.xlsx")
    write_book(
        path, out, months, profile, cfg, voice_aht, chat_aht,
        voice_n, chat_n, voice_h, chat_h, hpw, status, mail, mail_n, mail_h,
    )
    print(f"Файл: {path}")
    return path


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--quarter", action="store_true")
    parser.add_argument("--year", action="store_true")
    args = parser.parse_args()
    if args.quarter and args.year:
        sys.exit("Укажите один горизонт: --quarter или --year")
    main("quarter" if args.quarter else "year" if args.year else None)
