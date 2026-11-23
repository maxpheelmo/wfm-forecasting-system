# -*- coding: utf-8 -*-
"""Год: CR даёт вход, доля 6 недель переводит его в очередь. Где CR нет — обычная модель.

Кнопка — --open-year, сама не публикует. Год в кнопку не зашит.
Запуск без --open-year — старая проверка апреля–октября 2026, не прогноз года.
"""
import datetime as dt
import json
import os

import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import aht as A
import forecast_core as F
import wfm_forecast as W
import wfm_settings as S
import wfm_store
from wfm_common import report_path, stamp

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CR_PATH = os.path.join(ROOT, "inputs", "CR 2024-2026.xlsx")
PLAN_PATH = os.path.join(ROOT, "inputs", "Бизнес-план_КЦ_(обновлен).xlsx")
MAU_PATH = os.path.join(ROOT, "inputs", "MAU план_факт .xlsx")
VOICE_REACH = 0.87          # только старая проверка апреля–октября 2026. Кнопка года это не использует.
BASELINE_CHAT_BOT = 0.53    # лист «Цель 1. Чат»: baseline. В траектории столбцы с августа.
PLAN_MONTH1 = pd.Timestamp("2026-08-01")   # №142: месяц 1 плана = август 2026


def _num(v):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def read_cr():
    """Штуки и база из «Нагрузка 24-26». CR = штуки / база того же месяца."""
    df = pd.read_excel(CR_PATH, sheet_name="Нагрузка 24-26", header=None)
    header = df.iloc[0, 1:]
    voice = df.iloc[1, 1:]
    chat = df.iloc[2, 1:]
    base = df.iloc[4, 1:]
    rows = []
    for i, raw in enumerate(header):
        if pd.isna(raw):
            continue
        month = pd.Timestamp(raw).normalize()
        b, vv, cc = _num(base.iloc[i]), _num(voice.iloc[i]), _num(chat.iloc[i])
        if b is None or b <= 0 or vv is None or cc is None:
            continue
        rows.append(dict(month=month, base=b, voice_cell=vv, chat_cell=cc,
                         voice_cr=vv / b, chat_cr=cc / b))
    return pd.DataFrame(rows).set_index("month").sort_index()


def read_scenario2():
    """Доля чатов, закрытых ботом. Сценарий 2, без LLM. Месяц 1 = август 2026."""
    df = pd.read_excel(PLAN_PATH, sheet_name="Цель 1. Чат", header=None)
    hit = df.index[df.iloc[:, 1].astype(str).str.startswith("Сценарий 2")]
    if len(hit) != 1:
        raise RuntimeError("На листе «Цель 1. Чат» не одна строка сценария 2")
    row = df.loc[hit[0]]
    # заголовок месяцев — строкой выше
    hdr = df.loc[hit[0] - 1]
    out = {}
    month = PLAN_MONTH1
    for c in range(2, len(row)):
        val = _num(row.iloc[c])
        label = hdr.iloc[c]
        if val is None or pd.isna(label):
            continue
        out[month] = val
        month = month + pd.offsets.MonthBegin(1)
    if out.get(PLAN_MONTH1) != BASELINE_CHAT_BOT:
        raise RuntimeError(f"Август сценария 2 = {out.get(PLAN_MONTH1)}, ждали {BASELINE_CHAT_BOT}")
    return out


def read_mau_plan():
    """План MAU, только месяцы, которые есть в файле. Пустой месяц не заполняем."""
    df = pd.read_excel(MAU_PATH, sheet_name=0, header=None)
    # блок «ПЛАН»: строка заголовка с датами, ниже сегменты, последняя строка — итог
    plan_at = df.index[df.iloc[:, 0].astype(str).str.strip().str.lower().eq("channel_name")]
    # вторая такая строка — план (первая — факт)
    if len(plan_at) < 2:
        raise RuntimeError("В MAU нет второго блока channel_name (план)")
    hdr_i = int(plan_at[-1])
    header = df.iloc[hdr_i]
    months = {}
    for c in range(5, header.shape[0]):
        raw = header.iloc[c]
        if pd.isna(raw):
            continue
        months[c] = pd.Timestamp(raw).normalize()
    body = df.iloc[hdr_i + 1:]
    # итог — строка без названия канала, с числами
    total = None
    seg_sum = {c: 0.0 for c in months}
    for _, row in body.iterrows():
        name = row.iloc[0]
        vals = {c: _num(row.iloc[c]) for c in months}
        if all(v is None for v in vals.values()):
            continue
        if pd.isna(name) or str(name).strip() == "":
            total = {months[c]: vals[c] for c in months}
            continue
        for c, v in vals.items():
            if v is not None:
                seg_sum[c] += v
    seg = {months[c]: seg_sum[c] for c in months}
    if total is None:
        raise RuntimeError("В плане MAU нет строки итога")
    for m, v in total.items():
        if v is None or abs(v - seg[m]) > 1:
            raise RuntimeError(f"Итог MAU {m:%m.%Y} {v} не сходится с суммой сегментов {seg[m]:.0f}")
    return pd.Series(total).sort_index()


def store_daily(con):
    h = W.load_history(con, ext=False)
    h["voice_queue"] = h["voice_entry_queue"] + h["voice_direct"]
    h["chat_queue"] = h["chat_entry_queue"]
    return h


def chat_bot(month, scenario):
    """Доля, закрытая ботом. До августа в траектории нет столбца — baseline листа."""
    m = pd.Timestamp(month).normalize()
    if m in scenario:
        return scenario[m], "сценарий 2, лист «Цель 1. Чат»"
    if m < PLAN_MONTH1:
        return BASELINE_CHAT_BOT, "до августа столбца нет, baseline листа 53%"
    raise RuntimeError(f"Нет сценария 2 на {m:%m.%Y}")


def build_check(cr, scenario, h, ev, holidays):
    rows = []
    for month in pd.date_range("2026-04-01", "2026-10-01", freq="MS"):
        if month not in cr.index:
            raise RuntimeError(f"В CR нет {month:%m.%Y}")
        sl = h[(h.index >= month) & (h.index < month + pd.offsets.MonthBegin(1))]
        bot, bot_src = chat_bot(month, scenario)
        reach = 1.0 - bot
        cell_v = float(cr.loc[month, "voice_cell"])
        cell_c = float(cr.loc[month, "chat_cell"])
        base = float(cr.loc[month, "base"])
        menu = float(sl["voice_entry"].sum(min_count=1))
        # сентябрь и октябрь: ячейка уже операторы, вход — меню хранилища
        if month >= pd.Timestamp("2026-09-01"):
            entry = menu
            entry_src = "меню хранилища"
        else:
            entry = cell_v
            entry_src = "ячейка CR, это вход"
        voice_fc = entry * VOICE_REACH
        chat_fc = cell_c * reach
        v_fact = float(sl["voice_queue"].sum(min_count=1))
        c_fact = float(sl["chat_queue"].sum(min_count=1))
        v2v, v2c, v2_note = _v2_month(h, ev, holidays, month)
        rows.append(dict(
            month=month, days=int(sl["chat_entry"].notna().sum()),
            base=base,
            chat_cell=cell_c, chat_cr=cell_c / base,
            bot=bot, bot_src=bot_src, chat_reach=reach, chat_fc=chat_fc,
            chat_fact=c_fact, chat_entry_store=float(sl["chat_entry"].sum(min_count=1)),
            chat_offered=float(sl["fact_chat_queue"].sum(min_count=1)),
            chat_err=c_fact / chat_fc - 1.0,
            v2_chat=v2c, v2_chat_err=(c_fact / v2c - 1.0) if v2c else None,
            voice_cell=cell_v, voice_entry=entry, voice_src=entry_src,
            voice_cr=entry / base, voice_fc=voice_fc,
            voice_fact=v_fact,
            voice_menu=menu,
            voice_reached=float(sl["voice_entry_queue"].sum(min_count=1)),
            voice_direct=float(sl["voice_direct"].sum(min_count=1)),
            voice_robot=float(sl["voice_robot"].sum(min_count=1)),
            voice_offered=float(sl["fact_voice_queue"].sum(min_count=1)),
            voice_err=v_fact / voice_fc - 1.0,
            voice_vs_entry=v_fact / entry - 1.0,
            v2_voice=v2v, v2_voice_err=(v_fact / v2v - 1.0) if v2v else None,
            v2_note=v2_note,
        ))
    return pd.DataFrame(rows)


def _v2_month(h, ev, holidays, month):
    """Текущая модель v2: прогноз месяца с последнего дня прошлого месяца, будущие акции скрыты."""
    origin = month - pd.Timedelta(days=1)
    have = h.dropna(subset=["voice_entry"]).index
    have = have[have <= origin]
    if have.empty:
        return None, None, "нет истории"
    origin = have.max()
    cfg = dict(promo_reserve=True)
    try:
        daily, _ = F.forecast_daily(h, origin, month, month + pd.offsets.MonthBegin(1) - pd.Timedelta(days=1),
                                    cfg, W._bt_events(ev, origin, cfg), holidays)
    except Exception as e:
        return None, None, str(e)
    sl = daily[(daily.index >= month) & (daily.index < month + pd.offsets.MonthBegin(1))]
    # факт чата до очереди есть не с 1 апреля: сравниваем v2 только там, где ряд уже есть — сумму за месяц как есть
    note = f"точка {origin:%d.%m.%Y}"

    def ready(col):
        st = F.series_start(h[col].dropna())
        if st is None or origin < st + pd.Timedelta(days=42):
            return False
        return col in sl and sl[col].notna().any() and float(sl[col].sum()) > 0

    vv = float(sl["voice_queue"].sum()) if ready("voice_queue") else None
    cc = float(sl["chat_queue"].sum()) if ready("chat_queue") else None
    if cc is None:
        note += "; чат v2 не показан: до точки меньше 6 недель ряда «до очереди»"
    return vv, cc, note


def _hdr(ws, titles):
    fill = PatternFill("solid", fgColor="1F3864")
    font = Font(bold=True, color="FFFFFF")
    for c, t in enumerate(titles, 1):
        cell = ws.cell(1, c, t)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 36
    ws.auto_filter.ref = f"A1:{get_column_letter(len(titles))}1"
    ws.freeze_panes = "A2"
    for c in range(1, len(titles) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 18


def write_check(path, rows, origin):
    wb = Workbook()
    ws = wb.active
    ws.title = "Апрель-октябрь"
    titles = [
        "Месяц", "Дней чата в хранилище", "База CR",
        "Чат CR, шт", "Чат CR = шт / база", "Бот закрывает, сценарий 2", "Откуда доля бота",
        "До оператора, чат", "Формула, чат очереди", "Факт чат очереди, уник.",
        "Факт / формула − 1, чат", "Факт чат вход, хранилище", "Очередь offered, чат",
        "v2 чат очереди", "Факт / v2 − 1, чат",
        "Ячейка голос CR", "Вход голоса для формулы", "Откуда вход", "Вход / база",
        "Формула, голос очереди = вход × доля", "Факт голос очереди, уник.",
        "Факт / формула − 1, голос", "Факт / вход − 1, голос",
        "Меню хранилища", "Дошло из меню", "Прямые", "Робот", "Очередь offered, голос",
        "v2 голос очереди", "Факт / v2 − 1, голос", "Точка v2",
    ]
    _hdr(ws, titles)
    for i, r in rows.iterrows():
        vals = [
            r["month"].strftime("%m.%Y"), r["days"], round(r["base"]),
            round(r["chat_cell"]), r["chat_cr"], r["bot"], r["bot_src"],
            r["chat_reach"], round(r["chat_fc"], 1), round(r["chat_fact"], 1),
            r["chat_err"], round(r["chat_entry_store"], 1), round(r["chat_offered"], 1),
            None if r["v2_chat"] is None else round(r["v2_chat"], 1), r["v2_chat_err"],
            round(r["voice_cell"]), round(r["voice_entry"], 1), r["voice_src"], r["voice_cr"],
            round(r["voice_fc"], 1), round(r["voice_fact"], 1),
            r["voice_err"], r["voice_vs_entry"],
            round(r["voice_menu"], 1), round(r["voice_reached"], 1), round(r["voice_direct"], 1),
            round(r["voice_robot"], 1), round(r["voice_offered"], 1),
            None if r["v2_voice"] is None else round(r["v2_voice"], 1), r["v2_voice_err"], r["v2_note"],
        ]
        for c, v in enumerate(vals, 1):
            ws.cell(i + 2, c, v)
        for c in (5, 6, 8, 11, 15, 19, 22, 23, 30):
            ws.cell(i + 2, c).number_format = "0.0%"
    # итог по штукам
    n = len(rows) + 2
    ws.cell(n, 1, "апрель–октябрь, сумма")
    for col, key in ((9, "chat_fc"), (10, "chat_fact"), (20, "voice_fc"), (21, "voice_fact")):
        ws.cell(n, col, round(float(rows[key].sum()), 1))
    ws.cell(n, 11, float(rows["chat_fact"].sum()) / float(rows["chat_fc"].sum()) - 1)
    ws.cell(n, 22, float(rows["voice_fact"].sum()) / float(rows["voice_fc"].sum()) - 1)
    ws.cell(n, 11).number_format = "0.0%"
    ws.cell(n, 22).number_format = "0.0%"
    for c in range(1, len(titles) + 1):
        ws.cell(n, c).font = Font(bold=True)

    w3 = wb.create_sheet("Как считал")
    notes = [
        "Проверка формулы на апрель–октябрь 2026. Это не прогноз года.",
        f"Срез хранилища до {origin:%d.%m.%Y}. Факт очереди — уникальные сессии: чат reached_queue, голос дошло из меню + прямые.",
        "«Очередь offered» — v_interval_volume, первое попадание. Это не знаменатель формулы.",
        "Чат файла CR — все начатые чаты КЦ. До оператора = эти штуки × (1 − сценарий 2). Октябрьский факт доли не брал.",
        "Апрель–июль 2026 в траектории сценария 2 столбца нет. Взял baseline того же листа, 53%. Август тоже 53% (месяц 1).",
        "Голос файла CR до августа включительно — вход в меню. Сентябрь и октябрь ячейка уже операторы: вход этих месяцев из меню хранилища.",
        "Формула голоса на этих месяцах = вход × доля до очереди.",
        "За план бота не ручаюсь. Спорт не добавлял.",
        "v2 — текущая модель с запасом на акции, точка прогноза последний день прошлого месяца, акции с началом после точки скрыты.",
        "Доля = числитель / знаменатель из колонок этого файла. Ошибка = факт / формула − 1.",
    ]
    for i, t in enumerate(notes, 1):
        w3.cell(i, 1, t)
    w3.column_dimensions["A"].width = 120
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)


def recent_share(h, origin, weeks=6):
    """Доля до оператора по недавней очереди. Их таблицу сюда не подставляем."""
    lo = origin - pd.Timedelta(weeks=weeks)
    sl = h[(h.index > lo) & (h.index <= origin)]

    def ratio(num, den):
        d = float(sl[den].sum())
        if d <= 0:
            raise RuntimeError(f"Нет знаменателя {den} за {weeks} недель до {origin:%d.%m.%Y}")
        return float(sl[num].sum()) / d

    return {"chat_queue": ratio("chat_queue", "chat_entry"), "voice_queue": ratio("voice_queue", "voice_entry")}


def usual_daily(h, ev, holidays, start, end):
    """Обычная модель v2 на дни, где плана клиентов нет."""
    origin = h.dropna(subset=["voice_entry", "chat_entry"]).index.max()
    cfg = dict(promo_reserve=True)
    daily, log = F.forecast_daily(h, origin, pd.Timestamp(start), pd.Timestamp(end), cfg, ev, holidays)
    return daily, log, origin


def draft_daily(cr, mau, start, end, h, holidays, v2, share, ref_year=None, direct_share=0.0):
    """Месяц с планом клиентов: план × CR того же месяца × наша доля.
    Месяца в плане нет: день в день обычная модель. Внутри месяца CR — профиль v2."""
    origin = h.dropna(subset=["voice_entry", "chat_entry"]).index.max()
    days = pd.date_range(start, end, freq="D")
    out = pd.DataFrame(index=days, columns=["voice_entry", "chat_entry", "voice_queue", "chat_queue", "mail", "voice_direct"], dtype=float)
    hol = set(pd.DatetimeIndex(holidays).normalize())
    detail = []
    weights = {}
    for col in ("voice_entry", "chat_entry"):
        series = h[col].dropna()
        mask, _ = F.detect_spikes(series, F.DEFAULTS)
        weights[col] = F.weekday_index(series, mask.reindex(series.index, fill_value=False), origin, 12)
    cr_col = {"voice_entry": "voice_cell", "chat_entry": "chat_cell"}
    queue_of = {"voice_entry": "voice_queue", "chat_entry": "chat_queue"}
    for month in pd.date_range(start, end, freq="MS"):
        mdays = pd.date_range(month, month + pd.offsets.MonthBegin(1) - pd.Timedelta(days=1), freq="D")
        mdays = mdays[(mdays >= start) & (mdays <= end)]
        has = month in mau.index and pd.notna(mau.loc[month])
        row = dict(month=month, source="CR" if has else "обычная модель")
        if not has:
            for col in ("voice_entry", "chat_entry", "voice_queue", "chat_queue", "mail", "voice_direct"):
                part = v2.reindex(mdays)[col].astype(float)
                if part.isna().any():
                    raise RuntimeError(f"Обычная модель не заполнила {col} за {month:%m.%Y}")
                out.loc[mdays, col] = part.values
                row[col] = float(part.sum())
            detail.append(row)
            continue
        prev_year = int(month.year) - 1 if ref_year is None else int(ref_year)
        prev = pd.Timestamp(year=prev_year, month=int(month.month), day=1)
        if prev not in cr.index:
            raise RuntimeError(f"В CR нет того же месяца {prev:%m.%Y}")
        base_prev = float(cr.loc[prev, "base"])
        row["mau"] = float(mau.loc[month])
        for col in ("voice_entry", "chat_entry"):
            ratio = float(cr.loc[prev, cr_col[col]]) / base_prev
            entry = row["mau"] * ratio
            wdi = weights[col]
            w = pd.Series([float(wdi[6] if d in hol else wdi[d.weekday()]) for d in mdays], index=mdays)
            day_entry = entry * w / float(w.sum())
            out.loc[mdays, col] = day_entry.values
            out.loc[mdays, queue_of[col]] = (day_entry * share[queue_of[col]]).values
            row[col] = entry
            row[queue_of[col]] = entry * share[queue_of[col]]
        part = v2.reindex(mdays)["mail"].astype(float)
        if part.isna().any():
            raise RuntimeError(f"Обычная модель не заполнила почту за {month:%m.%Y}")
        out.loc[mdays, "mail"] = part.values
        row["mail"] = float(part.sum())
        out.loc[mdays, "voice_direct"] = out.loc[mdays, "voice_entry"].astype(float) * direct_share
        row["voice_direct"] = float(out.loc[mdays, "voice_direct"].sum())
        detail.append(row)
    return out, origin, pd.DataFrame(detail)


def write_draft(con, daily, origin, start, end, mode, comment, holidays, h, passport=None):
    lo, hi = origin - pd.Timedelta(weeks=8), origin
    it = wfm_store.query(W.SQL_PROFILE.format(lo=f"date '{lo:%Y-%m-%d}'", hi=f"date '{hi:%Y-%m-%d}'"), con)
    bad = set()
    for col in ("voice_queue", "chat_queue", "mail"):
        m, _ = F.detect_spikes(h[col].dropna(), F.DEFAULTS)
        bad |= set(m[m].index)
    prof = F.profile_30(it, "v", "channel", sorted(bad), holidays)
    cols = [c for c in ("voice_entry", "chat_entry", "voice_queue", "chat_queue", "mail") if c in daily.columns]
    mapping = {k: W.CHANNELS[k] for k in cols}
    intervals = F.to_intervals(daily, prof, mapping, holidays)
    ad = A.load(con)
    am = A.fit(ad, origin, getattr(S, "AHT_CFG", {}), holidays)
    slots = intervals["interval_start"].drop_duplicates().sort_values()
    aht_iv = A.predict(am, slots, holidays)
    aht_iv["channel"] = "AHT " + aht_iv["channel"].map(A.CHANNELS) + ", с"
    intervals = pd.concat([intervals, aht_iv.rename(columns={"aht_sec": "value"})], ignore_index=True)
    run_id = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    params = {
        "method": "CR × MAU × наша доля; без плана клиентов — обычная модель",
        "origin": str(pd.Timestamp(origin).date()),
        "mail": "обычная модель на все месяцы, бота нет", "sport": "нет",
        "bot_table": "не используется",
        "aht": {k: round(v["level"], 1) for k, v in am.items()},
    }
    if passport:
        params.update(passport)
    params["months"] = {d.strftime("%Y-%m"): {c: round(float(v), 1) for c, v in row.items()}
                        for d, row in daily.resample("MS").sum().iterrows()}
    now = pd.Timestamp.now().floor("s")
    run = pd.DataFrame([dict(
        run_id=run_id, created_at=now, kind=f"модель: {mode}",
        period_from=pd.Timestamp(start), period_to=pd.Timestamp(end),
        scenario="CR × наша доля, пустые месяцы — v2", status="черновик",
        params=json.dumps(params, ensure_ascii=False), comment=comment,
        data_to=origin, model="Год: CR×MAU×наша доля, без клиентов — обычная модель. Не таблица ГРКЦ.",
        published_at=pd.NaT)])
    root = W.store_root()
    for name, df in (("forecast_run", run), ("forecast_interval", intervals.assign(run_id=run_id))):
        d = os.path.join(root, name)
        os.makedirs(d, exist_ok=True)
        tmp = os.path.join(d, f"run_{run_id}.parquet.tmp")
        pq.write_table(pa.Table.from_pandas(df, preserve_index=False), tmp, compression="zstd")
        os.replace(tmp, os.path.join(d, f"run_{run_id}.parquet"))
    return run_id, am


def _kc(entry, on_operator, direct, side_share):
    """Вход в КЦ и доля без оператора. На оператора не меняется."""
    kc = float(entry) * (1 - side_share) + float(direct)
    return kc, (W._without_share(kc, float(on_operator)) if kc else None)


def fill_year_days(ws, daily, holidays, side_share, chat_share):
    """Лист «По дням»: та же раскладка, из которой складывается месяц."""
    titles = ["Дата", "День недели", "Праздник",
              "Вход в КЦ, голос", "На оператора, голос", "Без оператора, голос",
              "Вход в КЦ, чат", "На оператора, чат", "Без оператора, чат", "Почта"]
    _hdr(ws, titles)
    hol = set(pd.DatetimeIndex(holidays).normalize())
    frame = daily.sort_index()
    for i, (ts, r) in enumerate(frame.iterrows(), start=2):
        ve, ce = float(r["voice_entry"]), float(r["chat_entry"])
        vg, cg = float(r["voice_queue"]), float(r["chat_queue"])
        day = pd.Timestamp(ts).normalize()
        direct = float(r["voice_direct"]) if "voice_direct" in r and pd.notna(r["voice_direct"]) else 0.0
        kc_v, wo_v = _kc(ve, vg, direct, side_share)
        kc_c, wo_c = _kc(ce, cg, 0.0, chat_share)
        vals = [day.date(), F.WD[day.weekday()], "да" if day in hol else "",
                round(kc_v, 1), round(vg, 1), wo_v,
                round(kc_c, 1), round(cg, 1), wo_c, round(float(r["mail"]), 1)]
        for c, v in enumerate(vals, 1):
            ws.cell(i, c, v)
        ws.cell(i, 1).number_format = "DD.MM.YYYY"
        for c in (4, 5, 7, 8, 10):
            ws.cell(i, c).number_format = "#,##0.0"
        ws.cell(i, 6).number_format = "0.0%"
        ws.cell(i, 9).number_format = "0.0%"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(titles))}{len(frame) + 1}"


def _fmt(x, digits=0):
    return f"{float(x):,.{digits}f}".replace(",", " ").replace(".", ",")


def _pct(x):
    return f"{float(x):.1%}".replace(".", ",")


def _year_steps(detail, share):
    """Месяц с CR: клиенты × частота × доля. С числами первого такого месяца."""
    lines = ["Как считали этот прогон."]
    cr = detail[detail["source"] == "CR"] if "source" in detail.columns else detail.iloc[0:0]
    if len(cr) and "mau" in cr.columns and pd.notna(cr.iloc[0].get("mau")):
        r = cr.iloc[0]
        mau = float(r["mau"])
        ve, vg = float(r["voice_entry"]), float(r["voice_queue"])
        ce, cg = float(r["chat_entry"]), float(r["chat_queue"])
        label = pd.Timestamp(r["month"]).strftime("%m.%Y")
        lines.append(
            f"{label}, голос. Берём {_fmt(mau)} клиентов. Умножаем на частоту того же месяца прошлого года "
            f"{_pct(ve / mau if mau else 0)}: {_fmt(ve)} входов. "
            f"Умножаем на {_pct(share['voice_queue'])} — кто дошёл до оператора за 6 недель. "
            f"Колонка «На оператора, голос»: {_fmt(vg)}."
        )
        lines.append(
            f"{label}, чат. Те же {_fmt(mau)} клиентов × частота {_pct(ce / mau if mau else 0)} = {_fmt(ce)} входов. "
            f"× {_pct(share['chat_queue'])} = колонка «На оператора, чат» {_fmt(cg)}."
        )
    lines.append("Месяц без CR модель пишет очередь сразу. Долю 6 недель туда второй раз не умножаем.")
    return lines


def write_year_excel(detail, run_id, share, daily, holidays, side, chat_side):
    """Книга, которую открывает кнопка «План на год»: месяцы и те же дни."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Месяцы"
    titles = ["Месяц", "Откуда",
              "Вход в КЦ, голос", "На оператора, голос", "Без оператора, голос",
              "Вход в КЦ, чат", "На оператора, чат", "Без оператора, чат", "Почта"]
    _hdr(ws, titles)
    for i, r in detail.reset_index(drop=True).iterrows():
        ve, ce = float(r["voice_entry"]), float(r["chat_entry"])
        vg, cg = float(r["voice_queue"]), float(r["chat_queue"])
        direct = float(r["voice_direct"]) if "voice_direct" in r and pd.notna(r["voice_direct"]) else 0.0
        kc_v, wo_v = _kc(ve, vg, direct, side["share"])
        kc_c, wo_c = _kc(ce, cg, 0.0, chat_side["share"])
        vals = [pd.Timestamp(r["month"]).strftime("%m.%Y"), r["source"],
                round(kc_v, 1), round(vg, 1), wo_v,
                round(kc_c, 1), round(cg, 1), wo_c,
                round(float(r["mail"]), 1)]
        for c, v in enumerate(vals, 1):
            ws.cell(i + 2, c, v)
        for c in (3, 4, 6, 7, 9):
            ws.cell(i + 2, c).number_format = "#,##0.0"
        ws.cell(i + 2, 5).number_format = "0.0%"
        ws.cell(i + 2, 8).number_format = "0.0%"
    note = wb.create_sheet("Параметры")
    lines = _year_steps(detail, share) + [
        f"Черновик {run_id}. Сам не публикуется.",
        "Один принцип: вход в КЦ — линия очереди, на оператора — кто дошёл до очереди. Без оператора = (вход в КЦ − на оператора) / вход в КЦ.",
        "На оператора эта доля не режет. Прямые звонки есть и во входе, и на операторе.",
        (f"Голос: из входа вынуты звонки меню на смежные отделы, не дошедшие до рабочей очереди. "
         f"За {side['lo']:%d.%m.%Y}–{side['hi']:%d.%m.%Y} таких {side['side']:.0f} из {side['menu']:.0f} = {side['share']:.1%}."),
        (f"Чат: из входа вынуты «чат VIP и смежный отдел», «чат отдела претензий» и их завершение, если сессия не встала в рабочей очереди. "
         f"За {chat_side['lo']:%d.%m.%Y}–{chat_side['hi']:%d.%m.%Y} таких {chat_side['side']:.0f} из {chat_side['menu']:.0f} = {chat_side['share']:.1%}. "
         "Почта: вход = на оператора, без оператора 0%."),
        f"Доля до оператора — факт последних 6 недель: голос {share['voice_queue']:.1%} входов, чат {share['chat_queue']:.1%}. CR даёт вход, доля переводит его в очередь. Обычная модель сразу пишет очередь.",
        "Таблица бота из бизнес-плана не используется.",
        "Лист «По дням»: месяц по CR разложен по дням недели, праздник как воскресенье. Месяц обычной модели — её день.",
        "Почта — обычная модель на каждый день. CR на письма не умножается. Письмо сразу в очереди, без оператора 0%.",
    ]
    for i, t in enumerate(lines, 1):
        note.cell(i, 1, t)
    note.column_dimensions["A"].width = 130
    days = wb.create_sheet("По дням", 1)
    fill_year_days(days, daily, holidays, side["share"], chat_side["share"])
    path = report_path(f"Прогноз год {stamp()}.xlsx")
    wb.save(path)
    return path


def main(check_only=False, check_path=None):
    cr = read_cr()
    scenario = read_scenario2()
    con = wfm_store.connect()
    h = store_daily(con)
    origin = h.dropna(subset=["voice_entry", "chat_entry"]).index.max()
    ev = W.load_events()
    holidays = pd.DatetimeIndex(pd.to_datetime(getattr(S, "HOLIDAYS", [])))
    rows = build_check(cr, scenario, h, ev, holidays)
    stamp = dt.datetime.now().strftime("%d.%m.%Y %H-%M")
    path = check_path or os.path.join(ROOT, "reports", "checks", f"Проверка формулы CR {stamp}.xlsx")
    write_check(path, rows, origin)
    print(f"Проверка: {path}")
    print(f"Чат: факт {rows.chat_fact.sum():.1f} / формула {rows.chat_fc.sum():.1f}")
    print(f"Голос: факт {rows.voice_fact.sum():.1f} / формула {rows.voice_fc.sum():.1f}")

    if check_only:
        con.close()
        return
    con.close()


def month_spans(months):
    """Сплошные месяцы одной строкой: 01.2028–06.2028. Дырка — отдельный кусок."""
    months = sorted(pd.to_datetime(list(months)))
    if not months:
        return "нет"
    spans, start, prev = [], months[0], months[0]
    for m in months[1:]:
        if m == (prev + pd.offsets.MonthBegin(1)):
            prev = m
            continue
        spans.append((start, prev))
        start = prev = m
    spans.append((start, prev))
    return ", ".join(f"{a:%m.%Y}" if a == b else f"{a:%m.%Y}–{b:%m.%Y}" for a, b in spans)


def open_year():
    """Год. CR даёт вход, доля переводит его в очередь. Где CR нет — обычная модель. Не публикуем."""
    cr = read_cr()
    mau = read_mau_plan()
    con = wfm_store.connect()
    h = store_daily(con)
    full = W.load_history(con, ext=True)
    full["voice_queue"] = full["voice_entry_queue"] + full["voice_direct"]
    full["chat_queue"] = full["chat_entry_queue"]
    ev = W.load_events()
    holidays = pd.DatetimeIndex(pd.to_datetime(getattr(S, "HOLIDAYS", [])))
    start_ts, end_ts = W.year_period()
    start, end = f"{start_ts:%Y-%m-%d}", f"{end_ts:%Y-%m-%d}"
    v2, log, origin = usual_daily(full, ev, holidays, start, end)
    share = recent_share(h, origin)
    sl = h[(h.index > origin - pd.Timedelta(weeks=6)) & (h.index <= origin)]
    direct_share = float(sl["voice_direct"].sum()) / float(sl["voice_entry"].sum())
    side = W.voice_funnel_side(con, origin)
    chat_side = W.chat_funnel_side(con, origin)
    daily, org, detail = draft_daily(cr, mau, start, end, h, holidays, v2, share, ref_year=start_ts.year - 1,
                                     direct_share=direct_share)
    yid, am = write_draft(
        con, daily, org, start, end, "год",
        "черновик года: CR даёт вход, доля переводит его в очередь; где CR нет — обычная модель. Не публиковать",
        holidays, h,
        passport={"share_chat": round(share["chat_queue"], 4), "share_voice": round(share["voice_queue"], 4),
                  "share_weeks": 6, "v2_chat_share": (log.get("chat_share") or {}).get("share"),
                  "v2_voice_share": (log.get("voice_share") or {}).get("share")})
    con.close()
    path = write_year_excel(detail, yid, share, daily, holidays, side, chat_side)
    aht_weeks = int(getattr(S, "AHT_CFG", {}).get("level_weeks", 8))
    aht_bits = [f"{name.lower()} {model['level']:.1f} с" for name, model in am.items()]
    cr_months = detail.loc[detail["source"] == "CR", "month"]
    other_months = detail.loc[detail["source"] != "CR", "month"]
    print(f"Черновик {yid}. Сам не публикуется.")
    print(f"CR есть: {month_spans(cr_months)}")
    print(f"CR нет: {month_spans(other_months)}")
    print(f"AHT — медиана факта за {aht_weeks} недель до {origin:%d.%m.%Y}: " + ", ".join(aht_bits) + ".")
    print(f"Файл: {path}")


if __name__ == "__main__":
    import sys
    if "--open-year" in sys.argv:
        open_year()
    else:
        only = "--check-only" in sys.argv
        fixed = os.path.join(ROOT, "reports", "checks", "Проверка формулы CR 10.11.2026 17-45.xlsx")
        main(check_only=only, check_path=fixed if only else None)
