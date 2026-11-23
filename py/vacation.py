# -*- coding: utf-8 -*-
"""Окна отпуска на 2027. Не прогноз и не CP-SAT.

Читает состав из «Графики Q4.xlsx», лист «План графиков» (без дней и без колонки «Квота»),
и книгу потребности на последний опубликованный прогноз года, листы «30 мин» и «Месяцы».
Ёмкость линии — сумма ставок. Люди заново не считаются.
"""
import os
import sys
from datetime import date, datetime, timedelta

import openpyxl
import pandas as pd
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import wfm_store  # noqa: E402
from wfm_common import report_path, stamp  # noqa: E402
from wfm_settings import HOLIDAYS  # noqa: E402

ROOT = os.path.dirname(HERE)
ROSTER_PATH = os.path.join(ROOT, "inputs", "Графики Q4.xlsx")
NEED_DIR = os.path.join(ROOT, "reports", "need")
VACATION_DAYS = 28
PART_DAYS = (7, 14, 7)
SPARE_ROWS = 5
GROUPS = ("линия", "201", "дежурный", "РС")

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
BAND = PatternFill("solid", fgColor="F3F6FB")
SECTION_FILL = PatternFill("solid", fgColor="1F3864")
SECTION_FONT = Font(bold=True, color="FFFFFF")
LABEL_FONT = Font(color="1F3864")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
PART_FILL = (
    PatternFill("solid", fgColor="D6EAF8"),
    PatternFill("solid", fgColor="FDEBD0"),
    PatternFill("solid", fgColor="D5F5E3"),
)
OFF_FILL = PatternFill("solid", fgColor="C6EFCE")
OFF_FONT = Font(color="006100")
WARN_FILL = PatternFill("solid", fgColor="F8CBAD")
WEEKEND_FILL = PatternFill("solid", fgColor="D6DCE4")
UNLOCK = Protection(locked=False)
WEEKDAYS = ("пн", "вт", "ср", "чт", "пт", "сб", "вс")

THIN = Border(
    left=Side(style="thin", color="D0D7E2"),
    right=Side(style="thin", color="D0D7E2"),
    top=Side(style="thin", color="D0D7E2"),
    bottom=Side(style="thin", color="D0D7E2"),
)


def _one(value):
    return round(float(value) + 1e-9, 1)


def _shift(value):
    if isinstance(value, bool):
        return ""
    if isinstance(value, (int, float)):
        return str(int(value))
    return str(value or "").strip()


def _num(value):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    return float(value)


def load_roster(path):
    """Логин только как признак «строка есть». Имена и дни не сохраняются."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if "План графиков" not in wb.sheetnames:
        sys.exit("В книге графиков нет листа «План графиков»")
    ws = wb["План графиков"]
    left = list(ws.iter_rows(min_row=3, max_col=12, values_only=True))
    wb.close()
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["План графиков"]
    right = list(ws.iter_rows(min_row=3, min_col=112, max_col=112, values_only=True))
    wb.close()
    if len(left) != len(right):
        sys.exit(f"Лист «План графиков»: шапка {len(left)} строк, ставка {len(right)}")

    groups = {
        "линия": {"n": 0, "rate": 0.0, "full": 0},
        "201": {"n": 0, "rate": 0.0},
        "дежурный": {"n": 0, "rate": 0.0},
        "РС": {"n": 0, "rate": 0.0},
    }
    logins = 0
    people = []
    for row, rate_cell in zip(left, right):
        login = row[1]
        if not isinstance(login, str) or not login.strip():
            continue
        logins += 1
        channel = str(row[9] or "").strip()
        shift = _shift(row[6])
        rate = _num(rate_cell[0])
        if rate is None:
            sys.exit("В колонке «Ставка» пусто")
        if channel == "чат/звонок" and shift == "201":
            key = "201"
        elif channel == "чат/звонок":
            key = "линия"
        elif channel == "дежурный":
            key = "дежурный"
        elif channel == "РС":
            key = "РС"
        else:
            sys.exit(f"Канал вне линии, 201, дежурных и РС: {channel or 'пусто'}")
        groups[key]["n"] += 1
        groups[key]["rate"] += rate
        if key == "линия" and _num(row[11]) == 11 and rate == 1:
            groups["линия"]["full"] += 1
        people.append({"group": key, "shift": shift, "rate": round(rate, 2)})

    line = groups["линия"]
    got = (
        logins,
        line["n"],
        round(line["rate"], 2),
        line["full"],
        groups["201"]["n"],
        groups["дежурный"]["n"],
        groups["РС"]["n"],
    )
    expect = (42, 34, 32.07, 27, 2, 4, 2)
    if got != expect:
        sys.exit(
            "Состав не сошёлся с журналом №173: "
            f"логины {got[0]}, линия {got[1]}, ставки {got[2]}, полных {got[3]}, "
            f"201 {got[4]}, дежурные {got[5]}, РС {got[6]}"
        )
    return {
        "logins": logins,
        "line_n": line["n"],
        "line_rate": round(line["rate"], 2),
        "full": line["full"],
        "m201_n": groups["201"]["n"],
        "m201_rate": round(groups["201"]["rate"], 2),
        "duty_n": groups["дежурный"]["n"],
        "duty_rate": round(groups["дежурный"]["rate"], 2),
        "rs_n": groups["РС"]["n"],
        "rs_rate": round(groups["РС"]["rate"], 2),
        "people": people,
    }


def _as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str) and len(value) >= 10 and value[2] == ".":
        return datetime.strptime(value[:10], "%d.%m.%Y").date()
    return None


def load_need(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    for name in ("30 мин", "Месяцы"):
        if name not in wb.sheetnames:
            sys.exit(f"В потребности нет листа «{name}»")
    ws = wb["30 мин"]
    rows = ws.iter_rows(values_only=True)
    header = [str(c or "") for c in next(rows)]
    try:
        i_date = header.index("Дата")
        i_voice = header.index("Голос: в линии")
        i_chat = header.index("Чат: в линии")
    except ValueError:
        sys.exit("На листе «30 мин» нет даты или колонок «в линии»")
    daily = {}
    for row in rows:
        day = _as_date(row[i_date])
        if day is None:
            continue
        voice, chat = _num(row[i_voice]), _num(row[i_chat])
        if voice is None or chat is None:
            sys.exit(f"Пустая «в линии» на {day:%d.%m.%Y}")
        daily[day] = daily.get(day, 0.0) + (voice + chat) * 0.5

    ws = wb["Месяцы"]
    rows = ws.iter_rows(values_only=True)
    header = [str(c or "") for c in next(rows)]
    try:
        i_month = header.index("Месяц")
        i_need = next(
            header.index(name) for name in ("Всего: люди с абсентеизмом", "Всего: люди с внепланом")
            if name in header
        )
    except (ValueError, StopIteration):
        sys.exit("На листе «Месяцы» нет колонки «Всего: люди с абсентеизмом»")
    months = []
    for row in rows:
        label = row[i_month]
        need = _num(row[i_need])
        if label is None or need is None:
            continue
        months.append((str(label), need))
    wb.close()
    if len(daily) != 365 or len(months) != 12:
        sys.exit(f"В потребности дней {len(daily)}, месяцев {len(months)}. Ждал 365 и 12")
    return daily, months


def _span(daily, start, days):
    return sum(daily[start + timedelta(days=i)] for i in range(days))


def _check_windows(daily):
    year = min(daily)
    if year.year != 2027 or year != date(2027, 1, 1):
        sys.exit("Потребность начинается не с 01.01.2027")


def _book_run_id(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if "Параметры" not in wb.sheetnames:
        wb.close()
        return None
    found = None
    for row in wb["Параметры"].iter_rows(max_col=2, values_only=True):
        if row and row[0] == "Номер" and row[1]:
            found = str(row[1]).strip()
            break
    wb.close()
    return found


def _year_need_path():
    """Книга потребности на последний опубликованный прогноз года."""
    runs = wfm_store.query(
        "select run_id, created_at, status from forecast_run where kind = 'модель: год'"
    )
    if runs.empty:
        sys.exit("Нет прогноза на год")
    runs["created_at"] = pd.to_datetime(runs["created_at"])
    published = runs[runs["status"] == "опубликован"]
    if published.empty:
        sys.exit("Нет опубликованного прогноза на год")
    run_id = str(published.sort_values("created_at").iloc[-1]["run_id"])
    if not os.path.isdir(NEED_DIR):
        sys.exit(f"Нет папки потребности: {NEED_DIR}")
    hits = []
    for name in os.listdir(NEED_DIR):
        if not name.startswith("Потребность год ") or not name.endswith(".xlsx"):
            continue
        if name.startswith("~$"):
            continue
        path = os.path.join(NEED_DIR, name)
        if _book_run_id(path) == run_id:
            hits.append(path)
    if not hits:
        sys.exit(f"Нет книги потребности на опубликованный год {run_id}")
    hits.sort(key=os.path.getmtime)
    return hits[-1], run_id


def _paint(ws, formats):
    headers = [c.value for c in ws[1]]
    for cell in ws[1]:
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = THIN
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{max(ws.max_row, 2)}"
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        stripe = row[0].row % 2 == 0
        for cell in row:
            header = headers[cell.column - 1]
            fmt = formats.get(header)
            if fmt and isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                cell.number_format = fmt
            if stripe:
                cell.fill = BAND
            cell.border = THIN
            cell.alignment = Alignment(vertical="center")
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"


def _write_table(ws, headers, records, widths, formats):
    ws.append(headers)
    for rec in records:
        ws.append([rec.get(h) for h in headers])
    for i, header in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(header, 16)
    _paint(ws, formats)


def _periods(daily, length, holidays):
    """Все окна года. Соседние с одним чел.-ч внутри месяца — одна строка."""
    last = date(2027, 12, 31)
    windows = []
    day = date(2027, 1, 1)
    while day + timedelta(days=length - 1) <= last:
        end = day + timedelta(days=length - 1)
        hours = _one(_span(daily, day, length))
        has_holiday = any((day + timedelta(days=i)) in holidays for i in range(length))
        windows.append((day, end, hours, has_holiday))
        day += timedelta(days=1)

    glued = []

    def flush(buf):
        first, last_w = buf[0], buf[-1]
        if len(buf) == 1:
            period = f"{first[0]:%d.%m.%Y}–{first[1]:%d.%m.%Y}"
        else:
            period = f"любые {length} дней, начало {first[0]:%d.%m.%Y}–{last_w[0]:%d.%m.%Y}"
        month = f"{first[0]:%m.%Y}"
        if (first[0].year, first[0].month) != (last_w[1].year, last_w[1].month):
            month = f"{first[0]:%m.%Y}–{last_w[1]:%m.%Y}"
        return {
            "start": first[0],
            "Часть": length,
            "Месяц": month,
            "Период": period,
            "Чел.-ч": first[2],
            "Пометка": "праздник" if first[3] else None,
        }

    buf = []
    for window in windows:
        if buf:
            prev = buf[-1]
            same = (
                window[0] == prev[0] + timedelta(days=1)
                and window[2] == prev[2]
                and window[3] == prev[3]
                and window[0].month == prev[0].month
                and window[0].year == prev[0].year
            )
            if not same:
                glued.append(flush(buf))
                buf = []
        buf.append(window)
    if buf:
        glued.append(flush(buf))
    glued.sort(key=lambda row: (row["Чел.-ч"], row["start"]))
    for i, row in enumerate(glued, 1):
        row["Место"] = i
    return glued


def _cumulative(holidays):
    cum = {date(2026, 12, 31): 0}
    running = 0
    day = date(2027, 1, 1)
    while day <= date(2027, 12, 31):
        if day not in holidays:
            running += 1
        cum[day] = running
        day += timedelta(days=1)
    return cum


def _end_on(start, days, cum):
    target = cum[start - timedelta(days=1)] + days
    day = start
    while day <= date(2027, 12, 31):
        if cum[day] == target:
            return day
        day += timedelta(days=1)
    return None


def _check_form(daily, holidays):
    cum = _cumulative(holidays)
    if _end_on(date(2027, 5, 11), 7, cum) != date(2027, 5, 17):
        sys.exit("Конец 7 дней от 11.05.2027 посчитан неверно")
    if _end_on(date(2027, 1, 1), 7, cum) != date(2027, 1, 17):
        sys.exit("Праздники 1–10.01 не удлиняют часть")
    return cum


def _slots(people):
    ordered = []
    for group in GROUPS:
        ordered.extend(person for person in people if person["group"] == group)
    if len(ordered) != 42:
        sys.exit(f"В форме {len(ordered)} строк состава, ждал 42")
    ordered.extend({"group": "", "shift": "", "rate": None} for _ in range(SPARE_ROWS))
    return ordered


def _parameters(ws, need_name):
    blocks = [
        ("Куда писать", [
            ("Лист", "«Сотрудники». Жёлтые ячейки: имя и три даты начала", None),
            ("Части", "сначала 7 дней, потом 14, потом ещё 7. Вместе 28. Порядок частей на календаре любой", None),
            ("Строки", "порядок листа «План графиков». Имена не подставлены. Внизу 5 пустых строк", None),
            ("Защита", "листы «Сотрудники» и «Календарь» без пароля. Снять: Рецензирование → Снять защиту листа", None),
        ]),
        ("Что встаёт само", [
            ("Конец", "7 или 14 дней без праздника. Праздник внутри части удлиняет её до конца года", None),
            ("Чел.-ч", "люди «в линии», голос и чат, от начала до конца, получас × 0,5. Меньше — спокойнее эти дни", None),
            ("Список", "лист «Периоды»: те же чел.-ч у ровного блока 7 или 14 дней. Сверху легче. Праздник помечен", None),
            ("Календарь", "буква «от» и число людей группы в этот день. 201 и дежурные: двое в один день подсвечены", None),
            ("Проверка", "ок, не все части, пересечение, не влезает в год, не 28", None),
            ("Потолок линии", "не ставится", None),
        ]),
        ("Откуда", [
            ("Потребность", f"{need_name}, листы «30 мин» и «Месяцы». Последний опубликованный год", None),
            ("Праздники", "2027, настройки консоли", None),
            ("Год и квартал", "не пересчитывались", None),
        ]),
    ]
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 92
    row_i = 1
    for title, items in blocks:
        ws.merge_cells(start_row=row_i, start_column=1, end_row=row_i, end_column=2)
        head = ws.cell(row_i, 1, title)
        head.fill = SECTION_FILL
        head.font = SECTION_FONT
        head.alignment = Alignment(vertical="center", indent=1)
        ws.cell(row_i, 2).fill = SECTION_FILL
        ws.row_dimensions[row_i].height = 22
        row_i += 1
        for name, value, fmt in items:
            left = ws.cell(row_i, 1, name)
            left.font = LABEL_FONT
            left.alignment = Alignment(vertical="center", indent=1)
            left.border = THIN
            right = ws.cell(row_i, 2, value)
            right.alignment = Alignment(vertical="center", wrap_text=True)
            right.border = THIN
            if fmt and isinstance(value, (int, float)) and not isinstance(value, bool):
                right.number_format = fmt
            if isinstance(value, str) and len(value) > 70:
                ws.row_dimensions[row_i].height = 32
            row_i += 1
        row_i += 1
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False


def _end_formula(row, start_col, days_col):
    start = get_column_letter(start_col)
    days = get_column_letter(days_col)
    return (
        f'=IF({start}{row}="","",IFERROR(INDEX(Дни!$A$3:$A$367,'
        f'MATCH(IFERROR(VLOOKUP({start}{row}-1,Дни!$A$2:$C$367,3,FALSE),0)+{days}{row},'
        f'Дни!$C$3:$C$367,0)),""))'
    )


def _hours_formula(row, start_col, end_col):
    start = get_column_letter(start_col)
    end = get_column_letter(end_col)
    return (
        f'=IF({end}{row}="","",SUMIFS(Дни!$D$3:$D$367,'
        f'Дни!$A$3:$A$367,">="&{start}{row},Дни!$A$3:$A$367,"<="&{end}{row}))'
    )


def _check_formula(row):
    return (
        f'=IF(COUNT(F{row},J{row},N{row})=0,"",'
        f'IF(OR(AND(F{row}<>"",H{row}=""),AND(J{row}<>"",L{row}=""),AND(N{row}<>"",P{row}="")),'
        f'"не влезает в год",'
        f'IF(COUNT(F{row},J{row},N{row})<3,"не все части",'
        f'IF(OR(AND(F{row}<=L{row},J{row}<=H{row}),AND(F{row}<=P{row},N{row}<=H{row}),'
        f'AND(J{row}<=P{row},N{row}<=L{row})),"пересечение",'
        f'IF(R{row}<>28,"не 28","ок")))))'
    )


def _mark_formula(staff_row, col_letter):
    day = f"{col_letter}$4"
    terms = []
    for start, end in (("F", "H"), ("J", "L"), ("N", "P")):
        terms.append(
            f'AND(Сотрудники!${start}{staff_row}<>"",'
            f'Сотрудники!${start}{staff_row}<={day},'
            f'Сотрудники!${end}{staff_row}>={day})'
        )
    return '=IF(OR(' + ",".join(terms) + '),"от","")'


def _write_days(ws, daily, holidays, cum):
    ws.append(["Дата", "Праздник", "Накоплено", "Чел.-ч"])
    ws.append([date(2026, 12, 31), 0, 0, 0])
    day = date(2027, 1, 1)
    while day <= date(2027, 12, 31):
        ws.append([day, 1 if day in holidays else 0, cum[day], daily[day]])
        day += timedelta(days=1)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1):
        row[0].number_format = "DD.MM.YYYY"
    ws.sheet_state = "hidden"


def _write_staff(ws, slots):
    headers = [
        "№", "Группа", "Смена", "Ставка", "Имя",
        "7 дней, начало", "Дней", "Конец", "Чел.-ч",
        "14 дней, начало", "Дней", "Конец", "Чел.-ч",
        "7 дней, начало", "Дней", "Конец", "Чел.-ч",
        "Всего дней", "Проверка",
    ]
    ws.append(headers)
    part_cols = {6: 0, 7: 0, 8: 0, 9: 0, 10: 1, 11: 1, 12: 1, 13: 1, 14: 2, 15: 2, 16: 2, 17: 2}
    input_cols = {5, 6, 10, 14}
    first_spare = len(slots) - SPARE_ROWS + 2
    for i, person in enumerate(slots):
        row = i + 2
        ws.cell(row, 1, i + 1)
        ws.cell(row, 2, person["group"] or None)
        ws.cell(row, 3, person["shift"] or None)
        if person["rate"] is not None:
            ws.cell(row, 4, person["rate"])
        ws.cell(row, 7, PART_DAYS[0])
        ws.cell(row, 8, _end_formula(row, 6, 7))
        ws.cell(row, 9, _hours_formula(row, 6, 8))
        ws.cell(row, 11, PART_DAYS[1])
        ws.cell(row, 12, _end_formula(row, 10, 11))
        ws.cell(row, 13, _hours_formula(row, 10, 12))
        ws.cell(row, 15, PART_DAYS[2])
        ws.cell(row, 16, _end_formula(row, 14, 15))
        ws.cell(row, 17, _hours_formula(row, 14, 16))
        ws.cell(
            row, 18,
            f'=IF(COUNT(F{row},J{row},N{row})=0,"",(F{row}<>"")*G{row}+(J{row}<>"")*K{row}+(N{row}<>"")*O{row})',
        )
        ws.cell(row, 19, _check_formula(row))
        for col in range(1, 20):
            cell = ws.cell(row, col)
            cell.border = THIN
            cell.alignment = Alignment(vertical="center")
            if col in part_cols and col not in input_cols:
                cell.fill = PART_FILL[part_cols[col]]
            if col in input_cols or (row >= first_spare and col in (2, 3, 4)):
                cell.fill = INPUT_FILL
                cell.protection = UNLOCK
        ws.cell(row, 4).number_format = "0.00"
        for col in (6, 8, 10, 12, 14, 16):
            ws.cell(row, col).number_format = "DD.MM.YYYY"
        for col in (9, 13, 17):
            ws.cell(row, col).number_format = "#,##0.0"
    for cell in ws[1]:
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = THIN
    ws.row_dimensions[1].height = 36
    ws.freeze_panes = "F2"
    ws.auto_filter.ref = f"A1:S{len(slots) + 1}"
    widths = {
        "A": 6, "B": 14, "C": 12, "D": 10, "E": 28,
        "F": 16, "G": 8, "H": 14, "I": 12,
        "J": 16, "K": 8, "L": 14, "M": 12,
        "N": 16, "O": 8, "P": 14, "Q": 12,
        "R": 12, "S": 18,
    }
    for letter, width in widths.items():
        ws.column_dimensions[letter].width = width
    last = len(slots) + 1
    dates = DataValidation(
        type="date", operator="between", allow_blank=True,
        formula1="DATE(2027,1,1)", formula2="DATE(2027,12,31)",
        error="Дата с 01.01.2027 по 31.12.2027", errorTitle="Дата",
    )
    groups = DataValidation(
        type="list", formula1='"линия,201,дежурный,РС"', allow_blank=True,
        error="Группа: линия, 201, дежурный или РС", errorTitle="Группа",
    )
    ws.add_data_validation(dates)
    ws.add_data_validation(groups)
    dates.add(f"F2:F{last}")
    dates.add(f"J2:J{last}")
    dates.add(f"N2:N{last}")
    groups.add(f"B{first_spare}:B{last}")
    ws.conditional_formatting.add(
        f"S2:S{last}",
        FormulaRule(formula=['AND($S2<>"",$S2<>"ок")'], fill=WARN_FILL),
    )
    ws.protection.sheet = True
    ws.protection.autoFilter = False
    ws.protection.sort = False


def _write_calendar(ws, slots, holidays):
    days = []
    day = date(2027, 1, 1)
    while day <= date(2027, 12, 31):
        days.append(day)
        day += timedelta(days=1)
    first = 10
    last = first + len(slots) - 1
    labels = ("линия", "201", "дежурный", "РС")
    for col, day in enumerate(days, 4):
        letter = get_column_letter(col)
        ws.cell(2, col, day.day)
        ws.cell(3, col, WEEKDAYS[day.weekday()])
        date_cell = ws.cell(4, col, day)
        date_cell.number_format = "DD.MM.YYYY"
        head_fill = WARN_FILL if day in holidays else WEEKEND_FILL if day.weekday() >= 5 else HDR_FILL
        for row in (1, 2, 3):
            cell = ws.cell(row, col)
            cell.fill = head_fill
            cell.font = Font(bold=True, color="FFFFFF" if head_fill is HDR_FILL else "1F3864")
            cell.alignment = Alignment(horizontal="center")
        for offset, group in enumerate(labels):
            ws.cell(
                5 + offset, col,
                f'=COUNTIFS($B${first}:$B${last},"{group}",{letter}{first}:{letter}{last},"от")',
            )
        for i in range(len(slots)):
            ws.cell(first + i, col, _mark_formula(2 + i, letter))
        ws.column_dimensions[letter].width = 3.4
    month_start = 4
    current = days[0].month
    for index, day in enumerate(days + [None]):
        month = day.month if day else None
        if month != current:
            end_col = 3 + index
            ws.merge_cells(start_row=1, start_column=month_start, end_row=1, end_column=end_col)
            label = ws.cell(1, month_start, f"{current:02d}.2027")
            label.alignment = Alignment(horizontal="center")
            month_start = end_col + 1
            current = month
    ws.cell(5, 1, "В отпуске, линия")
    ws.cell(6, 1, "В отпуске, 201")
    ws.cell(7, 1, "В отпуске, дежурные")
    ws.cell(8, 1, "В отпуске, РС")
    for row, person_index in enumerate(range(len(slots)), first):
        staff_row = 2 + person_index
        ws.cell(row, 1, f'=IF(Сотрудники!A{staff_row}="","",Сотрудники!A{staff_row})')
        ws.cell(row, 2, f'=IF(Сотрудники!B{staff_row}="","",Сотрудники!B{staff_row})')
        ws.cell(row, 3, f'=IF(Сотрудники!E{staff_row}="","",Сотрудники!E{staff_row})')
    for row in range(5, 9):
        ws.cell(row, 1).font = LABEL_FONT
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    ws.row_dimensions[4].hidden = True
    ws.freeze_panes = "D10"
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 28
    day_range = f"D{first}:{get_column_letter(3 + len(days))}{last}"
    ws.conditional_formatting.add(
        day_range,
        CellIsRule(operator="equal", formula=['"от"'], fill=OFF_FILL, font=OFF_FONT),
    )
    for row in (6, 7):
        ws.conditional_formatting.add(
            f"D{row}:{get_column_letter(3 + len(days))}{row}",
            CellIsRule(operator="greaterThan", formula=["1"], fill=WARN_FILL),
        )
    ws.protection.sheet = True
    ws.sheet_view.zoomScale = 90


def write_book(path, rows, slots, daily, holidays, cum, need_name):
    wb = openpyxl.Workbook()
    staff = wb.active
    staff.title = "Сотрудники"
    _write_staff(staff, slots)
    _write_calendar(wb.create_sheet("Календарь"), slots, holidays)
    periods = wb.create_sheet("Периоды")
    _write_table(
        periods,
        ["Место", "Часть", "Месяц", "Период", "Чел.-ч", "Пометка"],
        rows,
        {"Место": 10, "Часть": 10, "Месяц": 18, "Период": 56, "Чел.-ч": 12, "Пометка": 14},
        {"Место": "0", "Часть": "0", "Чел.-ч": "#,##0.0"},
    )
    _parameters(wb.create_sheet("Параметры"), need_name)
    _write_days(wb.create_sheet("Дни"), daily, holidays, cum)
    wb.save(path)


def main():
    if not os.path.isfile(ROSTER_PATH):
        sys.exit(f"Нет файла состава: {ROSTER_PATH}")
    need_path, _run_id = _year_need_path()
    roster = load_roster(ROSTER_PATH)
    daily, _months = load_need(need_path)
    _check_windows(daily)
    holidays = {date.fromisoformat(item) for item in HOLIDAYS if item.startswith("2027")}
    cum = _check_form(daily, holidays)
    rows = _periods(daily, 7, holidays) + _periods(daily, 14, holidays)
    path = report_path(f"Потребность отпуска {stamp()}.xlsx")
    write_book(path, rows, _slots(roster["people"]), daily, holidays, cum, os.path.basename(need_path))
    print(path)


if __name__ == "__main__":
    main()
