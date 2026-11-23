# -*- coding: utf-8 -*-
"""
Импорт истории (tools\\import_history.bat): файл «Прогноз нагрузки КЦ» исторического файла прогноза
(листы «Прогноз_распределение_Голос/Чат»). Файл лежит в inputs\\.
  * дневной вход КЦ с 2024 г. -> DB\\ext_history\\ext_daily.parquet (пересчитан к нашему определению входа
    по коэффициенту на общих с источник днях; голос — только с VOICE_EXT_FROM, раньше была другая платформа);
  * акции -> events.xlsx: повторяющиеся крупные (PROMO_RECURRING) с действием «акция», остальные — «фон».
Берётся самая свежая версия файла (более старые пропускаются). Строки прошлых импортов в events.xlsx
заменяются, строки, добавленные вручную, не трогаются.
"""
import datetime as dt
import os
import re
import sys

import numpy as np
import openpyxl
import openpyxl.utils
import pandas as pd

import wfm_settings as S
from wfm_common import EVENTS_FILE, error_line, is_db_timeout, store_root
from forecast_core import series_start

HIST_SHEETS = {"Голос": "Прогноз_распределение_Голос", "Чат": "Прогноз_распределение_Чат"}
SERIES = {"Голос": "voice_entry", "Чат": "chat_entry"}


def _write(df, name, fname):
    d = os.path.join(store_root(), name)
    os.makedirs(d, exist_ok=True)
    path = os.path.join(d, fname)
    tmp = path + ".tmp"
    df.to_parquet(tmp, index=False)
    os.replace(tmp, path)
    return path


def _read(name, fname):
    p = os.path.join(store_root(), name, fname)
    return pd.read_parquet(p) if os.path.exists(p) else None


# ---------------------------------------------------------------------------
# 1. Исторический файл прогноза
# ---------------------------------------------------------------------------
def read_history_file(wb):
    rows, events = [], []
    for ch, sh in HIST_SHEETS.items():
        if sh not in wb.sheetnames:
            continue
        for r in wb[sh].iter_rows(min_row=2, max_col=13, values_only=True):
            d = r[5]
            if not hasattr(d, "year"):
                continue
            d = pd.Timestamp(d).normalize()
            if isinstance(r[6], (int, float)):
                rows.append((d, ch, float(r[6])))
            if ch == "Голос":                       # акции размечены одинаково на обоих листах — берём один
                for v in (r[10], r[11]):
                    if isinstance(v, str) and v.strip() and v.strip() != "-":
                        events.append((d, v.strip()))
    return pd.DataFrame(rows, columns=["day", "channel", "value"]), events


def promo_group(text):
    """Группа акции по ключевым словам (wfm_settings.PROMO_GROUPS), первая подходящая; иначе «Другое»."""
    t = str(text).lower()
    for g, v in S.PROMO_GROUPS.items():
        if any(k in t for k in v["keywords"]):
            return g
    return "Другое"


def promo_name(text):
    """-> (название, группа, действие). Сбои — «исключить из истории», остальное — «акция» с группой."""
    name = re.sub(r"\s*\d+$", "", text.strip())[:80]       # «Промо A 12» -> «Промо A (регулярное)»
    if any(k in text.lower() for k in S.OUTAGE_KEYWORDS):
        return name, "", "исключить из истории"
    return name, promo_group(text), "акция"


def group_events(events):
    """Дни с одинаковым названием подряд (разрыв ≤ 3 дней) -> один период. -> [(с, по, название, группа, действие)]."""
    by = {}
    for d, txt in events:
        by.setdefault(promo_name(txt), set()).add(d)
    out = []
    for (name, group, action), days in by.items():
        days = sorted(days)
        a = b = days[0]
        for d in days[1:]:
            if (d - b).days <= 3:
                b = d
            else:
                out.append((a, b, name, group, action))
                a = b = d
        out.append((a, b, name, group, action))
    return sorted(out)


def src_name(path):
    """Папка + имя файла (для комментария в календаре)."""
    return f"{os.path.basename(os.path.dirname(os.path.abspath(path)))}\\{os.path.basename(path)}"


def source_entry():
    """Дневной вход КЦ из хранилища (только источник, без внешней истории)."""
    import wfm_store
    import wfm_forecast
    con = wfm_store.connect()
    h = wfm_forecast.load_history(con, ext=False)
    con.close()
    return h


def import_history_file(path, wb):
    df, events = read_history_file(wb)
    if df.empty:
        print("   В файле нет листов «Прогноз_распределение_Голос/Чат» с данными — пропуск.")
        return "skip"
    new_end = df["day"].max()
    old = _read("ext_history", "ext_daily.parquet")
    old_end = pd.to_datetime(old["day"]).max() if old is not None and len(old) else None
    newest = old_end is None or new_end >= old_end
    if not newest:
        print(f"   Более старая версия (данные до {new_end:%d.%m.%Y}, уже загружены до {old_end:%d.%m.%Y}) — "
              f"история и акции не меняются.")
    else:
        h = source_entry()
        out = []
        for ch, g in df.groupby("channel"):
            s = g.drop_duplicates("day").set_index("day")["value"].sort_index()
            if ch == "Голос":
                s = s[s.index >= pd.Timestamp(S.VOICE_EXT_FROM)]
            source = h[SERIES[ch]].dropna()
            st = series_start(source)
            if st is not None:                       # сверяем только дни, когда ряд источник уже «встал» (+7 дней)
                source = source[source.index >= st + pd.Timedelta(days=7)]
            both = pd.concat([s.rename("ext"), source.rename("source")], axis=1, sort=True).dropna()
            both = both[(both["source"] > 0) & (both["ext"] > 0)]
            k = float((both["ext"] / both["source"]).median()) if len(both) >= 14 else 1.0
            flag = "" if 0.9 <= k <= 1.1 else "   <-- ПОДОЗРИТЕЛЬНО, пришлите вывод"
            print(f"   {ch}: {s.index.min():%d.%m.%Y} – {s.index.max():%d.%m.%Y}, {len(s)} дн.; "
                  f"общих с источник дней {len(both)}, коэффициент файл/источник = {k:.3f}{flag}")
            if len(both) < 14:
                print(f"   ВНИМАНИЕ: {ch} — мало общих с источник дней, коэффициент не посчитан (= 1).")
            out.append(pd.DataFrame({"day": s.index, "series": SERIES[ch], "value": (s / k).values,
                                     "source": src_name(path), "k": k}))
        path_out = _write(pd.concat(out), "ext_history", "ext_daily.parquet")
        print(f"   История записана (самая свежая версия файла): {path_out}")
        add_events(group_events(events), src_name(path))

    return "ok"


def add_group_validation(ws, col_idx):
    """Выпадающий список групп акций в колонке «Группа акции»."""
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    letter = get_column_letter(col_idx)
    ws.data_validations.dataValidation = [d for d in ws.data_validations.dataValidation
                                          if letter not in str(d.sqref)]
    dv = DataValidation(type="list", formula1='"' + ",".join(S.PROMO_GROUPS) + '"', allow_blank=True)
    dv.add(f"{letter}2:{letter}1000")
    ws.add_data_validation(dv)


def add_events(ev, src):
    """Акции из файла -> events.xlsx. Строки прошлых импортов (Комментарий «импорт из …») заменяются целиком,
    строки, добавленные вручную, не трогаются."""
    if not ev:
        return
    if not os.path.exists(EVENTS_FILE):
        import wfm_menu
        wfm_menu.make_events_template()
    wb = openpyxl.load_workbook(EVENTS_FILE)
    ws = wb.worksheets[0]
    hdr = [c.value for c in ws[1]]
    for h_ in ("Название", "Группа акции"):
        if h_ not in hdr:
            ws.cell(1, len(hdr) + 1, h_)
            hdr.append(h_)
    col = {h: i + 1 for i, h in enumerate(hdr)}
    ci = col.get("Комментарий")
    old_rows = [i for i in range(ws.max_row, 1, -1)
                if ci and str(ws.cell(i, ci).value or "").startswith("импорт из")]
    for i in old_rows:                                  # снизу вверх, чтобы номера строк не съезжали
        ws.delete_rows(i)
    have = set()
    for r in ws.iter_rows(min_row=2, values_only=True):    # ручные строки: не дублировать
        if r[0] is not None:
            nm = r[col["Название"] - 1] if len(r) >= col["Название"] else None
            have.add((str(nm or "").upper(), pd.Timestamp(r[0]).date()))
    n_new = 0
    for a, b, name, group, action in ev:
        if (name.upper(), a.date()) in have:
            continue
        row = ws.max_row + 1
        vals = {"Дата с": a.date(), "Дата по": b.date(), "Канал": "Все",
                "Тип": "сбой" if action == "исключить из истории" else "акция", "Действие": action,
                "Название": name, "Группа акции": group, "Комментарий": f"импорт из {src}"}
        for k, v in vals.items():
            if k in col:
                ws.cell(row, col[k], v)
        n_new += 1
    add_group_validation(ws, col["Группа акции"])
    ws.column_dimensions[openpyxl.utils.get_column_letter(col["Группа акции"])].width = 24
    wb.save(EVENTS_FILE)
    by_g = {}
    for e in ev:
        by_g[e[3] or "сбой"] = by_g.get(e[3] or "сбой", 0) + 1
    used = [g for g, v in S.PROMO_GROUPS.items() if v["use"]]
    print(f"   Календарь: прошлый импорт заменён ({len(old_rows)} строк удалено), записано {n_new} строк. Ваши строки не тронуты.")
    print("   По группам: " + ", ".join(f"{g} {n}" for g, n in by_g.items())
          + f". В прогнозе участвуют: {', '.join(used)}.")


# ---------------------------------------------------------------------------
def main(paths):
    if not paths and sys.stdin.isatty():
        print("Перетащите в это окно файлы или папку целиком и нажмите Enter (можно несколько, через пробел).")
        print("Просто Enter — взять файлы из папки inputs.")
        raw = input("> ").strip()
        if raw:
            paths = re.findall(r'"([^"]+)"|(\S+)', raw)
            paths = [a or b for a, b in paths]
        else:
            paths = [os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "inputs")]
    if not paths:
        print("Файлы не указаны.")
        return
    files = []
    for p in paths:
        p = p.strip().strip('"')
        if os.path.isdir(p):                       # папка — все Excel-файлы внутри, включая подпапки
            for root, _, names in os.walk(p):
                files += [os.path.join(root, n) for n in sorted(names) if n.lower().endswith((".xlsx", ".xlsm"))]
        else:
            files.append(p)
    ok = skipped = 0
    for p in files:
        name = os.path.basename(p)
        print(f"\n{name}")
        if name.startswith("~$"):
            print("   Временный файл Excel (открыт у кого-то) — пропуск.")
            skipped += 1
            continue
        if not os.path.exists(p):
            print("   Файл не найден.")
            skipped += 1
            continue
        if not name.lower().endswith((".xlsx", ".xlsm")):
            print("   Формат не поддерживается (нужен .xlsx или .xlsm; старый .xls — пересохранить в Excel).")
            skipped += 1
            continue
        try:
            wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        except Exception as e:
            print(f"   Не открылся: {e}")
            skipped += 1
            continue
        if any(sh in wb.sheetnames for sh in HIST_SHEETS.values()):
            res = import_history_file(p, wb)
        else:
            print("   Это не файл «Прогноз нагрузки КЦ» (нет листов «Прогноз_распределение_Голос/Чат») — пропуск.")
            res = "skip"
        wb.close()
        if res == "ok":
            ok += 1
        else:
            skipped += 1
    print(f"\nЗагружено файлов: {ok}; пропущено: {skipped}.")


if __name__ == "__main__":
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    try:
        main(sys.argv[1:])
    except Exception as e:
        if not is_db_timeout(e):
            import traceback
            traceback.print_exc()
        print("\n" + error_line(e))
        sys.exit(1)
