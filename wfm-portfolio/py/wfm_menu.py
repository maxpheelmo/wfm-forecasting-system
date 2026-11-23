# -*- coding: utf-8 -*-
"""
Консоль WFM (журнал решений №33). Запуск: wfm.bat в корне папки.

Главное меню — по одной кнопке на цикл:
  1. План на квартал     — свежие данные источника → прогноз квартала по 30 мин
  2. Недельный пересчёт  — свежие данные источника → прогноз на 2 недели
  3. План на год         — календарный год, с июля следующий. Сам не публикует и уже опубликованный год не затирает
  4. Точность            — сохранённые версии против факта, только по уже прошедшим дням
  5. Потребность         — 1 квартал (последний опубликованный) или 2 год (опубликованный год)
  6. Отпуска             — окна по готовой потребности года. Прогноз не пересчитывает
  7. Настройки           — данные источника, календарь, пароль, публикация, показатели людей

Пункты 1–4 догружают данные источника, если данные старше FRESH_HOURS. Пункты 5 и 6 читают уже посчитанные книги.

Структура папки:
  wfm.bat       — эта консоль (единственный файл в корне)
  py\\          — скрипты и SQL-представления (views); README.txt — описание папки
  tools\\       — батники разовых проверок (бэктест, проверка хранилища, обзор рядов, импорт истории)
  inputs\\      — входные данные помимо источника: events.xlsx, need.txt, состав, план клиентов, отсутствия
  DB\\          — хранилище данных (не трогать руками)
  reports\\     — отчёты по видам: forecast\\ (прогноз, бэктест, точность), need\\, checks\\, other\\
"""
import csv
import datetime as dt
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from wfm_common import EVENTS_FILE, INPUTS_DIR, REPORTS_DIR, error_line, is_db_timeout, store_root   # noqa: E402

FRESH_HOURS = 3          # данные источника свежее — не догружаем перед пунктом


# ---------------------------------------------------------------------------
# Состояние
# ---------------------------------------------------------------------------
def last_load():
    """Когда последний раз обновлялись данные (по журналу загрузчика)."""
    path = os.path.join(store_root(), "_log", "load_runs.csv")
    try:
        with open(path, encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f, delimiter=";"))
        return dt.datetime.fromisoformat(rows[-1]["run"]) if rows else None
    except Exception:
        return None


def last_published():
    """Последний опубликованный год и последний опубликованный квартал."""
    try:
        import pandas as pd
        d = os.path.join(store_root(), "forecast_run")
        files = [os.path.join(d, f) for f in os.listdir(d) if f.endswith(".parquet")]
        if not files:
            return None, None
        frames = []
        for f in files:
            df = pd.read_parquet(f)
            keep = [c for c in ("created_at", "published_at", "kind", "status") if c in df.columns]
            if "kind" in keep and "status" in keep and "created_at" in keep:
                frames.append(df[keep])
        if not frames:
            return None, None
        r = pd.concat(frames, ignore_index=True)
        r = r[r["status"] == "опубликован"].copy()
        r["created_at"] = pd.to_datetime(r["created_at"])
        if "published_at" not in r.columns:
            r["published_at"] = pd.NaT
        r["published_at"] = pd.to_datetime(r["published_at"])
        r["when"] = r["published_at"].fillna(r["created_at"])

        def latest(kind):
            part = r[r["kind"] == kind].sort_values("when")
            return part.iloc[-1] if len(part) else None

        return latest("модель: год"), latest("модель: квартал")
    except Exception:
        return None, None


def fmt(t):
    try:
        return f"{t:%d.%m.%Y %H:%M:%S}"
    except Exception:
        return str(t)


def migrate_events():
    """events.xlsx переехал из корня в inputs: перенести сразу, иначе прогноз пойдёт без календаря."""
    old = os.path.join(ROOT, "events.xlsx")
    if os.path.exists(old) and not os.path.exists(EVENTS_FILE):
        os.makedirs(INPUTS_DIR, exist_ok=True)
        shutil.move(old, EVENTS_FILE)
        print("events.xlsx перенесён из корня в папку inputs.")


# ---------------------------------------------------------------------------
# Экран
# ---------------------------------------------------------------------------
def line(ch="-"):
    print(ch * 66)


def header(title="WFM — консоль"):
    os.system("cls" if os.name == "nt" else "clear")
    line("=")
    print(f"  {title}")
    line("=")
    t = last_load()
    if t:
        print(f"  Данные источник: последнее обновление {fmt(t)}.")
    else:
        print("  Данные источник: обновлений ещё не было.")
    year, quarter = last_published()
    print("  Опубликованные прогнозы:")
    print("  Год от " + (fmt(year["when"]) if year is not None else "нет"))
    print("  Квартал от " + (fmt(quarter["when"]) if quarter is not None else "нет"))
    if not os.path.exists(EVENTS_FILE):
        print("  Календаря событий нет (Настройки → 2 создаст)")
    line()


def main_menu():
    header()
    print("  1. План на квартал     — прогноз квартала по 30 мин (голос, чат, почта)")
    print("  2. Недельный пересчёт  — ближайшие 2 недели по свежим неделям")
    print("  3. План на год         — календарный год, с июля следующий. Сам не публикует")
    print("  4. Точность            — сохранённые версии против факта. Будущие месяцы не оценивает")
    print("  5. Потребность         — люди по 30 мин. Спросит: 1 квартал или 2 год")
    print("  6. Отпуска             — окна 7, 14 и 7 дней. Год и квартал не пересчитывает")
    print()
    print("  7. Настройки")
    print("  0. Выход")
    line()
    print("  Разовые проверки (бэктест, хранилище, обзор рядов, импорт истории) — папка tools.")
    line()


def need_menu():
    header("WFM — потребность")
    print("  Сформировать потребность:")
    print("  1. Квартал — последний опубликованный квартал, люди на его даты")
    print("  2. Год     — опубликованный прогноз года")
    print("  0. Назад")
    line()


def settings_menu():
    header("WFM — настройки")
    print("  1. Обновить данные источника сейчас")
    print("  2. Календарь событий (events.xlsx): акции, сбои, смена режима")
    print("  3. Настройки модели и очередей (wfm_settings.py, откроется в Блокноте)")
    print("  4. Сохранить пароль от базы (один раз на компьютере)")
    print("  5. Открыть папку с отчётами")
    print("  6. Опубликовать версию прогноза (после проверки; по ней строятся потребность и графики)")
    print("  7. Показатели потребности (inputs\\need.txt, откроется в Блокноте)")
    print("  0. Назад")
    line()


# ---------------------------------------------------------------------------
# Действия
# ---------------------------------------------------------------------------
def run_script(script, args=()):
    print()
    code = subprocess.call([sys.executable, os.path.join(HERE, script), *args], cwd=HERE)
    if code == 0:
        print("\nГотово.")
    return code


def ask(prompt):
    return input(prompt).strip().lower() in ("д", "да", "y", "yes", "l")


def ensure_fresh():
    """Догрузить данные источника, если они старше FRESH_HOURS. -> можно продолжать?"""
    t = last_load()
    if t and (dt.datetime.now() - t).total_seconds() < FRESH_HOURS * 3600:
        print(f"\nДанные источник свежие ({fmt(t)}) — загрузку пропускаю.")
        return True
    print("\nОбновляю данные из источник...")
    if run_script("wfm_load.py") == 0:
        return True
    return ask("Продолжить на старых данных? (д/н): ")


def open_path(path, folder=False):
    """Открыть папку (создав её при необходимости) или файл в Windows; в других ОС — показать путь."""
    if folder:
        os.makedirs(path, exist_ok=True)
    if os.name == "nt":
        os.startfile(path)
    else:
        print(f"Путь: {path}")


def edit_settings():
    path = os.path.join(HERE, "wfm_settings.py")
    print("\nПосле правки сохраните файл; новые настройки действуют со следующего запуска пункта.")
    if os.name == "nt":
        subprocess.Popen(["notepad.exe", path])      # не startfile: .py запустился бы, а не открылся
    else:
        print(f"Путь: {path}")


def make_events_template():
    import pandas as pd
    from openpyxl.styles import Font, PatternFill
    rows = [
        {"Дата с": dt.date(2026, 10, 14), "Дата по": dt.date(2026, 10, 15), "Канал": "Все", "Тип": "сбой",
         "Действие": "исключить из истории", "Название": "сбой", "Группа акции": None, "Множитель": None,
         "Комментарий": "сбой 14–15.10"},
    ]
    with pd.ExcelWriter(EVENTS_FILE, engine="openpyxl") as w:
        df = pd.DataFrame(rows)
        df.to_excel(w, sheet_name="events", index=False)
        ws = w.book["events"]
        for c in ws[1]:
            c.fill, c.font = PatternFill("solid", fgColor="1F3864"), Font(bold=True, color="FFFFFF")
        for col, width in (("A", 12), ("B", 12), ("C", 10), ("D", 10), ("E", 22), ("F", 40), ("G", 24), ("H", 11),
                           ("I", 60)):
            ws.column_dimensions[col].width = width
        from import_history import add_group_validation
        add_group_validation(ws, 7)
        hs = w.book.create_sheet("Как заполнять")
        help_rows = [
            ("Что сюда писать", "Повторяющиеся крупные акции (прошлые и будущие) и сбои. Голосового робота модель видит "
                                "в данных сама. Мелкие акции и обычные всплески модель учитывает сама — их можно не писать."),
            ("Дата с / Дата по", "период (включительно); для смены режима «Дата по» пустая"),
            ("Канал", "Все / Голос / Чат / Почта"),
            ("Тип", "акция / бот / сбой / праздник / прочее (для справки)"),
            ("Действие", "«акция» — прошлые даты исключаются из уровня, будущие умножаются на эффект прошлых запусков "
                         "той же ГРУППЫ (один множитель на период, осторожно, если запусков мало); "
                         "«смена режима» — с этой даты доля до очереди меняется (бот чата, новый робот, которого нет в настройках), "
                         "в Множитель — новая доля или пусто; "
                         "«исключить из истории» — прошлый период не использовать (сбой); "
                         "«множитель» — будущие даты × Множитель; «фон» — для справки, модель не использует"),
            ("Название", "любое, как в рассылке (для себя) — модель на него не смотрит"),
            ("Группа акции", "выбрать из списка. По группе модель ищет прошлые запуски. Сейчас в прогнозе участвует "
                             "только «Промо A (регулярное)» (у остальных групп эффект нестабилен — бэктест 23.10); "
                             "если группу не выбрать — определится по словам в названии"),
            ("Множитель", "обычно пусто (модель считает сама). Если акция новая или крупнее обычной — укажите, "
                          "например 1,3: тогда она учтётся даже для группы, которая сама не участвует"),
            ("Пример", "21.02.2027 | 07.03.2027 | Все | акция | акция | Весенняя кампания | Промо A (регулярное) | | рассылка"),
            ("Пример", "10.03.2027 | 20.03.2027 | Все | акция | акция | Реактивация (рассылка) | Реактивация (рассылка) | 1,15 | "
                       "своя оценка"),
        ]
        for r, (a, b) in enumerate(help_rows, 1):
            hs.cell(r, 1, a).font = Font(bold=True)
            hs.cell(r, 2, b)
        hs.column_dimensions["A"].width, hs.column_dimensions["B"].width = 18, 140


def publish_menu():
    """Список последних версий прогноза и публикация выбранного черновика (№34)."""
    import wfm_forecast as WF
    r = WF.list_runs()
    if r.empty:
        print("\nВерсий прогноза ещё нет — сначала «План на квартал».")
        return
    r = r.head(10)
    print("\nПоследние версии прогноза:")
    for i, x in r.iterrows():
        dto = f"{pd_ts(x['data_to']):%d.%m.%Y}" if pd_ts(x["data_to"]) is not None else "—"
        kind = str(x["kind"]).replace("модель: ", "")
        print(f"  {i + 1}. {pd_ts(x['created_at']):%d.%m.%Y %H:%M}  {kind:<8} {pd_ts(x['period_from']):%d.%m.%Y}–"
              f"{pd_ts(x['period_to']):%d.%m.%Y}  данные по {dto}  [{x['status']}]")
    k = input("Номер версии для публикации (Enter — отмена): ").strip()
    if not k.isdigit() or not 1 <= int(k) <= len(r):
        print("Отменено.")
        return
    x = r.iloc[int(k) - 1]
    if x["status"] == "опубликован":
        print("Эта версия уже опубликована — опубликованные не меняются (№34).")
        return
    print(f"Модель: {x.get('model') or '—'}")
    comment = input("Комментарий (например, «Q1 2027, проверено, к публикации»): ").strip()
    if not ask(f"Опубликовать версию от {pd_ts(x['created_at']):%d.%m.%Y %H:%M}? (д/н): "):
        print("Отменено.")
        return
    y = WF.publish_run(x["run_id"], comment)
    print(f"Опубликовано: версия от {pd_ts(y['created_at']):%d.%m.%Y %H:%M} (номер {y['run_id']}), {y['published_at']:%d.%m.%Y %H:%M}.")


def pd_ts(v):
    import pandas as pd
    try:
        t = pd.Timestamp(v)
        return None if pd.isna(t) else t
    except Exception:
        return None


def settings_loop():
    while True:
        settings_menu()
        choice = input("Выберите пункт и нажмите Enter: ").strip().lower()
        if choice in ("0", "", "q", "й"):
            return
        try:
            if choice == "1":
                run_script("wfm_load.py")
            elif choice == "2":
                if not os.path.exists(EVENTS_FILE):
                    make_events_template()
                    print("\nСоздан events.xlsx с примерами.")
                open_path(EVENTS_FILE)
                if os.name == "nt":
                    continue
            elif choice == "3":
                edit_settings()
            elif choice == "4":
                run_script("save_password.py")
            elif choice == "5":
                open_path(REPORTS_DIR, folder=True)
                continue
            elif choice == "6":
                publish_menu()
            elif choice == "7":
                path = os.path.join(INPUTS_DIR, "need.txt")
                print("\nПосле правки сохраните файл. Следующий запуск «Потребность» прочитает его заново.")
                if os.path.exists(path):
                    open_path(path)
                    if os.name == "nt":
                        continue
                else:
                    print(f"Нет файла: {path}")
            else:
                continue
        except Exception as e:                      # меню не должно закрываться из-за ошибки пункта
            print("\n" + error_line(e))
            if isinstance(e, PermissionError):
                print("Возможно, файл открыт в Excel — закройте его и повторите.")
        input("\nНажмите Enter, чтобы вернуться в меню...")


def main():
    for d in (INPUTS_DIR, REPORTS_DIR):
        os.makedirs(d, exist_ok=True)
    migrate_events()
    while True:
        main_menu()
        choice = input("Выберите пункт и нажмите Enter: ").strip().lower()
        try:
            if choice == "1":
                if ensure_fresh():
                    run_script("wfm_forecast.py")
            elif choice == "2":
                if ensure_fresh():
                    run_script("wfm_forecast.py", ["--weekly"])
            elif choice == "3":
                if ensure_fresh():
                    run_script("cr_volume.py", ["--open-year"])
            elif choice == "4":
                if ensure_fresh():
                    run_script("accuracy.py")
            elif choice == "5":
                need_menu()
                sub = input("Выберите пункт и нажмите Enter: ").strip().lower()
                if sub == "1":
                    print("\nПотребность квартала. Прогноз заново не считается.")
                    run_script("need.py", ["--quarter"])
                elif sub == "2":
                    print("\nПотребность года. Опубликованный квартал не меняется.")
                    run_script("need.py", ["--year"])
                elif sub in ("0", "", "q", "й"):
                    continue
            elif choice == "6":
                print("\nОкна отпуска. Год и квартал эта кнопка не меняет.")
                run_script("vacation.py")
            elif choice == "7":
                settings_loop()
                continue
            elif choice in ("0", "q", "й"):
                return
            else:
                continue
        except Exception as e:
            print("\n" + error_line(e))
            if isinstance(e, PermissionError):
                print("Возможно, файл открыт в Excel — закройте его и повторите.")
        input("\nНажмите Enter, чтобы вернуться в меню...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
    except Exception as e:
        if not is_db_timeout(e):
            import traceback
            traceback.print_exc()
        print("\n" + error_line(e))
        input("Нажмите Enter, чтобы закрыть...")
