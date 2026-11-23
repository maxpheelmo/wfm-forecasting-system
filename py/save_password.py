# -*- coding: utf-8 -*-
"""
Один раз на этом компьютере: сохранить пароль от реплики PostgreSQL
в хранилище учётных данных Windows.

Пароль сначала проверяется подключением к базе (только чтение) и сохраняется,
только если подключение удалось. В файлах пароля нет.

Посмотреть или удалить: Панель управления → Диспетчер учётных данных →
Учётные данные Windows → имя сервиса из KEYRING_SERVICE (по умолчанию «wfm_source_db»).

Все скрипты консоли берут пароль отсюда автоматически. Переменная SOURCE_PG_PASSWORD,
если задана, имеет приоритет над хранилищем (удобно в CI, не храните её в файлах).
"""
import getpass
import sys

from wfm_common import error_line
import wfm_settings as S


def main():
    import keyring
    import psycopg

    default_user = S.SOURCE_DB["user"]
    user = input(f"Логин БД [{default_user}]: ").strip() or default_user
    pw = getpass.getpass(f"Пароль для {user} (ввод не отображается): ")
    print("Проверяю подключение...")
    try:
        db = dict(S.SOURCE_DB, user=user)
        with psycopg.connect(**db, password=pw, connect_timeout=15,
                             options="-c default_transaction_read_only=on") as conn:
            conn.execute("select 1")
    except Exception as e:
        print(error_line(e))
        print("Пароль не сохранён.")
        sys.exit(1)
    keyring.set_password(S.KEYRING_SERVICE, user, pw)
    print(f"Подключение есть. Пароль для {user} сохранён в хранилище учётных данных Windows "
          f"({S.KEYRING_SERVICE}).")
    if user != default_user:
        print(f"Логин не {default_user}: перед запуском задайте SOURCE_PG_USER={user}.")


if __name__ == "__main__":
    main()
