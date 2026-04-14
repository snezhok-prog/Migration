# Universal Fairs Migration

Скрипт миграции для реестров ярмарок с унифицированным каркасом запуска, checkpoints/resume, operator-mode и логированием в стиле `universal_myPet`.

## Структура

- `migration.py` — основной entrypoint миграции
- `fair_migration.py` — реализация миграции ярмарок
- `rollback.py` — откат записей по success-логам
- `_api.py` — API клиент и авторизация (runtime URL по профилям)
- `_profiles.py` — профили стендов `dev|psi|prod|custom`
- `_excel_input.py` — поиск XLSM книг для single/mass режима
- `_state.py` — checkpoints/resume состояние
- `_logger.py` — script/success/fail/user/rollback логгеры
- `_config.py` — базовая конфигурация
- `_utils.py` — вспомогательные функции
- `_templates.py` — шаблоны сущностей
- `files/` — папка с файлами для загрузки
- `logs/` — логи запусков

## Быстрый запуск

Из каталога `universal_YarmarkiNTOrinky/Ярмарки миграция`:

```bash
python migration.py --profile dev
```

## Режимы запуска

- `--mode auto|single|mass` — выбор одной или нескольких книг
- `--workbook "<file.xlsm>"` — совместимый запуск одной книги
- `--workbooks "a.xlsm;b.xlsm"` — явный список книг
- `--files-dir "<dir>"` — базовая папка файлов
- `--files-map "book1.xlsm=dir1;book2.xlsm=dir2"` — привязка книга->папка
- `--ask-files-always` — спрашивать папку файлов для каждой книги

## Авторизация и профили

- `--profile dev|psi|prod|custom`
- `--base-url`, `--jwt-url`, `--ui-base-url` — переопределение URL
- `--auth-only` — проверить авторизацию и завершить
- `--skip-auth` — пропустить авторизацию (только с `--dry-run`)
- `--no-prompt` — читать `cookie.md`/`token.md` без вопросов

## Checkpoints / Resume

- `--resume` / `--no-resume`
- `--state-file state/checkpoints.json`
- `--reset-state` — очистить checkpoints перед запуском

Если предыдущий запуск был оборван/неуспешен, в интерактивном режиме скрипт предложит:

1. продолжить с checkpoint
2. сбросить checkpoints
3. выйти

## Operator mode

```bash
python migration.py --profile dev --operator-mode
```

На ошибках строк можно выбрать `retry / skip / abort`.

## Dry-run

```bash
python migration.py --profile dev --dry-run --skip-auth
```

## Откат

```bash
python rollback.py --profile dev
```

С фильтром success-логов:

```bash
python rollback.py --profile dev --success-log-glob "logs/success_log-2026-04-14_*.txt"
```

## Логи

- `logs/script_creation_log-*.txt` — технический лог
- `logs/user_log-*.txt` — лог для оператора
- `logs/success_log-*.txt` — успешно созданные записи
- `logs/fail_log-*.txt` — ошибки строк
- `logs/script_rollback_log-*.txt` — лог отката
