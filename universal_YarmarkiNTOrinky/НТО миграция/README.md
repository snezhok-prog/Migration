# Universal NTO Migration

Скрипт миграции для НТО с унифицированным каркасом запуска, чекпоинтами, resume, operator-mode и логированием в стиле `universal_myPet`.

## Структура

- `migration.py` — основной entrypoint (рекомендуемый запуск)
- `nto_migration.py` — реализация миграции
- `_api.py` — API клиент и авторизация (runtime URL по профилям)
- `_profiles.py` — профили стендов `dev|psi|prod|custom`
- `_excel_input.py` — поиск XLSM книг
- `_state.py` — checkpoints/resume
- `_logger.py` — script/success/fail/user логи
- `_vba_nto.py` / `_nto_mappings.py` — маппинг данных (не менялся)
- `rollback.py` — откат записей по success логам
- `state/checkpoints.json` — checkpoint состояние
- `logs/*.txt` — журналы запусков

## Быстрый запуск

Из папки `universal_YarmarkiNTOrinky/НТО миграция`:

```bash
python migration.py --profile dev
```

## Основные режимы

- `--mode auto|single|mass` — выбор одной или нескольких книг
- `--workbooks "a.xlsm;b.xlsm"` — явный список книг
- `--files-dir "..."` — базовая папка файлов
- `--files-map "book1.xlsm=dir1;book2.xlsm=dir2"` — привязка книга->папка
- `--ask-files-always` — всегда спрашивать папку файлов для каждой книги

## Авторизация

- `--profile dev|psi|prod|custom`
- `--base-url`, `--jwt-url`, `--ui-base-url` — переопределение URL
- `--auth-only` — проверить авторизацию и завершить
- `--skip-auth` — пропустить авторизацию (только с `--dry-run`)
- `--no-prompt` — брать cookie/token из `cookie.md`/`token.md` без вопросов

## Resume / checkpoints

- `--resume` / `--no-resume`
- `--state-file state/checkpoints.json`
- `--reset-state` — очистить checkpoints перед стартом

Если предыдущий запуск был незавершен, при интерактивном режиме предлагается:
- продолжить с checkpoint
- сбросить checkpoint и начать заново
- выйти

## Operator mode

- `--operator-mode` — на ошибке строки: `retry / skip / abort`

Пример:

```bash
python migration.py --profile dev --operator-mode
```

## Dry-run

```bash
python migration.py --profile dev --dry-run --skip-auth
```

## Откат

Откат по success логам:

```bash
python rollback.py --profile dev
```

Можно переопределить шаблон логов:

```bash
python rollback.py --profile dev --success-log-glob "logs/success_log-2026-04-14_*.txt"
```
