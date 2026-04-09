# universal_myPet — миграция «Мой питомец» (Python)

`universal_myPet` — рабочий проект миграции по сервису «Мой питомец».
Скрипт объединяет:
- бизнес-логику и маппинг старых JS/VBA,
- удобный CLI и устойчивость (resume, operator flow),
- автоматизацию запросов без Postman.

## Что умеет

- Читает данные напрямую из `.xlsm` (листы 2/3/4) по логике `old makros.vba`.
- Создает записи во всех целевых коллекциях:
  - `animalCatchOrderRegistryCollection`
  - `animalsRecordsCollectionTwo`
  - `animalCatchActRegistryCollection`
  - `myPetAnimalCardReestr`
  - `animalTransferActRegistryCollection`
  - `animalReleaseActRegistryEntry`
- Загружает файлы через `POST /api/v1/storage/upload` (multipart), с fallback в base64 (настройка в `_config.py`).
- Формирует `ROLLBACK_BODY.json` и детальные логи (`Created`, `Errors`, success/fail logs).
- Поддерживает resume (`state/checkpoints.json`) и повторный запуск «с места остановки».
- Поддерживает массовый запуск по нескольким книгам и привязку разных папок `files` к каждой книге.

## Важные поведенческие моменты

- В интерактивном запуске (`без --no-interactive`) скрипт:
  - спрашивает папку файлов для книги,
  - включает operator-flow (retry/skip/abort) по ошибкам строк.
- В неинтерактивном запуске (`--no-interactive`) operator prompts отключены.
- Ошибки загрузки файлов (включая `413 Request Entity Too Large`) явно попадают в `Errors` и в `fail_log`.

## Быстрый старт

Установка:

```bash
pip install -r requirements.txt
```

Проверка авторизации:

```bash
python migration.py --profile dev --auth-only
```

Обычная миграция (интерактивно):

```bash
python migration.py
```

Dry-run без API:

```bash
python migration.py --dry-run --skip-auth
```

## Переключение стенда (DEV / PSI / PROD)

Переключение делается флагом `--profile`:

```bash
python migration.py --profile dev
python migration.py --profile psi
python migration.py --profile prod
```

Для очистки коллекций:

```bash
python clear_collections.py --profile dev
python clear_collections.py --profile psi
python clear_collections.py --profile prod
```

Если нужен нестандартный стенд:

```bash
python migration.py --profile custom --base-url "https://your-stand" --jwt-url "https://your-stand/jwt/"
```

Профили и URL заданы в `_profiles.py`.

## Массовая миграция и папки файлов

Пример с явной привязкой workbook -> папка файлов:

```bash
python migration.py --profile dev --mode mass --workbooks "book1.xlsm;book2.xlsm" --files-map "book1.xlsm=one;book2.xlsm=two"
```

Если `--files-map` не указан и запуск интерактивный, скрипт попросит выбрать папку файлов.

## Operator режим

Явно включить operator-mode можно флагом:

```bash
python migration.py --profile dev --operator-mode
```

В operator-flow доступны действия:
- `Retry`
- `Skip`
- `Abort`

## Resume / checkpoints

По умолчанию resume включен (кроме `--dry-run`).

Полезные флаги:

```bash
python migration.py --reset-state
python migration.py --no-resume
python migration.py --state-file "custom_checkpoints.json"
```

## Очистка и откат

Очистка коллекций:

```bash
python clear_collections.py --profile dev
python clear_collections.py --profile dev --dry-run
```

Откат:

```bash
python rollback.py
# или по частям
python rollback_orders.py
python rollback_stray.py
python rollback_cards.py
```

## Логи

- `logs/script_creation_log-*.txt` — основной лог запуска.
- `logs/success_log-*.txt` — успешные записи.
- `logs/fail_log-*.txt` — ошибки.
- `ROLLBACK_BODY.json` — данные для отката.

## Совместимость

Скрипт адаптирован под продовый набор библиотек, включая:
- `requests==2.32.3`
- `openpyxl==2.6.2`

В `_excel_input.py` добавлен `numpy`-compat patch для старого `openpyxl`.

## Сверка паритета

Матрица соответствия старой логики JS/VBA и Python-реализации:

- `MAPPING_PARITY.md`
