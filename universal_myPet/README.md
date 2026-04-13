# universal_myPet — миграция «Мой питомец» (Python)

`universal_myPet` — рабочий проект миграции по сервису «Мой питомец».
Скрипт объединяет:
- бизнес-логику и маппинг старых JS/VBA,
- удобный CLI и устойчивость (resume, operator flow),
- автоматизацию запросов без Postman.

Подробный эксплуатационный гайд: [OPS_GUIDE_RU.md](./OPS_GUIDE_RU.md).

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
- Формирует отдельный пользовательский журнал `logs/user_log-*.txt` для службы эксплуатации:
  - таблица/лист/строка,
  - поле и файл с ошибкой,
  - действие после ошибки (повтор/пропуск/остановка/автопродолжение),
  - мигрировалась ли строка,
  - подсказки по resume и rollback.
- Поддерживает resume (`state/checkpoints.json`) и повторный запуск «с места остановки».
- Поддерживает массовый запуск по нескольким книгам и привязку разных папок `files` к каждой книге.

## Важные поведенческие моменты

- В интерактивном запуске (`без --no-interactive`) скрипт:
  - спрашивает папку файлов для книги,
  - включает operator-flow (retry/skip/abort) по ошибкам строк.
- В неинтерактивном запуске (`--no-interactive`) operator prompts отключены.
- Ошибки загрузки файлов (включая `413 Request Entity Too Large`) явно попадают в `Errors` и в `fail_log`.
- При частичных upload-ошибках скрипт продолжает попытки загрузки остальных файлов строки, а затем возвращает агрегированную ошибку по неуспешным файлам.

## Быстрый старт

Установка:

```bash
pip install -r requirements.txt
```

Проверка авторизации:

```bash
python migration.py --profile psi --auth-only
```

Обычная миграция (интерактивно):

```bash
python migration.py
```

Dry-run без API:

```bash
python migration.py --dry-run --skip-auth
```

По умолчанию профиль запуска: `psi`.

## Переключение стенда (PSI / DEV / PROD)

Переключение делается флагом `--profile`:

```bash
python migration.py --profile psi
python migration.py --profile dev
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

Важно: межреестровые ссылки в данных (`...RecordLink`, `animalRegistryLink`) теперь формируются по публичному UI-домену профиля
(`psi.pgs.gosuslugi.ru`, `pgs.gosuslugi.ru`, `iam...dev`), а не по внутреннему `*-inner` API-домену.
При `--profile custom` можно явно задать публичный домен флагом:

```bash
python migration.py --profile custom --base-url "https://internal-stand" --jwt-url "https://internal-stand/jwt/" --ui-base-url "https://public-ui-stand"
```

## Передача в эксплуатацию

Перед передачей на новую ВМ:
- заполните `token.md` и `cookie.md` актуальными значениями для нужного стенда;
- убедитесь, что установлен Python и зависимости из `requirements.txt`;
- запускайте из папки проекта, пути внутри скрипта относительные (привязки к вашему ПК не требуются).

Базовая команда для PSI:

```bash
python migration.py --profile psi --no-prompt --no-interactive
```

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
- `Повторить` (`п`, `retry`)
- `Пропустить` (`пр`, `skip`) — строка не будет мигрирована в этой попытке
- `Дальше` для нефатальных ошибок (`д`, `continue`)
- `Остановить` (`о`, `abort`)
- При `Остановить` скрипт явно пишет, что прогресс сохранен в `state/checkpoints.json`.
- Если основная запись строки уже создана, checkpoint сохраняется даже при ошибке загрузки файла (например, `413`), чтобы при повторном запуске не создать дубль.

### Режим "ошибка по строке"

Этот режим уже реализован в operator-flow:

- при ошибке строки оператор выбирает: `Повторить` / `Пропустить` / `Остановить`;
- при нефатальной ошибке после частичной записи строки: `Дальше` / `Остановить`;
- каждое решение фиксируется в `user_log` с итогом по строке (`мигрирована`/`не мигрирована`).

Важно:
- если ошибка случилась после создания основной записи (например, на загрузке файла), строка может быть частично мигрирована;
- при остановке скрипт пишет в `user_log`, как продолжить с `checkpoints` и как сделать откат.

## Resume / checkpoints

По умолчанию resume включен (кроме `--dry-run`).

При обычном интерактивном запуске скрипт автоматически проверяет незавершенную прошлую миграцию и предлагает:
- `Продолжить по checkpoint`;
- `Сбросить checkpoint и начать заново`;
- `Выйти`.

После полного успешного завершения миграции checkpoints очищаются автоматически.  
После остановки/ошибки checkpoints сохраняются.

Полезные флаги:

```bash
python migration.py --reset-state
python migration.py --no-resume
python migration.py --state-file "custom_checkpoints.json"
```

Для нестандартного стенда отдельно задайте публичный URL для записываемых ссылок:

```bash
python migration.py --profile custom --ui-base-url "https://public-ui-stand"
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
- `logs/user_log-*.txt` — понятный журнал для эксплуатации (что сломалось, где, и что сделал скрипт/оператор).
- `ROLLBACK_BODY.json` — данные для отката.

## Совместимость

Скрипт адаптирован под продовый набор библиотек, включая:
- `requests==2.32.3`
- `openpyxl==2.6.2`

В `_excel_input.py` добавлен `numpy`-compat patch для старого `openpyxl`.

## Сверка паритета

Матрица соответствия старой логики JS/VBA и Python-реализации:

- `MAPPING_PARITY.md`
