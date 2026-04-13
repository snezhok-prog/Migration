# Практический гайд по скрипту migration.py (universal_myPet)

Документ для службы эксплуатации: как запускать миграцию, как реагировать на ошибки и как безопасно продолжать/откатывать.

## 1. Что делает скрипт

`migration.py` переносит данные из Excel (листы 2/3/4) или JSON-part файлов в ПГС:
- заказ-наряды на отлов,
- реестр животных без владельцев,
- карточки животных,
- связанные акты (отлов, передача, выпуск, смерть, передача владельцу),
- загрузка файлов/видео в поля реестров.

## 2. Минимальные требования

- Python 3.11+ (допустим и выше, если проходят проверки).
- Установленные зависимости из `requirements.txt`.
- Актуальные `cookie.md` и `token.md` в папке проекта.

Установка:

```bash
pip install -r requirements.txt
```

## 3. Файлы, которые должны быть готовы перед запуском

- Excel-книга (`*.xlsm`) в корне проекта.
- Папка `files` с вложенными папками по книгам (например `files/one`, `files/two`).
- `cookie.md` и `token.md` с актуальными значениями для целевого стенда.

## 4. Основные режимы запуска

## 4.1 Проверка авторизации

```bash
python migration.py --profile psi --auth-only
```

Использовать перед боевым запуском. Если не проходит — обновить cookie/token.

## 4.2 Полная миграция (обычный запуск)

```bash
python migration.py --profile psi
```

Интерактивный режим по умолчанию:
- выбор книги,
- выбор папки файлов,
- operator-решения при ошибках строки.

## 4.3 Неинтерактивный запуск (для операторов/CI)

```bash
python migration.py --profile psi --no-prompt --no-interactive
```

- cookie/token берутся только из файлов;
- нет вопросов в консоли;
- ошибки строк логируются, миграция продолжает работу по текущим правилам.

## 4.4 Dry-run (без записи в API)

```bash
python migration.py --dry-run --skip-auth --no-interactive --limit 1
```

Проверяет разбор входных данных и режим запуска, но не создает записи в ПГС.

## 4.5 Массовый запуск по нескольким книгам

```bash
python migration.py --profile psi --mode mass \
  --workbooks "book1.xlsm;book2.xlsm" \
  --files-map "book1.xlsm=one;book2.xlsm=two" \
  --no-prompt --no-interactive
```

## 5. Resume / checkpoints (повторная входимость)

По умолчанию resume включен.

Полезные флаги:

```bash
python migration.py --reset-state
python migration.py --state-file state/custom_checkpoints.json
python migration.py --no-resume
```

Рекомендация перед чистым тестом:

```bash
python migration.py --profile dev --reset-state --state-file state/dev_test.json
```

## 6. Operator mode: что делать при ошибке строки

В интерактивном режиме оператор получает выбор:
- `П` (повторить строку),
- `Пр` (пропустить строку в текущей попытке),
- `О` (остановить миграцию с сохранением прогресса).

Для нефатальной ошибки после частичной записи строки:
- `Д` (идти дальше),
- `О` (остановить).

Важно:
- если ошибка произошла после создания основной записи, строка может быть частично мигрирована;
- это явно фиксируется в `user_log`.

## 7. Логи и где смотреть результат

- `logs/script_creation_log-*.txt` — технический подробный лог.
- `logs/success_log-*.txt` — успешные сущности.
- `logs/fail_log-*.txt` — JSON ошибок.
- `logs/user_log-*.txt` — человеко-читаемый отчет для эксплуатации:
  - таблица/лист/строка,
  - поле/файл,
  - краткая и полная ошибка,
  - что сделано после ошибки,
  - мигрирована ли строка,
  - как продолжить/как откатить.
- `ROLLBACK_BODY.json` — данные для rollback.

## 8. Типовые сценарии и действия

## 8.1 Ошибка 413 Request Entity Too Large

Симптом:
- загрузка файла/видео не проходит, в логах `code=413`.

Что делать:
1. Зафиксировать в отчете как инфраструктурный лимит.
2. Уточнить лимит на стенде (nginx/backend).
3. Повторить миграцию после увеличения лимита/уменьшения файла.

## 8.2 sourceKind=none / No upload source

Симптом:
- файл не найден локально по пути.

Что делать:
1. Проверить структуру `files/<папка>`.
2. Проверить имена файлов в Excel.
3. Проверить, что архив распакован без искажений пути/имени.

## 8.3 Сессия устарела / auth не проходит

Что делать:
1. Обновить `cookie.md` и `token.md`.
2. Запустить `--auth-only`.
3. После успешного auth запускать миграцию.

## 8.4 Нужно продолжить после остановки

```bash
python migration.py --profile psi --no-prompt --no-interactive
```

Скрипт продолжит на базе `state/checkpoints.json` (или вашего `--state-file`).

## 8.5 Нужно откатить созданные записи

```bash
python rollback.py
# или точечно:
python rollback_orders.py
python rollback_stray.py
python rollback_cards.py
```

## 9. Рекомендуемая последовательность на новом стенде (PSI)

1. Проверить структуру каталога (`*.xlsm`, `files/*`, `cookie.md`, `token.md`).
2. `python migration.py --profile psi --auth-only`
3. Быстрый dry-run:
   `python migration.py --dry-run --skip-auth --no-interactive --limit 1`
4. Боевой запуск:
   `python migration.py --profile psi --no-prompt --no-interactive`
5. Проверить `logs/user_log-*.txt` и `logs/script_creation_log-*.txt`.
6. При необходимости — повтор с resume или rollback.

## 10. Полезные флаги (кратко)

- `--profile {custom,dev,psi,prod}` — выбор стенда.
- `--ui-base-url` — публичный домен для межреестровых ссылок в данных (`...RecordLink`), если API-домен внутренний.
- `--auth-only` — только проверка авторизации.
- `--dry-run` — без create/update запросов.
- `--skip-auth` — только с dry-run.
- `--no-prompt` — брать cookie/token только из файлов.
- `--no-interactive` — отключить интерактивные вопросы.
- `--operator-mode` — принудительно включить operator-flow.
- `--limit N` — ограничить число строк на каждый раздел.
- `--resume` / `--no-resume` — управление checkpoints.
- `--reset-state` — очистить namespace state перед запуском.
- `--state-file` — отдельный файл checkpoints.
- `--mode single|mass|auto`, `--workbooks`, `--files-map`, `--ask-files-always` — управление входными книгами.

---
Если нужен быстрый чек перед передачей: сначала `--auth-only`, затем `--limit 1`, затем полный запуск.
