# MAPPING PARITY: JS/VBA -> universal_myPet

Документ фиксирует соответствие старой рабочей логики (JS + VBA) новой реализации в Python.

## Источники

- JS:
  - `Скрипт записи заказ-нарядов на отлов животных без владельцев.js`
  - `Скрипт записи животных без владельцев.js`
  - `Скрипт записи карточек учета животных без владельцев.js`
- VBA:
  - `old makros.vba`
    - `ConvertCatchOrdersToJSON`
    - `ConvertStrayAnimalsToJSON`
    - `ConvertAnimalCardsToJSON`

## 1) Реестры и целевые коллекции

| Старый сценарий | Коллекция | В Python |
|---|---|---|
| Заказ-наряды | `animalCatchOrderRegistryCollection` | `process_order_rows` в `migration.py` |
| Животные без владельцев | `animalsRecordsCollectionTwo` | `process_stray_rows` в `migration.py` |
| Акт отлова (из животных) | `animalCatchActRegistryCollection` | `process_stray_rows` (ветка create act) |
| Карточки учета | `myPetAnimalCardReestr` | `process_card_rows` |
| Акт передачи (ловец) | `animalTransferActRegistryCollection` | `process_card_rows` (handover-with-catcher) |
| Акт передачи (приют) | `animalTransferActRegistryCollection` | `process_card_rows` (handover-with-shelter) |
| Акт выпуска | `animalReleaseActRegistryEntry` | `process_card_rows` (release-act) |
| Акт падежа | `animalReleaseActRegistryEntry` | `process_card_rows` (death-act) |
| Акт передачи владельцу | `animalReleaseActRegistryEntry` | `process_card_rows` (transfer-owner-act) |

Итог: набор коллекций совпадает с JS-логикой.

## 2) Источник данных и парсинг Excel

### VBA -> Python (ключевые соответствия)

- VBA `BuildOrderRow` -> Python `_parse_catch_rows` (`_excel_input.py`)
  - диапазон колонок для строк заказ-нарядов: до 151.
- VBA `BuildStrayRow` -> Python `_parse_stray_rows`
  - диапазон колонок для животных: до 36.
- VBA `BuildCardRow` + Parse* функций -> Python `_build_card_row`
  - диапазон колонок карточки: до 428.
  - сохранены групповые блоки событий (дегельминтизация, дезинсекция, вакцинация, стерилизация, выпуск/передача/падеж и др.).

Итог: Excel-парсинг перенесен по тем же листам/диапазонам, что и VBA.

## 3) Поведение upload

JS делал upload через `uploadInBase64`.

Python делает:
- приоритетный upload файла через `POST /api/v1/storage/upload` (multipart),
- fallback в `uploadInBase64` при включенном флаге.

Настройки в `_config.py`:
- `PREFER_FILES_DIR_UPLOAD = True`
- `ALLOW_BASE64_FALLBACK = True`

Итог: сценарий улучшен, обратная совместимость с base64 сохранена.

## 4) Организации и fallback

JS: поиск организаций, fallback/default behavior.

Python:
- lookup организаций сохранён,
- `DEFAULT_ORG_ENABLED` поддержан,
- строгий/мягкий режим поиска через `ORG_STRICT_SEARCH_BY_NAME_OGRN`.

## 5) Откат и аудит

JS: формирование rollback body.

Python:
- единый `ROLLBACK_BODY.json`,
- логи `Created/Errors`,
- отдельные rollback-скрипты.

## 6) Проверка фактического паритета (DEV smoke)

Проверено в `universal_myPet`:
- `clear_collections.py` очистил все 6 коллекций.
- `migration.py` (mass, 2 workbook, files-map one/two) создал записи во всех коллекциях.
- `VERIFY`:
  - первичный прогон: `checked=18 missing_or_error=0`
  - повторный прогон (resume): строки пропущены как `resumed=true`.

## 7) Нефункциональные требования

Требования выполнены:

1. Автоматизация запросов:
- полный цикл без Postman.

2. Логирование:
- журналы по запуску + success/fail + rollback.

3. Повторная входимость:
- checkpoints + восстановление после прерывания.

Дополнительно:
- mass/single/auto режимы,
- выбор книг и папок файлов,
- operator mode для интерактивного восстановления.
