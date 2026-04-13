# RKN012 Migration Tool

Инструмент для миграции данных в систему Роскомнадзора (RKN012) через API.

## 📋 Описание

Данный проект представляет собой набор утилит для автоматизированного создания обращений, под услуг, субъектов и документов в системе RKN012 через REST API.

## 🏗️ Структура проекта

```
├── _config.py        # Конфигурация проекта
├── _logger.py        # Настройка логирования
├── _utils.py         # Вспомогательные функции
├── _api.py           # Функции для работы с API
├── _templates.py     # Шаблоны субъектов (ЮЛ и ИП)
├── migration.py      # Основной скрипт миграции обращений
├── fair_migration.py # Скрипт миграции данных ярмарок
├── rollback.py       # Скрипт для отката миграции
├── files/            # Директория для файлов миграции
├── logs/             # Директория для логов
└── README.md         # Документация
```

## 🎪 Миграция данных ярмарок

### Описание

Скрипт `fair_migration.py` предназначен для миграции данных о ярмарках из Excel файла напрямую в API системы без промежуточного создания JSON файлов.

### Структура Excel файла

Excel файл должен содержать следующие листы:
- `4. Реестр ярмарок` - информация о ярмарках и организаторах
- `2. Реестр мест` - данные о местах размещения ярмарок  
- `3. Реестр разрешений` - разрешения на организацию ярмарок

**Важно:** Названия столбцов находятся в 4-й строке, данные начинаются с 6-й строки.

### Запуск миграции

```bash
# Тестовый режим (без реального API)
python fair_migration.py

# Боевой режим (требуется настройка аутентификации)
# Изменить TEST = False в _config.py
python fair_migration.py
```

### Создаваемые структуры

Скрипт создает следующие структуры в API:

1. **Ярмарки** (`informatsiya_o_yarmarke_1` + `organizator_1_1`)
2. **Места** (`dannye_po_reestru_mest_dlya_razmescheniya_yarmarok`)
3. **Разрешения** (`razreshenie_na_pravo_organizatsii_yarmarki`)

### Логирование

Результаты миграции записываются в файлы:
- `logs/script_creation_log-*.txt` - общий лог
- `logs/success_log-*.txt` - успешные операции
- `logs/fail_log-*.txt` - ошибки

## 📦 Требования

- Python 3.8+
- Библиотеки:
  - `requests`
  - `pandas`
  - `openpyxl` (для чтения Excel)

Установка зависимостей:
```bash
pip install requests pandas openpyxl
```

## ⚙️ Конфигурация

### _config.py

Основные настройки в файле `_config.py`:

```python
# URL API
BASE_URL = "https://iam.torknd-customer.dev.pd15.digitalgov.mtp"

# Имя Excel файла с данными
EXCEL_FILE_NAME = "ТЕСТ МИГРАЦИЯ 012.xlsx"

# Коды стандартов
STANDARD_CODES = {
    "Уведомление о вводе сети связи в эксплуатацию": "40692",
}

# Настройки подразделения
UNIT = {
    "id": "6650527c3000227496944b6b",
    "name": "ООО Агентство \"Полилог\"",
    "ogrn": "1027706014874",
}

# Поддерживаемые расширения файлов
SUPPORTED_EXTENSIONS = [
    '.pdf', '.xml', '.doc', '.docx', '.xls', '.xlsx',
    '.jpg', '.jpeg', '.png', '.zip', '.txt', '.rtf',
    # ... и другие
]
```

## 🔐 Авторизация

Для работы с API необходимо:

1. Открыть браузер и войти в систему
2. Открыть DevTools (F12) → Network
3. Выполнить любой запрос к API
4. Скопировать заголовок `Cookie` из Request Headers
5. Скопировать JWT токен

При запуске скрипта будет запрошен ввод:
```
Cookie: PLATFORM_SESSION=...; XSRF-TOKEN=...; ...
JWT токен: Bearer eyJhbGc...
```

## 📖 Основные функции

### Работа с Excel

```python
from _utils import read_excel

df = read_excel("path/to/file.xlsx")
```

Читает Excel файл, пропускает первые 4 строки, все значения как строки.

### Форматирование телефонов

```python
from _utils import format_phone, format_multiple_phones

# Одиночный телефон
phone = format_phone("89141234567")
# Результат: "+7 (914) 123 45 67"

# Несколько телефонов (разделены ;)
phones = format_multiple_phones("89141234567; 89147654321")
# Результат: ["+7 (914) 123 45 67", "+7 (914) 765 43 21"]
```

### Работа с датами

```python
from _utils import to_iso_date, parse_date_to_birthday_obj

# Преобразование в ISO формат
iso_date = to_iso_date("31.12.2024")
# Результат: "2024-12-31T00:00:00.000+0300"

# Парсинг даты рождения
birthday = parse_date_to_birthday_obj("31.12.1984")
# Результат: {
#   "date": {"year": 1984, "month": 12, "day": 31},
#   "jsDate": "1984-12-30T21:00:00.000Z",
#   "formatted": "31.12.1984",
#   "epoc": 473288400
# }
```

### Загрузка файлов

```python
from _api import upload_file, delete_file_from_storage

# Загрузка файла
result = upload_file(
    session=session,
    logger=logger,
    file_path="/path/to/file.pdf",
    entry_name="RKN012Appeals",
    entry_id="appeal_id_123",
    entity_field_path=""
)

# Удаление файла
success = delete_file_from_storage(
    session=session,
    logger=logger,
    file_id="file_id_123"
)
```

### Создание обращения

```python
from _api import (
    create_appeal_data,
    create_subservice_data,
    create_subject_data,
    create_mainElement_data,
    create_appeal_with_entities
)

# Создание данных обращения
appeal_data = create_appeal_data(
    unit=unit_info,
    data={"number": "123", "pin": "0000"}
)

# Создание под услуги
subservice_data = create_subservice_data(
    subserviceTemplate=template,
    data={"additional_field": "value"}
)

# Создание субъекта (ЮЛ или ИП)
from _templates import SUBJECT_UL, SUBJECT_IP

subject_data = create_subject_data(
    template=SUBJECT_UL,  # или SUBJECT_IP
    data={"xsdData": {"phone": "+7 (999) 123 45 67"}}
)

# Создание всех сущностей сразу
success, appeal, subservice, subject, document = create_appeal_with_entities(
    session=session,
    logger=logger,
    appeal_data=appeal_data,
    subservice_data=subservice_data,
    subject_data=subject_data,
    document_data=document_data,
    files_contents=[(base64_content, "file.pdf")]
)
```

### Удаление записей

```python
from _api import delete_from_collection, find_in_collection

# Удаление из коллекции
success = delete_from_collection(
    session=session,
    logger=logger,
    data={
        "_id": "record_id",
        "guid": "record_guid",
        "parentEntries": "RKN012Appeals.subservices"
    }
)
```

## 📊 Шаблоны субъектов

### Юридическое лицо (SUBJECT_UL)

```python
{
    "kind": {
        "subKind": {
            "name": "Юридическое лицо",
            "specialTypeId": "ulApplicant"
        }
    },
    "data": {
        "organization": {
            "opf": {...},
            "shortName": "Название организации",
            "name": "Полное название",
            "ogrn": "0000000000000",
            "inn": "0000000000",
            "kpp": "000000000",
            "registrationAddress": {
                "fullAddress": "Адрес"
            }
        }
    }
}
```

### Индивидуальный предприниматель (SUBJECT_IP)

```python
{
    "kind": {
        "subKind": {
            "name": "Индивидуальный предприниматель",
            "specialTypeId": "ipApplicant"
        }
    },
    "data": {
        "person": {
            "lastName": "Фамилия",
            "firstName": "Имя",
            "middleName": "Отчество",
            "birthday": {...},
            "documentType": [...],
            "documentSeries": "12",
            "documentNumber": "1231231",
            "inn": "000000000000",
            "ogrn": "322440000000311"
        }
    }
}
```

## 📝 Логирование

Проект использует три типа логгеров:

1. **Основной логгер** (`setup_logger`) - логирует все операции
2. **Логгер успешных операций** (`setup_success_logger`) - только успешные миграции
3. **Логгер ошибок** (`setup_fail_logger`) - только ошибки
4. **Логгер отката** (`setup_rollback_logger`) - операции удаления

Пример использования:
```python
from _logger import setup_logger

logger = setup_logger()
logger.info("Начало миграции")
logger.error("Произошла ошибка")
logger.debug("Отладочная информация")
```

## 🔧 Вспомогательные функции

### Конвертация в JSON-сериализуемый формат

```python
from _utils import jsonable

# Преобразует numpy типы, pandas Timestamp и т.д.
data = jsonable(complex_object)
json_str = json.dumps(data)
```

### Поиск файлов

```python
from _utils import find_file_in_dir

# Поиск файла в директории
file_path = find_file_in_dir(
    files_dir="/path/to/files",
    filename_hint="document"  # с расширением или без
)
```

### Работа с base64

```python
from _utils import read_file_as_base64

# Чтение файла в base64
b64_content = read_file_as_base64("/path/to/file.pdf")
```

## 🌐 API Endpoints

Основные используемые endpoints:

- `POST /api/v1/search/subservices` - поиск под услуг
- `POST /api/v1/search/organizations` - поиск организаций
- `POST /api/v1/create/{collection}` - создание записи
- `PUT /api/v1/update/{collection}` - обновление записи
- `DELETE /api/v1/delete/{collection}` - удаление записи
- `POST /api/v1/storage/upload` - загрузка файла
- `DELETE /api/v1/storage/remove` - удаление файла

## ⚠️ Важные замечания

1. **SSL Verification**: Отключена (`session.verify = False`) для работы с dev-сервером
2. **Повторные попытки**: API запросы автоматически повторяются при ошибках 401, 403, 500
3. **Кодировка**: Все логи сохраняются в UTF-8
4. **Временная зона**: Даты конвертируются в UTC+3 (московское время)

## 📄 Лицензия

Внутренний инструмент для миграции данных.

## 👥 Поддержка

По вопросам обращайтесь к разработчикам проекта.