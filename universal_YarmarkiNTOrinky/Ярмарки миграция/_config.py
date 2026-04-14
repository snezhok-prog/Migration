import os

# Стендозависимые настройки
TEST = False
BASE_URL = "https://iam.torknd-customer.dev.pd15.digitalgov.mtp"
UI_BASE_URL = BASE_URL
JWT_URL = BASE_URL.rstrip("/") + "/jwt/"
VERIFY_SSL = False
AUTO_JWT = False
SAVE_AUTH = False
AUTH_TEST_COLLECTION = "organizations"
REAUTH_ON_401 = False
REAUTH_RETRIES = 1
EXCEL_FILE_NAME = "Форма_для_миграции_Ярмарки.xlsm"
EXCEL_INPUT_GLOB = "*.xlsm"
EXCEL_LISTS = [
    "2. Реестр мест",
    "3. Реестр разрешений",
    "4. Реестр ярмарок"
]
RECORDS_TEMPLATES = {
    "4. Реестр ярмарок": {
        "guid": None,
        "unit": None,
        "informatsiya_o_yarmarke_1": {
            "nomer_yarmarki": None,
            "naimenovanie_yarmarki": None,
            "tip_yarmarki": None,
            "spetsializatsiya_yarmarki": None,
            "period_provedeniya": None,
            "mesto_provedeniya": None,
            "organizator": None
        },
        "organizator_1_1": {
            "tip_organizatora": None,
            "naimenovanie_organizatora": None,
            "inn": None,
            "ogrn": None,
            "adres_organizatora": None,
            "telefon_organizatora": None,
            "email_organizatora": None
        }
    },
    "2. Реестр мест": {
        "guid": None,
        "unit": None,
        "generalInformation": {
            "subject": None,
            "disctrict": None
        },
        "placeMarketInfo": {
            "marketSchemeNumber": None,
            "coordinates": None,
            "landArea": None,
            "marketType": None,
            "marketSpecialization": None,
            "marketSpecializationOther": None,
            "statusFairPlace": None,
            "statusFairPlaceSpecific": None,
            "statusPlaceFair": None,
            "blockCadNumber": {
                "cadNumber": None,
                "landAddress": None
            },
            "cadsObjects": {
                "cadObjNum": None,
                "objectAddress": None
            }
        },
        "marketInfo": []
    },
    "3. Реестр разрешений": {
        "guid": None,
        "unit": None,
        "razreshenie_na_pravo_organizatsii_yarmarki": {
            "nomer_razresheniya": None,
            "data_vydachi": None,
            "organizator": None,
            "naimenovanie_yarmarki": None,
            "mesto_provedeniya": None,
            "srok_deystviya": None,
            "status": None,
            "file": None
        }
    }
}
STANDARD_CODES = {
    "Уведомление о вводе сети связи в эксплуатацию": "40692",  # ПС - 42088, ДЕВ - 40692
}
METAREGLAMENT = "RKN012"
FAIR_COLLECTION = "Fair"
FAIR_MESTO_COLLECTION = "FairMesto"
FAIR_PERMITS_COLLECTION = "FairPermits"
RECORDS_COLLECTION = "reestrpermitsReestr"
UNIT = {
    "id": "6650527c3000227496944b6b",
    "name": "ООО Агентство \"Полилог\"",
    "ogrn": "1027706014874",
    "region": {
        "code": "39",
        "name": "Калининградская область"
    },
    "shortName": "Агентство \"Полилог\""
}
# Пути к директориям
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
FILES_DIR = os.path.join(SCRIPT_DIR, "files")
LOGS_DIR = os.path.join(SCRIPT_DIR, "logs")
TOKEN_FILE = "token.md"
COOKIE_FILE = "cookie.md"
STATE_DIR = os.path.join(SCRIPT_DIR, "state")
STATE_FILE = os.path.join(STATE_DIR, "checkpoints.json")
RESUME_BY_DEFAULT = True
ROLLBACK_BODY_FILE = os.path.join(SCRIPT_DIR, "ROLLBACK_BODY.json")
# Паттерн названий логов удачно мигрированных записей
SUCCESS_LOG_PATTERN = os.path.join(LOGS_DIR, "success_log-*.txt")
# Поддерживаемые расширения файлов дял миграции
SUPPORTED_EXTENSIONS = [
    '.pdf', '.xml', '.doc', '.docx', '.xls', '.xlsx',
    '.jpg', '.jpeg', '.png', '.zip', '.txt', '.rtf',
    '.gif', '.bmp', '.tiff', '.tif', '.webp',
    '.odt', '.ods', '.odp', '.odg', '.odf',
    '.csv', '.json', '.html', '.htm', '.css', '.js',
    '.mp3', '.wav', '.mp4', '.avi', '.mov', '.mkv',
    '.psd', '.ai', '.eps', '.svg',
    '.sql', '.db', '.sqlite', '.dbf',
    '.msg', '.eml', '.pst',
    '.dwg', '.dxf',
    '.heic', '.raw'
]
# Шаблон дел миграции
APPEAL_SETTINGS = {
    "parentEntries": "RKN012Appeals",
    "number": "н/н",
    "pin": "0000",
    "status": {
        "code": "completePositive",
        "name": "Положительное решение"
    },
    "executor": {
        "name": "",
        "email": "",
        "position": ""
    },
    "dateFinish": "2026-01-01T00:00:00.001+0000",
    "registryEntryType": {
        "code": "RKN012_record",
        "name": "Запись реестра сетей связи",
        "actions": []
    },
    "operationType": "registration",
    "unit": {
        "id": "5fffe1e13a956b0001280ff8",
        "name": "Федеральная служба по надзору в сфере связи, информационных технологий и массовых коммуникаций (Роскомнадзор)",
        "ogrn": "1087746736296",
        "region": {
            "code": "77",
            "name": "г. Москва"
        },
        "shortName": "ЦА РКН"
    }
}

