from __future__ import annotations

import os


TEST = False

BASE_URL = "https://iam.torknd-customer.dev.pd15.digitalgov.mtp"
UI_BASE_URL = BASE_URL
VERIFY_SSL = False
JWT_URL = BASE_URL.rstrip("/") + "/jwt/"
AUTO_JWT = True
SAVE_AUTH = False
AUTH_TEST_COLLECTION = "organizations"
REAUTH_ON_401 = True
REAUTH_RETRIES = 2

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
FILES_DIR = os.path.join(SCRIPT_DIR, "files")
LOGS_DIR = os.path.join(SCRIPT_DIR, "logs")
TOKEN_FILE = "token.md"
COOKIE_FILE = "cookie.md"
STATE_DIR = os.path.join(SCRIPT_DIR, "state")
STATE_FILE = os.path.join(STATE_DIR, "checkpoints.json")
RESUME_BY_DEFAULT = True
ROLLBACK_BODY_FILE = os.path.join(SCRIPT_DIR, "ROLLBACK_BODY.json")
SUCCESS_LOG_PATTERN = os.path.join(LOGS_DIR, "success_log-*.txt")
EXCEL_FILE_NAME = "Форма_для_миграции_НТО.xlsm"
EXCEL_INPUT_GLOB = "*.xlsm"

SHEET_MESTO = "2. Реестр мест"
SHEET_TORGI = "3. Реестр торгов"
EXCEL_LISTS = [SHEET_MESTO, SHEET_TORGI]

NTO_MESTO_COLLECTION = "NTOmesto"
NTO_NSI_COLLECTION = "nsiLocalObjectNTO"
TORGI_COLLECTION = "reestrbiddingReestr"

DEFAULT_ORG = {
    "id": "5982897d-bf70-4e95-8030-a54cafae4b30",
    "_id": "6650527c3000227496944b6b",
    "guid": "5982897d-bf70-4e95-8030-a54cafae4b30",
    "name": "ООО Агентство \"Полилог\"",
    "shortName": "Агентство \"Полилог\"",
    "ogrn": "1027706014874",
    "kpp": "772201001",
    "email": "cityhall@klgd.ru",
    "site": "http://www.klgd.ru",
    "phone": "(4012) 92-33-80,92-30-71",
    "regions": {
        "_id": "5e7c84507e0daf0001872342",
        "code": "39",
        "name": "Калининградская область",
    },
}
