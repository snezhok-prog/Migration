import os

# Stand configuration
TEST = True
BASE_URL = "https://iam.torknd-customer.dev.pd15.digitalgov.mtp"
JWT_URL = BASE_URL.rstrip("/") + "/jwt/"
VERIFY_SSL = False
AUTH_TEST_COLLECTION = "organizations"
AUTO_JWT = True
SAVE_AUTH = True
REAUTH_ON_AUTH_ERROR = True
REAUTH_RETRIES = 2

# Collections/endpoints
TARGET_COLLECTION = "animalsRecordsCollectionTwo"
ACT_COLLECTION = "animalCatchActRegistryCollection"
ORDER_COLLECTION = "animalCatchOrderRegistryCollection"
CARD_COLLECTION = "myPetAnimalCardReestr"
RELEASE_COLLECTION = "animalReleaseActRegistryEntry"
TRANSFER_ACT_COLLECTION = "animalTransferActRegistryCollection"
STORAGE_UPLOAD_PATH = "/api/v1/storage/uploadInBase64"
STORAGE_UPLOAD_FILE_PATH = "/api/v1/storage/upload"

# UI paths used for cross-links between registries
UI_CATCH_ORDER_EDIT_PATH = "/myPet/myPetReestrs/animalCatchOrder/edit"
UI_ANIMAL_EDIT_PATH = "/myPet/myPetReestrs/animalsRecords/edit"
UI_CATCH_ACT_EDIT_PATH = "/myPet/myPetReestrs/animalCatchActRegistry/edit"
UI_RELEASE_ACT_EDIT_PATH = "/myPet/myPetReestrs/animalReleaseActRegistry/edit"
UI_TRANSFER_ACT_EDIT_PATH = "/myPet/myPetReestrs/animalTransferActRegistry/edit"

# Migration behavior
ENABLE_FILE_UPLOADS = True
DRY_RUN_LOG_UPLOAD_TARGETS = False
STOP_ON_FIRST_CREATE_ERROR = True
STOP_ON_FIRST_FATAL_UPLOAD = False
VERIFY_CREATED = True

# Upload strategy:
# - if True, script first tries to find a file in FILES_DIR and upload it via /api/v1/storage/upload
# - if local file is not found (or flag disabled), it can fallback to base64 upload according to flag below
PREFER_FILES_DIR_UPLOAD = True
ALLOW_BASE64_FALLBACK = True

# Organization lookup
DEFAULT_ORG_ENABLED = False
ORG_STRICT_SEARCH_BY_NAME_OGRN = False
DEFAULT_ORG = {
    "id": "5982897d-bf70-4e95-8030-a54cafae4b30",
    "_id": "6650527c3000227496944b6b",
    "name": "ООО Агентство \"Полилог\"",
    "shortName": "Агентство \"Полилог\"",
    "ogrn": "1027706014874",
    "inn": "7700000000",
    "regions": {"code": "61", "name": "Ростовская область"},
}

# Regional codes (as in JS scripts)
REGION_CODE = {
    "калининградская область": "39",
    "челябинская область": "74",
    "ростовская область": "61",
    "санкт-петербург": "78",
    "город санкт-петербург": "78",
    "москва": "77",
    "город москва": "77",
    "республика татарстан": "16",
    "татарстан": "16",
    "город казань": "16",
    "казань": "16",
}

# Project paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
FILES_DIR = os.path.join(SCRIPT_DIR, "files")
LOGS_DIR = os.path.join(SCRIPT_DIR, "logs")
STATE_DIR = os.path.join(SCRIPT_DIR, "state")
STATE_FILE = os.path.join(STATE_DIR, "checkpoints.json")

# Input files from Excel macros
STRAY_PART_GLOB = "stray_animals_registry_part*.json"
CATCH_PART_GLOB = "catch_orders_registry_part*.json"
CARD_PART_GLOB = "animal_cards_registry_part*.json"

# Direct Excel input (preferred)
USE_EXCEL_INPUT = True
EXCEL_INPUT_FILE = ""
EXCEL_INPUT_GLOB = "*.xlsm"
EXCEL_DATA_START_ROW = 6
RESUME_BY_DEFAULT = True

SUCCESS_LOG_PATTERN = os.path.join(LOGS_DIR, "success_log-*.txt")
ROLLBACK_BODY_PATH = os.path.join(SCRIPT_DIR, "ROLLBACK_BODY.json")
