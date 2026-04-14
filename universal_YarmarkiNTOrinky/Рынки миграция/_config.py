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
EXCEL_FILE_NAME = "Форма_для_миграции_Рынки.xlsm"
EXCEL_INPUT_GLOB = "*.xlsm"
EXCEL_LISTS = [
    "2. Реестр разрешений", 
    "3. Реестр рынков"
]
RECORDS_TEMPLATES = {
    "2. Реестр разрешений": {},
    "3. Реестр рынков": {
        "guid": None,
        "unit": None,
        "generalInformation":{
            "Subject": None,
            "Disctrict": None
        },
        "MeetingProtocol":{
            "ProtocolName": None,
            "protocolNumberRenewal": None,
            "protocolDate": None,
            "CommissionMeetingLocation": None,
            "CommissionChairFullName": None,
            "CommissionChairPosition": None,
            "CommissionSecretaryFullName": None,
            "CommissionSecretaryPos": None,
            "CommissionCompositionFormat": [],
            "DescriptionPerformancesFormat": [],
            "decisionType": None,
            "commissionProtocolFileReissue1": None,
            "commissionProtocolFileReissue2": None,
            "commissionProtocolFileReissue": None
        },
        "resolution": {
            "resolutionName": None,
            "resolutionNumber": None,
            "resolutionDate": None,
            "resolutionApprovalFile": None,
            "resolutionDenialFile": None,
            "resolutionExtensionFile": None,
            "resolutionRefusalFile": None,
            "resolutionReissueFile": None
        },
        "positiveDecisionNotification": {
            "positiveDecisionNotificationName": None,
            "positiveDecisionNotificationNumber": None,
            "positiveDecisionNotificationDate": None,
            "positiveDecisionDetails": None,
            "positiveDecisionApprovalFile": None,
            "positiveDecisionExtensionFile": None,
            "positiveDecisionReissueFile": None
        },
        "denialNotification": {
            "denialNotificationName": None,
            "denialNotificationNumber": None,
            "denialNotificationDate": None,
            "denialReasons": None,
            "denialIssueFile": None,
            "denialExtensionFile": None,
            "denialReissueFile": None
        },
        "permission": {
            "PermissionStatus": None,
            "PermissionNumber": None,
            "PermissionStartDate": None,
            "administrationName": None,
            "permissionEffectiveStartDate": None,
            "PermissionEndDate": None,
            "permissionExtensionDate": None,
            "reissuePermissionFile": None
        },
        "suspensionDocumentScourtDecision":{
            "registrationNumber": None,
            "DecisionDate": None,
            "suspensionRegulationNumber": None,
            "suspensionRegulationDate": None,
            "suspensionNotificationFile": None
        },
        "violationCorrectionSuspensionDocuments": {
            "correctionNotificationNumber": None,
            "correctionNotificationDate": None,
            "permissionResumptionDate": None,
            "resumptionNotificationNumber": None,
            "resumptionNotificationDate": None,
            "resumptionRegulationNumber": None,
            "resumptionRegulationDate": None,
            "violationCorrectionNotificationFile": None,
            "resumptionNotificationFile": None,
            "resumptionRegulationFile": None
        },
        "omsuPetitionSuspensionDocuments": {
            "iniciator": None,
            "cancellReason": None,
            "petitionNumber": None,
            "petitionDate": None,
            "courtDecisionRegistrationNumber": None,
            "courtDecisionDate": None,
            "annulmentDate": None,
            "cancellationNotificationNumber": None,
            "cancellationNotificationDate": None,
            "npaAnnulmentNpaNumber": None,
            "npaAnnulmentNpaDate": None,
            "cancelDoc": None,
            "omsuPetitionFile": None,
            "courtDecisionForAnnulmentFile": None,
            "cancellationNotificationFile": None,
            "npaAnnulmentFile": None
        }
    }
}
STANDARD_CODES = {
    "Уведомление о вводе сети связи в эксплуатацию": "40692",  # ПСИ - 42088, ДЕВ - 40692
}
METAREGLAMENT = "RKN012"
LICENSES_COLLECTION = "RKN012_Licenses"
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

