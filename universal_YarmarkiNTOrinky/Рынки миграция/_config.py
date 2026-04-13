import os

# РЎС‚РµРЅРґРѕР·Р°РІРёСЃРёРјС‹Рµ РЅР°СЃС‚СЂРѕР№РєРё
TEST = False
BASE_URL = "https://iam.torknd-customer.dev.pd15.digitalgov.mtp"
EXCEL_FILE_NAME = "Р¤РѕСЂРјР°_РґР»СЏ_РјРёРіСЂР°С†РёРё_Р С‹РЅРєРё.xlsm"
EXCEL_LISTS = [
    "2. Р РµРµСЃС‚СЂ СЂР°Р·СЂРµС€РµРЅРёР№", 
    "3. Р РµРµСЃС‚СЂ СЂС‹РЅРєРѕРІ"
]
RECORDS_TEMPLATES = {
    "2. Р РµРµСЃС‚СЂ СЂР°Р·СЂРµС€РµРЅРёР№": {},
    "3. Р РµРµСЃС‚СЂ СЂС‹РЅРєРѕРІ": {
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
    "РЈРІРµРґРѕРјР»РµРЅРёРµ Рѕ РІРІРѕРґРµ СЃРµС‚Рё СЃРІСЏР·Рё РІ СЌРєСЃРїР»СѓР°С‚Р°С†РёСЋ": "40692",  # РџРЎР - 42088, Р”Р•Р’ - 40692
}
METAREGLAMENT = "RKN012"
LICENSES_COLLECTION = "RKN012_Licenses"
RECORDS_COLLECTION = "reestrpermitsReestr"
UNIT = {
    "id": "6650527c3000227496944b6b",
    "name": "РћРћРћ РђРіРµРЅС‚СЃС‚РІРѕ \"РџРѕР»РёР»РѕРі\"",
    "ogrn": "1027706014874",
    "region": {
        "code": "39",
        "name": "РљР°Р»РёРЅРёРЅРіСЂР°РґСЃРєР°СЏ РѕР±Р»Р°СЃС‚СЊ"
    },
    "shortName": "РђРіРµРЅС‚СЃС‚РІРѕ \"РџРѕР»РёР»РѕРі\""
}
# РџСѓС‚Рё Рє РґРёСЂРµРєС‚РѕСЂРёСЏРј
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
FILES_DIR = os.path.join(SCRIPT_DIR, "files")
LOGS_DIR = os.path.join(SCRIPT_DIR, "logs")
TOKEN_FILE = "token.md"
COOKIE_FILE = "cookie.md"
STATE_DIR = os.path.join(SCRIPT_DIR, "state")
STATE_FILE = os.path.join(STATE_DIR, "checkpoints.json")
RESUME_BY_DEFAULT = True
# РџР°С‚С‚РµСЂРЅ РЅР°Р·РІР°РЅРёР№ Р»РѕРіРѕРІ СѓРґР°С‡РЅРѕ РјРёРіСЂРёСЂРѕРІР°РЅРЅС‹С… Р·Р°РїРёСЃРµР№
SUCCESS_LOG_PATTERN = os.path.join(LOGS_DIR, "success_log-*.txt")
# РџРѕРґРґРµСЂР¶РёРІР°РµРјС‹Рµ СЂР°СЃС€РёСЂРµРЅРёСЏ С„Р°Р№Р»РѕРІ РґСЏР» РјРёРіСЂР°С†РёРё
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
# РЁР°Р±Р»РѕРЅ РґРµР» РјРёРіСЂР°С†РёРё
APPEAL_SETTINGS = {
    "parentEntries": "RKN012Appeals",
    "number": "РЅ/РЅ",
    "pin": "0000",
    "status": {
        "code": "completePositive",
        "name": "РџРѕР»РѕР¶РёС‚РµР»СЊРЅРѕРµ СЂРµС€РµРЅРёРµ"
    },
    "executor": {
        "name": "",
        "email": "",
        "position": ""
    },
    "dateFinish": "2026-01-01T00:00:00.001+0000",
    "registryEntryType": {
        "code": "RKN012_record",
        "name": "Р—Р°РїРёСЃСЊ СЂРµРµСЃС‚СЂР° СЃРµС‚РµР№ СЃРІСЏР·Рё",
        "actions": []
    },
    "operationType": "registration",
    "unit": {
        "id": "5fffe1e13a956b0001280ff8",
        "name": "Р¤РµРґРµСЂР°Р»СЊРЅР°СЏ СЃР»СѓР¶Р±Р° РїРѕ РЅР°РґР·РѕСЂСѓ РІ СЃС„РµСЂРµ СЃРІСЏР·Рё, РёРЅС„РѕСЂРјР°С†РёРѕРЅРЅС‹С… С‚РµС…РЅРѕР»РѕРіРёР№ Рё РјР°СЃСЃРѕРІС‹С… РєРѕРјРјСѓРЅРёРєР°С†РёР№ (Р РѕСЃРєРѕРјРЅР°РґР·РѕСЂ)",
        "ogrn": "1087746736296",
        "region": {
            "code": "77",
            "name": "Рі. РњРѕСЃРєРІР°"
        },
        "shortName": "Р¦Рђ Р РљРќ"
    }
}

