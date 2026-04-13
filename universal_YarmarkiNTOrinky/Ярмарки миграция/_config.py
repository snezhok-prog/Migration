import os

# РЎС‚РµРЅРґРѕР·Р°РІРёСЃРёРјС‹Рµ РЅР°СЃС‚СЂРѕР№РєРё
TEST = False  # РР·РјРµРЅРµРЅРѕ РґР»СЏ С‚РµСЃС‚РёСЂРѕРІР°РЅРёСЏ  # РР·РјРµРЅРµРЅРѕ РґР»СЏ С‚РµСЃС‚РёСЂРѕРІР°РЅРёСЏ
BASE_URL = "https://iam.torknd-customer.dev.pd15.digitalgov.mtp"
EXCEL_FILE_NAME = "Р¤РѕСЂРјР°_РґР»СЏ_РјРёРіСЂР°С†РёРё_РЇСЂРјР°СЂРєРё.xlsm"
EXCEL_LISTS = [
    "2. Р РµРµСЃС‚СЂ РјРµСЃС‚",
    "3. Р РµРµСЃС‚СЂ СЂР°Р·СЂРµС€РµРЅРёР№", 
    "4. Р РµРµСЃС‚СЂ СЏСЂРјР°СЂРѕРє"
]
RECORDS_TEMPLATES = {
    "4. Р РµРµСЃС‚СЂ СЏСЂРјР°СЂРѕРє": {
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
    "2. Р РµРµСЃС‚СЂ РјРµСЃС‚": {
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
    "3. Р РµРµСЃС‚СЂ СЂР°Р·СЂРµС€РµРЅРёР№": {
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
    "РЈРІРµРґРѕРјР»РµРЅРёРµ Рѕ РІРІРѕРґРµ СЃРµС‚Рё СЃРІСЏР·Рё РІ СЌРєСЃРїР»СѓР°С‚Р°С†РёСЋ": "40692",  # РџРЎР - 42088, Р”Р•Р’ - 40692
}
METAREGLAMENT = "RKN012"
FAIR_COLLECTION = "Fair"
FAIR_MESTO_COLLECTION = "FairMesto"
FAIR_PERMITS_COLLECTION = "FairPermits"
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

