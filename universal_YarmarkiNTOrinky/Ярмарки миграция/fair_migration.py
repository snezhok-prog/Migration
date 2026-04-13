import argparse
import copy
import json
import os
import re
import sys
import traceback
import warnings
from pathlib import Path

import pandas as pd
from urllib3.exceptions import InsecureRequestWarning

from _api import (
    api_request,
    delete_file_from_storage,
    delete_from_collection,
    setup_session,
    upload_file,
)
from _config import (
    BASE_URL,
    EXCEL_FILE_NAME,
    EXCEL_LISTS,
    FAIR_COLLECTION,
    FAIR_MESTO_COLLECTION,
    FAIR_PERMITS_COLLECTION,
    FILES_DIR,
    RECORDS_TEMPLATES,
    RESUME_BY_DEFAULT,
    SCRIPT_DIR,
    STATE_FILE,
    TEST,
    UNIT,
)
from _logger import setup_fail_logger, setup_logger, setup_success_logger, setup_user_logger
from _state import ResumeState
from _utils import find_file_in_dir, generate_guid, jsonable, read_excel


NSI_LOCAL_OBJECT_FAIR_COLLECTION = "nsiLocalObjectFair"
DEFAULT_ORG = copy.deepcopy(UNIT)

SHEET_FAIR = "4. Р РµРµСЃС‚СЂ СЏСЂРјР°СЂРѕРє"
SHEET_MESTO = "2. Р РµРµСЃС‚СЂ РјРµСЃС‚"
SHEET_PERMITS = "3. Р РµРµСЃС‚СЂ СЂР°Р·СЂРµС€РµРЅРёР№"


def norm_str(value):
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    text = str(value).strip()
    return text if text else None


def norm_key(value):
    return (norm_str(value) or "").replace("\xa0", " ").strip().lower()


def pick_first_non_empty(*values):
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        return value
    return None


def is_empty_val(value):
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    return False


def has_any_non_null(obj):
    if not isinstance(obj, dict):
        return False
    for value in obj.values():
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        if isinstance(value, dict):
            if has_any_non_null(value):
                return True
            continue
        if isinstance(value, list):
            if any(
                item is not None
                and (not isinstance(item, str) or item.strip())
                and (not isinstance(item, dict) or has_any_non_null(item))
                for item in value
            ):
                return True
            continue
        return True
    return False


def parse_bool(value):
    if isinstance(value, bool):
        return value
    return norm_key(value) in {"РґР°", "true", "1", "yes", "y"}


def pad2(value):
    return str(value).zfill(2)


def parse_date_to_iso(value):
    if value is None:
        return None
    if isinstance(value, (int, float)) and not pd.isna(value):
        try:
            dt = pd.Timestamp("1899-12-30") + pd.to_timedelta(float(value), unit="D")
            return dt.strftime("%Y-%m-%d")
        except Exception:
            return None

    text = norm_str(value)
    if not text:
        return None

    match_iso = re.match(r"^(\d{4})-(\d{2})-(\d{2})", text)
    if match_iso:
        return f"{match_iso.group(1)}-{match_iso.group(2)}-{match_iso.group(3)}"

    match_ru = re.match(r"^(\d{1,2})[./](\d{1,2})[./](\d{4})$", text)
    if match_ru:
        return f"{match_ru.group(3)}-{pad2(match_ru.group(2))}-{pad2(match_ru.group(1))}"

    return text


def as_date_ru_or_dash(value):
    if value is None:
        return "-"
    if isinstance(value, (int, float)) and not pd.isna(value):
        try:
            dt = pd.Timestamp("1899-12-30") + pd.to_timedelta(float(value), unit="D")
            return dt.strftime("%d.%m.%Y")
        except Exception:
            return "-"

    text = norm_str(value)
    if not text:
        return "-"

    match_iso = re.match(r"^(\d{4})-(\d{2})-(\d{2})", text)
    if match_iso:
        return f"{match_iso.group(3)}.{match_iso.group(2)}.{match_iso.group(1)}"

    match_ru = re.match(r"^(\d{1,2})[./](\d{1,2})[./](\d{4})$", text)
    if match_ru:
        return f"{pad2(match_ru.group(1))}.{pad2(match_ru.group(2))}.{match_ru.group(3)}"

    return "-"


def dash_str(value):
    return norm_str(value) or "-"


def dash(value):
    if value is None:
        return "-"
    if isinstance(value, str):
        return value.strip() or "-"
    return value


def extract_postal_code_and_rest(text):
    if not text:
        return {"postalCode": None, "fullAddress": None}

    match = re.match(r"^(\d{6})\s*,\s*(.+)$", text)
    if match:
        return {"postalCode": match.group(1), "fullAddress": match.group(2).strip() or None}

    match = re.match(r"^(\d{6})\s+(.+)$", text)
    if match:
        return {"postalCode": match.group(1), "fullAddress": match.group(2).strip() or None}

    return {"postalCode": None, "fullAddress": text}


def parse_address_to_obj(value):
    if value is None:
        return {"fullAddress": None, "postalCode": None}

    if isinstance(value, dict):
        full_address = norm_str(value.get("fullAddress") or value.get("address") or value.get("value"))
        postal_code = norm_str(value.get("postalCode") or value.get("zip"))
        extracted = extract_postal_code_and_rest(full_address)
        return {
            "postalCode": postal_code or extracted["postalCode"],
            "fullAddress": extracted["fullAddress"],
        }

    return extract_postal_code_and_rest(norm_str(value))


def address_obj_to_one_line(addr_obj):
    if not isinstance(addr_obj, dict):
        return dash_str(addr_obj)
    postal_code = norm_str(addr_obj.get("postalCode"))
    full_address = norm_str(addr_obj.get("fullAddress"))
    if postal_code and full_address:
        return f"{postal_code}, {full_address}"
    return full_address or postal_code or "-"


def join_non_empty(values, sep="; "):
    if not isinstance(values, list):
        return "-"
    cleaned = [str(v).strip() for v in values if v is not None and str(v).strip()]
    return sep.join(cleaned) if cleaned else "-"


def latin_key(text, to_lower=True):
    if not text:
        return ""
    translit = {
        "Р°": "a",
        "Р±": "b",
        "РІ": "v",
        "Рі": "g",
        "Рґ": "d",
        "Рµ": "e",
        "С‘": "e",
        "Р¶": "zh",
        "Р·": "z",
        "Рё": "i",
        "Р№": "y",
        "Рє": "k",
        "Р»": "l",
        "Рј": "m",
        "РЅ": "n",
        "Рѕ": "o",
        "Рї": "p",
        "СЂ": "r",
        "СЃ": "s",
        "С‚": "t",
        "Сѓ": "u",
        "С„": "f",
        "С…": "kh",
        "С†": "ts",
        "С‡": "ch",
        "С€": "sh",
        "С‰": "shch",
        "СЉ": "",
        "С‹": "y",
        "СЊ": "",
        "СЌ": "e",
        "СЋ": "yu",
        "СЏ": "ya",
    }
    text = "".join(translit.get(ch.lower(), ch) for ch in str(text))
    text = re.sub(r"[\s\-.(),/]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text.lower() if to_lower else text


MAP_MARKET_TYPE = {
    "specialized": {"code": "specialized", "name": "РЎРїРµС†РёР°Р»РёР·РёСЂРѕРІР°РЅРЅР°СЏ"},
    "universal": {"code": "universal", "name": "РЈРЅРёРІРµСЂСЃР°Р»СЊРЅР°СЏ"},
}


def map_market_type(value):
    raw = norm_str(value)
    if not raw:
        return None
    key = norm_key(raw)
    if key in {"specialized", norm_key("РЎРїРµС†РёР°Р»РёР·РёСЂРѕРІР°РЅРЅР°СЏ")}:
        return MAP_MARKET_TYPE["specialized"]
    if key in {"universal", norm_key("РЈРЅРёРІРµСЂСЃР°Р»СЊРЅР°СЏ")}:
        return MAP_MARKET_TYPE["universal"]
    return {"code": None, "name": raw}


MAP_MARKET_SPEC = {
    "agricultural": {"code": "agricultural", "name": "РЎРµР»СЊСЃРєРѕС…РѕР·СЏР№СЃС‚РІРµРЅРЅР°СЏ"},
    "fleaMarkets": {"code": "fleaMarkets", "name": "Р‘Р»РѕС€РёРЅС‹Рµ СЂС‹РЅРєРё"},
    "food": {"code": "food", "name": "РџСЂРѕРґРѕРІРѕР»СЊСЃС‚РІРµРЅРЅР°СЏ"},
    "industrial": {"code": "industrial", "name": "РџСЂРѕРјС‹С€Р»РµРЅРЅР°СЏ"},
    "other": {"code": "other", "name": "РРЅР°СЏ"},
    "specializedSales": {"code": "specializedSales", "name": "РџСЂРѕРґР°Р¶Р° РѕРїСЂРµРґРµР»РµРЅРЅРѕРіРѕ РІРёРґР° С‚РѕРІР°СЂРѕРІ"},
    "vernissage": {"code": "vernissage", "name": "Р’РµСЂРЅРёСЃР°Р¶"},
    "winery": {"code": "winery", "name": "Р’РёРЅРѕРґРµР»СЊС‡РµСЃРєР°СЏ РїСЂРѕРґСѓРєС†РёСЏ"},
}


def map_market_specialization(value):
    raw = norm_str(value)
    if not raw:
        return None
    key = norm_key(raw)
    ru_to_code = [
        ("РЎРµР»СЊСЃРєРѕС…РѕР·СЏР№СЃС‚РІРµРЅРЅР°СЏ", "agricultural"),
        ("Р‘Р»РѕС€РёРЅС‹Рµ СЂС‹РЅРєРё", "fleaMarkets"),
        ("РџСЂРѕРґРѕРІРѕР»СЊСЃС‚РІРµРЅРЅР°СЏ", "food"),
        ("РџСЂРѕРјС‹С€Р»РµРЅРЅР°СЏ", "industrial"),
        ("РРЅР°СЏ", "other"),
        ("РџСЂРѕРґР°Р¶Р° РѕРїСЂРµРґРµР»РµРЅРЅРѕРіРѕ РІРёРґР° С‚РѕРІР°СЂРѕРІ", "specializedSales"),
        ("Р’РµСЂРЅРёСЃР°Р¶", "vernissage"),
        ("Р’РёРЅРѕРґРµР»СЊС‡РµСЃРєР°СЏ РїСЂРѕРґСѓРєС†РёСЏ", "winery"),
    ]
    for ru_name, code in ru_to_code:
        if key == norm_key(ru_name):
            return MAP_MARKET_SPEC[code]
    if "СЃРµР»СЊСЃРєРѕС…" in key:
        return MAP_MARKET_SPEC["agricultural"]
    if "Р±Р»РѕС€" in key:
        return MAP_MARKET_SPEC["fleaMarkets"]
    if "РїСЂРѕРґРѕРІ" in key:
        return MAP_MARKET_SPEC["food"]
    if "РїСЂРѕРјС‹С€" in key:
        return MAP_MARKET_SPEC["industrial"]
    if "РІРµСЂРЅРёСЃР°Р¶" in key:
        return MAP_MARKET_SPEC["vernissage"]
    if "РІРёРЅРѕРґ" in key:
        return MAP_MARKET_SPEC["winery"]
    if "РёРЅР°СЏ" in key:
        return MAP_MARKET_SPEC["other"]
    return {"code": None, "name": raw}


MAP_STATUS_FAIR_PLACE = {
    "approved": {"code": "approved", "name": "РЈС‚РІРµСЂР¶РґРµРЅРѕ"},
    "draft": {"code": "draft", "name": "Р§РµСЂРЅРѕРІРёРє"},
    "liquid": {"code": "liquid", "name": "Р›РёРєРІРёРґРёСЂРѕРІР°РЅРѕ"},
}


def map_status_fair_place(value):
    raw = norm_str(value)
    if not raw:
        return None
    key = norm_key(raw)
    if key == norm_key("РЈС‚РІРµСЂР¶РґРµРЅРѕ"):
        return MAP_STATUS_FAIR_PLACE["approved"]
    if key == norm_key("Р§РµСЂРЅРѕРІРёРє"):
        return MAP_STATUS_FAIR_PLACE["draft"]
    if key == norm_key("Р›РёРєРІРёРґРёСЂРѕРІР°РЅРѕ"):
        return MAP_STATUS_FAIR_PLACE["liquid"]
    return {"code": None, "name": raw}


MAP_STATUS_FAIR_PLACE_SPECIFIC = {
    "approvedFree": {"code": "approvedFree", "name": "РЎРІРѕР±РѕРґРЅРѕ", "parentCode": "approved"},
    "approvedUsed": {"code": "approvedUsed", "name": "РСЃРїРѕР»СЊР·СѓРµС‚СЃСЏ", "parentCode": "approved"},
    "draftApplicant": {"code": "draftApplicant", "name": "РџСЂРµРґР»РѕР¶РµРЅРѕ Р·Р°СЏРІРёС‚РµР»РµРј", "parentCode": "draft"},
    "draftMun": {"code": "draftMun", "name": "РџСЂРµРґР»РѕР¶РµРЅРѕ РјСѓРЅРёС†РёРїР°Р»РёС‚РµС‚РѕРј", "parentCode": "draft"},
}


def map_status_fair_place_specific(value, status_fair_place):
    raw = norm_str(value)
    if not raw:
        return None
    key = norm_key(raw)
    if key == norm_key("РЎРІРѕР±РѕРґРЅРѕ"):
        return MAP_STATUS_FAIR_PLACE_SPECIFIC["approvedFree"]
    if key == norm_key("РСЃРїРѕР»СЊР·СѓРµС‚СЃСЏ"):
        return MAP_STATUS_FAIR_PLACE_SPECIFIC["approvedUsed"]
    if key == norm_key("РџСЂРµРґР»РѕР¶РµРЅРѕ Р·Р°СЏРІРёС‚РµР»РµРј"):
        return MAP_STATUS_FAIR_PLACE_SPECIFIC["draftApplicant"]
    if key == norm_key("РџСЂРµРґР»РѕР¶РµРЅРѕ РјСѓРЅРёС†РёРїР°Р»РёС‚РµС‚РѕРј"):
        return MAP_STATUS_FAIR_PLACE_SPECIFIC["draftMun"]
    parent_code = status_fair_place.get("code") if status_fair_place else None
    return {"code": None, "name": raw, "parentCode": parent_code}


MAP_MARKET_STATUS = {
    "Р°РєС‚РёРІРЅР°": {"code": "active", "name": "РђРєС‚РёРІРЅР°"},
    "РґРµР№СЃС‚РІСѓРµС‚": {"code": "active", "name": "Р”РµР№СЃС‚РІСѓРµС‚"},
    "РѕС‚РјРµРЅРµРЅР°": {"code": "cancelled", "name": "РћС‚РјРµРЅРµРЅР°"},
    "С‡РµСЂРЅРѕРІРёРє": {"code": "draft", "name": "Р§РµСЂРЅРѕРІРёРє"},
    "Р·Р°РІРµСЂС€РµРЅР°": {"code": "finished", "name": "Р—Р°РІРµСЂС€РµРЅР°"},
    "Р·Р°РїР»Р°РЅРёСЂРѕРІР°РЅР°": {"code": "planned", "name": "Р—Р°РїР»Р°РЅРёСЂРѕРІР°РЅР°"},
}


def map_market_status(value):
    raw = norm_str(value)
    if not raw:
        return None
    return MAP_MARKET_STATUS.get(norm_key(raw), {"code": None, "name": raw})


MAP_MARKET_FREQUENCY_PARENT = {
    "СЂРµРіСѓР»СЏСЂРЅР°СЏ": {"code": "regular", "name": "Р РµРіСѓР»СЏСЂРЅР°СЏ"},
    "СЂР°Р·РѕРІР°СЏ": {"code": "single", "name": "Р Р°Р·РѕРІР°СЏ"},
}


def map_market_frequency_parent(value):
    raw = norm_str(value)
    if not raw:
        return None
    return MAP_MARKET_FREQUENCY_PARENT.get(norm_key(raw), {"code": None, "name": raw})


MAP_MARKET_VARIATION_REGULAR = {
    "РїРѕСЃС‚РѕСЏРЅРЅРѕ РґРµР№СЃС‚РІСѓСЋС‰Р°СЏ": {"code": "regularPermanent", "name": "РџРѕСЃС‚РѕСЏРЅРЅРѕ РґРµР№СЃС‚РІСѓСЋС‰Р°СЏ", "parentCode": "regular"},
    "СЃРµР·РѕРЅРЅР°СЏ": {"code": "regularSeasonal", "name": "РЎРµР·РѕРЅРЅР°СЏ", "parentCode": "regular"},
    "РІС‹С…РѕРґРЅРѕРіРѕ РґРЅСЏ": {"code": "regularWeekend", "name": "Р’С‹С…РѕРґРЅРѕРіРѕ РґРЅСЏ", "parentCode": "regular"},
    "РµР¶РµРЅРµРґРµР»СЊРЅР°СЏ": {"code": "regularWeekly", "name": "Р•Р¶РµРЅРµРґРµР»СЊРЅР°СЏ", "parentCode": "regular"},
}

MAP_MARKET_VARIATION_SINGLE = {
    "РїСЂР°Р·РґРЅРёС‡РЅР°СЏ": {"code": "singleFestive", "name": "РџСЂР°Р·РґРЅРёС‡РЅР°СЏ", "parentCode": "single"},
    "СЃРµР·РѕРЅРЅР°СЏ": {"code": "singleSeasonal", "name": "РЎРµР·РѕРЅРЅР°СЏ", "parentCode": "single"},
    "С‚РµРјР°С‚РёС‡РµСЃРєР°СЏ": {"code": "singleThematic", "name": "РўРµРјР°С‚РёС‡РµСЃРєР°СЏ", "parentCode": "single"},
}


def map_market_variation_child(value, market_frequency_parent):
    raw = norm_str(value)
    if not raw or not market_frequency_parent:
        return None
    key = norm_key(raw)
    if market_frequency_parent.get("code") == "regular":
        return MAP_MARKET_VARIATION_REGULAR.get(key, {"code": None, "name": raw, "parentCode": "regular"})
    if market_frequency_parent.get("code") == "single":
        return MAP_MARKET_VARIATION_SINGLE.get(key, {"code": None, "name": raw, "parentCode": "single"})
    return None


MAP_DAY_OF_WEEK = {
    "Р±РµР· РІС‹С…РѕРґРЅС‹С…": {"code": "AllWeek", "name": "Р‘РµР· РІС‹С…РѕРґРЅС‹С…"},
    "РїРѕРЅРµРґРµР»СЊРЅРёРє": {"code": "Monday", "name": "РџРѕРЅРµРґРµР»СЊРЅРёРє"},
    "РІС‚РѕСЂРЅРёРє": {"code": "Tuesday", "name": "Р’С‚РѕСЂРЅРёРє"},
    "СЃСЂРµРґР°": {"code": "Wednesday", "name": "РЎСЂРµРґР°"},
    "С‡РµС‚РІРµСЂРі": {"code": "Thursday", "name": "Р§РµС‚РІРµСЂРі"},
    "РїСЏС‚РЅРёС†Р°": {"code": "Friday", "name": "РџСЏС‚РЅРёС†Р°"},
    "СЃСѓР±Р±РѕС‚Р°": {"code": "Saturday", "name": "РЎСѓР±Р±РѕС‚Р°"},
    "РІРѕСЃРєСЂРµСЃРµРЅСЊРµ": {"code": "Sunday", "name": "Р’РѕСЃРєСЂРµСЃРµРЅСЊРµ"},
}

MAP_MARKET_PURPOSE = {
    "РїСЂРѕРґРІРёР¶РµРЅРёРµ С†РµРЅРЅРѕСЃС‚РµР№ РЅР°С†РёРѕРЅР°Р»СЊРЅРѕР№ РєСѓР»СЊС‚СѓСЂС‹ СЃСЂРµРґРё РѕС‚РµС‡РµСЃС‚РІРµРЅРЅС‹С… Рё РёРЅРѕСЃС‚СЂР°РЅРЅС‹С… РїРѕСЃРµС‚РёС‚РµР»РµР№": {
        "code": "culturePromotion",
        "name": "РџСЂРѕРґРІРёР¶РµРЅРёРµ С†РµРЅРЅРѕСЃС‚РµР№ РЅР°С†РёРѕРЅР°Р»СЊРЅРѕР№ РєСѓР»СЊС‚СѓСЂС‹ СЃСЂРµРґРё РѕС‚РµС‡РµСЃС‚РІРµРЅРЅС‹С… Рё РёРЅРѕСЃС‚СЂР°РЅРЅС‹С… РїРѕСЃРµС‚РёС‚РµР»РµР№",
    },
    "РёРЅР°СЏ": {"code": "other", "name": "РРЅР°СЏ"},
    "СЂР°СЃС€РёСЂРµРЅРёРµ РєР°РЅР°Р»РѕРІ СЃР±С‹С‚Р° РїСЂРѕРґСѓРєС†РёРё РѕС‚РµС‡РµСЃС‚РІРµРЅРЅС‹С…, СЂРµРіРёРѕРЅР°Р»СЊРЅС‹С…, Р»РѕРєР°Р»СЊРЅС‹С… С‚РѕРІР°СЂРѕРїСЂРѕРёР·РІРѕРґРёС‚РµР»РµР№, РІ С‚РѕРј С‡РёСЃР»Рµ Рё РЅР° РјРµР¶РґСѓРЅР°СЂРѕРґРЅРѕРј СѓСЂРѕРІРЅРµ": {
        "code": "marketExpansion",
        "name": "Р Р°СЃС€РёСЂРµРЅРёРµ РєР°РЅР°Р»РѕРІ СЃР±С‹С‚Р° РїСЂРѕРґСѓРєС†РёРё РѕС‚РµС‡РµСЃС‚РІРµРЅРЅС‹С…, СЂРµРіРёРѕРЅР°Р»СЊРЅС‹С…, Р»РѕРєР°Р»СЊРЅС‹С… С‚РѕРІР°СЂРѕРїСЂРѕРёР·РІРѕРґРёС‚РµР»РµР№, РІ С‚РѕРј С‡РёСЃР»Рµ Рё РЅР° РјРµР¶РґСѓРЅР°СЂРѕРґРЅРѕРј СѓСЂРѕРІРЅРµ",
    },
    "СЃРѕР·РґР°РЅРёРµ РєРѕРјС„РѕСЂС‚РЅРѕР№ РїРѕС‚СЂРµР±РёС‚РµР»СЊСЃРєРѕР№ СЃСЂРµРґС‹": {"code": "consumerEnvironment", "name": "РЎРѕР·РґР°РЅРёРµ РєРѕРјС„РѕСЂС‚РЅРѕР№ РїРѕС‚СЂРµР±РёС‚РµР»СЊСЃРєРѕР№ СЃСЂРµРґС‹"},
    "РїРѕРґРґРµСЂР¶РєР° РѕС‚РµС‡РµСЃС‚РІРµРЅРЅС‹С… С‚РѕРІР°СЂРѕРїСЂРѕРёР·РІРѕРґРёС‚РµР»РµР№ РІ СЂРµР°Р»РёР·Р°С†РёРё СЃРѕР±СЃС‚РІРµРЅРЅРѕР№ РїСЂРѕРґСѓРєС†РёРё": {
        "code": "domesticSalesSupport",
        "name": "РџРѕРґРґРµСЂР¶РєР° РѕС‚РµС‡РµСЃС‚РІРµРЅРЅС‹С… С‚РѕРІР°СЂРѕРїСЂРѕРёР·РІРѕРґРёС‚РµР»РµР№ РІ СЂРµР°Р»РёР·Р°С†РёРё СЃРѕР±СЃС‚РІРµРЅРЅРѕР№ РїСЂРѕРґСѓРєС†РёРё",
    },
    "РѕР±РµСЃРїРµС‡РµРЅРёРµ Р·РЅР°РєРѕРјСЃС‚РІР° СЃ РЅР°С†РёРѕРЅР°Р»СЊРЅРѕР№ РёР»Рё РјРµСЃС‚РЅРѕР№ РёР»Рё СЂРµРіРёРѕРЅР°Р»СЊРЅРѕР№ РєСѓР»СЊС‚СѓСЂРѕР№, РєСѓС…РЅРµР№, С‚СЂР°РґРёС†РёСЏРјРё": {
        "code": "culturalAwareness",
        "name": "РћР±РµСЃРїРµС‡РµРЅРёРµ Р·РЅР°РєРѕРјСЃС‚РІР° СЃ РЅР°С†РёРѕРЅР°Р»СЊРЅРѕР№ РёР»Рё РјРµСЃС‚РЅРѕР№ РёР»Рё СЂРµРіРёРѕРЅР°Р»СЊРЅРѕР№ РєСѓР»СЊС‚СѓСЂРѕР№, РєСѓС…РЅРµР№, С‚СЂР°РґРёС†РёСЏРјРё",
    },
    "С„РѕСЂРјРёСЂРѕРІР°РЅРёРµ СЌС„С„РµРєС‚РёРІРЅРѕР№ РєРѕРЅРєСѓСЂРµРЅС‚РЅРѕР№ СЃСЂРµРґС‹": {"code": "competitiveEnvironment", "name": "Р¤РѕСЂРјРёСЂРѕРІР°РЅРёРµ СЌС„С„РµРєС‚РёРІРЅРѕР№ РєРѕРЅРєСѓСЂРµРЅС‚РЅРѕР№ СЃСЂРµРґС‹"},
}

MAP_PERMISSION_STATUS = {
    "РґРµР№СЃС‚РІСѓРµС‚": {"code": "Working", "name": "Р”РµР№СЃС‚РІСѓРµС‚"},
    "РїСЂРёРѕСЃС‚Р°РЅРѕРІР»РµРЅРѕ": {"code": "Stop", "name": "РџСЂРёРѕСЃС‚Р°РЅРѕРІР»РµРЅРѕ"},
    "Р°РЅРЅСѓР»РёСЂРѕРІР°РЅРѕ": {"code": "Annul", "name": "РђРЅРЅСѓР»РёСЂРѕРІР°РЅРѕ"},
    "РЅРµ РґРµР№СЃС‚РІСѓРµС‚": {"code": "doesNotWork", "name": "РќРµ РґРµР№СЃС‚РІСѓРµС‚"},
    "С‡РµСЂРЅРѕРІРёРє": {"code": "Draft", "name": "Р§РµСЂРЅРѕРІРёРє"},
}


def map_dict_value(value, mapping):
    raw = norm_str(value)
    if not raw:
        return None
    if isinstance(value, dict) and ("code" in value or "name" in value):
        return value
    hit = mapping.get(norm_key(raw))
    return copy.deepcopy(hit) if hit else {"code": None, "name": raw}


def map_day_of_week(value):
    return map_dict_value(value, MAP_DAY_OF_WEEK)


def map_market_purpose(value):
    return map_dict_value(value, MAP_MARKET_PURPOSE)


def map_permission_status(value):
    return map_dict_value(value, MAP_PERMISSION_STATUS)


def map_org_state_form(value):
    raw = norm_str(value)
    if not raw:
        return None
    return {"code": None, "name": raw}


def to_time_hhmm(value):
    if value is None:
        return None
    text = norm_str(value)
    if not text or text == "-":
        return None
    if re.match(r"^\d{1,2}:\d{2}$", text):
        return text
    try:
        number = float(text)
    except Exception:
        return text
    minutes = round(number * 24 * 60)
    return f"{str((minutes // 60) % 24).zfill(2)}:{str(minutes % 60).zfill(2)}"


def format_date_to_dmy(value):
    text = norm_str(value)
    if not text:
        return None
    match_iso = re.match(r"^(\d{4})-(\d{2})-(\d{2})", text)
    if match_iso:
        return f"{match_iso.group(3)}.{match_iso.group(2)}.{match_iso.group(1)}"
    match_ru = re.match(r"^(\d{1,2})[./](\d{1,2})[./](\d{4})$", text)
    if match_ru:
        return f"{pad2(match_ru.group(1))}.{pad2(match_ru.group(2))}.{match_ru.group(3)}"
    match_y = re.match(r"^(\d{4})[./](\d{1,2})[./](\d{1,2})$", text)
    if match_y:
        return f"{pad2(match_y.group(3))}.{pad2(match_y.group(2))}.{match_y.group(1)}"
    return text


def as_date_or_dash_dmy(value):
    return format_date_to_dmy(value) or "-"


def set_by_path(obj, path, value):
    parts = []
    for segment in str(path).split("."):
        match = re.match(r"^([^\[]+)((\[\d+\])*)$", segment)
        if not match:
            parts.append(segment)
            continue
        parts.append(match.group(1))
        for idx in re.findall(r"\[(\d+)\]", match.group(2) or ""):
            parts.append(int(idx))

    current = obj
    i = 0
    while i < len(parts):
        key = parts[i]
        last = i == len(parts) - 1
        next_key = parts[i + 1] if i + 1 < len(parts) else None

        if isinstance(key, str):
            if isinstance(next_key, int):
                if key not in current or not isinstance(current[key], list):
                    current[key] = []
                while len(current[key]) <= next_key:
                    current[key].append({})
                if i + 1 == len(parts) - 1:
                    current[key][next_key] = value
                    return
                current = current[key][next_key]
                i += 2
                continue
            if last:
                current[key] = value
                return
            if key not in current or not isinstance(current[key], dict):
                current[key] = {}
            current = current[key]
        else:
            if last:
                current[key] = value
                return
            current = current[key]
        i += 1


def build_day_on_block_from_row(row):
    result = []
    seen = set()
    for index in range(1, 101):
        day = norm_str(row.get(f"{index}. Р”РµРЅСЊ РЅРµРґРµР»Рё, РєРѕС‚РѕСЂС‹Р№ РѕС‚Р»РёС‡Р°РµС‚СЃСЏ РѕС‚ РѕСЃРЅРѕРІРЅРѕРіРѕ"))
        start = to_time_hhmm(row.get(f"{index}. Р’СЂРµРјСЏ РЅР°С‡Р°Р»Р° СЂР°Р±РѕС‚С‹ СЏСЂРјР°СЂРєРё"))
        end = to_time_hhmm(row.get(f"{index}. Р’СЂРµРјСЏ РѕРєРѕРЅС‡Р°РЅРёСЏ СЂР°Р±РѕС‚С‹ СЏСЂРјР°СЂРєРё"))
        if not day and not start and not end:
            if index > 10:
                break
            continue
        item = {
            "dayOn": map_day_of_week(day),
            "marketOpeningTimeOther": start,
            "marketClosingTimeOther": end,
        }
        key = f"{(item['dayOn'] or {}).get('code')}|{start}|{end}"
        if key in seen:
            continue
        seen.add(key)
        result.append(item)
    return result


def build_block_day_off_from_row(row):
    result = []
    seen = set()
    for index in range(1, 101):
        day = norm_str(row.get(f"{index}. Р’С‹С…РѕРґРЅРѕР№ РґРµРЅСЊ"))
        if not day:
            if index > 10:
                break
            continue
        item = {"dayOff": map_day_of_week(day)}
        key = (item["dayOff"] or {}).get("code") or (item["dayOff"] or {}).get("name")
        if key in seen:
            continue
        seen.add(key)
        result.append(item)
    return result


def infer_variation_from_frequency_name(value):
    key = norm_key(value)
    if key in {"РїСЂР°Р·РґРЅРёС‡РЅР°СЏ", "С‚РµРјР°С‚РёС‡РµСЃРєР°СЏ"}:
        return {"code": "single", "name": "Р Р°Р·РѕРІР°СЏ"}
    if key in {"РїРѕСЃС‚РѕСЏРЅРЅРѕ РґРµР№СЃС‚РІСѓСЋС‰Р°СЏ", "РІС‹С…РѕРґРЅРѕРіРѕ РґРЅСЏ", "РµР¶РµРЅРµРґРµР»СЊРЅР°СЏ"}:
        return {"code": "regular", "name": "Р РµРіСѓР»СЏСЂРЅР°СЏ"}
    return None


def map_market_frequency(value, variation_obj):
    raw = norm_str(value)
    if not raw:
        return None
    key = norm_key(raw)
    if key == "СЃРµР·РѕРЅРЅР°СЏ":
        parent_code = (variation_obj or {}).get("code")
        if parent_code == "regular":
            return {"code": "regularSeasonal", "name": "РЎРµР·РѕРЅРЅР°СЏ", "parentCode": "regular"}
        if parent_code == "single":
            return {"code": "singleSeasonal", "name": "РЎРµР·РѕРЅРЅР°СЏ", "parentCode": "single"}
        return {"code": None, "name": raw, "parentCode": None}
    child_map = {
        "РїРѕСЃС‚РѕСЏРЅРЅРѕ РґРµР№СЃС‚РІСѓСЋС‰Р°СЏ": {"code": "regularPermanent", "name": "РџРѕСЃС‚РѕСЏРЅРЅРѕ РґРµР№СЃС‚РІСѓСЋС‰Р°СЏ", "parentCode": "regular"},
        "РІС‹С…РѕРґРЅРѕРіРѕ РґРЅСЏ": {"code": "regularWeekend", "name": "Р’С‹С…РѕРґРЅРѕРіРѕ РґРЅСЏ", "parentCode": "regular"},
        "РµР¶РµРЅРµРґРµР»СЊРЅР°СЏ": {"code": "regularWeekly", "name": "Р•Р¶РµРЅРµРґРµР»СЊРЅР°СЏ", "parentCode": "regular"},
        "РїСЂР°Р·РґРЅРёС‡РЅР°СЏ": {"code": "singleFestive", "name": "РџСЂР°Р·РґРЅРёС‡РЅР°СЏ", "parentCode": "single"},
        "С‚РµРјР°С‚РёС‡РµСЃРєР°СЏ": {"code": "singleThematic", "name": "РўРµРјР°С‚РёС‡РµСЃРєР°СЏ", "parentCode": "single"},
    }
    hit = child_map.get(key)
    return copy.deepcopy(hit) if hit else {"code": None, "name": raw, "parentCode": None}


def get_market_indices(row):
    indices = set()
    for key in row.keys():
        match = re.match(r"^(\d+)\.\s", str(key))
        if match:
            indices.add(int(match.group(1)))
    return sorted(indices)


def resolve_unit_for_mesto(row, session, logger):
    default_org = copy.deepcopy(DEFAULT_ORG)
    if TEST or session is None:
        return default_org

    ogrn = norm_str(row.get("РћР“Р Рќ СѓРїРѕР»РЅРѕРјРѕС‡РµРЅРЅРѕРіРѕ РѕСЂРіР°РЅР°"))
    if not ogrn:
        logger.warning("РћР“Р Рќ СѓРїРѕР»РЅРѕРјРѕС‡РµРЅРЅРѕРіРѕ РѕСЂРіР°РЅР° РѕС‚СЃСѓС‚СЃС‚РІСѓРµС‚, РёСЃРїРѕР»СЊР·СѓРµС‚СЃСЏ UNIT РёР· РєРѕРЅС„РёРіСѓСЂР°С†РёРё")
        return default_org

    body = {"search": {"search": [{"field": "ogrn", "operator": "eq", "value": ogrn}]}, "size": 2}
    try:
        response = api_request(session, logger, "post", f"{BASE_URL}/api/v1/search/organizations", json=body, max_retries=1)
        if response.status_code != 200:
            logger.warning(f"РџРѕРёСЃРє РѕСЂРіР°РЅРёР·Р°С†РёРё РїРѕ РћР“Р Рќ {ogrn} РІРµСЂРЅСѓР» HTTP {response.status_code}, РёСЃРїРѕР»СЊР·СѓРµС‚СЃСЏ UNIT РёР· РєРѕРЅС„РёРіСѓСЂР°С†РёРё")
            return default_org
        payload = response.json()
        items = payload.get("content") or []
        if len(items) == 1:
            return items[0]
        if len(items) > 1:
            logger.warning(f"РџРѕ РћР“Р Рќ {ogrn} РЅР°Р№РґРµРЅРѕ РЅРµСЃРєРѕР»СЊРєРѕ РѕСЂРіР°РЅРёР·Р°С†РёР№, РёСЃРїРѕР»СЊР·СѓРµС‚СЃСЏ UNIT РёР· РєРѕРЅС„РёРіСѓСЂР°С†РёРё")
            return default_org
        logger.warning(f"РџРѕ РћР“Р Рќ {ogrn} РѕСЂРіР°РЅРёР·Р°С†РёСЏ РЅРµ РЅР°Р№РґРµРЅР°, РёСЃРїРѕР»СЊР·СѓРµС‚СЃСЏ UNIT РёР· РєРѕРЅС„РёРіСѓСЂР°С†РёРё")
    except Exception as exc:
        logger.warning(f"РћС€РёР±РєР° РїРѕРёСЃРєР° РѕСЂРіР°РЅРёР·Р°С†РёРё РїРѕ РћР“Р Рќ {ogrn}: {exc}. РСЃРїРѕР»СЊР·СѓРµС‚СЃСЏ UNIT РёР· РєРѕРЅС„РёРіСѓСЂР°С†РёРё")
    return default_org


def build_organizer_blocks(row, index):
    choose_authority = parse_bool(row.get(f"{index}. Р’С‹Р±СЂР°С‚СЊ РѕСЂРіР°РЅ РІР»Р°СЃС‚Рё РІ РєР°С‡РµСЃС‚РІРµ РѕСЂРіР°РЅРёР·Р°С‚РѕСЂР°"))
    status_legal = norm_str(row.get(f"{index}. РџСЂР°РІРѕРІРѕР№ СЃС‚Р°С‚СѓСЃ РѕСЂРіР°РЅРёР·Р°С‚РѕСЂР°"))

    ul_info = {
        "fullName": norm_str(row.get(f"{index}. РџРѕР»РЅРѕРµ РЅР°РёРјРµРЅРѕРІР°РЅРёРµ")),
        "ogrn": norm_str(row.get(f"{index}. РћР“Р Рќ")),
        "inn": norm_str(row.get(f"{index}. РРќРќ")),
        "phoneNumber": norm_str(row.get(f"{index}. РќРѕРјРµСЂ С‚РµР»РµС„РѕРЅР°")),
        "email": norm_str(row.get(f"{index}. Р­Р»РµРєС‚СЂРѕРЅРЅР°СЏ РїРѕС‡С‚Р°")),
    }
    ip_info = {
        "nameIP": norm_str(row.get(f"{index}. РќР°РёРјРµРЅРѕРІР°РЅРёРµ РРџ")),
        "ogrnIP": norm_str(row.get(f"{index}. РћР“Р РќРРџ")),
        "inn": norm_str(row.get(f"{index}. РРќРќ.1")),
        "phoneNumber": norm_str(row.get(f"{index}. РќРѕРјРµСЂ С‚РµР»РµС„РѕРЅР°.1")),
        "email": norm_str(row.get(f"{index}. Р­Р»РµРєС‚СЂРѕРЅРЅР°СЏ РїРѕС‡С‚Р°.1")),
    }
    authority_info = {
        "fullName": norm_str(row.get(f"{index}. РћСЂРіР°РЅРёР·Р°С‚РѕСЂ СЏСЂРјР°СЂРєРё")),
        "phoneNumber": norm_str(row.get(f"{index}. РќРѕРјРµСЂ С‚РµР»РµС„РѕРЅР°.2")),
        "email": norm_str(row.get(f"{index}. Р­Р»РµРєС‚СЂРѕРЅРЅР°СЏ РїРѕС‡С‚Р°.2")),
    }

    ul_has = has_any_non_null(ul_info)
    ip_has = has_any_non_null(ip_info)

    if choose_authority:
        return {
            "chooseAuthorityAsOrganizator": True,
            "statusLegal": status_legal,
            "blockOrganizerULInfo": None,
            "blockOrganizerIPInfo": None,
            "blockOrganizerAuthority": authority_info if has_any_non_null(authority_info) else None,
        }

    status_key = norm_key(status_legal)
    if ul_has and (not ip_has or status_key == norm_key("Р®СЂРёРґРёС‡РµСЃРєРѕРµ Р»РёС†Рѕ")):
        return {
            "chooseAuthorityAsOrganizator": False,
            "statusLegal": status_legal or "Р®СЂРёРґРёС‡РµСЃРєРѕРµ Р»РёС†Рѕ",
            "blockOrganizerULInfo": ul_info,
            "blockOrganizerIPInfo": None,
            "blockOrganizerAuthority": None,
        }

    if ip_has and (not ul_has or "РїСЂРµРґРїСЂРёРЅРёРј" in status_key):
        return {
            "chooseAuthorityAsOrganizator": False,
            "statusLegal": status_legal or "РРЅРґРёРІРёРґСѓР°Р»СЊРЅС‹Р№ РїСЂРµРґРїСЂРёРЅРёРјР°С‚РµР»СЊ",
            "blockOrganizerULInfo": None,
            "blockOrganizerIPInfo": ip_info,
            "blockOrganizerAuthority": None,
        }

    return {
        "chooseAuthorityAsOrganizator": False,
        "statusLegal": status_legal,
        "blockOrganizerULInfo": ul_info if ul_has else None,
        "blockOrganizerIPInfo": ip_info if ip_has else None,
        "blockOrganizerAuthority": authority_info if has_any_non_null(authority_info) else None,
    }


def build_market_info_items(row):
    market_info = []
    for index in get_market_indices(row):
        item_seed = {
            "marketNumber": norm_str(row.get(f"{index}. РќРѕРјРµСЂ СЏСЂРјР°СЂРєРё")),
            "permissionNumber": norm_str(row.get(f"{index}. РќРѕРјРµСЂ СЂР°Р·СЂРµС€РµРЅРёСЏ")),
            "marketName": norm_str(row.get(f"{index}. РќР°РёРјРµРЅРѕРІР°РЅРёРµ СЏСЂРјР°СЂРєРё")),
            "startDate": norm_str(row.get(f"{index}. Р”Р°С‚Р° РЅР°С‡Р°Р»Р° СЏСЂРјР°СЂРєРё")),
            "endDate": norm_str(row.get(f"{index}. Р”Р°С‚Р° РѕРєРѕРЅС‡Р°РЅРёСЏ СЏСЂРјР°СЂРєРё")),
            "openingTime": norm_str(row.get(f"{index}. Р’СЂРµРјСЏ РЅР°С‡Р°Р»Р° СЏСЂРјР°СЂРєРё")),
            "closingTime": norm_str(row.get(f"{index}. Р’СЂРµРјСЏ РѕРєРѕРЅС‡Р°РЅРёСЏ СЏСЂРјР°СЂРєРё")),
            "status": norm_str(row.get(f"{index}. РЎС‚Р°С‚СѓСЃ СЏСЂРјР°СЂРєРё")),
            "frequency": norm_str(row.get(f"{index}. РџРµСЂРёРѕРґРёС‡РЅРѕСЃС‚СЊ РїСЂРѕРІРµРґРµРЅРёСЏ СЏСЂРјР°СЂРєРё")),
            "variation": norm_str(row.get(f"{index}. Р’РёРґ СЏСЂРјР°СЂРєРё")),
            "placeCount": norm_str(row.get(f"{index}. РљРѕР»РёС‡РµСЃС‚РІРѕ С‚РѕСЂРіРѕРІС‹С… РјРµСЃС‚")),
            "placeCountFree": norm_str(row.get(f"{index}. РљРѕР»РёС‡РµСЃС‚РІРѕ С‚РѕСЂРіРѕРІС‹С… РјРµСЃС‚ РЅР° Р±РµР·РІРѕР·РјРµР·РґРЅРѕР№ РѕСЃРЅРѕРІРµ")),
        }
        org = build_organizer_blocks(row, index)
        has_meaningful_org = bool(
            org.get("chooseAuthorityAsOrganizator")
            or org.get("statusLegal")
            or org.get("blockOrganizerULInfo")
            or org.get("blockOrganizerIPInfo")
            or org.get("blockOrganizerAuthority")
        )

        if not has_any_non_null(item_seed) and not has_meaningful_org:
            continue

        market_frequency = map_market_frequency_parent(item_seed["frequency"])
        item = {
            "marketNumber": item_seed["marketNumber"],
            "permissionNumber": item_seed["permissionNumber"],
            "marketName": item_seed["marketName"],
            "blockMarketDates": {
                "startDate": parse_date_to_iso(item_seed["startDate"]),
                "endDate": parse_date_to_iso(item_seed["endDate"]),
            },
            "blockMarketOperatingTime": None,
            "marketStatus": map_market_status(item_seed["status"]),
            "marketFrequency": market_frequency,
            "marketVariation": map_market_variation_child(item_seed["variation"], market_frequency),
            "placeCount": item_seed["placeCount"],
            "placeCountFree": item_seed["placeCountFree"],
            "chooseAuthorityAsOrganizator": bool(org["chooseAuthorityAsOrganizator"]),
            "statusLegal": org["statusLegal"],
            "blockOrganizerULInfo": org["blockOrganizerULInfo"],
            "blockOrganizerIPInfo": org["blockOrganizerIPInfo"],
            "blockOrganizerAuthority": org["blockOrganizerAuthority"],
        }

        operating_time = {
            "marketOpeningTime": item_seed["openingTime"],
            "marketClosingTime": item_seed["closingTime"],
        }
        if has_any_non_null(operating_time):
            item["blockMarketOperatingTime"] = operating_time

        if not has_any_non_null(item["blockMarketDates"]):
            item["blockMarketDates"] = {"startDate": None, "endDate": None}
        if item["statusLegal"] is None:
            item.pop("statusLegal")
        if item["blockOrganizerULInfo"] is None:
            item.pop("blockOrganizerULInfo")
        if item["blockOrganizerIPInfo"] is None:
            item.pop("blockOrganizerIPInfo")
        if item["blockOrganizerAuthority"] is None:
            item.pop("blockOrganizerAuthority")
        if item["blockMarketOperatingTime"] is None:
            item.pop("blockMarketOperatingTime")

        market_info.append(item)
    return market_info


def process_fair_sheet(row, logger, session=None):
    unit_resolved = resolve_unit_for_mesto(row, session if not TEST else None, logger)
    general_information = {
        "subject": norm_str(row.get("РЎСѓР±СЉРµРєС‚ Р Р¤")),
        "disctrict": norm_str(row.get("РњСѓРЅРёС†РёРїР°Р»СЊРЅС‹Р№ СЂР°Р№РѕРЅ/РѕРєСЂСѓРі, РіРѕСЂРѕРґСЃРєРѕР№ РѕРєСЂСѓРі РёР»Рё РІРЅСѓС‚СЂРёРіРѕСЂРѕРґСЃРєР°СЏ С‚РµСЂСЂРёС‚РѕСЂРёСЏ")),
    }
    unit = {
        "id": unit_resolved.get("_id") or unit_resolved.get("id") or DEFAULT_ORG.get("id"),
        "_id": unit_resolved.get("_id") or DEFAULT_ORG.get("_id") or DEFAULT_ORG.get("id"),
        "guid": unit_resolved.get("guid") or DEFAULT_ORG.get("guid"),
        "name": norm_str(unit_resolved.get("name")) or DEFAULT_ORG.get("name"),
        "shortName": unit_resolved.get("shortName") or DEFAULT_ORG.get("shortName"),
        "ogrn": norm_str(row.get("РћР“Р Рќ СѓРїРѕР»РЅРѕРјРѕС‡РµРЅРЅРѕРіРѕ РѕСЂРіР°РЅР°")) or norm_str(unit_resolved.get("ogrn")) or DEFAULT_ORG.get("ogrn"),
        "inn": norm_str(row.get("РРќРќ СѓРїРѕР»РЅРѕРјРѕС‡РµРЅРЅРѕРіРѕ РѕСЂРіР°РЅР°")) or norm_str(unit_resolved.get("inn")),
        "kpp": unit_resolved.get("kpp") or DEFAULT_ORG.get("kpp"),
        "email": unit_resolved.get("email") or DEFAULT_ORG.get("email"),
        "site": unit_resolved.get("site") or DEFAULT_ORG.get("site"),
        "phone": unit_resolved.get("phone") or DEFAULT_ORG.get("phone"),
        "regions": unit_resolved.get("regions") or DEFAULT_ORG.get("regions"),
    }

    market_variation = map_market_frequency_parent(row.get("РџРµСЂРёРѕРґРёС‡РЅРѕСЃС‚СЊ РїСЂРѕРІРµРґРµРЅРёСЏ СЏСЂРјР°СЂРєРё"))
    if not market_variation:
        market_variation = infer_variation_from_frequency_name(row.get("Р’РёРґ СЏСЂРјР°СЂРєРё"))
    market_frequency = map_market_frequency(row.get("Р’РёРґ СЏСЂРјР°СЂРєРё"), market_variation)

    no_days_off = parse_bool(row.get("Р‘РµР· РІС‹С…РѕРґРЅС‹С…"))
    block_organizer_ul = {
        "fullName": norm_str(row.get("РџРѕР»РЅРѕРµ РЅР°РёРјРµРЅРѕРІР°РЅРёРµ")),
        "shortName": norm_str(row.get("РЎРѕРєСЂР°С‰С‘РЅРЅРѕРµ РЅР°РёРјРµРЅРѕРІР°РЅРёРµ")),
        "orgStateForm": map_org_state_form(row.get("РћСЂРіР°РЅРёР·Р°С†РёРѕРЅРЅРѕ-РїСЂР°РІРѕРІР°СЏ С„РѕСЂРјР°")),
        "ogrn": norm_str(row.get("РћР“Р Рќ")),
        "inn": norm_str(row.get("РРќРќ")),
        "kpp": norm_str(row.get("РљРџРџ")),
        "fioUL": norm_str(row.get("Р¤РРћ СЂСѓРєРѕРІРѕРґРёС‚РµР»СЏ")),
        "phoneNumber": norm_str(row.get("РќРѕРјРµСЂ С‚РµР»РµС„РѕРЅР°")),
        "email": norm_str(row.get("Р­Р»РµРєС‚СЂРѕРЅРЅР°СЏ РїРѕС‡С‚Р°")),
        "addressUL": parse_address_to_obj(row.get("Р®СЂРёРґРёС‡РµСЃРєРёР№ Р°РґСЂРµСЃ")),
        "addressFact": parse_address_to_obj(row.get("Р¤Р°РєС‚РёС‡РµСЃРєРёР№ Р°РґСЂРµСЃ")),
    }
    block_organizer_ip = {
        "nameIP": norm_str(row.get("РќР°РёРјРµРЅРѕРІР°РЅРёРµ РРџ")),
        "fioIP": None,
        "inn": norm_str(row.get("РРќРќ.1")),
        "ogrnIP": norm_str(row.get("РћР“Р РќРРџ")),
        "passportSeries": None,
        "passportNumber": None,
        "passportIssueDate": None,
        "passportAuthority": None,
        "passportCode": None,
        "placeBirth": None,
        "phoneNumber": None,
        "email": None,
        "addressReg": {"fullAddress": None, "postalCode": None},
        "addressPost": {"fullAddress": None, "postalCode": None},
    }
    block_organizer_authority = {
        "fullName": norm_str(row.get("РћСЂРіР°РЅРёР·Р°С‚РѕСЂ СЏСЂРјР°СЂРєРё")),
        "phoneNumber": norm_str(row.get("РќРѕРјРµСЂ С‚РµР»РµС„РѕРЅР°.1")),
        "email": norm_str(row.get("Р­Р»РµРєС‚СЂРѕРЅРЅР°СЏ РїРѕС‡С‚Р°.1")),
    }

    market_info = {
        "marketNumber": norm_str(row.get("РќРѕРјРµСЂ СЏСЂРјР°СЂРєРё")),
        "permissionNumber": norm_str(row.get("РќРѕРјРµСЂ СЂР°Р·СЂРµС€РµРЅРёСЏ")),
        "marketStatus": map_market_status(row.get("РЎС‚Р°С‚СѓСЃ СЏСЂРјР°СЂРєРё")),
        "marketName": norm_str(row.get("РќР°РёРјРµРЅРѕРІР°РЅРёРµ СЏСЂРјР°СЂРєРё")),
        "marketType": map_market_type(row.get("РўРёРї СЏСЂРјР°СЂРєРё")),
        "marketSpecialization": map_market_specialization(row.get("РЎРїРµС†РёР°Р»РёР·Р°С†РёСЏ СЏСЂРјР°СЂРєРё")),
        "marketSpecializationOther": norm_str(row.get("РРЅР°СЏ СЃРїРµС†РёР°Р»РёР·Р°С†РёСЏ СЏСЂРјР°СЂРєРё")),
        "marketPurpose": map_market_purpose(row.get("Р¦РµР»СЊ СЏСЂРјР°СЂРєРё")),
        "marketPurposeOther": norm_str(row.get("РРЅР°СЏ С†РµР»СЊ СЏСЂРјР°СЂРєРё")),
        "marketVariation": market_variation,
        "marketFrequency": market_frequency,
        "marketArea": norm_str(row.get("РџР»РѕС‰Р°РґСЊ СЏСЂРјР°СЂРєРё, РєРІ. Рј")),
        "blockMarketDates": {
            "startDate": parse_date_to_iso(row.get("РЎСЂРѕРєРё РїСЂРѕРІРµРґРµРЅРёСЏ СЏСЂРјР°СЂРєРё: РґР°С‚Р° РЅР°С‡Р°Р»Р°")),
            "endDate": parse_date_to_iso(row.get("РЎСЂРѕРєРё РїСЂРѕРІРµРґРµРЅРёСЏ СЏСЂРјР°СЂРєРё: РґР°С‚Р° РѕРєРѕРЅС‡Р°РЅРёСЏ")),
        },
        "blockMarketOperatingTime": {
            "marketOpeningTime": to_time_hhmm(row.get("Р РµР¶РёРј СЂР°Р±РѕС‚С‹ СЏСЂРјР°СЂРєРё: РІСЂРµРјСЏ РЅР°С‡Р°Р»Р°")),
            "marketClosingTime": to_time_hhmm(row.get("Р РµР¶РёРј СЂР°Р±РѕС‚С‹ СЏСЂРјР°СЂРєРё: РІСЂРµРјСЏ РѕРєРѕРЅС‡Р°РЅРёСЏ")),
        },
        "dayOnBlock": build_day_on_block_from_row(row),
        "dayOffDayOff": copy.deepcopy(MAP_DAY_OF_WEEK["Р±РµР· РІС‹С…РѕРґРЅС‹С…"]) if no_days_off is True else None,
        "BlockDayOff": [] if no_days_off is True else build_block_day_off_from_row(row),
        "sanitaryDayOfMonth": norm_str(row.get("РЎР°РЅРёС‚Р°СЂРЅС‹Р№ РґРµРЅСЊ РјРµСЃСЏС†Р°")),
        "placeCount": norm_str(row.get("РљРѕР»РёС‡РµСЃС‚РІРѕ С‚РѕСЂРіРѕРІС‹С… РјРµСЃС‚")),
        "placeCountFree": norm_str(row.get("РљРѕР»РёС‡РµСЃС‚РІРѕ С‚РѕСЂРіРѕРІС‹С… РјРµСЃС‚ РЅР° Р±РµР·РІРѕР·РјРµР·РґРЅРѕР№ РѕСЃРЅРѕРІРµ")),
        "placeNumber": norm_str(row.get("РќРѕРјРµСЂ РјРµСЃС‚Р° РґР»СЏ СЂР°Р·РјРµС‰РµРЅРёСЏ СЏСЂРјР°СЂРѕРє")),
        "blockCadNumber": [],
        "cadsObjects": [],
    }

    cad_land = norm_str(row.get("РљР°РґР°СЃС‚СЂРѕРІС‹Р№ РЅРѕРјРµСЂ Р·РµРјРµР»СЊРЅРѕРіРѕ СѓС‡Р°СЃС‚РєР°"))
    addr_land = row.get("РђРґСЂРµСЃ Р·РµРјРµР»СЊРЅРѕРіРѕ СѓС‡Р°СЃС‚РєР°")
    if cad_land or norm_str(addr_land):
        market_info["blockCadNumber"].append({"cadObjNum": cad_land, "landAddress": parse_address_to_obj(addr_land)})

    cad_obj = norm_str(row.get("РљР°РґР°СЃС‚СЂРѕРІС‹Р№ РЅРѕРјРµСЂ РѕР±СЉРµРєС‚Р° РЅРµРґРІРёР¶РёРјРѕСЃС‚Рё"))
    addr_obj = row.get("РђРґСЂРµСЃ РѕР±СЉРµРєС‚Р° РЅРµРґРІРёР¶РёРјРѕСЃС‚Рё")
    if cad_obj or norm_str(addr_obj):
        market_info["cadsObjects"].append({"cadNumber": cad_obj, "objectAddress": parse_address_to_obj(addr_obj)})

    payload = {
        "guid": generate_guid(),
        "parentEntries": FAIR_COLLECTION,
        "generalInformation": general_information,
        "unit": unit,
        "marketInfo": market_info,
        "chooseAuthorityAsOrganizator": parse_bool(row.get("Р’С‹Р±СЂР°С‚СЊ РѕСЂРіР°РЅ РІР»Р°СЃС‚Рё РІ РєР°С‡РµСЃС‚РІРµ РѕСЂРіР°РЅРёР·Р°С‚РѕСЂР°")),
        "statusLegal": norm_str(row.get("РџСЂР°РІРѕРІРѕР№ СЃС‚Р°С‚СѓСЃ РѕСЂРіР°РЅРёР·Р°С‚РѕСЂР°")),
        "blockOrganizerULInfo": block_organizer_ul,
        "blockOrganizerIPInfo": block_organizer_ip,
    }
    if has_any_non_null(block_organizer_authority):
        payload["blockOrganizerAuthority"] = block_organizer_authority
    return payload


def process_mesto_sheet(row, logger, session=None):
    unit_resolved = resolve_unit_for_mesto(row, session, logger)
    subject = norm_str(row.get("РЎСѓР±СЉРµРєС‚ Р Р¤"))
    district = norm_str(row.get("РњСѓРЅРёС†РёРїР°Р»СЊРЅС‹Р№ СЂР°Р№РѕРЅ/РѕРєСЂСѓРі, РіРѕСЂРѕРґСЃРєРѕР№ РѕРєСЂСѓРі РёР»Рё РІРЅСѓС‚СЂРёРіРѕСЂРѕРґСЃРєР°СЏ С‚РµСЂСЂРёС‚РѕСЂРёСЏ"))

    general_information = {
        "subject": subject,
        "disctrict": district,
    }

    unit = {
        "id": unit_resolved.get("_id") or unit_resolved.get("id") or DEFAULT_ORG.get("id"),
        "_id": unit_resolved.get("_id") or DEFAULT_ORG.get("_id") or DEFAULT_ORG.get("id"),
        "guid": unit_resolved.get("guid") or DEFAULT_ORG.get("guid"),
        "name": unit_resolved.get("name") or DEFAULT_ORG.get("name"),
        "shortName": unit_resolved.get("shortName") or DEFAULT_ORG.get("shortName"),
        "ogrn": norm_str(row.get("РћР“Р Рќ СѓРїРѕР»РЅРѕРјРѕС‡РµРЅРЅРѕРіРѕ РѕСЂРіР°РЅР°")) or norm_str(unit_resolved.get("ogrn")),
        "inn": norm_str(row.get("РРќРќ СѓРїРѕР»РЅРѕРјРѕС‡РµРЅРЅРѕРіРѕ РѕСЂРіР°РЅР°")) or norm_str(unit_resolved.get("inn")),
        "kpp": unit_resolved.get("kpp") or DEFAULT_ORG.get("kpp"),
        "email": unit_resolved.get("email") or DEFAULT_ORG.get("email"),
        "site": unit_resolved.get("site") or DEFAULT_ORG.get("site"),
        "phone": unit_resolved.get("phone") or DEFAULT_ORG.get("phone"),
        "regions": unit_resolved.get("regions") or DEFAULT_ORG.get("regions"),
    }

    status_fair_place = map_status_fair_place(row.get("РЎС‚Р°С‚СѓСЃ РїСЂРѕРµРєС‚РЅРѕРіРѕ РјРµСЃС‚Р°"))
    place_market_info = {
        "marketSchemeNumber": norm_str(row.get("РќРѕРјРµСЂ РјРµСЃС‚Р° РґР»СЏ СЂР°Р·РјРµС‰РµРЅРёСЏ СЏСЂРјР°СЂРѕРє")),
        "coordinates": norm_str(row.get("Р“РµРѕРєРѕРѕСЂРґРёРЅР°С‚С‹ С‚РѕС‡РєРё СЂР°Р·РјРµС‰РµРЅРёСЏ РѕСЂРіР°РЅРёР·СѓРµРјРѕР№ СЏСЂРјР°СЂРєРё")),
        "landArea": pick_first_non_empty(
            norm_str(row.get("РџР»РѕС‰Р°РґСЊ РјРµСЃС‚Р° РґР»СЏ СЂР°Р·РјРµС‰РµРЅРёСЏ СЏСЂРјР°СЂРєРё, РєРІ. Рј")),
            norm_str(row.get("РџР»РѕС‰Р°РґСЊ РјРµСЃС‚Р° РґР»СЏ СЂР°Р·РјРµС‰РµРЅРёСЏ СЏСЂРјР°СЂРєРё, РєРІ.Рј")),
        ),
        "marketType": map_market_type(row.get("РўРёРї СЏСЂРјР°СЂРєРё")),
        "marketSpecialization": map_market_specialization(row.get("РЎРїРµС†РёР°Р»РёР·Р°С†РёСЏ СЏСЂРјР°СЂРєРё")),
        "marketSpecializationOther": norm_str(row.get("РРЅР°СЏ СЃРїРµС†РёР°Р»РёР·Р°С†РёСЏ СЏСЂРјР°СЂРєРё")),
        "statusFairPlace": status_fair_place,
        "statusFairPlaceSpecific": map_status_fair_place_specific(row.get("РЎРїРµС†РёС„РёРєР°С†РёСЏ СЃС‚Р°С‚СѓСЃР° РїСЂРѕРµРєС‚РЅРѕРіРѕ РјРµСЃС‚Р°"), status_fair_place),
        "statusPlaceFair": norm_str(row.get("РЎС‚Р°С‚СѓСЃ РІ Р РµРµСЃС‚СЂРµ РјРµСЃС‚ РЅР° СЂР°Р·РјРµС‰РµРЅРёСЏ СЏСЂРјР°СЂРѕРє")),
        "blockCadNumber": [],
        "cadsObjects": [],
    }

    cad_land = norm_str(row.get("РљР°РґР°СЃС‚СЂРѕРІС‹Р№ РЅРѕРјРµСЂ Р·РµРјРµР»СЊРЅРѕРіРѕ СѓС‡Р°СЃС‚РєР°"))
    addr_land = norm_str(row.get("РђРґСЂРµСЃ Р·РµРјРµР»СЊРЅРѕРіРѕ СѓС‡Р°СЃС‚РєР°"))
    if not is_empty_val(cad_land) or not is_empty_val(addr_land):
        place_market_info["blockCadNumber"].append(
            {"cadNumber": cad_land, "landAddress": parse_address_to_obj(addr_land)}
        )

    cad_obj = norm_str(row.get("РљР°РґР°СЃС‚СЂРѕРІС‹Р№ РЅРѕРјРµСЂ РѕР±СЉРµРєС‚Р° РЅРµРґРІРёР¶РёРјРѕСЃС‚Рё"))
    addr_obj = norm_str(row.get("РђРґСЂРµСЃ РѕР±СЉРµРєС‚Р° РЅРµРґРІРёР¶РёРјРѕСЃС‚Рё"))
    if not is_empty_val(cad_obj) or not is_empty_val(addr_obj):
        place_market_info["cadsObjects"].append(
            {"cadObjNum": cad_obj, "objectAddress": parse_address_to_obj(addr_obj)}
        )

    return {
        "guid": generate_guid(),
        "generalInformation": general_information,
        "unit": unit,
        "placeMarketInfo": place_market_info,
        "marketInfo": build_market_info_items(row),
    }


def build_nsi_local_object_fair_from_payload(payload, market_item=None):
    general_information = payload.get("generalInformation") or {}
    place_market_info = payload.get("placeMarketInfo") or {}
    unit = payload.get("unit") or {}
    market = market_item or None

    land_cad_nums = [item.get("cadNumber") for item in place_market_info.get("blockCadNumber", []) if not is_empty_val(item.get("cadNumber"))]
    land_addr_str = [
        address_obj_to_one_line(item.get("landAddress"))
        for item in place_market_info.get("blockCadNumber", [])
        if address_obj_to_one_line(item.get("landAddress")) != "-"
    ]
    obj_cad_nums = [item.get("cadObjNum") for item in place_market_info.get("cadsObjects", []) if not is_empty_val(item.get("cadObjNum"))]
    obj_addr_str = [
        address_obj_to_one_line(item.get("objectAddress"))
        for item in place_market_info.get("cadsObjects", [])
        if address_obj_to_one_line(item.get("objectAddress")) != "-"
    ]

    full_address_out = join_non_empty(obj_addr_str) if obj_addr_str else join_non_empty(land_addr_str)
    cad_number_out = join_non_empty(obj_cad_nums) if obj_cad_nums else join_non_empty(land_cad_nums)

    status_name = norm_str((place_market_info.get("statusFairPlace") or {}).get("name") or (place_market_info.get("statusFairPlace") or {}).get("code"))
    status_specific_name = norm_str((place_market_info.get("statusFairPlaceSpecific") or {}).get("name") or (place_market_info.get("statusFairPlaceSpecific") or {}).get("code"))
    project_status_out = f"{status_name}: {status_specific_name}" if status_name and status_specific_name else (status_name or status_specific_name or "-")

    type_code = (place_market_info.get("marketType") or {}).get("code")
    type_name = (place_market_info.get("marketType") or {}).get("name")
    is_universal = type_code == "universal" or norm_key(type_name) == norm_key("РЈРЅРёРІРµСЂСЃР°Р»СЊРЅР°СЏ")

    specialization_out = "-"
    if not is_universal:
        spec_obj = place_market_info.get("marketSpecialization") or {}
        spec_name = norm_str(spec_obj.get("name") or spec_obj.get("code"))
        if spec_name:
            is_other = spec_obj.get("code") == "other" or norm_key(spec_name) in {norm_key("РРЅР°СЏ"), "other"}
            specialization_out = dash_str(place_market_info.get("marketSpecializationOther")) if is_other else dash_str(spec_name)

    organizer_name = "-"
    organizer_inn = "-"
    organizer_ogrn = "-"
    organizer_ogrnip = "-"
    organizer_number = "-"
    organizer_email = "-"

    if market:
        if market.get("blockOrganizerULInfo"):
            ul_info = market["blockOrganizerULInfo"]
            organizer_name = dash_str(ul_info.get("fullName"))
            organizer_inn = dash_str(ul_info.get("inn"))
            organizer_ogrn = dash_str(ul_info.get("ogrn"))
            organizer_number = dash_str(ul_info.get("phoneNumber"))
            organizer_email = dash_str(ul_info.get("email"))
        elif market.get("blockOrganizerIPInfo"):
            ip_info = market["blockOrganizerIPInfo"]
            organizer_name = dash_str(ip_info.get("nameIP"))
            organizer_inn = dash_str(ip_info.get("inn"))
            organizer_ogrnip = dash_str(ip_info.get("ogrnIP"))
        elif market.get("blockOrganizerAuthority"):
            authority_info = market["blockOrganizerAuthority"]
            organizer_name = dash_str(authority_info.get("fullName"))
            organizer_number = dash_str(authority_info.get("phoneNumber"))
            organizer_email = dash_str(authority_info.get("email"))

    nsi_guid = generate_guid()
    fair_type_name = (place_market_info.get("marketType") or {}).get("name") or (place_market_info.get("marketType") or {}).get("code")

    out = {
        "LayerId": "3",
        "Layer": "РЇСЂРјР°СЂРєРё",
        "dictionaryType": "local",
        "dictionaryUnitId": dash_str(unit.get("_id") or unit.get("id")),
        "autokey": nsi_guid,
        "code": nsi_guid,
        "ObjectID": nsi_guid,
        "parentEntries": NSI_LOCAL_OBJECT_FAIR_COLLECTION,
        "Subject": dash_str(general_information.get("subject")),
        "Disctrict": dash_str(general_information.get("disctrict")),
        "FairType": dash_str(fair_type_name),
        "SpecializationFair": dash_str(specialization_out),
        "NumberFair": dash_str(market.get("marketNumber")) if market else "-",
        "GeoCoordinates": dash_str(place_market_info.get("coordinates")),
        "CadNumber": dash_str(cad_number_out),
        "PermissionNumber": dash_str(market.get("permissionNumber")) if market else "-",
        "PermissionStartDate": "-",
        "PermissionEndDate": "-",
        "TitleFair": dash_str(market.get("marketName")) if market else "-",
        "NumberFairLocation": dash_str(place_market_info.get("marketSchemeNumber")),
        "FullAddress": dash_str(full_address_out),
        "ProjectStatus": dash_str(project_status_out),
        "LandArea": dash(place_market_info.get("landArea")),
        "FairStatus": dash_str((market.get("marketStatus") or {}).get("name") or (market.get("marketStatus") or {}).get("code")) if market else "-",
        "FairStartDate": as_date_ru_or_dash((market.get("blockMarketDates") or {}).get("startDate")) if market else "-",
        "FairEndDate": as_date_ru_or_dash((market.get("blockMarketDates") or {}).get("endDate")) if market else "-",
        "StartTime": dash_str((market.get("blockMarketOperatingTime") or {}).get("marketOpeningTime")) if market else "-",
        "EndTime": dash_str((market.get("blockMarketOperatingTime") or {}).get("marketClosingTime")) if market else "-",
        "OrganizerName": organizer_name,
        "OrganizerINN": organizer_inn,
        "OrganizerOGRN": organizer_ogrn,
        "OrganizerOGRNIP": organizer_ogrnip,
        "OrganizerNumber": organizer_number,
        "OrganizerEmail": organizer_email,
        "FrequencyFair": dash_str((market.get("marketFrequency") or {}).get("name") or (market.get("marketFrequency") or {}).get("code")) if market else "-",
        "FairKind": dash_str((market.get("marketVariation") or {}).get("name") or (market.get("marketVariation") or {}).get("code")) if market else "-",
        "NumberPlace": dash(market.get("placeCount", "-")) if market else "-",
    }

    for key, value in list(out.items()):
        if value is None or (isinstance(value, str) and not value.strip()):
            out[key] = "-"
    return out


def process_permits_sheet(row, logger):
    return {
        "guid": generate_guid(),
        "parentEntries": FAIR_PERMITS_COLLECTION,
        "unit": {
            "id": DEFAULT_ORG.get("_id") or DEFAULT_ORG.get("id"),
            "name": DEFAULT_ORG.get("name"),
            "shortName": DEFAULT_ORG.get("shortName"),
            "inn": DEFAULT_ORG.get("inn"),
            "ogrn": DEFAULT_ORG.get("ogrn"),
        },
        "generalInformation": {
            "subject": norm_str(row.get("РЎСѓР±СЉРµРєС‚ Р Р¤")),
            "disctrict": norm_str(row.get("РњСѓРЅРёС†РёРїР°Р»СЊРЅС‹Р№ СЂР°Р№РѕРЅ/РѕРєСЂСѓРі, РіРѕСЂРѕРґСЃРєРѕР№ РѕРєСЂСѓРі РёР»Рё РІРЅСѓС‚СЂРёРіРѕСЂРѕРґСЃРєР°СЏ С‚РµСЂСЂРёС‚РѕСЂРёСЏ")),
        },
        "permission": {
            "permissionNumber": norm_str(row.get("РќРѕРјРµСЂ СЂР°Р·СЂРµС€РµРЅРёСЏ")),
            "permissionStatus": map_permission_status(row.get("РЎС‚Р°С‚СѓСЃ СЂР°Р·СЂРµС€РµРЅРёСЏ")),
            "fairNumber": norm_str(row.get("РќРѕРјРµСЂ СЏСЂРјР°СЂРєРё")),
            "permissionDate": parse_date_to_iso(row.get("Р”Р°С‚Р° РІС‹РґР°С‡Рё СЂР°Р·СЂРµС€РµРЅРёСЏ")),
            "permissionStartDate": parse_date_to_iso(row.get("Р”Р°С‚Р° РЅР°С‡Р°Р»Р° РґРµР№СЃС‚РІРёСЏ СЂР°Р·СЂРµС€РµРЅРёСЏ")),
            "permissionEndDate": parse_date_to_iso(row.get("Р”Р°С‚Р° Р·Р°РІРµСЂС€РµРЅРёСЏ РґРµР№СЃС‚РІРёСЏ СЂР°Р·СЂРµС€РµРЅРёСЏ")),
            "permissionFile": None,
        },
    }


def create_record(session, logger, collection, payload):
    response = api_request(session, logger, "post", f"{BASE_URL}/api/v1/create/{collection}", json=jsonable(payload))
    if not response.ok:
        return None, response
    return response.json(), response


def log_success(success_logger, record):
    success_logger.info(json.dumps(record, ensure_ascii=False))


def find_fair_by_number(session, logger, fair_number, subject=None, district=None):
    fair_number = norm_str(fair_number)
    subject = norm_str(subject)
    district = norm_str(district)
    if not fair_number:
        return {"ok": False, "many": False, "entry": None, "tried": []}
    tried = []
    for field in ["NumberFair", "numberFair", "FairNumber", "fairNumber"]:
        tried.append(field)
        body = {"search": {"search": [{"field": field, "operator": "eq", "value": fair_number}]}, "size": 2}
        try:
            response = api_request(session, logger, "post", f"{BASE_URL}/api/v1/search/{NSI_LOCAL_OBJECT_FAIR_COLLECTION}", json=body, max_retries=1)
            if response.status_code != 200:
                continue
            content = (response.json() or {}).get("content") or []
            if subject is not None or district is not None:
                filtered = []
                for entry in content:
                    entry_subject = norm_str(entry.get("Subject"))
                    entry_district = norm_str(entry.get("Disctrict"))
                    if subject is not None and entry_subject != subject:
                        continue
                    if district is not None and entry_district != district:
                        continue
                    filtered.append(entry)
                content = filtered
            if len(content) == 1:
                return {"ok": True, "many": False, "entry": content[0], "field": field, "tried": tried}
            if len(content) > 1:
                return {"ok": False, "many": True, "entry": content, "field": field, "tried": tried}
        except Exception:
            continue
    return {"ok": False, "many": False, "entry": None, "tried": tried}


def update_fair_dates_from_permit(row, session, logger, success_logger, fail_logger, row_num):
    if TEST or session is None:
        return
    fair_number = norm_str(row.get("РќРѕРјРµСЂ СЏСЂРјР°СЂРєРё"))
    subject = norm_str(row.get("РЎСѓР±СЉРµРєС‚ Р Р¤"))
    district = norm_str(row.get("РњСѓРЅРёС†РёРїР°Р»СЊРЅС‹Р№ СЂР°Р№РѕРЅ/РѕРєСЂСѓРі, РіРѕСЂРѕРґСЃРєРѕР№ РѕРєСЂСѓРі РёР»Рё РІРЅСѓС‚СЂРёРіРѕСЂРѕРґСЃРєР°СЏ С‚РµСЂСЂРёС‚РѕСЂРёСЏ"))

    result = find_fair_by_number(session, logger, fair_number, subject=subject, district=district)
    if not result.get("ok"):
        if result.get("many"):
            logger.warning(
                f"Р”Р»СЏ СЏСЂРјР°СЂРєРё {fair_number} РЅР°Р№РґРµРЅРѕ РЅРµСЃРєРѕР»СЊРєРѕ Р·Р°РїРёСЃРµР№ nsiLocalObjectFair "
                f"СЃ Subject={subject!r} Рё Disctrict={district!r}"
            )
        else:
            logger.warning(
                f"РЇСЂРјР°СЂРєР° {fair_number} РЅРµ РЅР°Р№РґРµРЅР° РІ nsiLocalObjectFair "
                f"СЃ Subject={subject!r} Рё Disctrict={district!r}"
            )
        return

    found_fair = result["entry"]
    update_doc = {"_id": found_fair["_id"], "guid": found_fair["guid"]}
    if found_fair.get("auid"):
        update_doc["auid"] = found_fair["auid"]
    update_doc["PermissionStartDate"] = as_date_or_dash_dmy(row.get("Р”Р°С‚Р° РІС‹РґР°С‡Рё СЂР°Р·СЂРµС€РµРЅРёСЏ"))
    update_doc["PermissionEndDate"] = as_date_or_dash_dmy(row.get("Р”Р°С‚Р° Р·Р°РІРµСЂС€РµРЅРёСЏ РґРµР№СЃС‚РІРёСЏ СЂР°Р·СЂРµС€РµРЅРёСЏ"))

    response = api_request(
        session,
        logger,
        "put",
        f"{BASE_URL}/api/v1/update/{NSI_LOCAL_OBJECT_FAIR_COLLECTION}?mainId={found_fair['_id']}&guid={found_fair['guid']}",
        json=jsonable(update_doc),
    )
    if response.ok:
        log_success(success_logger, {"_id": found_fair["_id"], "guid": found_fair["guid"], "parentEntries": NSI_LOCAL_OBJECT_FAIR_COLLECTION})
        return
    logger.error(f"РћС€РёР±РєР° РѕР±РЅРѕРІР»РµРЅРёСЏ {NSI_LOCAL_OBJECT_FAIR_COLLECTION} РґР»СЏ РЅРѕРјРµСЂР° СЏСЂРјР°СЂРєРё {fair_number}: HTTP {response.status_code}")
    fail_logger.info(f"{SHEET_PERMITS}:{row_num}:fair_update")


def handle_permit_file_upload(row, rec_data, collection, record_json, session, logger, fail_logger, row_num, files_dir):
    file_hint = norm_str(row.get("Р Р°Р·СЂРµС€РµРЅРёРµ РЅР° РїСЂР°РІРѕ РѕСЂРіР°РЅРёР·Р°С†РёРё СЏСЂРјР°СЂРєРё"))
    file_path = find_file_in_dir(files_dir, file_hint) if file_hint else None
    file_field = "permission.permissionFile"

    if not file_path:
        return True

    logger.info(f"РќР°Р№РґРµРЅ С„Р°Р№Р» РґР»СЏ СЂР°Р·СЂРµС€РµРЅРёСЏ: {os.path.basename(file_path)}")
    file_object = upload_file(session, logger, file_path, collection, record_json["_id"], entity_field_path=file_field)
    if not file_object:
        logger.error("РћС€РёР±РєР° РїСЂРё Р·Р°РіСЂСѓР·РєРµ С„Р°Р№Р»Р° СЂР°Р·СЂРµС€РµРЅРёСЏ")
        delete_from_collection(session, logger, record_json)
        fail_logger.info(f"{SHEET_PERMITS}:{row_num}")
        return False

    update_payload = {
        "_id": record_json["_id"],
        "guid": rec_data["guid"],
        "permission": {**rec_data["permission"], "permissionFile": file_object},
    }
    update_url = f"{BASE_URL}/api/v1/update/{collection}?mainId={record_json['_id']}&guid={rec_data['guid']}"
    update_response = api_request(session, logger, "put", update_url, json=jsonable(update_payload))
    if update_response.ok:
        return True

    logger.error("РћС€РёР±РєР° РїСЂРё РѕР±РЅРѕРІР»РµРЅРёРё Р·Р°РїРёСЃРё СЃ С„Р°Р№Р»РѕРј")
    delete_file_from_storage(session, logger, file_object["_id"])
    delete_from_collection(session, logger, record_json)
    fail_logger.info(f"{SHEET_PERMITS}:{row_num}")
    return False


def parse_args():
    parser = argparse.ArgumentParser(description="Fairs migration runner")
    parser.add_argument("--workbook", default=os.path.join(SCRIPT_DIR, EXCEL_FILE_NAME), help="Path to XLSM workbook")
    parser.add_argument("--files-dir", default=FILES_DIR, help="Directory with attachment files")
    parser.add_argument(
        "--sheet",
        choices=["all", "fair", "mesto", "permits"],
        default="all",
        help="Sheets to process",
    )
    parser.add_argument("--limit", type=int, default=0, help="Max rows per sheet (0 = no limit)")
    parser.add_argument("--no-auth", action="store_true", help="Skip API authentication")
    parser.add_argument("--no-prompt", action="store_true", help="Do not ask for cookie/token in console")
    parser.add_argument("--no-interactive", action="store_true", help="Disable interactive prompts")
    parser.add_argument("--state-file", default=str(STATE_FILE), help="Path to checkpoints JSON")
    parser.add_argument("--reset-state", action="store_true", help="Clear checkpoints before run")
    parser.add_argument("--resume", dest="resume", action="store_true", default=RESUME_BY_DEFAULT, help="Resume from checkpoints")
    parser.add_argument("--no-resume", dest="resume", action="store_false", help="Ignore checkpoints")
    return parser.parse_args()


def _prompt_with_default(label, default_value, interactive):
    if not interactive:
        return default_value
    entered = input(f"{label} [{default_value}]: ").strip().strip('"').strip("'")
    return entered or default_value


def _ask_yes_no(prompt, default_yes=True):
    suffix = " [Y/n]: " if default_yes else " [y/N]: "
    answer = input(prompt + suffix).strip().lower()
    if not answer:
        return default_yes
    return answer in {"y", "yes"}


def _selected_lists(sheet_mode):
    mapping = {
        "fair": SHEET_FAIR,
        "mesto": SHEET_MESTO,
        "permits": SHEET_PERMITS,
    }
    if sheet_mode == "all":
        return list(EXCEL_LISTS)
    selected = mapping.get(sheet_mode)
    return [selected] if selected else list(EXCEL_LISTS)


def process_sheet_data(
    list_name,
    excel,
    session,
    logger,
    success_logger,
    fail_logger,
    *,
    workbook_path,
    files_dir,
    state,
    resume_enabled,
    limit,
    stats,
):
    collection_map = {
        SHEET_FAIR: FAIR_COLLECTION,
        SHEET_MESTO: FAIR_MESTO_COLLECTION,
        SHEET_PERMITS: FAIR_PERMITS_COLLECTION,
    }
    collection = collection_map.get(list_name)
    if not collection:
        logger.error(f"Unknown sheet: {list_name}")
        stats["failedRows"] += 1
        return

    rows_total = len(excel)
    for row_num, row in enumerate(excel.to_dict("records"), start=1):
        if limit and row_num > int(limit):
            logger.info("Достигнут лимит строк (%s) для листа %s", limit, list_name)
            break

        if resume_enabled:
            checkpoint = state.get(str(workbook_path), list_name, row_num)
            if isinstance(checkpoint, dict):
                stats["resumeSkips"] += 1
                logger.info("[RESUME][SKIP] sheet=%s row=%s _id=%s", list_name, row_num, checkpoint.get("_id"))
                continue

        stats["processedRows"] += 1
        logger.info(f"{list_name}: РѕР±СЂР°Р±РѕС‚РєР° СЃС‚СЂРѕРєРё {row_num}/{rows_total}")

        try:
            if list_name == SHEET_FAIR:
                rec_data = process_fair_sheet(row, logger, session=session if not TEST else None)
            elif list_name == SHEET_MESTO:
                rec_data = process_mesto_sheet(row, logger, session=session if not TEST else None)
            else:
                rec_data = process_permits_sheet(row, logger)

            if TEST:
                logger.info(f"TEST MODE {list_name}: {json.dumps(rec_data, ensure_ascii=False)}")
                if list_name == SHEET_MESTO:
                    markets = rec_data.get("marketInfo") or []
                    if not markets:
                        logger.info(f"TEST MODE NSI {list_name}: {json.dumps(build_nsi_local_object_fair_from_payload(rec_data, None), ensure_ascii=False)}")
                    else:
                        for market in markets:
                            logger.info(f"TEST MODE NSI {list_name}: {json.dumps(build_nsi_local_object_fair_from_payload(rec_data, market), ensure_ascii=False)}")
                continue

            record_json, response = create_record(session, logger, collection, rec_data)
            if not record_json:
                status_code = getattr(response, "status_code", "n/a")
                logger.error(f"РћС€РёР±РєР° РїСЂРё СЃРѕР·РґР°РЅРёРё Р·Р°РїРёСЃРё РІ {collection}: HTTP {status_code}")
                fail_logger.info(f"{list_name}:{row_num}")
                stats["failedRows"] += 1
                continue

            log_success(success_logger, {"_id": record_json["_id"], "guid": record_json["guid"], "parentEntries": collection})
            logger.info(f"РЎРѕР·РґР°РЅР° Р·Р°РїРёСЃСЊ РІ {collection}: {record_json['_id']}")

            if list_name == SHEET_PERMITS:
                permit_ok = handle_permit_file_upload(
                    row,
                    rec_data,
                    collection,
                    record_json,
                    session,
                    logger,
                    fail_logger,
                    row_num,
                    files_dir,
                )
                if not permit_ok:
                    stats["failedRows"] += 1
                    continue
                update_fair_dates_from_permit(row, session, logger, success_logger, fail_logger, row_num)
                state.mark_success(
                    workbook_path=str(workbook_path),
                    job_name=list_name,
                    row_idx=row_num,
                    collection=str(collection),
                    main_id=str(record_json.get("_id")),
                    guid=str(record_json.get("guid")),
                    had_errors=False,
                    error_count=0,
                )
                stats["createdRows"] += 1
                continue

            if list_name != SHEET_MESTO:
                state.mark_success(
                    workbook_path=str(workbook_path),
                    job_name=list_name,
                    row_idx=row_num,
                    collection=str(collection),
                    main_id=str(record_json.get("_id")),
                    guid=str(record_json.get("guid")),
                    had_errors=False,
                    error_count=0,
                )
                stats["createdRows"] += 1
                continue

            markets = rec_data.get("marketInfo") or []
            nsi_payloads = [build_nsi_local_object_fair_from_payload(rec_data, market) for market in markets] or [
                build_nsi_local_object_fair_from_payload(rec_data, None)
            ]

            nsi_failed = False
            for nsi_payload in nsi_payloads:
                nsi_json, nsi_response = create_record(session, logger, NSI_LOCAL_OBJECT_FAIR_COLLECTION, nsi_payload)
                if not nsi_json:
                    status_code = getattr(nsi_response, "status_code", "n/a")
                    logger.error(
                        f"РћС€РёР±РєР° РїСЂРё СЃРѕР·РґР°РЅРёРё Р·Р°РїРёСЃРё РІ {NSI_LOCAL_OBJECT_FAIR_COLLECTION}: HTTP {status_code}"
                    )
                    fail_logger.info(f"{list_name}:{row_num}:nsi")
                    stats["failedRows"] += 1
                    nsi_failed = True
                    break
                log_success(
                    success_logger,
                    {
                        "_id": nsi_json["_id"],
                        "guid": nsi_json["guid"],
                        "parentEntries": NSI_LOCAL_OBJECT_FAIR_COLLECTION,
                    },
                )
                logger.info(f"РЎРѕР·РґР°РЅР° Р·Р°РїРёСЃСЊ РІ {NSI_LOCAL_OBJECT_FAIR_COLLECTION}: {nsi_json['_id']}")
            if nsi_failed:
                continue
            state.mark_success(
                workbook_path=str(workbook_path),
                job_name=list_name,
                row_idx=row_num,
                collection=str(collection),
                main_id=str(record_json.get("_id")),
                guid=str(record_json.get("guid")),
                had_errors=False,
                error_count=0,
            )
            stats["createdRows"] += 1

        except Exception as exc:
            logger.error(f"РћС€РёР±РєР° РїСЂРё РѕР±СЂР°Р±РѕС‚РєРµ СЃС‚СЂРѕРєРё {row_num}: {exc}")
            fail_logger.info(f"{list_name}:{row_num}")
            logger.debug(traceback.format_exc())
            stats["failedRows"] += 1


def main():
    args = parse_args()
    warnings.filterwarnings("ignore", category=InsecureRequestWarning)

    logger = setup_logger()
    success_logger = setup_success_logger()
    fail_logger = setup_fail_logger()
    user_logger = setup_user_logger()

    interactive = not args.no_interactive and not args.no_prompt
    workbook_default = str(Path(args.workbook).expanduser().resolve())
    files_default = str(Path(args.files_dir).expanduser().resolve())
    workbook_path = Path(_prompt_with_default("Excel workbook", workbook_default, interactive)).expanduser().resolve()
    files_dir = str(Path(_prompt_with_default("Files directory", files_default, interactive)).expanduser().resolve())

    if not workbook_path.exists():
        logger.error("Workbook not found: %s", workbook_path)
        return 1

    state_path = Path(args.state_file).expanduser().resolve()
    state = ResumeState(path=state_path, namespace="fairs_migration", enabled=True)
    if args.reset_state:
        state.reset_namespace()
        logger.info("State reset: %s", state_path)

    resume_enabled = bool(args.resume)
    if resume_enabled and state.rows_count() > 0 and interactive:
        run_info = state.get_run_info()
        logger.info(
            "Найдены checkpoints: rows=%s status=%s startedAt=%s",
            state.rows_count(),
            run_info.get("status"),
            run_info.get("startedAt"),
        )
        if not _ask_yes_no("Продолжить с предыдущего checkpoints?", default_yes=True):
            state.clear_rows()
            resume_enabled = False

    state.begin_run(
        workbook=str(workbook_path),
        filesDir=files_dir,
        sheet=args.sheet,
        resume=resume_enabled,
    )

    stopped = False
    fatal_error = False
    stats = {
        "processedRows": 0,
        "createdRows": 0,
        "resumeSkips": 0,
        "failedRows": 0,
    }

    try:
        session = None
        if not TEST and not args.no_auth:
            session = setup_session(logger, no_prompt=(args.no_prompt or args.no_interactive))
            if session is None:
                fatal_error = True
                stats["failedRows"] += 1
                return 1

        excel_path = str(workbook_path)
        logger.info(f"Р§С‚РµРЅРёРµ С„Р°Р№Р»Р°: {excel_path}")
        user_logger.info(
            "START | workbook=%s | files=%s | sheet=%s | resume=%s",
            excel_path,
            files_dir,
            args.sheet,
            resume_enabled,
        )

        for list_name in _selected_lists(args.sheet):
            logger.info(f"РћР±СЂР°Р±РѕС‚РєР° Р»РёСЃС‚Р°: {list_name}")
            excel = read_excel(excel_path, skiprows=3, sheet_name=list_name)
            if excel is None:
                logger.error(f"Р›РёСЃС‚ {list_name} РЅРµ РЅР°Р№РґРµРЅ РІ С„Р°Р№Р»Рµ {excel_path}")
                stats["failedRows"] += 1
                continue

            excel = excel.iloc[1:].reset_index(drop=True)
            excel.columns = [str(column).strip() for column in excel.columns]
            logger.info(f"Р—Р°РіСЂСѓР¶РµРЅРѕ СЃС‚СЂРѕРє: {len(excel)}")

            process_sheet_data(
                list_name,
                excel,
                session,
                logger,
                success_logger,
                fail_logger,
                workbook_path=str(workbook_path),
                files_dir=files_dir,
                state=state,
                resume_enabled=resume_enabled,
                limit=args.limit,
                stats=stats,
            )

        logger.info("РњРёРіСЂР°С†РёСЏ Р·Р°РІРµСЂС€РµРЅР°")

    except KeyboardInterrupt:
        stopped = True
        logger.warning("Остановка по Ctrl+C")
    except Exception as exc:
        fatal_error = True
        logger.error(f"РљСЂРёС‚РёС‡РµСЃРєР°СЏ РѕС€РёР±РєР°: {exc}")
        logger.debug(traceback.format_exc())
    finally:
        if stopped:
            status = "stopped"
            state.finish_run(status="stopped", summary=stats, clear_rows=False)
        elif fatal_error or stats["failedRows"] > 0:
            status = "failed"
            state.finish_run(status="failed", summary=stats, clear_rows=False)
        else:
            status = "completed"
            state.finish_run(status="completed", summary=stats, clear_rows=True)

        user_logger.info("FINISH | status=%s | summary=%s", status, json.dumps(stats, ensure_ascii=False))
        logger.info(
            "Итог: status=%s created=%s failed=%s resumed_skips=%s",
            status,
            stats["createdRows"],
            stats["failedRows"],
            stats["resumeSkips"],
        )

    return 0 if not stopped and not fatal_error and stats["failedRows"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())

