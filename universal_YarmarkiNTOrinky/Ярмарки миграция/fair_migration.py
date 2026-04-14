import argparse
import copy
import json
import os
import re
import traceback
import warnings
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from urllib3.exceptions import InsecureRequestWarning

from _api import (
    api_request,
    delete_file_from_storage,
    delete_from_collection,
    get_runtime_base_url,
    get_runtime_ui_base_url,
    set_runtime_urls,
    setup_session,
    upload_file,
)
from _config import (
    BASE_URL,
    EXCEL_FILE_NAME,
    EXCEL_INPUT_GLOB,
    EXCEL_LISTS,
    FAIR_COLLECTION,
    FAIR_MESTO_COLLECTION,
    FAIR_PERMITS_COLLECTION,
    FILES_DIR,
    JWT_URL,
    RECORDS_TEMPLATES,
    RESUME_BY_DEFAULT,
    SCRIPT_DIR,
    STATE_FILE,
    TEST,
    UI_BASE_URL,
    UNIT,
)
from _excel_input import discover_excel_files
from _logger import setup_fail_logger, setup_logger, setup_success_logger, setup_user_logger
from _profiles import PROFILES
from _state import ResumeState
from _utils import (
    find_file_in_dir,
    generate_guid,
    jsonable,
    parse_key_value_mapping,
    parse_path_list,
    read_excel,
)


NSI_LOCAL_OBJECT_FAIR_COLLECTION = "nsiLocalObjectFair"
DEFAULT_ORG = copy.deepcopy(UNIT)

SHEET_FAIR = "4. Реестр ярмарок"
SHEET_MESTO = "2. Реестр мест"
SHEET_PERMITS = "3. Реестр разрешений"


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
    return norm_key(value) in {"да", "true", "1", "yes", "y"}


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
        "а": "a",
        "б": "b",
        "в": "v",
        "г": "g",
        "д": "d",
        "е": "e",
        "ё": "e",
        "ж": "zh",
        "з": "z",
        "и": "i",
        "й": "y",
        "к": "k",
        "л": "l",
        "м": "m",
        "н": "n",
        "о": "o",
        "п": "p",
        "р": "r",
        "с": "s",
        "т": "t",
        "у": "u",
        "ф": "f",
        "х": "kh",
        "ц": "ts",
        "ч": "ch",
        "ш": "sh",
        "щ": "shch",
        "ъ": "",
        "ы": "y",
        "ь": "",
        "э": "e",
        "ю": "yu",
        "я": "ya",
    }
    text = "".join(translit.get(ch.lower(), ch) for ch in str(text))
    text = re.sub(r"[\s\-.(),/]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text.lower() if to_lower else text


MAP_MARKET_TYPE = {
    "specialized": {"code": "specialized", "name": "Специализированная"},
    "universal": {"code": "universal", "name": "Универсальная"},
}


def map_market_type(value):
    raw = norm_str(value)
    if not raw:
        return None
    key = norm_key(raw)
    if key in {"specialized", norm_key("Специализированная")}:
        return MAP_MARKET_TYPE["specialized"]
    if key in {"universal", norm_key("Универсальная")}:
        return MAP_MARKET_TYPE["universal"]
    return {"code": None, "name": raw}


MAP_MARKET_SPEC = {
    "agricultural": {"code": "agricultural", "name": "Сельскохозяйственная"},
    "fleaMarkets": {"code": "fleaMarkets", "name": "Блошиные рынки"},
    "food": {"code": "food", "name": "Продовольственная"},
    "industrial": {"code": "industrial", "name": "Промышленная"},
    "other": {"code": "other", "name": "ная"},
    "specializedSales": {"code": "specializedSales", "name": "Продажа определенного вида товаров"},
    "vernissage": {"code": "vernissage", "name": "Вернисаж"},
    "winery": {"code": "winery", "name": "Винодельческая продукция"},
}


def map_market_specialization(value):
    raw = norm_str(value)
    if not raw:
        return None
    key = norm_key(raw)
    ru_to_code = [
        ("Сельскохозяйственная", "agricultural"),
        ("Блошиные рынки", "fleaMarkets"),
        ("Продовольственная", "food"),
        ("Промышленная", "industrial"),
        ("ная", "other"),
        ("Продажа определенного вида товаров", "specializedSales"),
        ("Вернисаж", "vernissage"),
        ("Винодельческая продукция", "winery"),
    ]
    for ru_name, code in ru_to_code:
        if key == norm_key(ru_name):
            return MAP_MARKET_SPEC[code]
    if "сельскох" in key:
        return MAP_MARKET_SPEC["agricultural"]
    if "блош" in key:
        return MAP_MARKET_SPEC["fleaMarkets"]
    if "продов" in key:
        return MAP_MARKET_SPEC["food"]
    if "промыш" in key:
        return MAP_MARKET_SPEC["industrial"]
    if "вернисаж" in key:
        return MAP_MARKET_SPEC["vernissage"]
    if "винод" in key:
        return MAP_MARKET_SPEC["winery"]
    if "иная" in key:
        return MAP_MARKET_SPEC["other"]
    return {"code": None, "name": raw}


MAP_STATUS_FAIR_PLACE = {
    "approved": {"code": "approved", "name": "Утверждено"},
    "draft": {"code": "draft", "name": "Черновик"},
    "liquid": {"code": "liquid", "name": "Ликвидировано"},
}


def map_status_fair_place(value):
    raw = norm_str(value)
    if not raw:
        return None
    key = norm_key(raw)
    if key == norm_key("Утверждено"):
        return MAP_STATUS_FAIR_PLACE["approved"]
    if key == norm_key("Черновик"):
        return MAP_STATUS_FAIR_PLACE["draft"]
    if key == norm_key("Ликвидировано"):
        return MAP_STATUS_FAIR_PLACE["liquid"]
    return {"code": None, "name": raw}


MAP_STATUS_FAIR_PLACE_SPECIFIC = {
    "approvedFree": {"code": "approvedFree", "name": "Свободно", "parentCode": "approved"},
    "approvedUsed": {"code": "approvedUsed", "name": "спользуется", "parentCode": "approved"},
    "draftApplicant": {"code": "draftApplicant", "name": "Предложено заявителем", "parentCode": "draft"},
    "draftMun": {"code": "draftMun", "name": "Предложено муниципалитетом", "parentCode": "draft"},
}


def map_status_fair_place_specific(value, status_fair_place):
    raw = norm_str(value)
    if not raw:
        return None
    key = norm_key(raw)
    if key == norm_key("Свободно"):
        return MAP_STATUS_FAIR_PLACE_SPECIFIC["approvedFree"]
    if key == norm_key("спользуется"):
        return MAP_STATUS_FAIR_PLACE_SPECIFIC["approvedUsed"]
    if key == norm_key("Предложено заявителем"):
        return MAP_STATUS_FAIR_PLACE_SPECIFIC["draftApplicant"]
    if key == norm_key("Предложено муниципалитетом"):
        return MAP_STATUS_FAIR_PLACE_SPECIFIC["draftMun"]
    parent_code = status_fair_place.get("code") if status_fair_place else None
    return {"code": None, "name": raw, "parentCode": parent_code}


MAP_MARKET_STATUS = {
    "активна": {"code": "active", "name": "Активна"},
    "действует": {"code": "active", "name": "Действует"},
    "отменена": {"code": "cancelled", "name": "Отменена"},
    "черновик": {"code": "draft", "name": "Черновик"},
    "завершена": {"code": "finished", "name": "Завершена"},
    "запланирована": {"code": "planned", "name": "Запланирована"},
}


def map_market_status(value):
    raw = norm_str(value)
    if not raw:
        return None
    return MAP_MARKET_STATUS.get(norm_key(raw), {"code": None, "name": raw})


MAP_MARKET_FREQUENCY_PARENT = {
    "регулярная": {"code": "regular", "name": "Регулярная"},
    "разовая": {"code": "single", "name": "Разовая"},
}


def map_market_frequency_parent(value):
    raw = norm_str(value)
    if not raw:
        return None
    return MAP_MARKET_FREQUENCY_PARENT.get(norm_key(raw), {"code": None, "name": raw})


MAP_MARKET_VARIATION_REGULAR = {
    "постоянно действующая": {"code": "regularPermanent", "name": "Постоянно действующая", "parentCode": "regular"},
    "сезонная": {"code": "regularSeasonal", "name": "Сезонная", "parentCode": "regular"},
    "выходного дня": {"code": "regularWeekend", "name": "Выходного дня", "parentCode": "regular"},
    "еженедельная": {"code": "regularWeekly", "name": "Еженедельная", "parentCode": "regular"},
}

MAP_MARKET_VARIATION_SINGLE = {
    "праздничная": {"code": "singleFestive", "name": "Праздничная", "parentCode": "single"},
    "сезонная": {"code": "singleSeasonal", "name": "Сезонная", "parentCode": "single"},
    "тематическая": {"code": "singleThematic", "name": "Тематическая", "parentCode": "single"},
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
    "без выходных": {"code": "AllWeek", "name": "Без выходных"},
    "понедельник": {"code": "Monday", "name": "Понедельник"},
    "вторник": {"code": "Tuesday", "name": "Вторник"},
    "среда": {"code": "Wednesday", "name": "Среда"},
    "четверг": {"code": "Thursday", "name": "Четверг"},
    "пятница": {"code": "Friday", "name": "Пятница"},
    "суббота": {"code": "Saturday", "name": "Суббота"},
    "воскресенье": {"code": "Sunday", "name": "Воскресенье"},
}

MAP_MARKET_PURPOSE = {
    "продвижение ценностей национальной культуры среди отечественных и иностранных посетителей": {
        "code": "culturePromotion",
        "name": "Продвижение ценностей национальной культуры среди отечественных и иностранных посетителей",
    },
    "иная": {"code": "other", "name": "ная"},
    "расширение каналов сбыта продукции отечественных, региональных, локальных товаропроизводителей, в том числе и на международном уровне": {
        "code": "marketExpansion",
        "name": "Расширение каналов сбыта продукции отечественных, региональных, локальных товаропроизводителей, в том числе и на международном уровне",
    },
    "создание комфортной потребительской среды": {"code": "consumerEnvironment", "name": "Создание комфортной потребительской среды"},
    "поддержка отечественных товаропроизводителей в реализации собственной продукции": {
        "code": "domesticSalesSupport",
        "name": "Поддержка отечественных товаропроизводителей в реализации собственной продукции",
    },
    "обеспечение знакомства с национальной или местной или региональной культурой, кухней, традициями": {
        "code": "culturalAwareness",
        "name": "Обеспечение знакомства с национальной или местной или региональной культурой, кухней, традициями",
    },
    "формирование эффективной конкурентной среды": {"code": "competitiveEnvironment", "name": "Формирование эффективной конкурентной среды"},
}

MAP_PERMISSION_STATUS = {
    "действует": {"code": "Working", "name": "Действует"},
    "приостановлено": {"code": "Stop", "name": "Приостановлено"},
    "аннулировано": {"code": "Annul", "name": "Аннулировано"},
    "не действует": {"code": "doesNotWork", "name": "Не действует"},
    "черновик": {"code": "Draft", "name": "Черновик"},
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
        day = norm_str(row.get(f"{index}. День недели, который отличается от основного"))
        start = to_time_hhmm(row.get(f"{index}. Время начала работы ярмарки"))
        end = to_time_hhmm(row.get(f"{index}. Время окончания работы ярмарки"))
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
        day = norm_str(row.get(f"{index}. Выходной день"))
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
    if key in {"праздничная", "тематическая"}:
        return {"code": "single", "name": "Разовая"}
    if key in {"постоянно действующая", "выходного дня", "еженедельная"}:
        return {"code": "regular", "name": "Регулярная"}
    return None


def map_market_frequency(value, variation_obj):
    raw = norm_str(value)
    if not raw:
        return None
    key = norm_key(raw)
    if key == "сезонная":
        parent_code = (variation_obj or {}).get("code")
        if parent_code == "regular":
            return {"code": "regularSeasonal", "name": "Сезонная", "parentCode": "regular"}
        if parent_code == "single":
            return {"code": "singleSeasonal", "name": "Сезонная", "parentCode": "single"}
        return {"code": None, "name": raw, "parentCode": None}
    child_map = {
        "постоянно действующая": {"code": "regularPermanent", "name": "Постоянно действующая", "parentCode": "regular"},
        "выходного дня": {"code": "regularWeekend", "name": "Выходного дня", "parentCode": "regular"},
        "еженедельная": {"code": "regularWeekly", "name": "Еженедельная", "parentCode": "regular"},
        "праздничная": {"code": "singleFestive", "name": "Праздничная", "parentCode": "single"},
        "тематическая": {"code": "singleThematic", "name": "Тематическая", "parentCode": "single"},
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

    ogrn = norm_str(row.get("ОГРН уполномоченного органа"))
    if not ogrn:
        logger.warning("ОГРН уполномоченного органа отсутствует, используется UNIT из конфигурации")
        return default_org

    body = {"search": {"search": [{"field": "ogrn", "operator": "eq", "value": ogrn}]}, "size": 2}
    try:
        response = api_request(
            session,
            logger,
            "post",
            f"{get_runtime_base_url()}/api/v1/search/organizations",
            json=body,
            max_retries=1,
        )
        if response.status_code != 200:
            logger.warning(f"Поиск организации по ОГРН {ogrn} вернул HTTP {response.status_code}, используется UNIT из конфигурации")
            return default_org
        payload = response.json()
        items = payload.get("content") or []
        if len(items) == 1:
            return items[0]
        if len(items) > 1:
            logger.warning(f"По ОГРН {ogrn} найдено несколько организаций, используется UNIT из конфигурации")
            return default_org
        logger.warning(f"По ОГРН {ogrn} организация не найдена, используется UNIT из конфигурации")
    except Exception as exc:
        logger.warning(f"Ошибка поиска организации по ОГРН {ogrn}: {exc}. спользуется UNIT из конфигурации")
    return default_org


def build_organizer_blocks(row, index):
    choose_authority = parse_bool(row.get(f"{index}. Выбрать орган власти в качестве организатора"))
    status_legal = norm_str(row.get(f"{index}. Правовой статус организатора"))

    ul_info = {
        "fullName": norm_str(row.get(f"{index}. Полное наименование")),
        "ogrn": norm_str(row.get(f"{index}. ОГРН")),
        "inn": norm_str(row.get(f"{index}. НН")),
        "phoneNumber": norm_str(row.get(f"{index}. Номер телефона")),
        "email": norm_str(row.get(f"{index}. Электронная почта")),
    }
    ip_info = {
        "nameIP": norm_str(row.get(f"{index}. Наименование П")),
        "ogrnIP": norm_str(row.get(f"{index}. ОГРНП")),
        "inn": norm_str(row.get(f"{index}. НН.1")),
        "phoneNumber": norm_str(row.get(f"{index}. Номер телефона.1")),
        "email": norm_str(row.get(f"{index}. Электронная почта.1")),
    }
    authority_info = {
        "fullName": norm_str(row.get(f"{index}. Организатор ярмарки")),
        "phoneNumber": norm_str(row.get(f"{index}. Номер телефона.2")),
        "email": norm_str(row.get(f"{index}. Электронная почта.2")),
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
    if ul_has and (not ip_has or status_key == norm_key("Юридическое лицо")):
        return {
            "chooseAuthorityAsOrganizator": False,
            "statusLegal": status_legal or "Юридическое лицо",
            "blockOrganizerULInfo": ul_info,
            "blockOrganizerIPInfo": None,
            "blockOrganizerAuthority": None,
        }

    if ip_has and (not ul_has or "предприним" in status_key):
        return {
            "chooseAuthorityAsOrganizator": False,
            "statusLegal": status_legal or "ндивидуальный предприниматель",
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
            "marketNumber": norm_str(row.get(f"{index}. Номер ярмарки")),
            "permissionNumber": norm_str(row.get(f"{index}. Номер разрешения")),
            "marketName": norm_str(row.get(f"{index}. Наименование ярмарки")),
            "startDate": norm_str(row.get(f"{index}. Дата начала ярмарки")),
            "endDate": norm_str(row.get(f"{index}. Дата окончания ярмарки")),
            "openingTime": norm_str(row.get(f"{index}. Время начала ярмарки")),
            "closingTime": norm_str(row.get(f"{index}. Время окончания ярмарки")),
            "status": norm_str(row.get(f"{index}. Статус ярмарки")),
            "frequency": norm_str(row.get(f"{index}. Периодичность проведения ярмарки")),
            "variation": norm_str(row.get(f"{index}. Вид ярмарки")),
            "placeCount": norm_str(row.get(f"{index}. Количество торговых мест")),
            "placeCountFree": norm_str(row.get(f"{index}. Количество торговых мест на безвозмездной основе")),
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
        "subject": norm_str(row.get("Субъект РФ")),
        "disctrict": norm_str(row.get("Муниципальный район/округ, городской округ или внутригородская территория")),
    }
    unit = {
        "id": unit_resolved.get("_id") or unit_resolved.get("id") or DEFAULT_ORG.get("id"),
        "_id": unit_resolved.get("_id") or DEFAULT_ORG.get("_id") or DEFAULT_ORG.get("id"),
        "guid": unit_resolved.get("guid") or DEFAULT_ORG.get("guid"),
        "name": norm_str(unit_resolved.get("name")) or DEFAULT_ORG.get("name"),
        "shortName": unit_resolved.get("shortName") or DEFAULT_ORG.get("shortName"),
        "ogrn": norm_str(row.get("ОГРН уполномоченного органа")) or norm_str(unit_resolved.get("ogrn")) or DEFAULT_ORG.get("ogrn"),
        "inn": norm_str(row.get("НН уполномоченного органа")) or norm_str(unit_resolved.get("inn")),
        "kpp": unit_resolved.get("kpp") or DEFAULT_ORG.get("kpp"),
        "email": unit_resolved.get("email") or DEFAULT_ORG.get("email"),
        "site": unit_resolved.get("site") or DEFAULT_ORG.get("site"),
        "phone": unit_resolved.get("phone") or DEFAULT_ORG.get("phone"),
        "regions": unit_resolved.get("regions") or DEFAULT_ORG.get("regions"),
    }

    market_variation = map_market_frequency_parent(row.get("Периодичность проведения ярмарки"))
    if not market_variation:
        market_variation = infer_variation_from_frequency_name(row.get("Вид ярмарки"))
    market_frequency = map_market_frequency(row.get("Вид ярмарки"), market_variation)

    no_days_off = parse_bool(row.get("Без выходных"))
    block_organizer_ul = {
        "fullName": norm_str(row.get("Полное наименование")),
        "shortName": norm_str(row.get("Сокращённое наименование")),
        "orgStateForm": map_org_state_form(row.get("Организационно-правовая форма")),
        "ogrn": norm_str(row.get("ОГРН")),
        "inn": norm_str(row.get("НН")),
        "kpp": norm_str(row.get("КПП")),
        "fioUL": norm_str(row.get("ФО руководителя")),
        "phoneNumber": norm_str(row.get("Номер телефона")),
        "email": norm_str(row.get("Электронная почта")),
        "addressUL": parse_address_to_obj(row.get("Юридический адрес")),
        "addressFact": parse_address_to_obj(row.get("Фактический адрес")),
    }
    block_organizer_ip = {
        "nameIP": norm_str(row.get("Наименование П")),
        "fioIP": None,
        "inn": norm_str(row.get("НН.1")),
        "ogrnIP": norm_str(row.get("ОГРНП")),
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
        "fullName": norm_str(row.get("Организатор ярмарки")),
        "phoneNumber": norm_str(row.get("Номер телефона.1")),
        "email": norm_str(row.get("Электронная почта.1")),
    }

    market_info = {
        "marketNumber": norm_str(row.get("Номер ярмарки")),
        "permissionNumber": norm_str(row.get("Номер разрешения")),
        "marketStatus": map_market_status(row.get("Статус ярмарки")),
        "marketName": norm_str(row.get("Наименование ярмарки")),
        "marketType": map_market_type(row.get("Тип ярмарки")),
        "marketSpecialization": map_market_specialization(row.get("Специализация ярмарки")),
        "marketSpecializationOther": norm_str(row.get("ная специализация ярмарки")),
        "marketPurpose": map_market_purpose(row.get("Цель ярмарки")),
        "marketPurposeOther": norm_str(row.get("ная цель ярмарки")),
        "marketVariation": market_variation,
        "marketFrequency": market_frequency,
        "marketArea": norm_str(row.get("Площадь ярмарки, кв. м")),
        "blockMarketDates": {
            "startDate": parse_date_to_iso(row.get("Сроки проведения ярмарки: дата начала")),
            "endDate": parse_date_to_iso(row.get("Сроки проведения ярмарки: дата окончания")),
        },
        "blockMarketOperatingTime": {
            "marketOpeningTime": to_time_hhmm(row.get("Режим работы ярмарки: время начала")),
            "marketClosingTime": to_time_hhmm(row.get("Режим работы ярмарки: время окончания")),
        },
        "dayOnBlock": build_day_on_block_from_row(row),
        "dayOffDayOff": copy.deepcopy(MAP_DAY_OF_WEEK["без выходных"]) if no_days_off is True else None,
        "BlockDayOff": [] if no_days_off is True else build_block_day_off_from_row(row),
        "sanitaryDayOfMonth": norm_str(row.get("Санитарный день месяца")),
        "placeCount": norm_str(row.get("Количество торговых мест")),
        "placeCountFree": norm_str(row.get("Количество торговых мест на безвозмездной основе")),
        "placeNumber": norm_str(row.get("Номер места для размещения ярмарок")),
        "blockCadNumber": [],
        "cadsObjects": [],
    }

    cad_land = norm_str(row.get("Кадастровый номер земельного участка"))
    addr_land = row.get("Адрес земельного участка")
    if cad_land or norm_str(addr_land):
        market_info["blockCadNumber"].append({"cadObjNum": cad_land, "landAddress": parse_address_to_obj(addr_land)})

    cad_obj = norm_str(row.get("Кадастровый номер объекта недвижимости"))
    addr_obj = row.get("Адрес объекта недвижимости")
    if cad_obj or norm_str(addr_obj):
        market_info["cadsObjects"].append({"cadNumber": cad_obj, "objectAddress": parse_address_to_obj(addr_obj)})

    payload = {
        "guid": generate_guid(),
        "parentEntries": FAIR_COLLECTION,
        "generalInformation": general_information,
        "unit": unit,
        "marketInfo": market_info,
        "chooseAuthorityAsOrganizator": parse_bool(row.get("Выбрать орган власти в качестве организатора")),
        "statusLegal": norm_str(row.get("Правовой статус организатора")),
        "blockOrganizerULInfo": block_organizer_ul,
        "blockOrganizerIPInfo": block_organizer_ip,
    }
    if has_any_non_null(block_organizer_authority):
        payload["blockOrganizerAuthority"] = block_organizer_authority
    return payload


def process_mesto_sheet(row, logger, session=None):
    unit_resolved = resolve_unit_for_mesto(row, session, logger)
    subject = norm_str(row.get("Субъект РФ"))
    district = norm_str(row.get("Муниципальный район/округ, городской округ или внутригородская территория"))

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
        "ogrn": norm_str(row.get("ОГРН уполномоченного органа")) or norm_str(unit_resolved.get("ogrn")),
        "inn": norm_str(row.get("НН уполномоченного органа")) or norm_str(unit_resolved.get("inn")),
        "kpp": unit_resolved.get("kpp") or DEFAULT_ORG.get("kpp"),
        "email": unit_resolved.get("email") or DEFAULT_ORG.get("email"),
        "site": unit_resolved.get("site") or DEFAULT_ORG.get("site"),
        "phone": unit_resolved.get("phone") or DEFAULT_ORG.get("phone"),
        "regions": unit_resolved.get("regions") or DEFAULT_ORG.get("regions"),
    }

    status_fair_place = map_status_fair_place(row.get("Статус проектного места"))
    place_market_info = {
        "marketSchemeNumber": norm_str(row.get("Номер места для размещения ярмарок")),
        "coordinates": norm_str(row.get("Геокоординаты точки размещения организуемой ярмарки")),
        "landArea": pick_first_non_empty(
            norm_str(row.get("Площадь места для размещения ярмарки, кв. м")),
            norm_str(row.get("Площадь места для размещения ярмарки, кв.м")),
        ),
        "marketType": map_market_type(row.get("Тип ярмарки")),
        "marketSpecialization": map_market_specialization(row.get("Специализация ярмарки")),
        "marketSpecializationOther": norm_str(row.get("ная специализация ярмарки")),
        "statusFairPlace": status_fair_place,
        "statusFairPlaceSpecific": map_status_fair_place_specific(row.get("Спецификация статуса проектного места"), status_fair_place),
        "statusPlaceFair": norm_str(row.get("Статус в Реестре мест на размещения ярмарок")),
        "blockCadNumber": [],
        "cadsObjects": [],
    }

    cad_land = norm_str(row.get("Кадастровый номер земельного участка"))
    addr_land = norm_str(row.get("Адрес земельного участка"))
    if not is_empty_val(cad_land) or not is_empty_val(addr_land):
        place_market_info["blockCadNumber"].append(
            {"cadNumber": cad_land, "landAddress": parse_address_to_obj(addr_land)}
        )

    cad_obj = norm_str(row.get("Кадастровый номер объекта недвижимости"))
    addr_obj = norm_str(row.get("Адрес объекта недвижимости"))
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
    is_universal = type_code == "universal" or norm_key(type_name) == norm_key("Универсальная")

    specialization_out = "-"
    if not is_universal:
        spec_obj = place_market_info.get("marketSpecialization") or {}
        spec_name = norm_str(spec_obj.get("name") or spec_obj.get("code"))
        if spec_name:
            is_other = spec_obj.get("code") == "other" or norm_key(spec_name) in {norm_key("ная"), "other"}
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
        "Layer": "Ярмарки",
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
            "subject": norm_str(row.get("Субъект РФ")),
            "disctrict": norm_str(row.get("Муниципальный район/округ, городской округ или внутригородская территория")),
        },
        "permission": {
            "permissionNumber": norm_str(row.get("Номер разрешения")),
            "permissionStatus": map_permission_status(row.get("Статус разрешения")),
            "fairNumber": norm_str(row.get("Номер ярмарки")),
            "permissionDate": parse_date_to_iso(row.get("Дата выдачи разрешения")),
            "permissionStartDate": parse_date_to_iso(row.get("Дата начала действия разрешения")),
            "permissionEndDate": parse_date_to_iso(row.get("Дата завершения действия разрешения")),
            "permissionFile": None,
        },
    }


def create_record(session, logger, collection, payload):
    response = api_request(
        session,
        logger,
        "post",
        f"{get_runtime_base_url()}/api/v1/create/{collection}",
        json=jsonable(payload),
    )
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
            response = api_request(
                session,
                logger,
                "post",
                f"{get_runtime_base_url()}/api/v1/search/{NSI_LOCAL_OBJECT_FAIR_COLLECTION}",
                json=body,
                max_retries=1,
            )
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
    fair_number = norm_str(row.get("Номер ярмарки"))
    subject = norm_str(row.get("Субъект РФ"))
    district = norm_str(row.get("Муниципальный район/округ, городской округ или внутригородская территория"))

    result = find_fair_by_number(session, logger, fair_number, subject=subject, district=district)
    if not result.get("ok"):
        if result.get("many"):
            logger.warning(
                f"Для ярмарки {fair_number} найдено несколько записей nsiLocalObjectFair "
                f"с Subject={subject!r} и Disctrict={district!r}"
            )
        else:
            logger.warning(
                f"Ярмарка {fair_number} не найдена в nsiLocalObjectFair "
                f"с Subject={subject!r} и Disctrict={district!r}"
            )
        return

    found_fair = result["entry"]
    update_doc = {"_id": found_fair["_id"], "guid": found_fair["guid"]}
    if found_fair.get("auid"):
        update_doc["auid"] = found_fair["auid"]
    update_doc["PermissionStartDate"] = as_date_or_dash_dmy(row.get("Дата выдачи разрешения"))
    update_doc["PermissionEndDate"] = as_date_or_dash_dmy(row.get("Дата завершения действия разрешения"))

    response = api_request(
        session,
        logger,
        "put",
        f"{get_runtime_base_url()}/api/v1/update/{NSI_LOCAL_OBJECT_FAIR_COLLECTION}?mainId={found_fair['_id']}&guid={found_fair['guid']}",
        json=jsonable(update_doc),
    )
    if response.ok:
        log_success(success_logger, {"_id": found_fair["_id"], "guid": found_fair["guid"], "parentEntries": NSI_LOCAL_OBJECT_FAIR_COLLECTION})
        return
    logger.error(f"Ошибка обновления {NSI_LOCAL_OBJECT_FAIR_COLLECTION} для номера ярмарки {fair_number}: HTTP {response.status_code}")
    fail_logger.info(f"{SHEET_PERMITS}:{row_num}:fair_update")


def handle_permit_file_upload(row, rec_data, collection, record_json, session, logger, fail_logger, row_num, files_dir):
    file_hint = norm_str(row.get("Разрешение на право организации ярмарки"))
    file_path = find_file_in_dir(files_dir, file_hint) if file_hint else None
    file_field = "permission.permissionFile"

    if not file_path:
        return True

    logger.info(f"Найден файл для разрешения: {os.path.basename(file_path)}")
    file_object = upload_file(session, logger, file_path, collection, record_json["_id"], entity_field_path=file_field)
    if not file_object:
        logger.error("Ошибка при загрузке файла разрешения")
        delete_from_collection(session, logger, record_json)
        fail_logger.info(f"{SHEET_PERMITS}:{row_num}")
        return False

    update_payload = {
        "_id": record_json["_id"],
        "guid": rec_data["guid"],
        "permission": {**rec_data["permission"], "permissionFile": file_object},
    }
    update_url = f"{get_runtime_base_url()}/api/v1/update/{collection}?mainId={record_json['_id']}&guid={rec_data['guid']}"
    update_response = api_request(session, logger, "put", update_url, json=jsonable(update_payload))
    if update_response.ok:
        return True

    logger.error("Ошибка при обновлении записи с файлом")
    delete_file_from_storage(session, logger, file_object["_id"])
    delete_from_collection(session, logger, record_json)
    fail_logger.info(f"{SHEET_PERMITS}:{row_num}")
    return False


@dataclass
class WorkbookRunSpec:
    workbook_path: str
    files_dir: str


def _console_block(title: str, lines: Optional[List[str]] = None, width: int = 92) -> str:
    safe_title = str(title or "").strip() or "Блок"
    content = ["", "=" * width, safe_title, "-" * width]
    for line in (lines or []):
        content.append(str(line))
    content.append("=" * width)
    return "\n".join(content)


def _format_iso_for_console(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return "-"
    try:
        normalized = text[:-1] + "+00:00" if text.endswith("Z") else text
        dt = datetime.fromisoformat(normalized)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone().strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return text


def _prompt_with_default(label: str, default_value: str, interactive: bool) -> str:
    if not interactive:
        return default_value
    entered = input(f"{label} [{default_value}]: ").strip().strip('"').strip("'")
    return entered or default_value


def _ask_yes_no(prompt: str, default_yes: bool = True) -> bool:
    suffix = " [Y/n]: " if default_yes else " [y/N]: "
    answer = input(prompt + suffix).strip().lower()
    if not answer:
        return default_yes
    return answer in {"y", "yes", "д", "да"}


def _ask_error_action() -> str:
    print("Ошибка строки. Действие: [r]etry / [s]kip / [a]bort")
    while True:
        answer = input("Выбор: ").strip().lower()
        if answer in {"r", "retry"}:
            return "retry"
        if answer in {"s", "skip"}:
            return "skip"
        if answer in {"a", "abort", "stop"}:
            return "abort"
        if not answer:
            return "abort"
        print("Введите r, s или a")


def _operator_action_on_row_error(*, operator_mode: bool, interactive: bool, logger, context: str) -> str:
    if not (operator_mode and interactive):
        return "skip"
    logger.warning("Ошибка обработки строки: %s", context)
    return _ask_error_action()


def _selected_lists(sheet_mode: str) -> List[str]:
    mapping = {
        "fair": SHEET_FAIR,
        "mesto": SHEET_MESTO,
        "permits": SHEET_PERMITS,
    }
    if sheet_mode == "all":
        return list(EXCEL_LISTS)
    selected = mapping.get(sheet_mode)
    return [selected] if selected else list(EXCEL_LISTS)


def _numeric_hints(text: str) -> List[str]:
    normalized = (" " + str(text or "").strip().lower() + " ")
    hints = set()
    for token in normalized.split():
        if token.isdigit():
            hints.add(token)
    word_map = {
        " one ": "1",
        " first ": "1",
        " two ": "2",
        " second ": "2",
        " three ": "3",
        " third ": "3",
        " один ": "1",
        " первый ": "1",
        " первая ": "1",
        " два ": "2",
        " второй ": "2",
        " вторая ": "2",
        " три ": "3",
        " третий ": "3",
        " третья ": "3",
    }
    for needle, number in word_map.items():
        if needle in normalized:
            hints.add(number)
    return sorted(hints)


def _choose_single_workbook(candidates: List[str], interactive: bool) -> Optional[str]:
    if not candidates:
        return None
    if len(candidates) == 1 or not interactive:
        return candidates[0]
    print("\nДоступные книги для миграции:")
    for idx, wb in enumerate(candidates, start=1):
        print(f"  {idx}) {wb}")
    raw = input("Выберите номер книги [1]: ").strip()
    if not raw:
        return candidates[0]
    try:
        selected = int(raw)
    except Exception:
        return candidates[0]
    if 1 <= selected <= len(candidates):
        return candidates[selected - 1]
    return candidates[0]


def _choose_mass_workbooks(candidates: List[str], interactive: bool) -> List[str]:
    if not candidates:
        return []
    if not interactive:
        return list(candidates)
    print("\nДоступные книги для миграции:")
    for idx, wb in enumerate(candidates, start=1):
        print(f"  {idx}) {wb}")
    raw = input("Введите номера книг через запятую или Enter для всех: ").strip()
    if not raw:
        return list(candidates)
    selected: List[str] = []
    for part in raw.split(","):
        token = part.strip()
        if not token:
            continue
        try:
            idx = int(token)
        except Exception:
            continue
        if 1 <= idx <= len(candidates):
            selected.append(candidates[idx - 1])
    return selected or list(candidates)


def _resolve_mode(mode: str, candidates: List[str], interactive: bool) -> str:
    if mode in {"single", "mass"}:
        return mode
    if len(candidates) <= 1 or not interactive:
        return "single"
    print("\nНайдено несколько книг XLSM.")
    print("  1) single - выбрать одну книгу")
    print("  2) mass   - обработать несколько/все")
    answer = input("Выберите режим [1]: ").strip().lower()
    if answer in {"2", "mass", "m", "м"}:
        return "mass"
    return "single"


def _infer_files_dir_for_workbook(
    *,
    workbook_path: str,
    files_root: str,
    files_map: Dict[str, str],
    interactive: bool,
    prompt_always: bool,
) -> str:
    workbook_abs = os.path.abspath(str(workbook_path))
    workbook_name = os.path.basename(workbook_abs)
    workbook_stem = os.path.splitext(workbook_name)[0]

    for key in (workbook_abs, workbook_name, workbook_stem):
        if key in files_map:
            target = str(files_map[key]).strip()
            if os.path.isabs(target):
                return os.path.abspath(target)
            return os.path.abspath(os.path.join(files_root, target))

    subdirs: List[str] = []
    if os.path.isdir(files_root):
        for name in sorted(os.listdir(files_root), key=lambda x: x.lower()):
            candidate = os.path.abspath(os.path.join(files_root, name))
            if os.path.isdir(candidate):
                subdirs.append(candidate)

    auto_candidate = os.path.abspath(files_root)
    same_name_dir = os.path.abspath(os.path.join(files_root, workbook_stem))
    if os.path.isdir(same_name_dir):
        auto_candidate = same_name_dir
    elif subdirs:
        stem_norm = workbook_stem.strip().lower()
        matched: List[str] = []
        exact_found = False
        for subdir in subdirs:
            name = os.path.basename(subdir).strip().lower()
            if not name:
                continue
            if name == stem_norm:
                auto_candidate = subdir
                exact_found = True
                break
            if (" " + name + " ") in (" " + stem_norm + " ") or stem_norm.endswith(" " + name):
                matched.append(subdir)

        if not exact_found:
            if len(matched) == 1:
                auto_candidate = matched[0]
            else:
                workbook_hints = set(_numeric_hints(stem_norm))
                hint_matches: List[str] = []
                if workbook_hints:
                    for subdir in subdirs:
                        sub_hints = set(_numeric_hints(os.path.basename(subdir)))
                        if workbook_hints.intersection(sub_hints):
                            hint_matches.append(subdir)
                if len(hint_matches) == 1:
                    auto_candidate = hint_matches[0]

            if auto_candidate == os.path.abspath(files_root):
                if len(subdirs) == 2:
                    default_like = [
                        folder
                        for folder in subdirs
                        if os.path.basename(folder).strip().lower() in {"one", "1", "default", "main"}
                    ]
                    if len(default_like) == 1:
                        auto_candidate = default_like[0]
                elif len(subdirs) == 1:
                    auto_candidate = subdirs[0]

    if not interactive or not subdirs:
        return auto_candidate

    options = [os.path.abspath(files_root)] + subdirs
    default_idx = 0
    for idx, path in enumerate(options):
        if os.path.abspath(path) == os.path.abspath(auto_candidate):
            default_idx = idx
            break

    print(f"\nКнига: {workbook_name}")
    print("Выберите папку с файлами:")
    print(f"  0) {os.path.abspath(files_root)} (корень)")
    for idx, folder in enumerate(subdirs, start=1):
        print(f"  {idx}) {os.path.basename(folder)}")

    raw = input(f"Номер папки [{default_idx}]: ").strip()
    if not raw:
        return options[default_idx]
    try:
        selected_idx = int(raw)
    except Exception:
        return options[default_idx]
    if 0 <= selected_idx < len(options):
        return options[selected_idx]
    return options[default_idx]


def _resolve_workbook_specs(args: argparse.Namespace, interactive: bool, logger) -> Tuple[List[WorkbookRunSpec], str]:
    explicit = parse_path_list(args.workbooks)
    if not explicit and args.workbook:
        explicit.append(args.workbook)
    explicit_raw = ";".join(explicit)

    candidates = discover_excel_files(SCRIPT_DIR, explicit_files=explicit_raw, pattern=EXCEL_INPUT_GLOB)
    if not candidates:
        candidates = discover_excel_files(SCRIPT_DIR, explicit_files="", pattern=EXCEL_INPUT_GLOB)
    if not candidates:
        raise FileNotFoundError("Книги XLSM не найдены")

    resolved_mode = _resolve_mode(args.mode, candidates, interactive)
    if resolved_mode == "single":
        chosen = _choose_single_workbook(candidates, interactive)
        selected = [chosen] if chosen else []
    else:
        selected = _choose_mass_workbooks(candidates, interactive)

    if not selected:
        raise RuntimeError("Не выбрано ни одной книги для миграции")

    files_root = str(Path(args.files_dir).expanduser().resolve())
    files_map = parse_key_value_mapping(args.files_map)
    prompt_files_for_each = bool(
        interactive
        and not files_map
        and (bool(args.ask_files_always) or args.mode in {"mass", "auto"})
    )

    specs: List[WorkbookRunSpec] = []
    for workbook_path in selected:
        workbook_abs = str(Path(workbook_path).expanduser().resolve())
        files_dir = _infer_files_dir_for_workbook(
            workbook_path=workbook_abs,
            files_root=files_root,
            files_map=files_map,
            interactive=interactive,
            prompt_always=prompt_files_for_each,
        )
        specs.append(WorkbookRunSpec(workbook_path=workbook_abs, files_dir=files_dir))

    logger.info("Выбрано книг: %s", len(specs))
    for idx, spec in enumerate(specs, start=1):
        logger.info("  %s) workbook=%s | files=%s", idx, spec.workbook_path, spec.files_dir)
    return specs, resolved_mode


def _resolve_runtime(args: argparse.Namespace) -> Tuple[str, str, str, str]:
    if args.profile == "custom":
        base_url = (args.base_url or BASE_URL).strip()
        jwt_url = (args.jwt_url or JWT_URL or (base_url.rstrip("/") + "/jwt/")).strip()
        ui_base_url = (args.ui_base_url or UI_BASE_URL or base_url).strip()
    else:
        profile = PROFILES[args.profile]
        base_url = (args.base_url or profile.base_url).strip()
        jwt_url = (args.jwt_url or profile.jwt_url).strip()
        ui_base_url = (args.ui_base_url or profile.ui_base_url).strip()

    if not base_url:
        raise ValueError("base_url is empty")
    if not jwt_url:
        jwt_url = base_url.rstrip("/") + "/jwt/"
    if not ui_base_url:
        ui_base_url = base_url
    return args.profile, base_url.rstrip("/"), jwt_url, ui_base_url.rstrip("/")


def _choose_resume_strategy(*, state: ResumeState, logger, user_logger, interactive: bool) -> str:
    if not state.enabled:
        return "continue"
    rows_count = state.rows_count()
    if rows_count <= 0:
        return "continue"

    run_info = state.get_run_info()
    status = str(run_info.get("status") or "").strip().lower()
    incomplete_statuses = {"running", "stopped", "aborted", "interrupted", "failed", "error", ""}
    if status not in incomplete_statuses:
        return "continue"

    last_checkpoint = run_info.get("lastCheckpoint") if isinstance(run_info.get("lastCheckpoint"), dict) else {}
    lines = [
        "Обнаружены незавершенные checkpoints предыдущего запуска.",
        "",
        f"Статус прошлого запуска : {status or 'неизвестно (старый формат state)'}",
        f"Начало запуска          : {_format_iso_for_console(run_info.get('startedAt'))}",
        f"Конец запуска           : {_format_iso_for_console(run_info.get('finishedAt'))}",
        f"Профиль / стенд         : {run_info.get('profile') or '-'} / {run_info.get('baseUrl') or '-'}",
        f"Строк в checkpoint      : {rows_count}",
        f"Последняя позиция       : {last_checkpoint.get('job') or '-'} / row={last_checkpoint.get('row') or '-'}",
        f"Последняя запись _id    : {last_checkpoint.get('_id') or '-'}",
    ]
    block = _console_block("RESUME: найдено незавершенное состояние", lines)
    logger.warning("%s", block)
    if user_logger:
        user_logger.info(block)

    if not interactive:
        logger.info("Интерактив выключен, автоматически продолжаем по checkpoint.")
        return "continue"

    prompt = "\n[RESUME] Выберите: [P]продолжить / [C]сбросить / [Q]выйти: "
    while True:
        try:
            raw = input(prompt)
        except EOFError:
            return "continue"
        choice = norm_str(raw)
        choice = choice.lower() if choice else ""
        if choice in {"", "p", "продолжить", "continue", "resume"}:
            return "continue"
        if choice in {"c", "сбросить", "reset", "сброс", "start"}:
            return "reset"
        if choice in {"q", "quit", "выйти", "exit", "abort"}:
            return "abort"


def _log_user_run_header(
    *,
    user_logger,
    profile: str,
    base_url: str,
    mode: str,
    interactive: bool,
    operator_mode: bool,
    state_file_path: str,
    success_log_path: str,
    fail_log_path: str,
) -> None:
    if user_logger is None:
        return
    lines = [
        "",
        "#" * 92,
        "===== СТАРТ МИГРАЦИИ ЯРМАРОК =====",
        "-" * 92,
        f"Профиль         : {profile}",
        f"Стенд           : {base_url}",
        f"Режим           : {mode}",
        f"Интерактивный   : {'да' if interactive else 'нет'}",
        f"Operator mode   : {'включен' if operator_mode else 'выключен'}",
        f"checkpoints     : {state_file_path}",
        f"Лог успехов     : {success_log_path or '-'}",
        f"Лог ошибок      : {fail_log_path or '-'}",
        "#" * 92,
    ]
    user_logger.info("\n".join(lines))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Миграция данных реестров ярмарок")
    parser.add_argument("--profile", choices=["custom", "dev", "psi", "prod"], default="dev")
    parser.add_argument("--base-url", default="")
    parser.add_argument("--jwt-url", default="")
    parser.add_argument("--ui-base-url", default="")

    parser.add_argument("--mode", choices=["auto", "single", "mass"], default="auto")
    parser.add_argument(
        "--workbook",
        default=str(Path(SCRIPT_DIR) / EXCEL_FILE_NAME),
        help="Путь к книге Excel (совместимость)",
    )
    parser.add_argument("--workbooks", default="", help="Явный список книг (разделитель ';' или новая строка)")
    parser.add_argument("--files-dir", default=FILES_DIR, help="Папка с файлами для загрузки")
    parser.add_argument("--files-map", default="", help="Связка книга->папка файлов: 'book1.xlsm=dir1;book2.xlsm=dir2'")
    parser.add_argument("--ask-files-always", action="store_true", help="Всегда спрашивать папку файлов для каждой книги")

    parser.add_argument("--sheet", choices=["all", "fair", "mesto", "permits"], default="all", help="Какие листы запускать")
    parser.add_argument("--limit", type=int, default=0, help="Ограничение по числу строк на лист")
    parser.add_argument("--auth-only", action="store_true", help="Проверить авторизацию и завершить работу")
    parser.add_argument("--skip-auth", action="store_true", help="Пропустить авторизацию. Допустимо только с --dry-run")
    parser.add_argument("--dry-run", action="store_true", help="Только собрать payload без записи в API")
    parser.add_argument("--no-auth", action="store_true", help=argparse.SUPPRESS)

    parser.add_argument("--operator-mode", action="store_true", help="На ошибке строки: retry/skip/abort")
    parser.add_argument("--no-prompt", action="store_true", help="Не запрашивать input, использовать значения из файлов")
    parser.add_argument("--no-interactive", action="store_true", help="Отключить интерактивный режим")

    parser.add_argument("--state-file", default=str(STATE_FILE), help="Путь к checkpoints JSON")
    parser.add_argument("--reset-state", action="store_true", help="Очистить checkpoints перед запуском")
    parser.add_argument("--resume", dest="resume", action="store_true", default=RESUME_BY_DEFAULT, help="Продолжать с checkpoints")
    parser.add_argument("--no-resume", dest="resume", action="store_false", help="Игнорировать checkpoints")
    return parser.parse_args()


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
    dry_run,
    operator_mode,
    interactive,
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
        return False

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
        logger.info("%s: обработка строки %s/%s", list_name, row_num, rows_total)
        while True:
            try:
                if list_name == SHEET_FAIR:
                    rec_data = process_fair_sheet(row, logger, session=session if not (TEST or dry_run) else None)
                elif list_name == SHEET_MESTO:
                    rec_data = process_mesto_sheet(row, logger, session=session if not (TEST or dry_run) else None)
                else:
                    rec_data = process_permits_sheet(row, logger)

                if TEST or dry_run:
                    logger.info("[DRY] %s", json.dumps(rec_data, ensure_ascii=False))
                    if list_name == SHEET_MESTO:
                        markets = rec_data.get("marketInfo") or []
                        if not markets:
                            logger.info(
                                "[DRY][NSI] %s",
                                json.dumps(build_nsi_local_object_fair_from_payload(rec_data, None), ensure_ascii=False),
                            )
                        else:
                            for market in markets:
                                logger.info(
                                    "[DRY][NSI] %s",
                                    json.dumps(build_nsi_local_object_fair_from_payload(rec_data, market), ensure_ascii=False),
                                )
                    break

                record_json, response = create_record(session, logger, collection, rec_data)
                if not record_json:
                    status_code = getattr(response, "status_code", "n/a")
                    logger.error("Ошибка при создании записи в %s: HTTP %s", collection, status_code)
                    fail_logger.info(f"{list_name}:{row_num}")
                    stats["failedRows"] += 1
                    action = _operator_action_on_row_error(
                        operator_mode=operator_mode,
                        interactive=interactive,
                        logger=logger,
                        context=f"{list_name}:{row_num}:create_http_{status_code}",
                    )
                    if action == "retry":
                        continue
                    if action == "abort":
                        return True
                    break

                log_success(success_logger, {"_id": record_json["_id"], "guid": record_json["guid"], "parentEntries": collection})
                logger.info("Создана запись в %s: %s", collection, record_json["_id"])

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
                        action = _operator_action_on_row_error(
                            operator_mode=operator_mode,
                            interactive=interactive,
                            logger=logger,
                            context=f"{list_name}:{row_num}:permit_upload",
                        )
                        if action == "retry":
                            continue
                        if action == "abort":
                            return True
                        break

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
                    break

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
                    break

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
                            "Ошибка при создании записи в %s: HTTP %s",
                            NSI_LOCAL_OBJECT_FAIR_COLLECTION,
                            status_code,
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
                    logger.info("Создана запись в %s: %s", NSI_LOCAL_OBJECT_FAIR_COLLECTION, nsi_json["_id"])

                if nsi_failed:
                    action = _operator_action_on_row_error(
                        operator_mode=operator_mode,
                        interactive=interactive,
                        logger=logger,
                        context=f"{list_name}:{row_num}:nsi_create",
                    )
                    if action == "retry":
                        continue
                    if action == "abort":
                        return True
                    break

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
                break
            except Exception as exc:
                logger.error("Ошибка при обработке строки %s: %s", row_num, exc)
                fail_logger.info(f"{list_name}:{row_num}")
                logger.debug(traceback.format_exc())
                stats["failedRows"] += 1
                action = _operator_action_on_row_error(
                    operator_mode=operator_mode,
                    interactive=interactive,
                    logger=logger,
                    context=f"{list_name}:{row_num}:exception",
                )
                if action == "retry":
                    continue
                if action == "abort":
                    return True
                break
    return False


def main():
    args = parse_args()
    warnings.filterwarnings("ignore", category=InsecureRequestWarning)

    logger = setup_logger()
    success_logger = setup_success_logger()
    fail_logger = setup_fail_logger()
    user_logger = setup_user_logger()

    interactive = not args.no_interactive and not args.no_prompt
    try:
        profile_name, runtime_base_url, runtime_jwt_url, runtime_ui_url = _resolve_runtime(args)
    except Exception as exc:
        logger.error("Runtime config error: %s", exc)
        return 1
    set_runtime_urls(base_url=runtime_base_url, jwt_url=runtime_jwt_url, ui_base_url=runtime_ui_url)

    skip_auth = bool(args.skip_auth or args.no_auth)
    if skip_auth and not args.dry_run:
        logger.error("--skip-auth can be used only with --dry-run")
        return 1
    if args.auth_only and skip_auth:
        logger.error("--auth-only cannot be used together with --skip-auth")
        return 1

    try:
        workbook_specs, resolved_mode = _resolve_workbook_specs(args, interactive, logger)
    except Exception as exc:
        logger.error("Workbook selection error: %s", exc)
        return 1

    state_path = Path(args.state_file).expanduser().resolve()
    state_namespace = f"fairs_migration:{profile_name}"
    state = ResumeState(path=state_path, namespace=state_namespace, enabled=True)
    if args.reset_state:
        state.reset_namespace()
        logger.info("State reset: %s", state_path)

    resume_enabled = bool(args.resume)
    if resume_enabled:
        strategy = _choose_resume_strategy(state=state, logger=logger, user_logger=user_logger, interactive=interactive)
        if strategy == "abort":
            logger.warning("Stopped by operator before start")
            return 1
        if strategy == "reset":
            state.clear_rows()
            resume_enabled = False
            logger.info("Checkpoints cleared before start")

    logger.info(
        "%s",
        _console_block(
            "Fairs migration start",
            [
                f"Profile      : {profile_name}",
                f"Base URL     : {get_runtime_base_url()}",
                f"UI URL       : {get_runtime_ui_base_url()}",
                f"Mode         : {resolved_mode}",
                f"Sheet mode   : {args.sheet}",
                f"Dry run      : {args.dry_run}",
                f"Skip auth    : {skip_auth}",
                f"Resume       : {resume_enabled}",
                f"State file   : {state_path}",
                f"Interactive  : {interactive}",
                f"Operator mode: {args.operator_mode}",
                f"Workbook(s)  : {len(workbook_specs)}",
            ],
        ),
    )
    _log_user_run_header(
        user_logger=user_logger,
        profile=profile_name,
        base_url=get_runtime_base_url(),
        mode=resolved_mode,
        interactive=interactive,
        operator_mode=bool(args.operator_mode),
        state_file_path=str(state_path),
        success_log_path=getattr(success_logger, "log_path", ""),
        fail_log_path=getattr(fail_logger, "log_path", ""),
    )

    state.begin_run(
        profile=profile_name,
        baseUrl=get_runtime_base_url(),
        uiBaseUrl=get_runtime_ui_base_url(),
        mode=resolved_mode,
        sheet=args.sheet,
        dryRun=bool(args.dry_run),
        resume=resume_enabled,
        operatorMode=bool(args.operator_mode),
        workbooks=[spec.workbook_path for spec in workbook_specs],
        filesMap={spec.workbook_path: spec.files_dir for spec in workbook_specs},
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
        if not skip_auth:
            session = setup_session(logger, no_prompt=(args.no_prompt or args.no_interactive))
            if session is None:
                raise RuntimeError("Auth failed")
            logger.info("Auth succeeded")

        if args.auth_only:
            logger.info("--auth-only: auth checked, migration skipped")
            state.finish_run(
                status="completed",
                summary={
                    "authOnly": True,
                    "workbooks": len(workbook_specs),
                    "createdRows": 0,
                    "processedRows": 0,
                    "resumeSkips": 0,
                    "failedRows": 0,
                },
                clear_rows=True,
            )
            return 0

        for spec in workbook_specs:
            workbook_path = Path(spec.workbook_path).expanduser().resolve()
            files_dir_active = str(Path(spec.files_dir).expanduser().resolve())
            if not workbook_path.exists():
                raise FileNotFoundError(f"Workbook not found: {workbook_path}")

            logger.info("%s", _console_block("Workbook", [f"Path: {workbook_path}", f"Files: {files_dir_active}"]))
            user_logger.info(
                "WORKBOOK_START | workbook=%s | files=%s | sheet=%s | resume=%s",
                str(workbook_path),
                files_dir_active,
                args.sheet,
                resume_enabled,
            )

            for list_name in _selected_lists(args.sheet):
                logger.info("Process sheet: %s | workbook: %s", list_name, workbook_path.name)
                excel = read_excel(str(workbook_path), skiprows=3, sheet_name=list_name)
                if excel is None:
                    raise RuntimeError(f"Cannot read sheet '{list_name}' from workbook {workbook_path}")
                excel = excel.iloc[1:].reset_index(drop=True)
                excel.columns = [str(column).strip() for column in excel.columns]
                logger.info("Rows loaded: %s", len(excel))

                stopped_here = process_sheet_data(
                    list_name,
                    excel,
                    session,
                    logger,
                    success_logger,
                    fail_logger,
                    workbook_path=str(workbook_path),
                    files_dir=files_dir_active,
                    state=state,
                    resume_enabled=resume_enabled,
                    limit=args.limit,
                    stats=stats,
                    dry_run=bool(args.dry_run),
                    operator_mode=bool(args.operator_mode),
                    interactive=interactive,
                )
                if stopped_here:
                    stopped = True
                    logger.warning("Stopped by operator on %s/%s", workbook_path.name, list_name)
                    break

            user_logger.info(
                "WORKBOOK_FINISH | workbook=%s | created=%s | failed=%s | resume_skips=%s",
                str(workbook_path),
                stats["createdRows"],
                stats["failedRows"],
                stats["resumeSkips"],
            )
            if stopped:
                break

    except KeyboardInterrupt:
        stopped = True
        logger.warning("Interrupted by Ctrl+C")
    except Exception as exc:
        fatal_error = True
        logger.error("Critical error: %s", exc)
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
            "Finish: status=%s created=%s failed=%s resumed_skips=%s",
            status,
            stats["createdRows"],
            stats["failedRows"],
            stats["resumeSkips"],
        )
    return 0 if not stopped and not fatal_error and stats["failedRows"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())

