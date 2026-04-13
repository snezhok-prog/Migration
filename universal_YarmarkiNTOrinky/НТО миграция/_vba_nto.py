from __future__ import annotations

import datetime
import re
from dataclasses import dataclass
from typing import Any, Dict, Iterable, List, Optional, Tuple

from _api import search_org_by_ogrn
from _config import DEFAULT_ORG
from _nto_mappings import (
    MAP_BIDDING_FORM,
    MAP_BIDDING_STATUS,
    MAP_NOTICE_TYPE,
    MAP_ORG_STATE_FORM,
    MAP_OWNERSHIP,
    MAP_SPECIAL_USE_ZONES,
)
from _utils import generate_guid, get_any, get_path, is_empty, split_tokens


FORCE_STRING_KEYS = {
    "bik",
    "oktmo",
    "inn_upolnomochennogo_organa",
    "inn_2",
    "inn_3",
    "inn_4",
    "kpp",
    "kpp_2",
    "kpp_3",
    "ogrn",
    "ogrn_2",
    "ogrn_upolnomochennogo_organa",
    "uin",
    "kbk",
    "kaznacheyskiy_schyot",
    "edinyy_kaznacheyskiy_schyot",
    "raschyotnyy_schyot",
    "korrespondentskiy_schyot",
    "nomer_izvescheniya",
    "nomer_protokola",
    "noticeSignedNumber",
    "ProtocolConsiderNumber",
}

_RU_TO_LATIN = {
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

_MAP_NTO_TYPE = {
    "автоцистерна": {"code": "TypeNTOAutoCisterna", "name": "Автоцистерна"},
    "автомагазин (торговый автофургон, автолавка)": {
        "code": "TypeNTOAutoMag",
        "name": "Автомагазин (торговый автофургон, автолавка)",
    },
    "бахчевой развал": {"code": "TypeNTOBaxRaz", "name": "Бахчевой развал"},
    "ёлочный базар": {"code": "TypeNTOElBaz", "name": "Ёлочный базар"},
    "киоск": {"code": "TypeNTOkiosk", "name": "Киоск"},
    "лоток": {"code": "TypeNTOLotok", "name": "Лоток"},
    "объект инд. образца": {"code": "TypeNTOObject", "name": "Объект инд. образца"},
    "отдельно стоящее нестационарное кафе": {
        "code": "TypeNTOOtdelCoffee",
        "name": "Отдельно стоящее нестационарное кафе",
    },
    "иной": {"code": "TypeNTOOther", "name": "Иной"},
    "павильон": {"code": "TypeNTOPav", "name": "Павильон"},
    "прилегающее нестационарное кафе": {
        "code": "TypeNTOPrilegCoffee",
        "name": "Прилегающее нестационарное кафе",
    },
    "шатер": {"code": "TypeNTOShater", "name": "Шатер"},
    "летнее кафе": {"code": "TypeNTOSummerCoffee", "name": "Летнее кафе"},
    "тележка": {"code": "TypeNTOTelezhka", "name": "Тележка"},
    "торговый автомат (вендинговый автомат)": {
        "code": "TypeNTOTorgAuto",
        "name": "Торговый автомат (вендинговый автомат)",
    },
    "торговая галерея": {"code": "TypeNTOTorgGal", "name": "Торговая галерея"},
    "трейлер": {"code": "TypeNTOTrailer", "name": "Трейлер"},
    "зонтик": {"code": "TypeNTOZontik", "name": "Зонтик"},
}

_MAP_SPECIALIZATION = {
    "непродовольственная": {"code": "SpecificNoProd", "name": "Непродовольственная"},
    "продовольственная": {"code": "SpecificProd", "name": "Продовольственная"},
}

_MAP_INFO_SMSP = {
    "исключительно смсп": {"code": "InfoSMSPIskl", "name": "Исключительно СМСП"},
    "требование не предъявляется": {
        "code": "InfoSMSPTrebNoPred",
        "name": "Требование не предъявляется",
    },
}

_MAP_PLACEMENT_SCHEDULE = {
    "по графику": {"code": "Scheduled", "name": "По графику"},
    "сезонно": {"code": "Seasonal", "name": "Сезонно"},
    "круглогодично": {"code": "YearRound", "name": "Круглогодично"},
}

_MAP_SEASON = {
    "осень": {"code": "Autumn", "name": "Осень"},
    "весна": {"code": "Spring", "name": "Весна"},
    "лето": {"code": "Summer", "name": "Лето"},
    "зима": {"code": "Winter", "name": "Зима"},
}

_MAP_RECORD_STATUS = {
    "подлежит утверждению нпа": {"code": "approvalNPA", "name": "Подлежит утверждению НПА"},
    "утверждено": {"code": "approved", "name": "Утверждено"},
    "утверждено, запланированы изменения, подлежащие утверждению нпа": {
        "code": "approvedPlannedChangesNPA",
        "name": "Утверждено, запланированы изменения, подлежащие утверждению НПА",
    },
    "черновик": {"code": "draft", "name": "Черновик"},
    "отклонено": {"code": "rejected", "name": "Отклонено"},
}

_MAP_SPEC_PROJECT_STATUS = {
    "существующее, торги не запущены (не подлежащие замене или переносу)": {
        "code": "StatusProjectPlaceApprovedExistingNoTenderNoReplacement",
        "name": "Существующее, торги не запущены (не подлежащие замене или переносу)",
        "parentId": "68ed183427eea1af1d547524",
    },
    "существующее, запущены торги (не подлежащие замене или переносу)": {
        "code": "StatusProjectPlaceApprovedExistingTenderNoReplacement",
        "name": "Существующее, запущены торги (не подлежащие замене или переносу)",
        "parentId": "68ed183427eea1af1d547524",
    },
    "занято (не подлежащие замене или переносу)": {
        "code": "StatusProjectPlaceApprovedOccupiedNoReplacement",
        "name": "Занято (не подлежащие замене или переносу)",
        "parentId": "68ed183427eea1af1d547524",
    },
    "занято (подлежащие замене или переносу)": {
        "code": "StatusProjectPlaceApprovedOccupiedWithReplacement",
        "name": "Занято (подлежащие замене или переносу)",
        "parentId": "68ed183427eea1af1d547524",
    },
    "предложено заявителем": {
        "code": "StatusProjectPlaceDraftProposedByApplicant",
        "name": "Предложено заявителем",
        "parentId": "68ed183d6c192c814d0d88ae",
    },
    "предложено муниципалитетом": {
        "code": "StatusProjectPlaceDraftProposedByMunicipality",
        "name": "Предложено муниципалитетом",
        "parentId": "68ed183d6c192c814d0d88ae",
    },
}

_MAP_NO_PROD = {
    "аксессуары": "SpecificNoProdAccessories",
    "автотовары": "SpecificNoProdAutoGoods",
    "книги": "SpecificNoProdBooks",
    "детские товары": "SpecificNoProdChildrenGoods",
    "одежда": "SpecificNoProdClothing",
    "строительные материалы, хозяйственные товары": "SpecificNoProdConstructionMaterials",
    "цветы": "SpecificNoProdFlowers",
    "продукты": "SpecificNoProdFood",
    "галантерейные товары": "SpecificNoProdHaberdashery",
    "хозяйственные товары": "SpecificNoProdHouseholdGoods",
    "бытовые услуги": "SpecificNoProdHouseholdServices",
    "сотовая связь": "SpecificNoProdMobileCommunications",
    "парфюмерия и косметические товары": "SpecificNoProdPerfumesCosmetics",
    "зоотовары": "SpecificNoProdPetGoods",
    "аптека": "SpecificNoProdPharmacy",
    "фотоуслуги": "SpecificNoProdPhotoServices",
    "почтомат": "SpecificNoProdPostomat",
    "печать": "SpecificNoProdPrinting",
    "пункт приёма вторичного сырья": "SpecificNoProdRecyclingCenter",
    "ремонт обуви": "SpecificNoProdRepairShoes",
    "сувениры / народные промыслы": "SpecificNoProdSouvenirs",
    "канцелярские товары": "SpecificNoProdStationery",
    "театральные билеты": "SpecificNoProdTheaterTickets",
    "шиномонтаж": "SpecificNoProdTireService",
    "проездные билеты": "SpecificNoProdTransportTickets",
    "ёлки, сосны, лапник": "SpecificNoProdTrees",
    "иное": "SpecificNoProdOther",
    "иной": "SpecificNoProdOther",
    "подарки": "SpecificNoProdGifts",
    "парфюмерные и косметические товары": "SpecificNoProdPerfumesCosmetics",
    "пункт приема вторичного сырья": "SpecificNoProdRecyclingCenter",
    "билеты на морской, речной транспорт": "SpecificNoProdSeaTickets",
    "сувениры/народные промыслы": "SpecificNoProdSouvenirs",
}

_MAP_PROD = {
    "выпечка": "SpecificProdBakery",
    "хлеб": "SpecificProdBread",
    "общественное питание": "SpecificProdCatering",
    "прохладительные напитки": "SpecificProdColdDrink",
    "кукуруза": "SpecificProdCorn",
    "молоко, молочная продукция": "SpecificProdDairy",
    "вода в розлив": "SpecificProdDraftWater",
    "фрукты": "SpecificProdFruits",
    "горячие напитки": "SpecificProdHotDrink",
    "мороженое": "SpecificProdIceCream",
    "бахчевые культуры": "SpecificProdMelons",
    "иной": "SpecificProdOther",
    "хлебобулочные изделия": "SpecificProdPastries",
    "снеки": "SpecificProdSnacks",
    "овощи": "SpecificProdVegetables",
}


@dataclass
class PendingUpload:
    filename: str
    target_path: str
    allow_external: bool = False


def coerce_string_ids(obj: Any) -> None:
    if not isinstance(obj, dict):
        return
    for key in list(obj.keys()):
        if key in FORCE_STRING_KEYS and obj[key] is not None:
            obj[key] = str(obj[key]).strip()


def _normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\u00A0", " ").strip()
    return re.sub(r"\s+", " ", text)


def norm_str(value: Any) -> str:
    return _normalize_text(value)


def norm_key(value: Any) -> str:
    return norm_str(value).lower()


def latin_key(value: Any, lowercase: bool = True) -> str:
    text = _normalize_text(value)
    if lowercase:
        text = text.lower()
    if not text:
        return ""
    parts: List[str] = []
    for ch in text:
        if ch in _RU_TO_LATIN:
            parts.append(_RU_TO_LATIN[ch])
        elif ch.isascii() and (ch.isalnum() or ch == "_"):
            parts.append(ch)
        else:
            parts.append("_")
    return re.sub(r"_+", "_", "".join(parts)).strip("_")


def as_date_or_null(value: Any) -> Optional[str]:
    if value is None:
        return None
    if isinstance(value, datetime.datetime):
        return value.date().isoformat()
    if isinstance(value, datetime.date):
        return value.isoformat()
    if isinstance(value, (int, float)):
        try:
            epoch = datetime.datetime(1899, 12, 30)
            return (epoch + datetime.timedelta(days=float(value))).date().isoformat()
        except Exception:
            return None
    text = norm_str(value)
    if not text:
        return None
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d.%m.%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    return None


def excel_time_to_hhmm(value: Any) -> Optional[str]:
    if value is None or value == "":
        return None
    if isinstance(value, datetime.time):
        return value.strftime("%H:%M")
    if isinstance(value, datetime.datetime):
        return value.strftime("%H:%M")
    if isinstance(value, str) and ":" in value:
        return value.strip()
    try:
        number = float(value)
    except Exception:
        number = None
    if number is None or not (0 <= number < 1):
        text = norm_str(value)
        if re.match(r"^\d{1,2}[.,]\d{2}$", text):
            left, right = re.split(r"[.,]", text)
            return f"{int(left):02d}:{int(right):02d}"
        return text or None
    total_minutes = round(number * 24 * 60)
    return f"{total_minutes // 60:02d}:{total_minutes % 60:02d}"


def to_number_or_null(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = norm_str(value).replace(" ", "").replace(",", ".")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def as_bool_yes_ru(value: Any) -> Optional[bool]:
    text = norm_key(value)
    if not text:
        return None
    if text in {"да", "yes", "true", "1"}:
        return True
    if text in {"нет", "no", "false", "0"}:
        return False
    return None


def addr_obj(value: Any) -> Optional[Dict[str, str]]:
    if value is None:
        return None

    def build(full_addr: Optional[str], postal_code: Optional[str]) -> Optional[Dict[str, str]]:
        out: Dict[str, str] = {}
        full = norm_str(full_addr)
        postal = norm_str(postal_code)
        if full:
            match = re.match(r"^\s*(\d{6})(?:\s*,\s*|\s+)(.*)$", full)
            if match:
                postal = postal or match.group(1)
                full = norm_str(match.group(2)) or full
            elif not postal:
                match2 = re.search(r"(\d{6})", full)
                if match2:
                    postal = match2.group(1)
                    full = norm_str(re.sub(re.escape(postal), "", full, count=1))
        if postal:
            out["postalCode"] = postal
        if full:
            out["fullAddress"] = full
        return out or None

    if isinstance(value, dict):
        return build(
            value.get("fullAddress") or value.get("full_address") or value.get("address"),
            value.get("postalCode") or value.get("postal_code") or value.get("postIndex"),
        )

    return build(value, None)


def _get_merged_value(ws, row: int, col: int) -> Any:
    if row <= 0:
        return None
    cell = ws.cell(row=row, column=col)
    if cell.value is not None:
        return cell.value
    for merged in ws.merged_cells.ranges:
        if merged.min_row <= row <= merged.max_row and merged.min_col <= col <= merged.max_col:
            return ws.cell(row=merged.min_row, column=merged.min_col).value
    return None


def _find_last_used_row(ws) -> int:
    for row in range(ws.max_row, 0, -1):
        for col in range(1, ws.max_column + 1):
            if not is_empty(ws.cell(row=row, column=col).value):
                return row
    return 1


def _last_used_col_in_row(ws, row: int) -> int:
    if row < 1:
        return 0
    for col in range(ws.max_column, 0, -1):
        if not is_empty(ws.cell(row=row, column=col).value):
            return col
    return 0


def _count_non_empty(ws, row: int, c1: int, c2: int) -> int:
    total = 0
    for col in range(c1, min(c2, ws.max_column) + 1):
        if not is_empty(_get_merged_value(ws, row, col)):
            total += 1
    return total


def _score_requirement_row(ws, row: int, c1: int, c2: int) -> int:
    if row < 1:
        return 0
    values: List[str] = []
    for col in range(c1, min(c2, ws.max_column) + 1):
        text = _normalize_text(_get_merged_value(ws, row, col))
        if text:
            values.append(text)
    if not values:
        return 0
    short_cnt = sum(1 for item in values if len(item) <= 4)
    freq: Dict[str, int] = {}
    for item in values:
        freq[item] = freq.get(item, 0) + 1
    max_freq = max(freq.values())
    non_empty = len(values)
    short_pct = 100.0 * short_cnt / non_empty
    repeat_pct = 100.0 * max_freq / non_empty
    return int(0.6 * short_pct + 0.4 * repeat_pct)


def detect_header_and_data(ws) -> Tuple[int, int, int, int]:
    first_col = 2
    last_used_row = _find_last_used_row(ws)
    for row in range(1, max(2, last_used_row)):
        last_col = _last_used_col_in_row(ws, row)
        if last_col < first_col:
            continue
        heads = _count_non_empty(ws, row, first_col, last_col)
        req = _score_requirement_row(ws, row + 1, first_col, last_col)
        if heads >= 3 and req >= 60:
            return row, first_col, last_col, row + 2
    return 0, first_col, ws.max_column, 0


def read_headers_unique(ws, header_row: int, first_col: int, last_col: int) -> List[str]:
    headers: List[str] = []
    seen: Dict[str, int] = {}
    for col in range(first_col, min(last_col, ws.max_column) + 1):
        raw = _normalize_text(ws.cell(row=header_row, column=col).value)
        if raw == "":
            raw = f"col{col - first_col + 1}"
        header = latin_key(raw, lowercase=True) or f"col{col - first_col + 1}"
        if header in seen:
            seen[header] += 1
            header = f"{header}_{seen[header]}"
        else:
            seen[header] = 1
        headers.append(header)
    return headers


def build_group_keys_2levels(ws, header_row: int, first_col: int, last_col: int) -> Tuple[List[str], List[str]]:
    row_l2 = header_row - 2
    row_l1 = header_row - 1
    prev2 = ""
    prev1 = ""
    groups_l2: List[str] = []
    groups_l1: List[str] = []
    for col in range(first_col, min(last_col, ws.max_column) + 1):
        raw2 = _normalize_text(_get_merged_value(ws, row_l2, col)) if row_l2 >= 1 else ""
        raw1 = _normalize_text(_get_merged_value(ws, row_l1, col)) if row_l1 >= 1 else ""
        if raw2 == "":
            raw2 = prev2
        if raw1 == "":
            raw1 = prev1
        if raw2:
            prev2 = raw2
        if raw1:
            prev1 = raw1
        groups_l2.append(latin_key(raw2, lowercase=True) if raw2 else "")
        groups_l1.append(latin_key(raw1, lowercase=True) if raw1 else "")
    return groups_l2, groups_l1


def _field_base(value: str) -> str:
    if not value:
        return ""
    match = re.search(r"_(\d+)$", value)
    return value[: match.start()] if match else value


def detect_arrays_by_duplicates(
    field_keys: List[str],
    groups_l2: List[str],
    groups_l1: List[str],
) -> Tuple[List[Optional[str]], List[int]]:
    size = len(field_keys)
    arr_name: List[Optional[str]] = [None] * size
    arr_idx: List[int] = [0] * size
    groups: Dict[str, List[int]] = {}
    for i in range(size):
        if groups_l1[i]:
            groups.setdefault(f"{groups_l2[i]}|{groups_l1[i]}", []).append(i)
    for cols in groups.values():
        count = len(cols)
        if count < 2:
            continue
        base = [_field_base(field_keys[i]) for i in cols]
        pos2 = 0
        for j in range(1, count):
            if base[j] == base[0]:
                pos2 = j + 1
                break
        if pos2 == 0:
            continue
        period = pos2 - 1
        if period <= 0 or count < 2 * period:
            continue
        if len(set(base[:period])) != period:
            continue
        if not all(base[i] == base[i + period] for i in range(count - period)):
            continue
        arr = groups_l1[cols[0]]
        for j, idx in enumerate(cols):
            arr_name[idx] = arr
            arr_idx[idx] = (j // period) + 1
    return arr_name, arr_idx


def _value_to_python(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value) if value.is_integer() else value
    if isinstance(value, datetime.datetime):
        return value
    if isinstance(value, datetime.date):
        return value.isoformat()
    if isinstance(value, datetime.time):
        return value.strftime("%H:%M")
    if isinstance(value, str):
        return value.strip()
    return value


def _row_has_any(values: List[Any]) -> bool:
    return any(not is_empty(v) for v in values)


def _legacy_key_alias(key: str) -> str:
    alias = str(key or "")
    alias = alias.replace("shch", "sch")
    alias = alias.replace("kh", "h")
    return alias


def _inject_legacy_aliases(obj: Any) -> Any:
    if isinstance(obj, list):
        return [_inject_legacy_aliases(item) for item in obj]
    if not isinstance(obj, dict):
        return obj

    out: Dict[str, Any] = {}
    for key, value in obj.items():
        nested = _inject_legacy_aliases(value)
        out[key] = nested
        alias = _legacy_key_alias(key)
        if alias != key and alias not in out:
            out[alias] = nested
    return out


def row_to_json_2level_fast(
    data: List[List[Any]],
    row_idx: int,
    field_keys: List[str],
    groups_l2: List[str],
    groups_l1: List[str],
    arr_name: List[Optional[str]],
    arr_idx: List[int],
) -> Dict[str, Any]:
    top_order: List[str] = []
    top_fields: Dict[str, List[Tuple[str, Any]]] = {}
    top_blocks: Dict[str, Dict[str, List[Tuple[str, Any]]]] = {}
    top_arrays: Dict[str, Dict[str, Dict[int, List[Tuple[str, Any]]]]] = {}

    for i, field_name in enumerate(field_keys):
        section = groups_l2[i].strip() if groups_l2[i] else ""
        sub_block = groups_l1[i].strip() if groups_l1[i] else ""
        final_name = field_name or f"col{i + 1}"
        value = _value_to_python(data[row_idx][i])
        top_key = section if section else (sub_block if sub_block and arr_idx[i] == 0 else "prochee")
        if top_key not in top_order:
            top_order.append(top_key)
        if arr_idx[i] > 0 and arr_name[i]:
            top_arrays.setdefault(top_key, {}).setdefault(arr_name[i], {}).setdefault(arr_idx[i], []).append((final_name, value))
        elif sub_block and section:
            top_blocks.setdefault(top_key, {}).setdefault(sub_block, []).append((final_name, value))
        else:
            top_fields.setdefault(top_key, []).append((final_name, value))

    result: Dict[str, Any] = {}
    for top_key in top_order:
        body: Dict[str, Any] = {}
        for name, value in top_fields.get(top_key, []):
            body[name] = value
        for sub_key, entries in top_blocks.get(top_key, {}).items():
            nested: Dict[str, Any] = {}
            for name, value in entries:
                nested[name] = value
            body[sub_key] = nested
        for arr_key, groups in top_arrays.get(top_key, {}).items():
            items: List[Dict[str, Any]] = []
            for idx in sorted(groups.keys()):
                item: Dict[str, Any] = {}
                for name, value in groups[idx]:
                    item[name] = value
                items.append(item)
            body[arr_key] = items
        result[top_key] = body
    return result


def iter_vba_rows(ws, row_limit: int = 0, start_row: Optional[int] = None) -> Iterable[Tuple[int, Dict[str, Any]]]:
    if start_row is not None:
        header_row = start_row - 1
        first_col = 2
        last_col = _last_used_col_in_row(ws, header_row)
        data_start = start_row
    else:
        header_row, first_col, last_col, data_start = detect_header_and_data(ws)
        if header_row <= 0:
            raise ValueError(f"Не удалось определить строки заголовков для листа '{ws.title}'")

    headers = read_headers_unique(ws, header_row, first_col, last_col)
    groups_l2, groups_l1 = build_group_keys_2levels(ws, header_row, first_col, last_col)
    arr_name, arr_idx = detect_arrays_by_duplicates(headers, groups_l2, groups_l1)
    last_row = _find_last_used_row(ws)
    if data_start > last_row:
        return

    emitted = 0
    for offset, row in enumerate(
        ws.iter_rows(
            min_row=data_start,
            max_row=last_row,
            min_col=first_col,
            max_col=last_col,
            values_only=True,
        )
    ):
        row_values = list(row)
        if not _row_has_any(row_values):
            continue
        payload = row_to_json_2level_fast([row_values], 0, headers, groups_l2, groups_l1, arr_name, arr_idx)
        payload = _inject_legacy_aliases(payload)
        yield data_start + offset, payload
        emitted += 1
        if row_limit and emitted >= row_limit:
            break


def build_unit_info(src: Any) -> Optional[Dict[str, Any]]:
    unit_name = norm_str(get_any(src, ["Структурное подразделение", "unit"], None))
    if not unit_name:
        return None
    return {"id": "", "_id": "", "guid": generate_guid(), "name": unit_name, "shortName": unit_name}


def build_org_info_fallback(src: Any) -> Dict[str, Any]:
    org = dict(DEFAULT_ORG)
    ogrn = norm_str(get_any(src, ["ОГРН организации", "ogrn_upolnomochennogo_organa", "orgn", "ogrn"], None))
    inn = norm_str(get_any(src, ["ИНН", "inn_upolnomochennogo_organa", "inn"], None))
    kpp = norm_str(get_any(src, ["КПП", "kpp"], None))
    name = norm_str(get_any(src, ["Наименование организации", "organization"], None))
    if ogrn:
        org["ogrn"] = ogrn
    if inn:
        org["inn"] = inn
    if kpp:
        org["kpp"] = kpp
    if name:
        org["name"] = name
        org["shortName"] = name
    return org


def build_org_info(src: Any, session, logger) -> Dict[str, Any]:
    ogrn = norm_str(get_any(src, ["ОГРН организации", "ogrn_upolnomochennogo_organa", "orgn", "ogrn"], None))
    org = search_org_by_ogrn(session, logger, ogrn) if session else None
    return org if org else build_org_info_fallback(src)


def map_nnto_type(value: Any) -> Optional[Dict[str, Any]]:
    key = norm_key(value)
    if not key:
        return None
    return _MAP_NTO_TYPE.get(key) or {"code": "TypeNTOOther", "name": norm_str(value)}


def map_nto_type_bidding(value: Any) -> Optional[Dict[str, Any]]:
    key = norm_key(value)
    if not key:
        return None
    return _MAP_NTO_TYPE.get(key) or {"code": None, "name": norm_str(value)}


def map_specialization(value: Any) -> Optional[Dict[str, Any]]:
    key = norm_key(value)
    if not key:
        return None
    return _MAP_SPECIALIZATION.get(key)


def map_product_type(value: Any) -> Optional[Dict[str, Any]]:
    text = norm_str(value)
    lower = text.lower()
    if not lower:
        return None
    if "товар" in lower:
        return {"code": "Goods", "name": "Товары"}
    if "услуг" in lower:
        return {"code": "Services", "name": "Услуги"}
    return {"code": "Other", "name": text}


def map_info_smsp(value: Any) -> Optional[Dict[str, Any]]:
    text = norm_str(value)
    if not text:
        return None
    return _MAP_INFO_SMSP.get(norm_key(text), {"code": None, "name": text})


def map_ownership_type(value: Any) -> Optional[Dict[str, Any]]:
    text = norm_str(value)
    if not text:
        return None
    return MAP_OWNERSHIP.get(norm_key(text), {"code": None, "name": text})


def map_placement_schedule(value: Any) -> Optional[Dict[str, Any]]:
    text = norm_str(value)
    if not text:
        return None
    return _MAP_PLACEMENT_SCHEDULE.get(norm_key(text), {"code": None, "name": text})


def map_season(value: Any) -> Optional[Dict[str, Any]]:
    text = norm_str(value)
    if not text:
        return None
    return _MAP_SEASON.get(norm_key(text), {"code": None, "name": text})


def map_project_status(value: Any) -> Optional[Dict[str, Any]]:
    text = norm_str(value)
    lower = text.lower()
    if not lower:
        return None
    if "утвержден" in lower:
        return {"code": "StatusProjectPlaceApproved", "name": "Утверждено", "_id": "68ed183427eea1af1d547524"}
    if "чернов" in lower:
        return {"code": "StatusProjectPlaceDraft", "name": "Черновик", "_id": "68ed183d6c192c814d0d88ae"}
    return {"code": "StatusProjectPlaceOther", "name": text}


def map_spec_project_status(value: Any) -> Optional[Dict[str, Any]]:
    text = norm_str(value)
    if not text:
        return None
    return _MAP_SPEC_PROJECT_STATUS.get(norm_key(text), {"code": None, "name": text, "parentId": None})


def map_record_status(value: Any) -> Optional[Dict[str, Any]]:
    text = norm_str(value)
    if not text:
        return None
    return _MAP_RECORD_STATUS.get(norm_key(text), {"code": None, "name": text})


def map_payment_period(value: Any) -> Optional[Dict[str, Any]]:
    text = norm_key(value)
    if not text:
        return None
    if text == "ежемесячно":
        return {"code": "monthly", "name": "Ежемесячно"}
    if text == "ежеквартально":
        return {"code": "quarterly", "name": "Ежеквартально"}
    if text == "ежегодно":
        return {"code": "yearly", "name": "Ежегодно"}
    if text == "единовременно":
        return {"code": "once", "name": "Единовременно"}
    if text == "monthly":
        return {"code": "monthly", "name": "Ежемесячно"}
    if text == "quarterly":
        return {"code": "quarterly", "name": "Ежеквартально"}
    if text == "yearly":
        return {"code": "yearly", "name": "Ежегодно"}
    if text == "once":
        return {"code": "once", "name": "Единовременно"}
    raw = norm_str(value)
    return {"code": raw, "name": raw}


def map_assort_no_prod(name: Any, other_text: Any) -> Dict[str, Any]:
    clean = norm_str(name)
    code = _MAP_NO_PROD.get(clean.lower()) if clean else None
    if clean.lower() in {"иной", "иное"}:
        code = "SpecificNoProdOther"
    item: Dict[str, Any] = {"Assort": {"name": clean, "code": code or None}}
    more = norm_str(other_text)
    if (code is None or code == "SpecificNoProdOther") and more:
        item["assortOther"] = more
    return item


def map_assort_prod(name: Any, other_text: Any) -> Dict[str, Any]:
    clean = norm_str(name)
    code = _MAP_PROD.get(clean.lower()) if clean else None
    if clean.lower() in {"иной", "иное"}:
        code = "SpecificProdOther"
    item: Dict[str, Any] = {"Assort": {"name": clean, "code": code or None}}
    more = norm_str(other_text)
    if (code is None or code == "SpecificProdOther") and more:
        item["assortOther"] = more
    return item


def map_assort_item(parent_code: Optional[str], name: str, other_text: Optional[str]) -> Dict[str, Any]:
    if parent_code == "SpecificProd":
        return map_assort_prod(name, other_text)
    if parent_code == "SpecificNoProd":
        return map_assort_no_prod(name, other_text)
    item: Dict[str, Any] = {"Assort": {"name": norm_str(name), "code": None}}
    more = norm_str(other_text)
    if more:
        item["assortOther"] = more
    return item


def map_org_state_form(value: Any) -> Optional[Dict[str, Any]]:
    key = norm_key(value)
    if not key:
        return None
    return MAP_ORG_STATE_FORM.get(key)


def map_notice_type(value: Any) -> Optional[Dict[str, Any]]:
    text = norm_str(value)
    if not text:
        return None
    return {"name": text, "code": MAP_NOTICE_TYPE.get(norm_key(text))}


def map_bidding_form(value: Any) -> Optional[Dict[str, Any]]:
    text = norm_str(value)
    if not text:
        return None
    key = norm_key(text)
    hit = MAP_BIDDING_FORM.get(key)
    if hit:
        return dict(hit)
    parent_code = "auction" if "аукцион" in key else "contest" if "конкурс" in key else None
    return {"code": None, "name": text, "parentCode": parent_code}


def map_specific(value: Any) -> Optional[Dict[str, Any]]:
    text = norm_str(value)
    if not text:
        return None
    return _MAP_SPECIALIZATION.get(norm_key(text), {"code": None, "name": text})


def map_bidding_status(value: Any) -> Optional[Dict[str, Any]]:
    text = norm_str(value)
    if not text:
        return None
    return {"name": text, "code": MAP_BIDDING_STATUS.get(norm_key(text))}


def map_special_use_zone(value: Any) -> Optional[Dict[str, Any]]:
    text = norm_str(value)
    if not text:
        return None
    return MAP_SPECIAL_USE_ZONES.get(norm_key(text), {"code": None, "name": text})


def _split_file_names(value: Any) -> List[str]:
    if value is None:
        return []

    def clean_name(raw: Any) -> str:
        text = norm_str(raw)
        if not text:
            return ""
        text = re.sub(r"[?#].*$", "", text)
        text = text.replace("\\", "/")
        return text.rsplit("/", 1)[-1].strip()

    if isinstance(value, str):
        names: List[str] = []
        for token in split_tokens(value, separators=";,\n"):
            cleaned = clean_name(token)
            if cleaned:
                names.append(cleaned)
        return names
    if isinstance(value, dict):
        for key in ("originalName", "fileName", "FileName", "filename", "name", "path", "url"):
            cleaned = clean_name(value.get(key))
            if cleaned:
                return [cleaned]
        return []
    if isinstance(value, list):
        names: List[str] = []
        for item in value:
            names.extend(_split_file_names(item))
        return names
    return []


def _collect_uploads_from_mapping(src: Any, mapping: List[Tuple[str, str]]) -> List[PendingUpload]:
    uploads: List[PendingUpload] = []
    for src_path, dst_path in mapping:
        value = get_path(src, src_path)
        allow_external = bool(isinstance(value, dict) and (value.get("path") or value.get("url")))
        for filename in _split_file_names(value):
            uploads.append(PendingUpload(filename=filename, target_path=dst_path, allow_external=allow_external))
    return uploads


def collect_pending_uploads_nto(src: Any) -> List[PendingUpload]:
    mapping = [
        (
            "dannye_po_reestru_mest_dlya_razmescheniya_nto.informatsiya_o_nto.kartograficheskiy_material_s_granitsami_predlagaemogo_mesta_raspolozheniya_nto",
            "ntoInformation.fileEnginTopogPlaceNew",
        ),
        (
            "dannye_po_reestru_mest_dlya_razmeshcheniya_nto.informatsiya_o_nto.kartograficheskiy_material_s_granitsami_predlagaemogo_mesta_raspolozheniya_nto",
            "ntoInformation.fileEnginTopogPlaceNew",
        ),
        (
            "dannye_po_reestru_mest_dlya_razmescheniya_nto.informatsiya_o_nto.foto_suschestvuyuschey_situatsii_bez_nto_pri_nalichii",
            "ntoInformation.photoSituationNTONew",
        ),
        (
            "dannye_po_reestru_mest_dlya_razmeshcheniya_nto.informatsiya_o_nto.foto_suschestvuyuschey_situatsii_bez_nto_pri_nalichii",
            "ntoInformation.photoSituationNTONew",
        ),
        (
            "dannye_po_reestru_mest_dlya_razmescheniya_nto.svedeniya_o_postanovlenii_o_vnesenii_izmeneniy_v_shemu_nto.postanovlenie_o_vnesenii_izmeneniy_v_shemu_nto",
            "resolutionBlock.npaAmendmentFile",
        ),
        (
            "dannye_po_reestru_mest_dlya_razmeshcheniya_nto.svedeniya_o_postanovlenii_o_vnesenii_izmeneniy_v_shemu_nto.postanovlenie_o_vnesenii_izmeneniy_v_shemu_nto",
            "resolutionBlock.npaAmendmentFile",
        ),
        (
            "informatsiya_o_dogovore_na_pravo_razmescheniya_nto.informatsiya_po_dogovoru.dogovor_na_pravo_razmescheniya_nto",
            "NTOContract.NTOPPlacementAgreement",
        ),
        (
            "informatsiya_o_dogovore_na_pravo_razmeshcheniya_nto.informatsiya_po_dogovoru.dogovor_na_pravo_razmescheniya_nto",
            "NTOContract.NTOPPlacementAgreement",
        ),
        (
            "informatsiya_o_dogovore_na_pravo_razmescheniya_nto.svedeniya_o_dopolnitelnom_soglashenii.dopolnitelnoe_soglashenie_k_dogovoru",
            "blockDopSogl[0].SupplementalAgreementIfAny",
        ),
        (
            "informatsiya_o_dogovore_na_pravo_razmeshcheniya_nto.svedeniya_o_dopolnitelnom_soglashenii.dopolnitelnoe_soglashenie_k_dogovoru",
            "blockDopSogl[0].SupplementalAgreementIfAny",
        ),
    ]
    uploads = _collect_uploads_from_mapping(src, mapping)
    files_src = get_any(src, ["Файлы", "files"], None)
    if files_src is not None:
        items = files_src if isinstance(files_src, list) else [files_src]
        for idx, item in enumerate(items):
            allow_external = bool(isinstance(item, dict) and (item.get("path") or item.get("url")))
            names = _split_file_names(item)
            if not names:
                continue
            uploads.append(PendingUpload(filename=names[0], target_path=f"files[{idx}]", allow_external=allow_external))

    deduped: List[PendingUpload] = []
    seen = set()
    for upload in uploads:
        key = (upload.target_path, upload.filename, upload.allow_external)
        if key in seen:
            continue
        seen.add(key)
        deduped.append(upload)
    return deduped


def collect_pending_uploads_torgov(src: Any) -> List[PendingUpload]:
    uploads: List[PendingUpload] = []

    def push_from_section(section: Any, *, aliases: List[str], pairs: List[Tuple[str, str]], dst_path: str) -> None:
        if not isinstance(section, dict):
            return
        for key in aliases:
            value = section.get(key)
            if value is None:
                continue
            allow_external = bool(isinstance(value, dict) and (value.get("path") or value.get("url")))
            for filename in _split_file_names(value):
                uploads.append(PendingUpload(filename=filename, target_path=dst_path, allow_external=allow_external))
        for _base_key, name_key in pairs:
            raw_name = section.get(name_key)
            if raw_name is None:
                continue
            for filename in _split_file_names({"filename": raw_name}):
                uploads.append(PendingUpload(filename=filename, target_path=dst_path, allow_external=False))

    notif = get_path(src, "dannye_po_reestru_torgov.izveschenie_i_dopolnitelnaya_dokumentatsiya") or {}
    reqs = get_path(src, "dannye_po_reestru_torgov.trebovaniya_k_zayavkam") or {}
    prot = get_path(src, "dannye_po_reestru_torgov.protokol_po_rezultatam_provedeniya_torgov") or {}

    push_from_section(
        notif,
        aliases=["izveschenie_o_provedenii_torgov", "agreedNotice", "agreed_notice", "agreed"],
        pairs=[("noticeFileBase64", "noticeFileName")],
        dst_path="GeneralInfoAuction.NotificationDocumentation.AgreedNotice",
    )
    push_from_section(
        notif,
        aliases=["proekt_zayavki_na_uchastie_v_torgah", "projectNotice", "project_notice"],
        pairs=[("projectNoticeBase64", "projectNoticeFileName")],
        dst_path="GeneralInfoAuction.NotificationDocumentation.ProjectNotice",
    )
    push_from_section(
        notif,
        aliases=["proekt_dogovora_na_pravo_razmeschenie_nto", "projectNoticeNTO", "project_notice_nto"],
        pairs=[("projectNoticeNTOBase64", "projectNoticeNTOFileName")],
        dst_path="GeneralInfoAuction.NotificationDocumentation.ProjectNoticeNTO",
    )
    push_from_section(
        notif,
        aliases=["izveschenie_ob_otkaze_ot_provedeniya_torgov", "modifiedNoticeRefusalAuction", "refusal_notice"],
        pairs=[("refusalNoticeBase64", "refusalNoticeFileName")],
        dst_path="GeneralInfoAuction.NotificationDocumentation.ModifiedNoticeRefusalAuction",
    )
    push_from_section(
        reqs,
        aliases=["trebovaniya_k_nto", "requirementsNTO", "requirements_nto"],
        pairs=[("requirementsNTOBase64", "requirementsNTOFileName")],
        dst_path="GeneralInfoAuction.ApplicationRequirements.RequirementsNTO",
    )
    push_from_section(
        prot,
        aliases=["protokol_po_rezultatam_provedeniya_torgov", "protocolConsider", "protocol_consider"],
        pairs=[("protocolFileBase64", "protocolFileName")],
        dst_path="InformationOnLots.AuctionResults[0].ProtocolConsider",
    )
    push_from_section(
        prot,
        aliases=["podpisannyy_dogovor_na_razmeschenie_nto", "signedNTO", "signed_nto"],
        pairs=[("signedNTOBase64", "signedNTOFileName")],
        dst_path="InformationOnLots.AuctionResults[0].SignedNTO",
    )
    push_from_section(
        prot,
        aliases=["schet_na_oplatu", "invoiceForPayment", "invoice_for_payment"],
        pairs=[("invoiceFileBase64", "invoiceFileName")],
        dst_path="InformationOnLots.AuctionResults[0].InvoiceForPayment[0]",
    )

    deduped: List[PendingUpload] = []
    seen = set()
    for upload in uploads:
        key = (upload.target_path, upload.filename, upload.allow_external)
        if key in seen:
            continue
        seen.add(key)
        deduped.append(upload)
    return deduped


def transform_row_to_registry(src: Any, session, logger) -> Tuple[Dict[str, Any], List[PendingUpload]]:
    block_reestr = get_any(
        src,
        ["dannye_po_reestru_mest_dlya_razmeshcheniya_nto", "dannye_po_reestru_mest_dlya_razmescheniya_nto"],
        {},
    ) or {}
    block_dogovor = get_any(
        src,
        ["informatsiya_o_dogovore_na_pravo_razmeshcheniya_nto", "informatsiya_o_dogovore_na_pravo_razmescheniya_nto"],
        {},
    ) or {}
    block = {**block_dogovor, **block_reestr}

    info_common = get_any(block, ["obschaya_informatsiya", "obshchaya_informatsiya"], {}) or {}
    info_nto = get_any(block, ["informatsiya_o_nto", "svedeniya_o_nto", "informaciya_o_nto"], {}) or {}
    info_res = block_reestr.get("svedeniya_o_postanovlenii_o_vnesenii_izmeneniy_v_shemu_nto", {}) or {}
    info_fl = block_reestr.get("informatsiya_ob_operatore_fizicheskoe_litso", {}) or {}
    info_ip = block_reestr.get("informatsiya_ob_operatore_individualnyy_predprinimatel", {}) or {}
    info_ul = block_reestr.get("informatsiya_ob_operatore_organizatsiya", {}) or {}
    info_dog = block_dogovor.get("informatsiya_po_dogovoru", {}) or {}
    info_dop = block_dogovor.get("svedeniya_o_dopolnitelnom_soglashenii", {}) or {}

    unit = build_unit_info(src) or build_unit_info(block) or build_unit_info(info_common)
    org = build_org_info(info_common, session, logger)

    general_information = {
        "Subject": norm_str(info_common.get("subekt_rf")) or None,
        "Disctrict": norm_str(info_common.get("munitsipalnyy_rayon_okrug_gorodskoy_okrug_ili_vnutrigorodskaya_territoriya")) or None,
    }

    season_arr = [mapped for mapped in (map_season(info_nto.get(key)) for key in ("1_sezon", "2_sezon", "3_sezon", "4_sezon")) if mapped]

    assort_values: List[str] = []
    other_values: List[str] = []

    def push_split(target: List[str], raw: Any) -> None:
        text = norm_str(raw)
        if not text:
            return
        for part in re.split(r"[;\r\n]+", text):
            candidate = part.strip()
            if candidate:
                target.append(candidate)

    if isinstance(info_nto.get("assortiment"), list):
        for item in info_nto.get("assortiment", []):
            push_split(assort_values, item)
    else:
        push_split(assort_values, info_nto.get("assortiment"))
    if isinstance(info_nto.get("inoy_assortiment"), list):
        for item in info_nto.get("inoy_assortiment", []):
            push_split(other_values, item)
    else:
        push_split(other_values, info_nto.get("inoy_assortiment"))

    def is_other_word(value: str) -> bool:
        return (value or "").strip().lower() in {"иной", "иное"}

    assort_block: List[Dict[str, Any]] = []
    specialization = map_specialization(info_nto.get("spetsializatsiya_nto"))
    parent_code = specialization.get("code") if specialization else None
    for value in assort_values:
        if is_other_word(value):
            if not other_values:
                assort_block.append(map_assort_item(parent_code, "Иной", None))
        else:
            assort_block.append(map_assort_item(parent_code, value, None))
    for value in other_values:
        assort_block.append(map_assort_item(parent_code, "Иной", value))

    raw_assortment = "; ".join(assort_values) if assort_values else norm_str(info_nto.get("assortiment"))
    raw_assortment_other = "; ".join(other_values) if other_values else norm_str(info_nto.get("inoy_assortiment"))

    spec_project_text = norm_str(info_nto.get("spetsifikatsiya_statusa_proektnogo_mesta"))
    spec_project_item = map_spec_project_status(spec_project_text) if spec_project_text else None
    spec_project_status = []
    if spec_project_text:
        spec_project_status.append(
            {
                "_id": None,
                "code": spec_project_item.get("code") if spec_project_item else None,
                "name": spec_project_item.get("name") if spec_project_item else spec_project_text,
                "parentId": spec_project_item.get("parentId") if spec_project_item else None,
            }
        )

    nto_information = {
        "ntoInfoID": info_nto.get("nomer_v_sheme") if info_nto.get("nomer_v_sheme") is not None else None,
        "GosuslugiData": None,
        "FullAddress": addr_obj(info_nto.get("adresnyy_orientir")),
        "infoAddress": norm_str(info_nto.get("dopolnitelnye_dannye_ob_adrese")) or None,
        "OtherInfo": norm_str(info_nto.get("dopolnitelnye_dannye_ob_adrese")) or None,
        "cadNumber": norm_str(info_nto.get("kadastrovyy_nomer")) or None,
        "suggestedWhom": norm_str(info_nto.get("kem_predlozheno_mesto")) or None,
        "GeoCoordinates": norm_str(info_nto.get("geokoordinaty_tochki_razmescheniya_nto")) or None,
        "GeoCoord1": norm_str(info_nto.get("geokoordinaty_poligona_razmescheniya_nto_tochka_1")) or None,
        "GeoCoord2": norm_str(info_nto.get("geokoordinaty_poligona_razmescheniya_nto_tochka_2")) or None,
        "GeoCoord3": norm_str(info_nto.get("geokoordinaty_poligona_razmescheniya_nto_tochka_3")) or None,
        "GeoCoord4": norm_str(info_nto.get("geokoordinaty_poligona_razmescheniya_nto_tochka_4")) or None,
        "GeoCoord5": norm_str(info_nto.get("geokoordinaty_poligona_razmescheniya_nto_tochka_5")) or None,
        "NtoType": map_nnto_type(info_nto.get("tip_nto")),
        "ntoTypeOther": norm_str(info_nto.get("inoy_tip_nto")) or None,
        "Specialization": specialization,
        "ntoInfoProductType": map_product_type(info_nto.get("vid_produktsii")),
        "infoSMSP": map_info_smsp(info_nto.get("informatsiya_ob_ispolzovanii_nto_subektami_malogo_i_srednego_predprinimatelstva")),
        "ntoInfoLandArea": info_nto.get("ploschad_uchastka_kv_m") if info_nto.get("ploschad_uchastka_kv_m") is not None else None,
        "ntoInfoObjectArea": info_nto.get("ploschad_obekta_kv_m") if info_nto.get("ploschad_obekta_kv_m") is not None else None,
        "ntoInfoLandOwnershipType": map_ownership_type(info_nto.get("vid_sobstvennosti_zemelnogo_uchastka")),
        "PlacementSchedule": map_placement_schedule(info_nto.get("grafik_razmescheniya")),
        "seasonPlacementSchedule": season_arr,
        "TradingData": norm_str(info_nto.get("svedeniya_o_torgovoy_protsedure")) or None,
        "ProjectStatus": map_project_status(info_nto.get("status_proektnogo_mesta")),
        "SpecProjectStatus": spec_project_status,
        "ntoInfoRecordStatus": map_record_status(info_nto.get("status_zapisi_v_reestre_mest_dlya_razmescheniya_nto")),
        "ntoInfoSchemeNumber": info_nto.get("nomer_v_sheme") if info_nto.get("nomer_v_sheme") is not None else None,
        "AssortBlock": assort_block,
        "ntoInfoSpecialUseZones": map_special_use_zone(info_nto.get("raspolozhenie_territorii_v_zonah_s_osobymi_usloviyami_ispolzovaniya")),
        "ntoInfoPlacementStartDate": as_date_or_null(info_nto.get("period_razmescheniya_data_nachala")),
        "ntoInfoPlacementEndDate": as_date_or_null(info_nto.get("period_razmescheniya_data_okonchaniya")),
        "Scheduled": norm_str(info_nto.get("grafik")) or None,
        "fileEnginTopogPlaceNew": None,
        "photoSituationNTONew": None,
        "rawAssortment": raw_assortment,
        "rawAssortmentOther": raw_assortment_other,
    }

    resolution_block = {
        "npaAmendmentNumber": norm_str(info_res.get("nomer_postanovleniya")),
        "npaAmendmentDate": as_date_or_null(info_res.get("data_postanovleniya")),
        "statusOperator": norm_str(info_res.get("pravovoy_status_operatora")),
        "npaAmendmentFile": None,
    }

    operator_information_fl = {
        "OperatorName": norm_str(info_fl.get("familiya_imya_otchestvo")),
        "OperatorBirthday": as_date_or_null(info_fl.get("data_rozhdeniya")),
        "OperatorSnils": norm_str(info_fl.get("snils")),
        "OperatorPassportSeries": norm_str(info_fl.get("seriya_pasporta")),
        "OperatorPassportNumber": norm_str(info_fl.get("nomer_pasporta")),
        "OperatorPassportWhom": norm_str(info_fl.get("kem_vydan_pasport")),
        "OperatorPassportDepartmentCode": norm_str(info_fl.get("kod_podrazdeleniya")),
        "OperatorPlaceBirth": norm_str(info_fl.get("mesto_rozhdeniya")),
        "OperatorNumber": norm_str(info_fl.get("nomer_telefona")),
        "OperatorEmail": norm_str(info_fl.get("elektronnaya_pochta")),
        "OperatorPermanentAddress": addr_obj(info_fl.get("adres_postoyannoy_registratsii")),
        "OperatorFactAddress": addr_obj(info_fl.get("adres_fakticheskogo_mesta_prozhivaniya")),
    }

    operator_information_ip = {
        "OperatorName": norm_str(info_ip.get("naimenovanie_ip")),
        "OperatorOGRNIP": norm_str(info_ip.get("ogrn_ip")),
        "OperatorINN": norm_str(info_ip.get("inn_ip")),
        "OperatorNumber": norm_str(info_ip.get("nomer_telefona_2")),
        "OperatorEmail": norm_str(info_ip.get("elektronnaya_pochta_2")),
        "OperatorPermanentAddress": addr_obj(info_ip.get("adres_postoyannoy_registratsii_2")),
        "OperatorFactAddress": addr_obj(info_ip.get("pochtovyy_adres")),
    }

    operator_information_ul = {
        "OperatorName": norm_str(info_ul.get("polnoe_naimenovanie_organizatsii")),
        "OperatorOGRN": norm_str(info_ul.get("ogrn")),
        "OperatorINN": norm_str(info_ul.get("inn")),
        "OperatorKPP": norm_str(info_ul.get("kpp")),
        "OperatorNumber": norm_str(info_ul.get("nomer_telefona_3")),
        "OperatorEmail": norm_str(info_ul.get("elektronnaya_pochta_3")),
        "OperatorUrAddress": addr_obj(info_ul.get("yuridicheskiy_adres")),
        "OperatorFactAddress": addr_obj(info_ul.get("fakticheskiy_adres")),
        "OrgStateForm": map_org_state_form(info_ul.get("organizatsionno_pravovaya_forma")),
    }

    operator_list = None
    if operator_information_ul.get("OperatorName"):
        operator_list = "Юридическое лицо"
    elif operator_information_ip.get("OperatorName"):
        operator_list = "Индивидуальный предприниматель"
    elif operator_information_fl.get("OperatorName"):
        operator_list = "Физическое лицо"

    nto_contract = {
        "ContractNumber": norm_str(info_dog.get("nomer_dogovora")),
        "ContractStartDate": as_date_or_null(info_dog.get("data_zaklyucheniya_dogovora")),
        "ContractPeriodStartDate": as_date_or_null(info_dog.get("data_nachala_deystviya_dogovora")),
        "ContractEndDate": as_date_or_null(info_dog.get("data_zaversheniya_deystviya_dogovora")),
        "ContractSubject": norm_str(info_dog.get("predmet_dogovora")),
        "TotalContractFee": info_dog.get("obschiy_razmer_platy_po_dogovoru_rub")
        if info_dog.get("obschiy_razmer_platy_po_dogovoru_rub") is not None
        else None,
        "PeriodicContractFee": info_dog.get("razmer_platy_za_period_rub")
        if info_dog.get("razmer_platy_za_period_rub") is not None
        else None,
        "PaymentPeriod": map_payment_period(info_dog.get("periodichnost_vneseniya_platezhey")),
        "NTOPPlacementAgreement": None,
    }

    block_dop_sogl: List[Dict[str, Any]] = []
    if any(
        info_dop.get(key) is not None
        for key in (
            "nomer_dopolnitelnogo_soglasheniya",
            "data_zaklyucheniya_dopolnitelnogo_soglasheniya",
            "prichina_zaklyucheniya_dopolnitelnogo_soglasheniya",
        )
    ):
        block_dop_sogl.append(
            {
                "AdditionalAgreementNumber": norm_str(info_dop.get("nomer_dopolnitelnogo_soglasheniya")),
                "AdditionalAgreementDate": as_date_or_null(info_dop.get("data_zaklyucheniya_dopolnitelnogo_soglasheniya")),
                "ReasonContract": norm_str(info_dop.get("prichina_zaklyucheniya_dopolnitelnogo_soglasheniya")),
                "SupplementalAgreementIfAny": None,
            }
        )

    payload = {
        "guid": generate_guid(),
        "unit": {
            "id": (unit.get("_id") or unit.get("id")) if unit else DEFAULT_ORG["_id"],
            "name": unit.get("name") if unit else DEFAULT_ORG["name"],
            "shortName": (unit.get("shortName") or unit.get("name")) if unit else DEFAULT_ORG["shortName"],
        },
        "parentEntries": "NTOmesto",
        "generalInformation": general_information,
        "ntoInformation": nto_information,
        "organization": org,
        "approvalDocuments": [],
        "resolutionBlock": resolution_block,
        "operatorInformationFL": operator_information_fl,
        "operatorInformationIP": operator_information_ip,
        "operatorInformationUL": operator_information_ul,
        "OperatorList": operator_list,
        "NTOContract": nto_contract,
        "blockDopSogl": block_dop_sogl,
    }
    return payload, collect_pending_uploads_nto(src)


def _to_nsi_str(value: Any) -> str:
    if value is None:
        return "-"
    text = str(value).strip()
    return text if text else "-"


def _iso_to_ru_date(value: Any) -> str:
    iso_value = as_date_or_null(value)
    if not iso_value:
        return "-"
    try:
        return datetime.datetime.strptime(iso_value, "%Y-%m-%d").strftime("%d.%m.%Y")
    except ValueError:
        return "-"


def build_nsi_local_object_nto_payload(nto_doc: Dict[str, Any]) -> Dict[str, Any]:
    guid = nto_doc.get("guid") or generate_guid()
    general = nto_doc.get("generalInformation") or {}
    nto_info = nto_doc.get("ntoInformation") or {}
    operator_fl = nto_doc.get("operatorInformationFL") or {}
    operator_ip = nto_doc.get("operatorInformationIP") or {}
    operator_ul = nto_doc.get("operatorInformationUL") or {}
    contract = nto_doc.get("NTOContract") or {}

    def is_other(name: str) -> bool:
        return bool(name and name.lower() in {"иной", "иное"})

    specialization_name = norm_str(get_any(nto_info, ["Specialization.name", "Specialization"]))
    assort_block = nto_info.get("AssortBlock") or []
    specialization_assortment: List[str] = []
    if isinstance(assort_block, list):
        for item in assort_block:
            name = norm_str(get_any(item, ["Assort.name", "Assort"]))
            if not name:
                continue
            if is_other(name):
                specialization_assortment.append(norm_str(item.get("assortOther")) or name)
            else:
                specialization_assortment.append(name)
    if not specialization_assortment:
        raw_assortment = norm_str(nto_info.get("rawAssortment")) or norm_str(nto_info.get("rawAssortmentOther"))
        if raw_assortment:
            specialization_assortment = [raw_assortment]

    if specialization_name and specialization_assortment:
        specialization_combined = f"{specialization_name}: {'; '.join(specialization_assortment)}"
    else:
        specialization_combined = specialization_name or "; ".join(specialization_assortment)

    placement_schedule_name = norm_str(get_any(nto_info, ["PlacementSchedule.name", "PlacementSchedule"]))
    placement_schedule_combined = placement_schedule_name
    if placement_schedule_name.lower() == "по графику":
        schedule = get_any(nto_info, ["Scheduled", "Scheduled.name"])
        placement_schedule_combined = f"{placement_schedule_name}: {norm_str(schedule)}" if schedule else placement_schedule_name
    elif placement_schedule_name.lower() == "сезонно":
        seasons = get_any(nto_info, ["seasonPlacementSchedule", "SeasonPlacementSchedule"])
        if isinstance(seasons, list):
            season_names: List[str] = []
            for item in seasons:
                name = norm_str(item.get("name")) if isinstance(item, dict) else norm_str(item)
                if name and name not in season_names:
                    season_names.append(name)
            placement_schedule_combined = f"{placement_schedule_name}: {', '.join(season_names)}" if season_names else placement_schedule_name

    project_status_name = norm_str(get_any(nto_info, ["ProjectStatus.name", "ProjectStatus"]))
    spec_project_status = nto_info.get("SpecProjectStatus") or []
    spec_proj_status_name = ""
    if isinstance(spec_project_status, list) and spec_project_status:
        item = spec_project_status[0]
        if isinstance(item, dict):
            spec_proj_status_name = norm_str(item.get("name"))
    project_status_combined = (
        f"{project_status_name}: {spec_proj_status_name}"
        if project_status_name and spec_proj_status_name
        else project_status_name or spec_proj_status_name
    )

    operator_name = ""
    operator_inn = ""
    operator_ogrn = ""
    operator_number = ""
    operator_email = ""
    if operator_ul.get("OperatorName"):
        operator_name = norm_str(operator_ul.get("OperatorName"))
        operator_inn = norm_str(operator_ul.get("OperatorINN"))
        operator_ogrn = norm_str(operator_ul.get("OperatorOGRN"))
        operator_number = norm_str(operator_ul.get("OperatorNumber"))
        operator_email = norm_str(operator_ul.get("OperatorEmail"))
    elif operator_ip.get("OperatorName"):
        operator_name = norm_str(operator_ip.get("OperatorName"))
        operator_inn = "-"
        operator_ogrn = norm_str(operator_ip.get("OperatorOGRNIP"))
        operator_number = "-"
        operator_email = "-"
    elif operator_fl.get("OperatorName"):
        operator_name = "Физическое лицо"
        operator_inn = "-"
        operator_ogrn = "-"
        operator_number = "-"
        operator_email = "-"

    unit_value = nto_doc.get("unit") or {}
    unit_id = unit_value.get("id") or unit_value.get("_id")
    location = get_any(nto_info, ["FullAddress.fullAddress", "FullAddress"]) or ""
    if isinstance(location, dict):
        location = get_any(location, ["fullAddress"]) or ""

    return {
        "guid": guid,
        "LayerId": "1",
        "Layer": "Нестационарные торговые объекты",
        "Subject": _to_nsi_str(get_any(general, ["Subject.name", "Subject"])),
        "Disctrict": _to_nsi_str(get_any(general, ["Disctrict.name", "Disctrict"])),
        "NtoType": _to_nsi_str(get_any(nto_info, ["NtoType.name", "NtoType"]) or norm_str(nto_info.get("ntoTypeOther"))),
        "Specialization": _to_nsi_str(specialization_combined),
        "NumberNTO": _to_nsi_str(nto_info.get("ntoInfoID") or ""),
        "GeoCoordinates": _to_nsi_str(get_any(nto_info, ["GeoCoordinates", "GeoCoordinates.name"])),
        "CadNumber": _to_nsi_str(get_any(nto_info, ["cadNumber", "CadastralNumber"])),
        "FullAddress": _to_nsi_str(location),
        "ProjectStatus": _to_nsi_str(project_status_combined),
        "PlacementSchedule": _to_nsi_str(placement_schedule_combined),
        "TradingData": _to_nsi_str(get_any(nto_info, ["TradingData", "TradingData.name"])),
        "ContractNumber": _to_nsi_str(contract.get("ContractNumber")),
        "ContractStartDate": _iso_to_ru_date(contract.get("ContractStartDate")),
        "ContractEndDate": _iso_to_ru_date(contract.get("ContractEndDate")),
        "OperatorName": _to_nsi_str(operator_name),
        "OperatorINN": _to_nsi_str(operator_inn),
        "OperatorOGRN": _to_nsi_str(operator_ogrn),
        "OperatorNumber": _to_nsi_str(operator_number),
        "OperatorEmail": _to_nsi_str(operator_email),
        "OtherInfo": _to_nsi_str(get_any(nto_info, ["OtherInfo", "OtherInfo.name"])),
        "dictionaryType": "local",
        "dictionaryUnitId": _to_nsi_str(unit_id),
        "autokey": _to_nsi_str(guid),
        "code": _to_nsi_str(guid),
        "ObjectID": _to_nsi_str(guid),
        "parentEntries": "nsiLocalObjectNTO",
    }


def _make_search_org_result(src: Any, session, logger) -> Optional[Dict[str, Any]]:
    ogrn = get_any(
        src,
        [
            "dannye_po_reestru_torgov.obschaya_informatsiya.ogrn_upolnomochennogo_organa",
            "dannye_po_reestru_torgov.obschaya_informatsiya.orgn",
            "ogrn",
        ],
        None,
    )
    return search_org_by_ogrn(session, logger, ogrn)


def _resolve_bidding_unit(src: Any, session, logger) -> Dict[str, Any]:
    org = _make_search_org_result(src, session, logger)
    if org:
        return {
            "id": org.get("_id") or org.get("id") or DEFAULT_ORG["_id"],
            "name": org.get("name") or DEFAULT_ORG["name"],
            "shortName": org.get("shortName") or org.get("name") or DEFAULT_ORG["shortName"],
        }
    return {"id": DEFAULT_ORG["_id"], "name": DEFAULT_ORG["name"], "shortName": DEFAULT_ORG["shortName"]}


def transform_row_to_bidding(src: Any, session, logger) -> Tuple[Dict[str, Any], List[PendingUpload]]:
    block = get_path(src, "dannye_po_reestru_torgov") or {}
    ob_inf = get_path(block, "obschaya_informatsiya") or {}
    init_inf = get_path(block, "initsiator_torgovoy_protsedury") or {}
    org_inf = get_path(block, "organizator_torgovoy_protsedury") or {}
    notif = get_path(block, "izveschenie_i_dopolnitelnaya_dokumentatsiya") or {}
    proc = get_path(block, "svedeniya_o_protsedure_provedeniya_torgov") or {}
    reqs = get_path(block, "trebovaniya_k_zayavkam") or {}
    lots = get_path(block, "informatsiya_o_lotah_dlya_vystavleniya_na_torgi") or {}
    winner = get_path(block, "pobeditel") or {}
    prot = get_path(block, "protokol_po_rezultatam_provedeniya_torgov") or {}

    for obj in (ob_inf, init_inf, org_inf, notif, proc, reqs, lots, winner, prot):
        coerce_string_ids(obj)

    general_information = {
        "Subject": norm_str(ob_inf.get("subekt_rf")),
        "Disctrict": norm_str(ob_inf.get("munitsipalnyy_rayon_okrug_gorodskoy_okrug_ili_vnutrigorodskaya_territoriya")),
    }

    general_info_auction = {
        "BiddingInitiator": {
            "NameInitiator": norm_str(init_inf.get("polnoe_naimenovanie")),
            "UrInitAddress": addr_obj(init_inf.get("yuridicheskiy_adres")),
            "FactInitAddress": addr_obj(init_inf.get("fakticheskiy_adres")),
            "EmailInit": norm_str(init_inf.get("elektronnaya_pochta")),
            "PhoneInit": norm_str(init_inf.get("kontaktnyy_telefon")),
            "ContactPersonInit": norm_str(init_inf.get("kontaktnoe_litso")),
        },
        "OrgBidding": {
            "NameOrganizer": norm_str(org_inf.get("polnoe_naimenovanie_organizatsii")),
            "UrOrgAddress": addr_obj(org_inf.get("yuridicheskiy_adres_organizatora")),
            "FactOrgAdress": addr_obj(org_inf.get("fakticheskiy_adres_organizatora")),
            "EmailOrg": norm_str(org_inf.get("adres_elektronnoy_pochty_organizatora")),
            "PhoneOrg": norm_str(org_inf.get("kontaktnyy_telefon_organizatora")),
            "ContactPersonOrg": norm_str(org_inf.get("kontaktnoe_litso_2")),
        },
        "NotificationDocumentation": {
            "NoticeType": map_notice_type(notif.get("tip_izvescheniya")),
            "BiddingForm": map_bidding_form(notif.get("forma_provedeniya_torgov")),
            "NoticeSignedNumber": norm_str(notif.get("nomer_izvescheniya")),
            "linkTorg": norm_str(notif.get("ssylka_na_torgovuyu_protseduru_url")),
            "AgreedNotice": None,
            "ProjectNotice": None,
            "ProjectNoticeNTO": None,
            "ModifiedNoticeRefusalAuction": None,
        },
        "ProcedureInfo": {
            "BasisBidding": norm_str(proc.get("osnovanie_dlya_provedeniya_torgov")),
            "StartOfferDate": as_date_or_null(proc.get("data_nachala_priema_zayavok")),
            "StartOfferTime": excel_time_to_hhmm(proc.get("vremya_nachala_priema_zayavok")),
            "EndOfferDate": as_date_or_null(proc.get("data_okonchaniya_priema_zayavok")),
            "EndOfferTime": excel_time_to_hhmm(proc.get("vremya_okonchaniya_priema_zayavok")),
            "StartReviewDate": as_date_or_null(proc.get("data_nachala_rassmotreniya_zayavok")),
            "StartReviewTime": excel_time_to_hhmm(proc.get("vremya_nachala_rassmotreniya_zayavok")),
            "EndReviewDate": as_date_or_null(proc.get("data_okonchaniya_rassmotreniya_zayavok")),
            "EndReviewTime": excel_time_to_hhmm(proc.get("vremya_okonchaniya_rassmotreniya_zayavok")),
            "BiddingDate": as_date_or_null(proc.get("data_provedeniya_torgov")),
            "BiddingTime": excel_time_to_hhmm(proc.get("vremya_provedeniya_torgov")),
            "ItogDate": as_date_or_null(proc.get("data_podvedeniya_itogov")),
            "ItogTime": excel_time_to_hhmm(proc.get("vremya_podvedeniya_itogov")),
            "CancelDate": as_date_or_null(
                get_any(
                    proc,
                    [
                        "data_do_kotoroy_organizator_vprave_otkazatsya_ot_provedeniya_torgov",
                        "data_do_kotoroy_organizator_vprave_otkazatsya_ot_provedenija_torgov",
                    ],
                    None,
                )
            ),
            "OfferPlaceE": norm_str(proc.get("naimenovanie_operatora_elektronnoy_ploschadki")),
            "BiddingPlaceE": norm_str(proc.get("elektronnaya_ploschadka_provedeniya_torgov_url")),
            "TermAndProcedureContract": norm_str(proc.get("srok_i_poryadok_zaklyucheniya_dogovora")),
        },
        "ApplicationRequirements": {
            "RequirementsMembers": norm_str(reqs.get("trebovaniya_k_uchastnikam")),
            "RequirementsDocs": norm_str(reqs.get("trebovaniya_k_dokumentam")),
            "CriteriaWinner": norm_str(reqs.get("kriterii_dlya_opredeleniya_pobeditelya")),
            "RequirementsNTO": None,
        },
    }

    lot_info: List[Dict[str, Any]] = []
    direct_lot = any(lots.get(field) is not None for field in ("nomer_nto_v_sheme", "adres", "geokoordinaty", "tip_nto", "spetsializatsiya"))
    if direct_lot:
        lot_info.append(
            {
                "LotName": norm_str(lots.get("naimenovanie_lota")),
                "InfoLot": norm_str(lots.get("opisanie_lota")),
                "NTOSchemeNumber": norm_str(lots.get("nomer_nto_v_sheme")),
                "LotAddress": addr_obj(lots.get("adres")),
                "GeoCoordinates": norm_str(lots.get("geokoordinaty")),
                "TypeNTO": map_nto_type_bidding(lots.get("tip_nto")),
                "ntoTypeOther": norm_str(lots.get("inoy_tip_nto")),
                "Specific": map_specific(lots.get("spetsializatsiya")),
                "ObjectRange": to_number_or_null(lots.get("ploschad_obekta_kv_m")),
                "LandRange": to_number_or_null(lots.get("ploschad_zemelnogo_uchastka_kv_m")),
                "PlacementScheduleType": map_placement_schedule(lots.get("grafik_razmescheniya")),
                "StartDate": as_date_or_null(lots.get("period_razmescheniya_data_nachala")),
                "EndDate": as_date_or_null(lots.get("period_razmescheniya_data_okonchaniya")),
                "StartPrice": to_number_or_null(lots.get("nachalnaya_minimalnaya_tsena_rub")),
                "StepAuction": to_number_or_null(lots.get("shag_auktsiona_rub_esli_tip_izvescheniya_auktsion")),
                "DepositSize": to_number_or_null(lots.get("razmer_zadatka_rub")),
                "nds": as_bool_yes_ru(lots.get("nalichie_nds")),
                "ndsSize": to_number_or_null(lots.get("stavka_nds")),
            }
        )
    else:
        for i in range(1, 101):
            base = f"{i}_"
            any_lot = any(lots.get(f"{base}{field}") is not None for field in ("nomer_nto_v_sheme", "adres", "geokoordinaty", "tip_nto", "spetsializatsiya"))
            if not any_lot:
                continue
            lot_info.append(
                {
                    "LotName": norm_str(lots.get(f"{base}naimenovanie_lota")),
                    "InfoLot": norm_str(lots.get(f"{base}opisanie_lota")),
                    "NTOSchemeNumber": norm_str(lots.get(f"{base}nomer_nto_v_sheme")),
                    "LotAddress": addr_obj(lots.get(f"{base}adres")),
                    "GeoCoordinates": norm_str(lots.get(f"{base}geokoordinaty")),
                    "TypeNTO": map_nto_type_bidding(lots.get(f"{base}tip_nto")),
                    "OtherType": norm_str(lots.get(f"{base}inoy_tip_nto")),
                    "Specific": map_specific(lots.get(f"{base}spetsializatsiya")),
                    "ObjectRange": to_number_or_null(lots.get(f"{base}ploschad_obekta_kv_m")),
                    "LandRange": to_number_or_null(lots.get(f"{base}ploschad_zemelnogo_uchastka_kv_m")),
                    "PlacementScheduleType": map_placement_schedule(lots.get(f"{base}grafik_razmescheniya")),
                    "StartDate": as_date_or_null(lots.get(f"{base}period_razmescheniya_data_nachala")),
                    "EndDate": as_date_or_null(lots.get(f"{base}period_razmescheniya_data_okonchaniya")),
                    "StartPrice": to_number_or_null(lots.get(f"{base}nachalnaya_minimalnaya_tsena_rub")),
                    "StepAuction": to_number_or_null(lots.get(f"{base}shag_auktsiona_rub_esli_tip_izvescheniya_auktsion")),
                    "DepositSize": to_number_or_null(lots.get(f"{base}razmer_zadatka_rub")),
                    "nds": as_bool_yes_ru(lots.get(f"{base}nalichie_nds")),
                    "ndsSize": to_number_or_null(lots.get(f"{base}stavka_nds")),
                }
            )

    auction_results: List[Dict[str, Any]] = []
    status_global = norm_str(lots.get("status_provedeniya_torgov"))
    if status_global:
        auction_results.append(
            {
                "NumberNTO": norm_str(lots.get("nomer_nto_v_sheme")),
                "BiddingStatus": map_bidding_status(status_global),
                "winner": {
                    "name": norm_str(winner.get("naimenovanie_pobeditelya")),
                    "PriceOffer": to_number_or_null(winner.get("predlozhenie_o_tsene_rub")),
                },
                "ProtocolConsiderNumber": norm_str(prot.get("nomer_protokola")),
                "ProtocolConsiderDate": as_date_or_null(prot.get("data_protokola")),
                "participants": [],
                "ProtocolConsider": None,
                "SignedNTO": None,
                "InvoiceForPayment": [],
            }
        )

    payload = {
        "guid": generate_guid(),
        "parentEntries": "reestrbiddingReestr",
        "unit": _resolve_bidding_unit(src, session, logger),
        "generalInformation": general_information,
        "GeneralInfoAuction": general_info_auction,
        "InformationOnLots": {"LotInfo": lot_info, "AuctionResults": auction_results},
    }
    return payload, collect_pending_uploads_torgov(src)
