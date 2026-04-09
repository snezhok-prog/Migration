import json
import re
import uuid
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from _config import REGION_CODE


def _fix_mojibake_cp1251_utf8(value: str) -> str:
    """
    Repair common mojibake where UTF-8 text was decoded as CP1251.
    Example: "РќРµ Р·РЅР°СЋ" -> "Не знаю".
    """
    text = str(value or "")
    for _ in range(2):
        try:
            repaired = text.encode("cp1251").decode("utf-8")
        except Exception:
            break
        if repaired == text:
            break
        text = repaired
    return text


def nz(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        s = _fix_mojibake_cp1251_utf8(value).strip()
    else:
        s = str(value).strip()
    if s.lower() in {"nan", "none"}:
        return ""
    return s


def as_string_or_null(value: Any):
    s = nz(value)
    return s if s else None


def norm_ru(value: Any) -> str:
    return re.sub(r"\s+", " ", nz(value).lower().replace("ё", "е")).strip()


def fix_mojibake_deep(value: Any):
    if isinstance(value, str):
        return _fix_mojibake_cp1251_utf8(value)
    if isinstance(value, list):
        return [fix_mojibake_deep(v) for v in value]
    if isinstance(value, dict):
        return {k: fix_mojibake_deep(v) for k, v in value.items()}
    return value

def safe_json(text: str):
    try:
        return json.loads(text)
    except Exception:
        return text


def generate_guid() -> str:
    return str(uuid.uuid4())


TZ = "+0300"


def _excel_serial_to_date(serial_str: str):
    try:
        value = float(serial_str.replace(",", "."))
    except Exception:
        return None
    if value < 20000 or value > 80000:
        return None
    base = datetime(1899, 12, 30)
    dt = base + timedelta(days=value)
    return dt.strftime("%Y-%m-%d")


def to_iso_z(date_str: Any):
    s = nz(date_str)
    if not s:
        return None

    # Already in ISO with timezone, keep as-is.
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}T.*[+-]\d{4}", s):
        return s

    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        date_part = s
    elif re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", s):
        dd, mm, yyyy = s.split(".")
        date_part = f"{yyyy}-{mm}-{dd}"
    else:
        date_part = _excel_serial_to_date(s)
        if not date_part:
            return None

    return f"{date_part}T00:00:00.000{TZ}"


def to_iso_z_datetime(date_str: Any, time_str: Any):
    base = to_iso_z(date_str)
    if not base:
        return None

    t = nz(time_str)
    if not t:
        return base

    m = re.fullmatch(r"(\d{1,2}):(\d{2})", t)
    if not m:
        return base

    hh = m.group(1).zfill(2)
    mm = m.group(2)
    return base.replace("T00:00:00.000", f"T{hh}:{mm}:00.000")


def to_millis_safe(value: Any) -> int:
    s = nz(value)
    if not s:
        return 0
    try:
        if s.endswith("Z"):
            dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
        else:
            dt = datetime.fromisoformat(s)
        return int(dt.timestamp() * 1000)
    except Exception:
        return 0


def parse_postal_code(full_address: Any) -> str:
    m = re.search(r"\b\d{6}\b", nz(full_address))
    return m.group(0) if m else ""


def base64_size_bytes(b64: Any) -> int:
    s = re.sub(r"\s+", "", nz(b64))
    eq = 2 if s.endswith("==") else 1 if s.endswith("=") else 0
    return max(0, (len(s) * 3 // 4) - eq)


def build_address(full_address: Any):
    s = nz(full_address)
    address = {
        "okato": "",
        "oktmo": "",
        "ifnsfl": "",
        "ifnsul": "",
        "country": "",
        "isNotFias": False,
        "postalCode": parse_postal_code(s),
        "regionCode": "",
        "fullAddress": s,
        "addressParts": [],
        "addressAsObject": None,
        "cadastralNumber": None,
        "isSpecialAddress": False,
        "unrecognizablePart": "",
    }

    norm_full = norm_ru(s)
    for key, code in REGION_CODE.items():
        if norm_ru(key) in norm_full:
            address["regionCode"] = code
            break
    return address


def build_minimal_address(full_address: Any, region_name: Any = None):
    s = nz(full_address)
    if not s:
        return None

    code = None
    if region_name:
        normalized_region = norm_ru(region_name)
        for key, value in REGION_CODE.items():
            if norm_ru(key) == normalized_region:
                code = value
                break

    return {
        "okato": None,
        "oktmo": None,
        "ifnsfl": None,
        "ifnsul": None,
        "country": "",
        "isNotFias": False,
        "postalCode": parse_postal_code(s) or None,
        "regionCode": code,
        "fullAddress": s,
        "cadastralNumber": None,
        "isSpecialAddress": False,
        "unrecognizablePart": "",
    }


_SEGMENT_RE = re.compile(r"^([^\[]+)((\[\d+\])*)$")
_INDEX_RE = re.compile(r"\[(\d+)\]")


def _path_to_parts(path: str):
    parts = []
    for segment in str(path).split("."):
        m = _SEGMENT_RE.match(segment)
        if not m:
            parts.append(segment)
            continue

        prop = m.group(1)
        idxs = m.group(2) or ""
        parts.append(prop)
        for mm in _INDEX_RE.finditer(idxs):
            parts.append(int(mm.group(1)))
    return parts


def set_by_path(obj: dict, path: str, value: Any):
    parts = _path_to_parts(path)
    cur = obj
    i = 0
    while i < len(parts):
        is_last = i == len(parts) - 1
        key = parts[i]
        next_key = parts[i + 1] if i + 1 < len(parts) else None

        if isinstance(key, str):
            if isinstance(next_key, int):
                if not isinstance(cur.get(key), list):
                    cur[key] = []

                arr = cur[key]
                while len(arr) <= next_key:
                    arr.append(None)

                if i + 1 == len(parts) - 1:
                    arr[next_key] = value
                    return

                if not isinstance(arr[next_key], dict):
                    arr[next_key] = {}
                cur = arr[next_key]
                i += 2
                continue

            if is_last:
                cur[key] = value
                return

            if not isinstance(cur.get(key), dict):
                cur[key] = {}
            cur = cur[key]
            i += 1
            continue

        if not isinstance(cur, list):
            raise ValueError("Expected list on path segment")

        while len(cur) <= key:
            cur.append(None)

        if is_last:
            cur[key] = value
            return

        if not isinstance(cur[key], dict):
            cur[key] = {}
        cur = cur[key]
        i += 1


def get_by_path(obj: Any, path: str):
    parts = _path_to_parts(path)
    cur = obj
    for key in parts:
        if isinstance(key, str):
            if not isinstance(cur, dict):
                return None
            cur = cur.get(key)
            continue
        if not isinstance(cur, list):
            return None
        if key < 0 or key >= len(cur):
            return None
        cur = cur[key]
    return cur


def make_boundary() -> str:
    return "----WebKitFormBoundary" + uuid.uuid4().hex


def build_multipart_body(boundary: str, filename: str, fields: dict, base64_content: str) -> str:
    crlf = "\r\n"
    chunks = []
    chunks.append(f"--{boundary}")
    chunks.append(f'Content-Disposition: form-data; name="file"; filename="{filename}"')
    chunks.append("Content-Type: application/octet-stream")
    chunks.append("Content-Transfer-Encoding: base64")
    chunks.append("")
    chunks.append(base64_content)

    for k, v in fields.items():
        chunks.append(f"--{boundary}")
        chunks.append(f'Content-Disposition: form-data; name="{k}"')
        chunks.append("")
        chunks.append("" if v is None else str(v))

    chunks.append(f"--{boundary}--")
    chunks.append("")
    return crlf.join(chunks)


def read_rows_json(path: str):
    fp = Path(path)
    if not fp.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    # utf-8-sig transparently strips BOM from files produced by some tools/macros
    raw = fp.read_text(encoding="utf-8-sig").strip()
    if not raw:
        return []
    data = json.loads(raw)
    if isinstance(data, dict):
        return [data]
    if not isinstance(data, list):
        raise ValueError("Input JSON must be an array or object")
    return data


def jsonable(obj):
    try:
        import numpy as np

        np_int = (np.integer,)
        np_float = (np.floating,)
        np_bool = (np.bool_,)
    except Exception:
        np = None
        np_int = tuple()
        np_float = tuple()
        np_bool = tuple()

    if obj is None:
        return None
    if isinstance(obj, datetime):
        return obj.isoformat()
    if np is not None and isinstance(obj, np_int):
        return int(obj)
    if np is not None and isinstance(obj, np_bool):
        return bool(obj)
    if np is not None and isinstance(obj, np_float):
        v = float(obj)
        return None if v != v else v
    if isinstance(obj, float) and obj != obj:
        return None
    if isinstance(obj, (str, int, float, bool)):
        return obj
    if isinstance(obj, list):
        return [jsonable(x) for x in obj]
    if isinstance(obj, dict):
        return {str(k): jsonable(v) for k, v in obj.items()}
    return str(obj)


