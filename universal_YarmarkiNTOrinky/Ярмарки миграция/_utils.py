import os
import re
import base64
import json
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional
from datetime import datetime, timezone, timedelta

import pandas as pd

from _config import SUPPORTED_EXTENSIONS


def nz(v) -> str:
    """Любое значение -> строка без NaN/None."""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip()


def split_sc(v):
    """Безопасный split по ';'."""
    s = nz(v)
    if not s:
        return []
    return [p.strip() for p in s.split(";") if p.strip()]


def jsonable(obj):
    """
    Делает объект JSON-сериализуемым:
    - numpy.int64/float64/bool_ -> обычные int/float/bool
    - NaN -> None
    - Timestamp/datetime -> строка ISO
    - dict/list -> рекурсивно
    """
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

    if isinstance(obj, pd.Timestamp):
        return obj.isoformat()
    if isinstance(obj, datetime):
        return obj.isoformat()

    if np is not None and isinstance(obj, np_int):
        return int(obj)
    if np is not None and isinstance(obj, np_bool):
        return bool(obj)
    if np is not None and isinstance(obj, np_float):
        val = float(obj)
        return None if (val != val) else val

    if isinstance(obj, (str, int, float, bool)):
        if isinstance(obj, float) and obj != obj:
            return None
        return obj

    if isinstance(obj, list):
        return [jsonable(x) for x in obj]

    if isinstance(obj, dict):
        return {str(k): jsonable(v) for k, v in obj.items()}

    return str(obj)


def generate_guid() -> str:
    return str(uuid.uuid4())


def to_iso_date(date_val):
    s = nz(date_val)
    if not s:
        return None
    try:
        dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
        if pd.isna(dt):
            return None
        return dt.strftime("%Y-%m-%dT00:00:00.000+0300")
    except Exception:
        return None


def read_excel(file_path, skiprows=4, sheet_name=None):
    """Читает Excel файл и возвращает DataFrame."""
    try:
        df = pd.read_excel(file_path, dtype=str, na_filter=False, skiprows=skiprows, sheet_name=sheet_name)
        return df
    except FileNotFoundError:
        print(f"Файл {file_path} не найден")
        return None
    except Exception as e:
        print(f"Ошибка чтения: {e}")
        return None


def parse_date_to_birthday_obj(date_str):
    s = nz(date_str)
    if not s:
        return None
    try:
        dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
        if pd.isna(dt):
            return None
        dt_local = dt.to_pydatetime().replace(tzinfo=timezone(timedelta(hours=3)))
    except Exception:
        return None

    year, month, day = dt_local.year, dt_local.month, dt_local.day
    formatted = f"{day:02d}.{month:02d}.{year}"
    dt_utc = dt_local.astimezone(timezone.utc)
    jsDate = dt_utc.strftime("%Y-%m-%dT%H:%M:%S.000Z")
    epoc = int(dt_utc.timestamp())

    return {
        "date": {"year": year, "month": month, "day": day},
        "jsDate": jsDate,
        "formatted": formatted,
        "epoc": epoc
    }


def format_phone(phone_str):
    if not phone_str:
        return ""

    digits = re.sub(r'\D', '', str(phone_str))

    if len(digits) == 11 and digits[0] in ('7', '8'):
        digits = '7' + digits[1:]
    elif len(digits) == 10:
        digits = '7' + digits
    else:
        return phone_str

    return f"+{digits[0]} ({digits[1:4]}) {digits[4:7]} {digits[7:9]} {digits[9:11]}"


def format_multiple_phones(phone_str):
    if pd.isna(phone_str) or not phone_str:
        return []

    phones = []
    for part in str(phone_str).split(';'):
        cleaned = part.strip()
        if cleaned:
            formatted = format_phone(cleaned)
            if formatted:
                phones.append(formatted)
    return phones


def read_file_as_base64(file_path):
    if not os.path.exists(file_path):
        return None
    try:
        with open(file_path, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")
    except Exception:
        return None


def make_boundary():
    return f"----WebKitFormBoundary{uuid.uuid4().hex}"


def build_multipart_body(boundary, filename, fields, base64_content):
    CRLF = "\r\n"
    body = []
    body.append(f"--{boundary}")
    body.append(f'Content-Disposition: form-data; name="file"; filename="{filename}"')
    body.append("Content-Type: application/octet-stream")
    body.append("Content-Transfer-Encoding: base64")
    body.append("")
    body.append(base64_content)

    for key, value in fields.items():
        body.append(f"--{boundary}")
        body.append(f'Content-Disposition: form-data; name="{key}"')
        body.append("")
        body.append(str(value) if value is not None else "")

    body.append(f"--{boundary}--")
    return CRLF.join(body)


def find_file_in_dir(files_dir, filename_hint):
    """
    Ищет файл в директории по имени (с или без расширения, регистронезависимо).
    """
    filename_hint = str(filename_hint).strip()
    if not filename_hint:
        return None
    
    # Тупой поиск напрямую
    candidate = os.path.join(files_dir, filename_hint)
    if os.path.isfile(candidate):
        return candidate

    # Если расширение не указано — перебираем все поддерживаемые
    base_name = filename_hint
    for ext in SUPPORTED_EXTENSIONS:
        candidate = os.path.join(files_dir, base_name + ext)
        if os.path.isfile(candidate):
            return candidate

    # Дополнительно: если имя содержит точку, пробуем как есть (регистронезависимо)
    if '.' in base_name:
        # Генерируем все комбинации регистра расширения (редко, но для надёжности)
        name_part, ext_part = os.path.splitext(base_name)
        candidate = os.path.join(files_dir, name_part)
        if os.path.isfile(candidate):
            return candidate

    return None


def find_document_group_by_mnemonic(document_groups, target_mnemonic="request"):
    for group in (document_groups or []):
        for branch_item in group.get("branch", []):
            if branch_item.get("mnemonic") == target_mnemonic:
                return branch_item
    return None


def parse_key_value_mapping(raw: str) -> Dict[str, str]:
    out: Dict[str, str] = {}
    text = str(raw or "").strip()
    if not text:
        return out
    for part in re.split(r"[;\n]+", text):
        piece = part.strip()
        if not piece or "=" not in piece:
            continue
        key, value = piece.split("=", 1)
        k = key.strip()
        v = value.strip()
        if k:
            out[k] = v
    return out


def parse_path_list(raw: str) -> List[str]:
    out: List[str] = []
    text = str(raw or "").strip()
    if not text:
        return out
    for part in re.split(r"[;\n]+", text):
        token = part.strip().strip('"').strip("'").strip()
        if token:
            out.append(token)
    return out


def resolve_local_file_path(filename: str, files_dir: Path, base_dir: Path) -> Optional[str]:
    candidate = str(filename or "").strip()
    if not candidate:
        return None

    if os.path.isabs(candidate) and os.path.isfile(candidate):
        return os.path.abspath(candidate)

    p1 = (files_dir / candidate).resolve()
    if p1.exists() and p1.is_file():
        return str(p1)

    p2 = (base_dir / candidate).resolve()
    if p2.exists() and p2.is_file():
        return str(p2)

    base_name = os.path.basename(candidate).lower()
    if files_dir.exists():
        for root, _, files in os.walk(str(files_dir)):
            for fn in files:
                if fn.lower() == base_name:
                    return str((Path(root) / fn).resolve())
    return None


def read_text_if_exists(path: Path) -> str:
    if not path.exists():
        return ""
    return path.read_text(encoding="utf-8", errors="ignore").replace("\ufeff", "").strip()


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def dump_json(path: Path, data: Any) -> None:
    write_text(path, json.dumps(data, ensure_ascii=False, indent=2))
