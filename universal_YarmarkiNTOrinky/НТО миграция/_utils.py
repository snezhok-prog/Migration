from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import datetime


def is_empty(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and value != value:
        return True
    if isinstance(value, str) and not value.strip():
        return True
    return False


def to_scalar(value: Any) -> Any:
    if is_empty(value):
        return None
    if isinstance(value, str):
        text = value.strip()
        return text if text else None
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def split_tokens(value: Any, separators: str = ";,\n") -> List[str]:
    text = str(to_scalar(value) or "")
    if not text:
        return []
    cleaned = text.replace("\r\n", "\n").replace("\r", "\n")
    for sep in separators:
        if sep == "\n":
            continue
        cleaned = cleaned.replace(sep, "\n")
    parts = []
    for chunk in cleaned.split("\n"):
        token = chunk.strip().strip('"').strip("'").strip()
        if token:
            parts.append(token)
    return parts


def read_text_if_exists(path: Path) -> str:
    if not path.exists():
        return ""
    return path.read_text(encoding="utf-8", errors="ignore").replace("\ufeff", "").strip()


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


_SEGMENT_RE = re.compile(r"^([^\[]+)((\[\d+\])*)$")
_INDEX_RE = re.compile(r"\[(\d+)\]")


def _path_to_parts(path: str) -> List[Any]:
    parts = []
    for segment in str(path).split("."):
        matched = _SEGMENT_RE.match(segment)
        if not matched:
            parts.append(segment)
            continue
        prop = matched.group(1)
        idxs = matched.group(2) or ""
        parts.append(prop)
        for mm in _INDEX_RE.finditer(idxs):
            parts.append(int(mm.group(1)))
    return parts


def set_by_path(obj: Dict[str, Any], path: str, value: Any) -> None:
    parts = _path_to_parts(path)
    cur = obj
    i = 0
    while i < len(parts):
        key = parts[i]
        is_last = i == len(parts) - 1
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
            raise ValueError("Expected list while setting path '%s'" % path)
        while len(cur) <= key:
            cur.append(None)
        if is_last:
            cur[key] = value
            return
        if not isinstance(cur[key], dict):
            cur[key] = {}
        cur = cur[key]
        i += 1


def get_path(obj: Any, path: str, default: Any = None) -> Any:
    parts = _path_to_parts(path)
    cur = obj
    for part in parts:
        if isinstance(cur, dict):
            cur = cur.get(part, default)
        elif isinstance(cur, list) and isinstance(part, int):
            if 0 <= part < len(cur):
                cur = cur[part]
            else:
                return default
        else:
            return default
    return cur


def _normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    text = text.replace("\u00A0", " ")
    return re.sub(r"\s+", " ", text)


def norm_str(value: Any) -> str:
    return _normalize_text(value)


def get_any(obj: Any, paths: List[str], default: Any = None) -> Any:
    for path in paths:
        value = get_path(obj, path)
        if value is not None and value != "":
            return value
    return default


def as_date_or_null(value: Any) -> Optional[str]:
    if value is None:
        return None
    if isinstance(value, datetime.datetime):
        return value.date().isoformat()
    if isinstance(value, datetime.date):
        return value.isoformat()
    if isinstance(value, (int, float)):
        return None
    text = _normalize_text(value)
    if not text:
        return None
    for fmt in ["%d.%m.%Y", "%Y-%m-%d", "%d.%m.%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"]:
        try:
            return datetime.datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    return None


def to_number_or_null(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = _normalize_text(value).replace(" ", "")
    text = text.replace(",", ".")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def addr_obj(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, dict):
        return value
    text = _normalize_text(value)
    if not text:
        return None
    return {"fullAddress": text}


def generate_guid() -> str:
    import uuid
    return str(uuid.uuid4())


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


def parse_cookie_pairs(raw_cookie: str) -> List[Tuple[str, str]]:
    out = []
    text = str(raw_cookie or "").replace("\ufeff", "")
    if not text.strip():
        return out

    header_match = re.search(
        r"(?ims)\bcookie\b\s*[:=]?\s*(.+?)(?:\n\s*[A-Za-z][A-Za-z0-9\-]*\s*(?::|$)|$)",
        text,
    )
    if header_match:
        cookie_chunk = header_match.group(1)
    else:
        cookie_chunk = text

    cleaned = cookie_chunk.replace("\r", " ").replace("\n", " ").strip()
    for piece in cleaned.split(";"):
        part = piece.strip()
        if not part:
            continue
        if part.lower().startswith("cookie:"):
            part = part.split(":", 1)[1].strip()
        if part.lower().startswith("cookie "):
            part = part[7:].strip()
        if "=" not in part:
            continue
        name, value = part.split("=", 1)
        name = name.strip()
        value = value.strip()
        if name and value and re.match(r"^[A-Za-z0-9_\-\.]+$", name):
            out.append((name, value))
    return out


def cookie_jar_to_string(cookie_jar) -> str:
    items = []
    for c in cookie_jar:
        name = getattr(c, "name", None)
        value = getattr(c, "value", None)
        if name:
            items.append("%s=%s" % (name, value or ""))
    return "; ".join(items)


def extract_jwt(text: str) -> str:
    candidate = str(text or "").replace("\ufeff", "").strip()
    if not candidate:
        return ""
    candidate = re.sub(r"^\s*bearer\s+", "", candidate, flags=re.IGNORECASE).strip()
    token_match = re.search(r"\b([A-Za-z0-9\-_]+\.[A-Za-z0-9\-_]+\.[A-Za-z0-9\-_]+)\b", candidate)
    if token_match:
        return token_match.group(1)
    return candidate


def extract_jwt_from_html(text: str) -> str:
    html = text or ""
    patterns = [
        r'<textarea[^>]*name=["\']token["\'][^>]*>(.*?)</textarea>',
        r'<textarea[^>]*name=["\']jwt["\'][^>]*>(.*?)</textarea>',
    ]
    for pattern in patterns:
        m = re.search(pattern, html, flags=re.IGNORECASE | re.DOTALL)
        if m:
            token = extract_jwt(m.group(1))
            if token:
                return token
    return extract_jwt(html)


def parse_key_value_mapping(raw: str) -> Dict[str, str]:
    out = {}
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
    out = []
    text = str(raw or "").strip()
    if not text:
        return out
    for part in re.split(r"[;\n]+", text):
        token = part.strip().strip('"').strip("'").strip()
        if token:
            out.append(token)
    return out


def dump_json(path: Path, data: Any) -> None:
    write_text(path, json.dumps(data, ensure_ascii=False, indent=2))
