import argparse
import glob
import json
import os
import re
import sys
import warnings
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional

from urllib3.exceptions import InsecureRequestWarning

from _api import (
    ApiCallError,
    create_record,
    get_runtime_base_url,
    search_collection,
    set_runtime_urls,
    setup_session,
    update_record,
    upload_file,
    upload_file_base64,
)
from _config import (
    ACT_COLLECTION,
    ALLOW_BASE64_FALLBACK,
    BASE_URL,
    CARD_COLLECTION,
    CARD_PART_GLOB,
    CATCH_PART_GLOB,
    DEFAULT_ORG,
    DEFAULT_ORG_ENABLED,
    DRY_RUN_LOG_UPLOAD_TARGETS,
    ENABLE_FILE_UPLOADS,
    EXCEL_DATA_START_ROW,
    EXCEL_INPUT_FILE,
    EXCEL_INPUT_GLOB,
    FILES_DIR,
    JWT_URL,
    ORG_STRICT_SEARCH_BY_NAME_OGRN,
    ORDER_COLLECTION,
    PREFER_FILES_DIR_UPLOAD,
    REGION_CODE,
    RELEASE_COLLECTION,
    RESUME_BY_DEFAULT,
    ROLLBACK_BODY_PATH,
    SCRIPT_DIR,
    STATE_FILE,
    STOP_ON_FIRST_CREATE_ERROR,
    STOP_ON_FIRST_FATAL_UPLOAD,
    STRAY_PART_GLOB,
    TARGET_COLLECTION,
    TRANSFER_ACT_COLLECTION,
    UI_ANIMAL_EDIT_PATH,
    UI_CATCH_ACT_EDIT_PATH,
    UI_CATCH_ORDER_EDIT_PATH,
    UI_RELEASE_ACT_EDIT_PATH,
    UI_TRANSFER_ACT_EDIT_PATH,
    USE_EXCEL_INPUT,
    VERIFY_CREATED,
)
from _excel_input import discover_excel_files, load_rows_from_excel
from _logger import setup_fail_logger, setup_logger, setup_success_logger
from _profiles import PROFILES
from _state import ResumeState
from _utils import (
    as_string_or_null,
    base64_size_bytes,
    build_address,
    build_minimal_address,
    get_by_path,
    generate_guid,
    norm_ru,
    read_rows_json,
    set_by_path,
    to_iso_z,
    to_iso_z_datetime,
    to_millis_safe,
)


PART_RE = re.compile(r"_part(\d+)\.json$", re.IGNORECASE)
ACTIVE_FILES_DIR = FILES_DIR
RUNTIME_OPERATOR_MODE = False


@dataclass
class WorkbookRunSpec:
    workbook_path: str
    files_dir: str


def set_active_files_dir(path: str):
    global ACTIVE_FILES_DIR
    if path:
        ACTIVE_FILES_DIR = os.path.abspath(path)
    else:
        ACTIVE_FILES_DIR = os.path.abspath(FILES_DIR)


def _parse_path_list(raw: str) -> List[str]:
    out = []
    text = str(raw or "").strip()
    if not text:
        return out
    for part in re.split(r"[;\n]+", text):
        token = part.strip().strip('"').strip("'").strip()
        if token:
            out.append(token)
    return out


def _parse_key_value_mapping(raw: str) -> Dict[str, str]:
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


def _resolve_explicit_workbook_path(raw: str, *, script_dir: str) -> str:
    token = str(raw or "").strip()
    if not token:
        return ""
    p = Path(token)
    if p.is_absolute():
        return str(p.resolve())

    candidates = [
        Path.cwd() / p,
        Path(script_dir) / p,
        Path(script_dir).parent / p,
    ]
    for candidate in candidates:
        try:
            resolved = candidate.resolve()
        except Exception:
            continue
        if resolved.exists() and resolved.is_file():
            return str(resolved)
    return str(candidates[0].resolve())


def _numeric_hints(text: str) -> List[str]:
    normalized = (" " + str(text or "").strip().lower() + " ")
    hints = set(re.findall(r"\d+", normalized))
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
    print("\nAvailable workbooks:")
    for i, wb in enumerate(candidates, start=1):
        print("  %s) %s" % (i, wb))
    raw = input("Choose workbook index [1]: ").strip()
    if not raw:
        return candidates[0]
    try:
        idx = int(raw)
    except Exception:
        return candidates[0]
    if 1 <= idx <= len(candidates):
        return candidates[idx - 1]
    return candidates[0]


def _choose_mass_workbooks(candidates: List[str], interactive: bool) -> List[str]:
    if not candidates:
        return []
    if not interactive:
        return list(candidates)
    print("\nAvailable workbooks:")
    for i, wb in enumerate(candidates, start=1):
        print("  %s) %s" % (i, wb))
    raw = input("Choose workbook indexes separated by comma, or Enter for ALL: ").strip()
    if not raw:
        return list(candidates)
    selected = []
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


def _infer_files_dir_for_workbook(
    *,
    workbook_path: str,
    files_root: str,
    files_map: Dict[str, str],
    interactive: bool,
    prompt_always: bool = False,
) -> str:
    wb_name = os.path.basename(workbook_path)
    wb_stem = os.path.splitext(wb_name)[0]

    for key in (wb_name, wb_stem):
        if key in files_map:
            target = files_map[key]
            if os.path.isabs(target):
                return os.path.abspath(target)
            return os.path.abspath(os.path.join(files_root, target))

    subdirs = []
    if os.path.isdir(files_root):
        for name in sorted(os.listdir(files_root), key=lambda x: x.lower()):
            p = os.path.abspath(os.path.join(files_root, name))
            if os.path.isdir(p):
                subdirs.append(p)

    auto_candidate = os.path.abspath(files_root)
    same_name_dir = os.path.abspath(os.path.join(files_root, wb_stem))
    if os.path.isdir(same_name_dir):
        auto_candidate = same_name_dir
    elif subdirs:
        stem_norm = wb_stem.strip().lower()
        matched = []
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
                wb_hints = set(_numeric_hints(stem_norm))
                hint_matches = []
                if wb_hints:
                    for subdir in subdirs:
                        sub_hints = set(_numeric_hints(os.path.basename(subdir)))
                        if wb_hints.intersection(sub_hints):
                            hint_matches.append(subdir)
                if len(hint_matches) == 1:
                    auto_candidate = hint_matches[0]
            if auto_candidate == os.path.abspath(files_root):
                if len(subdirs) == 2:
                    default_like = [x for x in subdirs if os.path.basename(x).strip().lower() in {"one", "1", "default", "main"}]
                    if len(default_like) == 1:
                        auto_candidate = default_like[0]
                elif len(subdirs) == 1:
                    auto_candidate = subdirs[0]

    if not interactive or not subdirs:
        return auto_candidate

    options = [os.path.abspath(files_root)] + subdirs
    default_idx = 0
    for i, p in enumerate(options):
        if os.path.abspath(p) == os.path.abspath(auto_candidate):
            default_idx = i
            break

    print("\nWorkbook: %s" % wb_name)
    print("Choose files folder:")
    print("  0) %s (root)" % os.path.abspath(files_root))
    for i, d in enumerate(subdirs, start=1):
        print("  %s) %s" % (i, os.path.basename(d)))
    raw = input("Folder index [%s]: " % default_idx).strip()
    if not raw:
        return options[default_idx]
    try:
        idx = int(raw)
    except Exception:
        return options[default_idx]
    if 0 <= idx < len(options):
        return options[idx]
    return options[default_idx]


def resolve_workbook_specs(
    *,
    mode: str,
    workbook_paths_arg: str,
    files_map_arg: str,
    interactive: bool,
    ask_files_always: bool,
) -> List[WorkbookRunSpec]:
    files_root = os.path.abspath(FILES_DIR)
    files_map = _parse_key_value_mapping(files_map_arg)

    explicit = _parse_path_list(workbook_paths_arg)
    if explicit:
        candidates = []
        for raw in explicit:
            candidate = _resolve_explicit_workbook_path(raw, script_dir=SCRIPT_DIR)
            if candidate and os.path.isfile(candidate):
                candidates.append(os.path.abspath(candidate))
    else:
        candidates = discover_excel_files(SCRIPT_DIR, explicit_files=EXCEL_INPUT_FILE, pattern=EXCEL_INPUT_GLOB)
    candidates = sorted(list(dict.fromkeys(candidates)), key=lambda p: os.path.basename(p).lower())

    if not candidates:
        return []

    if mode == "single":
        one = _choose_single_workbook(candidates, interactive=interactive)
        selected = [one] if one else []
    elif mode == "mass":
        selected = _choose_mass_workbooks(candidates, interactive=interactive)
    else:
        if len(candidates) == 1:
            selected = [candidates[0]]
        else:
            selected = _choose_mass_workbooks(candidates, interactive=interactive)

    prompt_files_for_each = bool(
        interactive
        and not files_map
        and (ask_files_always or mode in {"mass", "auto"})
    )

    out = []
    for workbook_path in selected:
        files_dir = _infer_files_dir_for_workbook(
            workbook_path=workbook_path,
            files_root=files_root,
            files_map=files_map,
            interactive=interactive,
            prompt_always=prompt_files_for_each,
        )
        out.append(WorkbookRunSpec(workbook_path=os.path.abspath(workbook_path), files_dir=os.path.abspath(files_dir)))
    return out


def dclone(obj):
    return json.loads(json.dumps(obj, ensure_ascii=False))


def ui_catch_order_link(_id):
    return f"{get_runtime_base_url()}{UI_CATCH_ORDER_EDIT_PATH}/{_id}"


def ui_animal_link(_id):
    return f"{get_runtime_base_url()}{UI_ANIMAL_EDIT_PATH}/{_id}"


def ui_catch_act_link(_id):
    return f"{get_runtime_base_url()}{UI_CATCH_ACT_EDIT_PATH}/{_id}"


def ui_release_link(_id):
    return f"{get_runtime_base_url()}{UI_RELEASE_ACT_EDIT_PATH}/{_id}"


def ui_transfer_act_link(_id):
    return f"{get_runtime_base_url()}{UI_TRANSFER_ACT_EDIT_PATH}/{_id}"


def rollback_body_payload(rollback_candidates):
    return [
        {
            "_id": item.get("_id"),
            "guid": item.get("guid"),
            "parentEntries": item.get("parentEntries"),
        }
        for item in rollback_candidates
    ]


def print_rollback(reason, rollback_candidates, logger):
    payload = rollback_body_payload(rollback_candidates)
    raw = json.dumps(payload, ensure_ascii=False, indent=2)
    logger.info("===== ROLLBACK BODY (%s) =====\n%s\n===== /ROLLBACK BODY =====", reason, raw)
    with open(ROLLBACK_BODY_PATH, "w", encoding="utf-8") as f:
        f.write(raw)


def _is_upload_error(exc):
    if isinstance(exc, ApiCallError):
        s = str(exc).lower()
        return "upload" in s or "загруз" in s
    return False


def build_file_placeholder(filename, size, entity_field_path, allow_external=False):
    return {
        "_id": "",
        "originalName": filename,
        "size": size,
        "isFile": True,
        "entityFieldPath": entity_field_path,
        "allowExternal": bool(allow_external),
    }


def _is_abs_path(p):
    try:
        return os.path.isabs(str(p))
    except Exception:
        return False


def resolve_local_file_path(filename):
    fn = as_string_or_null(filename)
    if not fn:
        return None

    # 1) absolute path from json
    if _is_abs_path(fn) and os.path.isfile(fn):
        return os.path.abspath(fn)

    # 2) relative path under files/
    direct = os.path.abspath(os.path.join(ACTIVE_FILES_DIR, fn))
    if os.path.isfile(direct):
        return direct

    # 3) relative path under script dir
    script_rel = os.path.abspath(os.path.join(SCRIPT_DIR, fn))
    if os.path.isfile(script_rel):
        return script_rel

    # 4) fallback under default files root
    default_root = os.path.abspath(os.path.join(FILES_DIR, fn))
    if os.path.isfile(default_root):
        return default_root

    # 5) recursive by basename under active files dir
    base = os.path.basename(fn).lower()
    for root, _, files in os.walk(ACTIVE_FILES_DIR):
        for cand in files:
            if cand.lower() == base:
                return os.path.abspath(os.path.join(root, cand))
    # 6) recursive by basename under default files root
    if os.path.abspath(ACTIVE_FILES_DIR) != os.path.abspath(FILES_DIR):
        for root, _, files in os.walk(FILES_DIR):
            for cand in files:
                if cand.lower() == base:
                    return os.path.abspath(os.path.join(root, cand))
    return None


def _upload_source(upload):
    local_path = as_string_or_null(upload.get("filePath")) or resolve_local_file_path(upload.get("filename"))
    has_local = bool(local_path and os.path.isfile(local_path))
    has_b64 = bool(as_string_or_null(upload.get("base64")))
    if PREFER_FILES_DIR_UPLOAD and has_local:
        return {"kind": "file", "filePath": local_path}
    if has_b64 and ALLOW_BASE64_FALLBACK:
        return {"kind": "base64"}
    if has_local:
        return {"kind": "file", "filePath": local_path}
    return {"kind": "none"}


def apply_uploads_to_doc(session, logger, collection, main_id, guid, doc_state, pending_uploads):
    if not pending_uploads:
        return doc_state

    for upload in pending_uploads:
        path = upload["path"]
        source = _upload_source(upload)
        if source["kind"] == "file":
            logger.info(
                "[UPLOAD][file] collection=%s id=%s path=%s file=%s",
                collection,
                main_id,
                path,
                source["filePath"],
            )
            uploaded = upload_file(
                session=session,
                logger=logger,
                file_path=source["filePath"],
                entry_name=collection,
                entry_id=main_id,
                entity_field_path=path,
                allow_external=bool(upload.get("allowExternal", False)),
            )
        elif source["kind"] == "base64":
            logger.info(
                "[UPLOAD][base64] collection=%s id=%s path=%s fileName=%s",
                collection,
                main_id,
                path,
                upload["filename"],
            )
            uploaded = upload_file_base64(
                session=session,
                logger=logger,
                entry_name=collection,
                entry_id=main_id,
                entity_field_path=path,
                filename=upload["filename"],
                base64_content=upload.get("base64"),
            )
        else:
            raise RuntimeError(
                f"No upload source for path={path}, file={upload.get('filename')}. "
                f"Add file to files/ or provide base64."
            )

        up = uploaded.get("data") if isinstance(uploaded, dict) else uploaded
        if not isinstance(up, dict):
            up = {}

        up_id = up.get("id")
        server_name = (
            up.get("originalName")
            or up.get("name")
            or (up_id.split("/")[-1] if isinstance(up_id, str) else None)
            or upload["filename"]
        )
        source_size = os.path.getsize(source["filePath"]) if source.get("kind") == "file" else base64_size_bytes(upload.get("base64"))
        response_file_id = up.get("_id") or up.get("fileId") or up.get("id") or (up_id if isinstance(up_id, str) else "")
        if not response_file_id:
            raise ApiCallError(
                "Upload response missing file id for path=%s file=%s" % (path, upload.get("filename")),
                data=up,
            )
        response_size_raw = up.get("size")
        try:
            response_size = int(float(response_size_raw)) if response_size_raw is not None else None
        except Exception:
            response_size = None
        if source.get("kind") == "file" and source_size > 0 and response_size is not None and response_size <= 0:
            raise ApiCallError(
                "Upload returned zero size for non-empty file path=%s file=%s" % (path, upload.get("filename")),
                code=413,
                data=up,
            )

        file_size_for_meta = response_size if (response_size is not None and response_size > 0) else source_size
        file_meta = {
            "_id": response_file_id,
            "size": file_size_for_meta,
            "isFile": True,
            "originalName": server_name,
            "allowExternal": bool(upload.get("allowExternal", False)),
            "entityFieldPath": path,
        }
        set_by_path(doc_state, path, file_meta)
        doc_state = update_record(session, logger, collection, main_id, guid, doc_state)

        persisted_meta = get_by_path(doc_state, path)
        persisted_id = persisted_meta.get("_id") if isinstance(persisted_meta, dict) else None
        if not persisted_id:
            raise ApiCallError(
                "Uploaded file not persisted with _id path=%s file=%s" % (path, upload.get("filename")),
                data={"uploadResponse": up, "persistedMeta": persisted_meta},
            )
        persisted_size_raw = persisted_meta.get("size") if isinstance(persisted_meta, dict) else None
        try:
            persisted_size = int(float(persisted_size_raw)) if persisted_size_raw is not None else None
        except Exception:
            persisted_size = None
        if source.get("kind") == "file" and source_size > 0 and (persisted_size is None or persisted_size <= 0):
            raise ApiCallError(
                "Uploaded file persisted with zero/missing size path=%s file=%s" % (path, upload.get("filename")),
                code=413,
                data={"uploadResponse": up, "persistedMeta": persisted_meta},
            )

    return doc_state


def append_success(success_logger, rollback_candidates, ok, payload):
    rollback_candidates.append(
        {
            "_id": payload["_id"],
            "guid": payload["guid"],
            "parentEntries": payload["parentEntries"],
        }
    )
    success_logger.info(
        json.dumps(
            {"_id": payload["_id"], "guid": payload["guid"], "parentEntries": payload["parentEntries"]},
            ensure_ascii=False,
        )
    )
    ok.append(payload)


def append_error(fail_logger, fail, payload):
    fail.append(payload)
    fail_logger.info(json.dumps(payload, ensure_ascii=False))


def serialize_exception(exc):
    if isinstance(exc, ApiCallError):
        out = {"message": str(exc)}
        if exc.code is not None:
            out["code"] = exc.code
        if exc.data is not None:
            out["data"] = exc.data
        return out
    return str(exc)


def log_processing_exception(logger, label: str, row_num: int, exc: Exception):
    if isinstance(exc, ApiCallError):
        logger.error("%s %s | %s", label, row_num, json.dumps(serialize_exception(exc), ensure_ascii=False))
    else:
        logger.exception("%s %s", label, row_num)


def part_sort_key(path):
    base = os.path.basename(path)
    m = PART_RE.search(base)
    if m:
        return int(m.group(1)), base
    return 10**9, base


def discover_input_files(part_glob, fallback_name=None):
    files = sorted(glob.glob(os.path.join(SCRIPT_DIR, part_glob)), key=part_sort_key)
    if files:
        return files
    if not fallback_name:
        return []
    fallback = os.path.join(SCRIPT_DIR, fallback_name)
    return [fallback] if os.path.exists(fallback) else []


def load_rows_from_files(files, logger, label):
    rows = []
    for path in files:
        part_rows = read_rows_json(path)
        rows.extend(part_rows)
        logger.info("[%s] file=%s rows=%s", label, os.path.basename(path), len(part_rows))
    logger.info("[%s] total rows=%s", label, len(rows))
    return rows


def verify_created_entries(session, logger, rollback_candidates):
    if not rollback_candidates:
        return
    checked = 0
    missing = 0
    for item in rollback_candidates:
        collection = item.get("parentEntries")
        main_id = item.get("_id")
        if not collection or not main_id:
            continue
        body = {
            "search": {
                "search": [
                    {
                        "andSubConditions": [
                            {"field": "_id", "operator": "eq", "value": str(main_id)},
                        ]
                    }
                ]
            }
        }
        try:
            data = search_collection(session, logger, collection, body)
            content = data.get("content") if isinstance(data, dict) else []
            found = isinstance(content, list) and len(content) > 0
            checked += 1
            if not found:
                missing += 1
                logger.warning("[VERIFY] NOT FOUND collection=%s _id=%s", collection, main_id)
        except Exception as exc:
            checked += 1
            missing += 1
            logger.warning("[VERIFY] ERROR collection=%s _id=%s err=%s", collection, main_id, exc)
    logger.info("[VERIFY] checked=%s missing_or_error=%s", checked, missing)


def pick_unit_short(unit):
    if not unit:
        return None
    return {
        "id": unit.get("id") or unit.get("_id"),
        "name": unit.get("name"),
        "shortName": unit.get("shortName"),
    }


def pick_unit_mini(unit):
    if not unit:
        return None
    return {
        "id": unit.get("id") or unit.get("_id") or None,
        "name": unit.get("name") or None,
        "shortName": unit.get("shortName") or None,
    }


def _lookup_normalized(mapping: Dict[str, Any], value: Any):
    normalized = norm_ru(value)
    if not normalized:
        return None
    for k, v in (mapping or {}).items():
        if norm_ru(k) == normalized:
            return v
    return None


def _lookup_normalized_startswith(mapping: Dict[str, Any], value: Any):
    normalized = norm_ru(value)
    if not normalized:
        return None
    for k, v in (mapping or {}).items():
        nk = norm_ru(k)
        if nk and normalized.startswith(nk):
            return v
    return None


def _region_code_by_name(region_name: Any):
    return _lookup_normalized(REGION_CODE, region_name)


def make_unit_from_excel(region_name, municipality, org_name, inn, ogrn):
    region_code = _region_code_by_name(region_name) if region_name else None
    return {
        "id": None,
        "name": org_name or None,
        "ogrn": str(ogrn or ""),
        "inn": str(inn or ""),
        "region": {"code": region_code, "name": region_name or None},
        "municipality": municipality or None,
        "shortName": None,
    }


def build_unit_from_org_record(org, fallback_unit):
    if not org and not fallback_unit:
        return None
    regions = (org or {}).get("regions") or (org or {}).get("region") or {}
    return {
        "id": (org or {}).get("_id") or (org or {}).get("id") or (fallback_unit or {}).get("id"),
        "name": (org or {}).get("name") or (fallback_unit or {}).get("name"),
        "ogrn": str((org or {}).get("ogrn") or (fallback_unit or {}).get("ogrn") or ""),
        "inn": str((org or {}).get("inn") or (fallback_unit or {}).get("inn") or ""),
        "region": {
            "code": regions.get("code") or ((fallback_unit or {}).get("region") or {}).get("code"),
            "name": regions.get("name") or ((fallback_unit or {}).get("region") or {}).get("name"),
        },
        "municipality": (fallback_unit or {}).get("municipality"),
        "shortName": (org or {}).get("shortName") or (fallback_unit or {}).get("shortName"),
    }


def search_org_strict_by_name_ogrn(session, logger, org_name, ogrn, role_label):
    name = as_string_or_null(org_name)
    orgn = as_string_or_null(ogrn)
    mode_label = "name/shortName + ogrn" if ORG_STRICT_SEARCH_BY_NAME_OGRN else "ogrn"

    if ORG_STRICT_SEARCH_BY_NAME_OGRN:
        if not name or not orgn:
            logger.warning("[ORG-LOOKUP][%s] strict mode skipped (name='%s', ogrn='%s')", role_label, name, orgn)
            return {"found": False, "reason": "no-input-strict", "mode": mode_label}
        search_body = {
            "search": {
                "search": [
                    {
                        "orSubConditions": [
                            {
                                "andSubConditions": [
                                    {"field": "name", "operator": "eq", "value": name},
                                    {"field": "ogrn", "operator": "eq", "value": orgn},
                                ]
                            },
                            {
                                "andSubConditions": [
                                    {"field": "shortName", "operator": "eq", "value": name},
                                    {"field": "ogrn", "operator": "eq", "value": orgn},
                                ]
                            },
                        ]
                    }
                ]
            }
        }
    else:
        if not orgn:
            logger.warning("[ORG-LOOKUP][%s] OGRN lookup skipped (ogrn='%s')", role_label, orgn)
            return {"found": False, "reason": "no-ogrn", "mode": mode_label}
        search_body = {
            "search": {
                "search": [
                    {
                        "andSubConditions": [
                            {"field": "ogrn", "operator": "eq", "value": orgn},
                        ]
                    }
                ]
            }
        }

    try:
        data = search_collection(session, logger, "organizations", search_body)
        entries = data.get("content") if isinstance(data, dict) else None
        entries = entries if isinstance(entries, list) else []
    except Exception as exc:
        logger.error("[ORG-LOOKUP][%s] search error (%s): %s", role_label, mode_label, exc)
        return {"found": False, "reason": "error", "error": str(exc), "mode": mode_label}

    if not entries:
        logger.warning("[ORG-LOOKUP][%s] not found (%s). name='%s', ogrn='%s'", role_label, mode_label, name, orgn)
        return {"found": False, "reason": "not-found", "mode": mode_label}
    if len(entries) > 1:
        logger.error("[ORG-LOOKUP][%s] multiple found (%s). name='%s', ogrn='%s', count=%s", role_label, mode_label, name, orgn, len(entries))
        return {"found": False, "reason": "multiple", "mode": mode_label}

    logger.info("[ORG-LOOKUP][%s] found (%s). id=%s", role_label, mode_label, entries[0].get("_id"))
    return {"found": True, "org": entries[0], "mode": mode_label}


def resolve_org_pair(
    session,
    logger,
    *,
    region_name,
    municipality,
    primary_name,
    primary_ogrn,
    primary_inn,
    secondary_name,
    secondary_ogrn,
    secondary_inn,
    secondary_role_label,
):
    result = {"unit": None, "units": [], "primaryFound": False, "secondaryFound": False}

    base_primary = make_unit_from_excel(region_name, municipality, primary_name, primary_inn, primary_ogrn)
    unit = base_primary
    primary_lookup = search_org_strict_by_name_ogrn(session, logger, primary_name, primary_ogrn, "authorizedOrg")
    if primary_lookup.get("found"):
        unit = build_unit_from_org_record(primary_lookup.get("org"), base_primary)
        result["primaryFound"] = True
    elif DEFAULT_ORG_ENABLED:
        unit = build_unit_from_org_record(DEFAULT_ORG, base_primary)
        result["primaryFound"] = True

    secondary_unit = None
    if secondary_name or secondary_ogrn:
        base_secondary = make_unit_from_excel(region_name, municipality, secondary_name, secondary_inn, secondary_ogrn)
        secondary_unit = base_secondary
        secondary_lookup = search_org_strict_by_name_ogrn(session, logger, secondary_name, secondary_ogrn, secondary_role_label)
        if secondary_lookup.get("found"):
            secondary_unit = build_unit_from_org_record(secondary_lookup.get("org"), base_secondary)
            result["secondaryFound"] = True
        elif DEFAULT_ORG_ENABLED:
            secondary_unit = build_unit_from_org_record(DEFAULT_ORG, base_secondary)
            result["secondaryFound"] = True

    units = []
    for cand in [unit, secondary_unit]:
        if not cand:
            continue
        key = "|".join([str(cand.get("id") or ""), str(cand.get("ogrn") or ""), str(cand.get("name") or "")])
        if not any("|".join([str(x.get("id") or ""), str(x.get("ogrn") or ""), str(x.get("name") or "")]) == key for x in units):
            units.append(cand)

    result["unit"] = unit
    result["units"] = units
    return result


def resolve_orgs_for_stray_row(session, logger, row):
    base = resolve_org_pair(
        session,
        logger,
        region_name=row.get("region"),
        municipality=row.get("municipality"),
        primary_name=row.get("authorizedOrgName"),
        primary_ogrn=row.get("ogrn"),
        primary_inn=row.get("inn"),
        secondary_name=row.get("catchOrgName"),
        secondary_ogrn=row.get("catchOrgOgrn"),
        secondary_inn=row.get("catchOrgInn"),
        secondary_role_label="catchOrg",
    )
    return {"unit": base["unit"], "units": base["units"], "authFound": base["primaryFound"], "catchFound": base["secondaryFound"]}


def resolve_orgs_for_order_row(session, logger, row):
    oi = row.get("orderInfo") or {}
    base = resolve_org_pair(
        session,
        logger,
        region_name=oi.get("region"),
        municipality=oi.get("municipality"),
        primary_name=oi.get("authorizedOrgName"),
        primary_ogrn=oi.get("ogrn"),
        primary_inn=oi.get("inn"),
        secondary_name=oi.get("catchOrgName"),
        secondary_ogrn=oi.get("catchOrgOgrn"),
        secondary_inn=oi.get("catchOrgInn"),
        secondary_role_label="catchOrg",
    )
    return {"unit": base["unit"], "units": base["units"], "authFound": base["primaryFound"], "catchFound": base["secondaryFound"]}


def resolve_orgs_for_card_row(session, logger, row):
    base = resolve_org_pair(
        session,
        logger,
        region_name=row.get("region"),
        municipality=row.get("municipality"),
        primary_name=row.get("authorizedOrgName"),
        primary_ogrn=row.get("ogrn"),
        primary_inn=row.get("inn"),
        secondary_name=row.get("shelterName"),
        secondary_ogrn=row.get("shelterOGRN"),
        secondary_inn=row.get("shelterINN"),
        secondary_role_label="shelterOrg",
    )
    return {"unit": base["unit"], "units": base["units"], "authFound": base["primaryFound"], "shelterFound": base["secondaryFound"]}


# ----------------------------- STRAY ANIMALS ---------------------------------

TYPE_STRAY = {
    "кошка": "cat",
    "кот": "cat",
    "собака": "dog",
    "пес": "dog",
    "пёс": "dog",
    "котенок": "kitten",
    "котёнок": "kitten",
    "щенок": "puppy",
}

SIZE_STRAY = {
    "маленький": "little",
    "маленькая": "little",
    "малый": "little",
    "средний": "medium",
    "средняя": "medium",
    "среднее": "medium",
    "большой": "big",
    "большая": "big",
    "крупный": "big",
}

STATUS_STRAY = {
    "не отловлено": "notCaptured",
    "на отлове": "onTheCapture",
    "в приюте": "atShelter",
    "в пункте временного содержания": "inThePointOfTemporaryContent",
    "передано": "passed",
    "выпущено": "released",
    "падеж": "died",
    "отловлено": "captured",
    "ожидает оформления заявки на отлов": "onTheCapture",
}


def map_type_stray(value):
    name = as_string_or_null(value)
    if not name:
        return None
    code = _lookup_normalized(TYPE_STRAY, name)
    return {"code": code, "name": name} if code else {"name": name}


def map_size_stray(value):
    name = as_string_or_null(value)
    if not name:
        return None
    code = _lookup_normalized_startswith(SIZE_STRAY, name)
    return {"code": code, "name": name} if code else {"name": name}


def map_clip_presence(value, clip_color):
    base = as_string_or_null(value)
    if not base and as_string_or_null(clip_color):
        return {"code": "presence", "name": "Да"}
    s = norm_ru(base or "")
    if not s:
        return None
    if s in {"да", "есть"} or "налич" in s:
        return {"code": "presence", "name": "Да"}
    if s == "нет" or "отсут" in s:
        return {"code": "absence", "name": "Нет"}
    if s == "не знаю" or "не знаю" in s:
        return {"code": "unknown", "name": "Не знаю"}
    return {"code": "unknown", "name": "Не знаю"}


def map_aggression_stray(value):
    s = norm_ru(value)
    if not s:
        return None
    if s == "была" or s.startswith("да"):
        return {"code": "presence", "name": "Была"}
    if s in {"не была", "не было"} or s.startswith("нет"):
        return {"code": "absence", "name": "Не было"}
    return None


def map_status_stray(value):
    s = norm_ru(value)
    if not s:
        return None
    code = (
        _lookup_normalized(STATUS_STRAY, s)
        or ("notCaptured" if "не отлов" in s else None)
        or ("onTheCapture" if "на отлов" in s else None)
        or ("atShelter" if "приют" in s else None)
        or ("inThePointOfTemporaryContent" if "временн" in s else None)
        or ("passed" if "передан" in s else None)
        or ("released" if "выпущ" in s else None)
        or ("died" if "падеж" in s else None)
        or ("captured" if "отловлен" in s else None)
    )
    return {"code": code, "name": value} if code else None


def validate_stray_row_before_create(row):
    missing = []
    if not as_string_or_null(row.get("type")):
        missing.append("type")
    if not as_string_or_null(row.get("locationAddress")):
        missing.append("locationAddress")
    if not as_string_or_null(row.get("animalNumber")) and not as_string_or_null(row.get("orderNumber")):
        missing.append("animalNumber|orderNumber")
    if missing:
        raise ValueError("VALIDATION: missing required fields: " + ", ".join(missing))


def build_animal_stray(row):
    clips_color = as_string_or_null(row.get("clipColor"))
    clip = map_clip_presence(row.get("clip"), clips_color)
    animal = {
        "clip": clip,
        "size": map_size_stray(row.get("size")),
        "type": map_type_stray(row.get("type")),
        "number": as_string_or_null(row.get("animalNumber")),
        "address": build_address(row.get("locationAddress") or ""),
        "benchmark": as_string_or_null(row.get("locationLandmark")),
        "clipsColor": clips_color,
        "coloration": as_string_or_null(row.get("coloration")),
        "additionalInfo": as_string_or_null(row.get("additionalInfo")),
        "unmotivatedAggression": map_aggression_stray(row.get("unmotivatedAggression")),
        "unmotivatedAggressionDescription": as_string_or_null(row.get("aggressionDescription")),
    }
    status_source = row.get("animalStatus") or row.get("orderStatus")
    st = map_status_stray(status_source)
    if st:
        animal["status"] = st
    animal["caught"] = bool(st and st.get("code") == "captured")
    return animal


def build_catch_info_stray(row):
    has_any = bool(
        row.get("orderNumber")
        or row.get("catchStartDate")
        or row.get("catchEndDate")
        or row.get("catchAddress")
        or row.get("catchVideoBase64")
        or row.get("catchActBase64")
        or row.get("catchActNumber")
        or row.get("catchActDate")
    )
    if not has_any:
        return None
    dt_end = to_iso_z_datetime(row.get("catchEndDate"), row.get("catchEndTime"))
    dt_beg = to_iso_z_datetime(row.get("catchStartDate"), row.get("catchStartTime"))
    catch_date = dt_end or dt_beg or None
    info = {
        "requestNumber": as_string_or_null(row.get("orderNumber")),
        "workOrderActRecordLink": as_string_or_null(row.get("workOrderActRecordLink")) or None,
        "catchRecordLink": as_string_or_null(row.get("catchRecordLink")) or None,
        "catchDate": catch_date,
        "catchAddress": build_address(row.get("catchAddress") or row.get("locationAddress") or ""),
    }
    if row.get("catchActNumber"):
        info["catchActNumber"] = as_string_or_null(row.get("catchActNumber"))
    if row.get("catchActDate"):
        info["catchActDate"] = to_iso_z(row.get("catchActDate"))
    return info


def collect_pending_uploads_stray(row):
    out = []
    photo_name = as_string_or_null(row.get("photoFileName")) or as_string_or_null(row.get("photo"))
    if photo_name:
        
        out.append({"path": "animal.photo[0]", "filename": photo_name, "base64": row.get("photoBase64"), "allowExternal": False})

    # note can be plain text, so use note as filename only when base64 is provided
    note_name = as_string_or_null(row.get("noteFileName")) or (
        as_string_or_null(row.get("note")) if row.get("noteBase64") else None
    )
    if note_name:
        out.append({"path": "animal.explanatoryNoteFile[0]", "filename": note_name, "base64": row.get("noteBase64"), "allowExternal": False})

    catch_video_name = as_string_or_null(row.get("catchVideoFileName")) or as_string_or_null(row.get("catchVideo"))
    if catch_video_name:
        out.append({"path": "catchInfo.catchVideoMigration[0]", "filename": catch_video_name, "base64": row.get("catchVideoBase64"), "allowExternal": True})

    catch_act_name = as_string_or_null(row.get("catchActFileName")) or as_string_or_null(row.get("catchAct"))
    if catch_act_name:
        out.append({"path": "catchInfo.catchActFile[0]", "filename": catch_act_name, "base64": row.get("catchActBase64"), "allowExternal": False})
    return out


def has_catch_act(row):
    return bool(
        row.get("catchActNumber")
        or row.get("catchActDate")
        or row.get("catchActBase64")
        or row.get("catchActFileName")
        or row.get("catchAct")
    )


def collect_act_pending_uploads(row):
    out = []
    # note can be plain text, so use note as filename only when base64 is provided
    note_name = as_string_or_null(row.get("noteFileName")) or (
        as_string_or_null(row.get("note")) if row.get("noteBase64") else None
    )
    if note_name:
        out.append({"path": "catchProcessInfo.explanatoryNoteFile[0]", "filename": note_name, "base64": row.get("noteBase64"), "allowExternal": False})

    catch_act_name = as_string_or_null(row.get("catchActFileName")) or as_string_or_null(row.get("catchAct"))
    if catch_act_name:
        out.append({"path": "actData.catchActFile[0]", "filename": catch_act_name, "base64": row.get("catchActBase64"), "allowExternal": False})
    return out


def build_stray_record(row, resolved_orgs):
    unit = resolved_orgs.get("unit")
    units = resolved_orgs.get("units") if isinstance(resolved_orgs.get("units"), list) else ([unit] if unit else [])
    doc = {
        "guid": generate_guid(),
        "unit": unit,
        "units": units,
        "animal": build_animal_stray(row),
        "entityType": None,
        "parentEntries": TARGET_COLLECTION,
    }
    catch_info = build_catch_info_stray(row)
    if catch_info:
        doc["catchInfo"] = catch_info
    return {"record": doc, "pendingUploads": collect_pending_uploads_stray(row)}


def find_catch_order_by_animal_and_request(session, logger, animal_number, request_number):
    an = as_string_or_null(animal_number)
    rn = as_string_or_null(request_number)
    if not an or not rn:
        return {"found": False, "reason": "no-input"}
    animal_fields = ["animalNumber", "animal.number.animal.number", "animal.number", "animal.animal.number", "animal.number.animalNumber"]
    request_fields = ["requestNumber", "orderNumber", "catchRequestInfo.requestNumber", "catchInfo.requestNumber"]
    ors = []
    for af in animal_fields:
        for rf in request_fields:
            ors.append({"andSubConditions": [{"field": af, "operator": "eq", "value": an}, {"field": rf, "operator": "eq", "value": rn}]})
    body = {"search": {"search": [{"orSubConditions": ors}]}}
    try:
        data = search_collection(session, logger, ORDER_COLLECTION, body)
        entries = data.get("content") if isinstance(data, dict) else []
        entries = entries if isinstance(entries, list) else []
        if not entries:
            return {"found": False, "reason": "not-found"}
        entries.sort(
            key=lambda x: (to_millis_safe(x.get("dateLastModification") or x.get("dateCreation")), int(x.get("auid") or 0)),
            reverse=True,
        )
        return {"found": True, "order": entries[0]}
    except Exception as exc:
        logger.warning("[ORDER-LOOKUP] error: %s", exc)
        return {"found": False, "reason": "error"}


def build_catch_act_record(row, unit, animal_doc, animal_id, catch_order_link):
    animal_obj = (animal_doc or {}).get("animal") or {}
    act_animal = {
        "number": {"animal": animal_obj},
        "animalRegistryLink": ui_animal_link(animal_id),
        "clip": animal_obj.get("clip"),
        "size": animal_obj.get("size"),
        "type": animal_obj.get("type"),
        "benchmark": animal_obj.get("benchmark"),
        "coloration": animal_obj.get("coloration"),
        "additionalInfo": animal_obj.get("additionalInfo"),
        "unmotivatedAggression": animal_obj.get("unmotivatedAggression"),
        "unmotivatedAggressionDescription": animal_obj.get("unmotivatedAggressionDescription"),
        "address": animal_obj.get("address"),
        "photo": animal_obj.get("photo"),
        "clipsColor": animal_obj.get("clipsColor"),
    }
    doc = {
        "guid": generate_guid(),
        "unit": pick_unit_short(unit),
        "parentEntries": ACT_COLLECTION,
        "animal": [act_animal],
        "catchRequestInfo": {
            "requestNumber": as_string_or_null(row.get("orderNumber")),
            "catchRequestRegistryLink": catch_order_link or "",
        },
        "catchProcessInfo": {
            "catchStartDate": to_iso_z(row.get("catchStartDate")),
            "catchStartTime": as_string_or_null(row.get("catchStartTime")),
            "catchEndDate": to_iso_z(row.get("catchEndDate")),
            "catchEndTime": as_string_or_null(row.get("catchEndTime")),
            "catchAddress": build_address(row.get("catchAddress") or row.get("locationAddress") or ""),
            "hunterFullName": as_string_or_null(row.get("catcherFIO")),
        },
        "actData": {
            "actDate": to_iso_z(row.get("catchActDate")) or to_iso_z(row.get("catchEndDate")) or None,
            "org": as_string_or_null(row.get("authorizedOrgName")) or as_string_or_null(row.get("catchOrgName")) or None,
            "actNumber": as_string_or_null(row.get("catchActNumber")),
        },
    }
    return {"record": doc, "pendingUploads": collect_act_pending_uploads(row)}


def process_stray_rows(session, logger, success_logger, fail_logger, rows, rollback_candidates, ok, fail):
    logger.info("[STRAY] start rows=%s", len(rows))
    for i, row in enumerate(rows):
        row_num = int(row.get("__row_num", i + 1))
        animal_number = row.get("animalNumber")
        try:
            validate_stray_row_before_create(row)
            resolved = resolve_orgs_for_stray_row(session, logger, row)
            if not resolved.get("authFound") and not resolved.get("catchFound") and not DEFAULT_ORG_ENABLED:
                append_error(
                    fail_logger,
                    fail,
                    {
                        "registry": "stray",
                        "index": row_num,
                        "animalNumber": animal_number,
                        "stage": "org-lookup",
                        "error": "no-organizations-found",
                    },
                )
                continue

            built = build_stray_record(row, resolved)
            created = create_record(session, logger, TARGET_COLLECTION, built["record"])
            main_id = created.get("_id")
            guid = created.get("guid") or built["record"]["guid"]
            if not main_id or not guid:
                raise RuntimeError("No _id/guid for created stray record")

            append_success(
                success_logger,
                rollback_candidates,
                ok,
                {
                    "registry": "stray",
                    "index": row_num,
                    "animalNumber": animal_number,
                    "_id": main_id,
                    "guid": guid,
                    "parentEntries": TARGET_COLLECTION,
                },
            )

            doc_state = created
            if ENABLE_FILE_UPLOADS and built["pendingUploads"]:
                if DRY_RUN_LOG_UPLOAD_TARGETS:
                    for u in built["pendingUploads"]:
                        logger.warning("[STRAY][DRY] upload planned: %s", {"path": u["path"], "filename": u["filename"]})
                else:
                    try:
                        doc_state = apply_uploads_to_doc(
                            session=session,
                            logger=logger,
                            collection=TARGET_COLLECTION,
                            main_id=main_id,
                            guid=guid,
                            doc_state=doc_state,
                            pending_uploads=built["pendingUploads"],
                        )
                    except Exception as upload_exc:
                        append_error(
                            fail_logger,
                            fail,
                            {
                                "registry": "stray",
                                "index": row_num,
                                "animalNumber": animal_number,
                                "stage": "upload/link",
                                "error": serialize_exception(upload_exc),
                            },
                        )
                        log_processing_exception(logger, "[STRAY] upload error:", row_num, upload_exc)
                        if _is_upload_error(upload_exc) and STOP_ON_FIRST_FATAL_UPLOAD and not RUNTIME_OPERATOR_MODE:
                            return True

            if has_catch_act(row):
                order_lookup = find_catch_order_by_animal_and_request(
                    session,
                    logger,
                    row.get("animalNumber"),
                    row.get("orderNumber"),
                )
                order_link = ""
                if order_lookup.get("found"):
                    order = order_lookup.get("order") or {}
                    order_link = ui_catch_order_link(order.get("_id") or order.get("id"))

                act_built = build_catch_act_record(row, resolved.get("unit"), doc_state, main_id, order_link)
                act_created = create_record(session, logger, ACT_COLLECTION, act_built["record"])
                act_id = act_created.get("_id")
                act_guid = act_created.get("guid") or act_built["record"]["guid"]
                if not act_id or not act_guid:
                    raise RuntimeError("No _id/guid for created catch act")

                append_success(
                    success_logger,
                    rollback_candidates,
                    ok,
                    {
                        "registry": "catch-act-from-stray",
                        "index": row_num,
                        "animalNumber": animal_number,
                        "_id": act_id,
                        "guid": act_guid,
                        "parentEntries": ACT_COLLECTION,
                    },
                )

                if ENABLE_FILE_UPLOADS and act_built["pendingUploads"]:
                    apply_uploads_to_doc(
                        session=session,
                        logger=logger,
                        collection=ACT_COLLECTION,
                        main_id=act_id,
                        guid=act_guid,
                        doc_state=act_created,
                        pending_uploads=act_built["pendingUploads"],
                    )

                refreshed_stray = _fetch_record_by_id(session, logger, TARGET_COLLECTION, main_id)
                if isinstance(refreshed_stray, dict):
                    doc_state = refreshed_stray
                doc_state.setdefault("catchInfo", {})
                doc_state["catchInfo"]["catchRecordLink"] = ui_catch_act_link(act_id)
                update_record(session, logger, TARGET_COLLECTION, main_id, guid, doc_state)

        except Exception as exc:
            append_error(
                fail_logger,
                fail,
                {"registry": "stray", "index": row_num, "animalNumber": animal_number, "error": serialize_exception(exc)},
            )
            log_processing_exception(logger, "[STRAY] row error:", row_num, exc)
            if _is_upload_error(exc):
                if STOP_ON_FIRST_FATAL_UPLOAD and not RUNTIME_OPERATOR_MODE:
                    return True
                continue
            if STOP_ON_FIRST_CREATE_ERROR and not RUNTIME_OPERATOR_MODE:
                return True
    return False


# ------------------------------ CATCH ORDERS ---------------------------------

PRIORITY_ORDER = {"высокий": "hight", "средний": "middle", "низкий": "low"}


def map_type_order(value):
    s = norm_ru(value)
    if "щен" in s:
        return {"code": "puppy", "name": "Щенок"}
    if "котен" in s or "котён" in s:
        return {"code": "kitten", "name": "Котенок"}
    if "кошк" in s or "кот" in s:
        return {"code": "cat", "name": "Кошка"}
    if "собак" in s or "пес" in s or "пёс" in s:
        return {"code": "dog", "name": "Собака"}
    return {"code": "dog", "name": as_string_or_null(value) or "Собака"}


def map_size_order(value):
    s = norm_ru(value)
    if "мал" in s:
        return {"code": "little", "name": "Маленький"}
    if "сред" in s:
        return {"code": "medium", "name": "Средний"}
    if "бол" in s or "круп" in s:
        return {"code": "big", "name": "Большой"}
    return {"code": "medium", "name": as_string_or_null(value) or "Средний"}


def map_clip_order(value):
    s = norm_ru(value)
    if s in {"да", "есть"} or "налич" in s:
        return {"code": "presence", "name": "Да"}
    if s == "нет" or "отсут" in s:
        return {"code": "absence", "name": "Нет"}
    return {"code": "unknown", "name": "Не знаю"}


def map_aggression_order(value):
    s = norm_ru(value)
    if s == "была" or "есть" in s or "набл" in s:
        return {"code": "presence", "name": "Была"}
    if s in {"не была", "не было"} or "нет" in s:
        return {"code": "absence", "name": "Не была"}
    return {"code": "absence", "name": "Не была"}


def map_status_order(value):
    s = norm_ru(value)
    if "на отлов" in s:
        return {"code": "inCatchProcess", "name": "На отлове"}
    if "не отлов" in s:
        return {"code": "notCaught", "name": "Не отловлено"}
    if "отлов" in s:
        return {"code": "caught", "name": "Отловлено"}
    return None


def validate_order_row_before_create(row):
    oi = row.get("orderInfo") or {}
    animals = row.get("animals")
    if not isinstance(animals, list) or not animals:
        raise ValueError("VALIDATION: animals[] is empty")
    if oi.get("orderNumber") is None or oi.get("orderNumber") == "":
        raise ValueError("VALIDATION: orderInfo.orderNumber is empty")
    for idx, item in enumerate(animals):
        if not as_string_or_null((item or {}).get("kind")):
            raise ValueError(f"VALIDATION: animals[{idx}].kind is empty")


def build_order_record(row, resolved):
    oi = row.get("orderInfo") or {}
    region_code = _region_code_by_name(oi.get("region")) if oi.get("region") else None
    unit = resolved.get("unit")
    units = resolved.get("units") if isinstance(resolved.get("units"), list) else ([unit] if unit else [])
    animals = []
    uploads = []

    for idx, item in enumerate(row.get("animals") or []):
        a = item or {}
        addr = {
            "okato": "",
            "oktmo": "",
            "country": "",
            "postalCode": "",
            "regionCode": region_code or "",
            "fullAddress": a.get("locationAddress") or None,
        }
        clip = map_clip_order(a.get("clip"))
        size = map_size_order(a.get("size"))
        kind = map_type_order(a.get("kind"))
        aggression = map_aggression_order(a.get("unmotivatedAggression"))
        status = map_status_order(a.get("status"))

        photo_name = as_string_or_null(a.get("photoFileName")) or as_string_or_null(a.get("photo"))
        if photo_name:
            uploads.append(
                {
                    "path": f"animal[{idx}].photo[0]",
                    "filename": photo_name,
                    "base64": a.get("photoBase64"),
                    "allowExternal": False,
                }
            )
        # note can be plain text, so use note as filename only when base64 is provided
        note_name = as_string_or_null(a.get("noteFileName")) or (as_string_or_null(a.get("note")) if a.get("noteBase64") else None)
        if note_name:
            uploads.append(
                {
                    "path": f"animal[{idx}].explanatoryNoteFile[0]",
                    "filename": note_name,
                    "base64": a.get("noteBase64"),
                    "allowExternal": False,
                }
            )

        animal_number = as_string_or_null(a.get("number"))
        animal_number_obj = {
            "animal": {
                "clip": clip,
                "clipsColor": as_string_or_null(a.get("clipColor")),
                "size": size,
                "type": kind,
                "number": animal_number,
                "address": addr,
                "benchmark": a.get("locationLandmark") or None,
                "coloration": a.get("color") or None,
                "additionalInfo": a.get("extraInfo") or None,
                "unmotivatedAggression": aggression,
                "unmotivatedAggressionDescription": as_string_or_null(a.get("aggressionDescription")),
            }
        }

        animals.append(
            {
                "clip": clip,
                "clipsColor": as_string_or_null(a.get("clipColor")),
                "size": size,
                "type": kind,
                "number": animal_number_obj,
                "address": addr,
                "benchmark": a.get("locationLandmark") or None,
                "coloration": a.get("color") or None,
                "additionalInfo": a.get("extraInfo") or None,
                "unmotivatedAggressionDescription": as_string_or_null(a.get("aggressionDescription")),
                "notificationNegative": as_string_or_null(a.get("note")),
                "catchOrderStatus": status,
                "unmotivatedAggression": aggression,
            }
        )

    doc = {
        "guid": generate_guid(),
        "unit": unit,
        "units": units,
        "animal": animals,
        "prioritet": _lookup_normalized(PRIORITY_ORDER, oi.get("priority")),
        "dayToOtlov": int(oi.get("catchDays")) if str(oi.get("catchDays") or "").isdigit() else None,
        "entityType": None,
        "parentEntries": ORDER_COLLECTION,
        "requestNumber": oi.get("orderNumber"),
    }
    return {"record": doc, "pendingUploads": uploads}


def process_order_rows(session, logger, success_logger, fail_logger, rows, rollback_candidates, ok, fail):
    logger.info("[CATCH-ORDER] start rows=%s", len(rows))
    for i, row in enumerate(rows):
        row_num = int(row.get("__row_num", i + 1))
        oi = row.get("orderInfo") or {}
        order_number = oi.get("orderNumber")
        try:
            validate_order_row_before_create(row)
            resolved = resolve_orgs_for_order_row(session, logger, row)
            if not resolved.get("authFound") and not resolved.get("catchFound") and not DEFAULT_ORG_ENABLED:
                append_error(
                    fail_logger,
                    fail,
                    {
                        "registry": "catch-order",
                        "index": row_num,
                        "orderNumber": order_number,
                        "stage": "org-lookup",
                        "error": "no-organizations-found",
                    },
                )
                continue

            built = build_order_record(row, resolved)
            created = create_record(session, logger, ORDER_COLLECTION, built["record"])
            main_id = created.get("_id")
            guid = created.get("guid") or built["record"]["guid"]
            if not main_id or not guid:
                raise RuntimeError("No _id/guid for created catch order")

            append_success(
                success_logger,
                rollback_candidates,
                ok,
                {
                    "registry": "catch-order",
                    "index": row_num,
                    "orderNumber": order_number,
                    "_id": main_id,
                    "guid": guid,
                    "parentEntries": ORDER_COLLECTION,
                },
            )

            if ENABLE_FILE_UPLOADS and built["pendingUploads"]:
                apply_uploads_to_doc(
                    session=session,
                    logger=logger,
                    collection=ORDER_COLLECTION,
                    main_id=main_id,
                    guid=guid,
                    doc_state=created,
                    pending_uploads=built["pendingUploads"],
                )

        except Exception as exc:
            append_error(
                fail_logger,
                fail,
                {"registry": "catch-order", "index": row_num, "orderNumber": order_number, "error": serialize_exception(exc)},
            )
            log_processing_exception(logger, "[CATCH-ORDER] row error:", row_num, exc)
            if _is_upload_error(exc):
                if STOP_ON_FIRST_FATAL_UPLOAD and not RUNTIME_OPERATOR_MODE:
                    return True
                continue
            if STOP_ON_FIRST_CREATE_ERROR and not RUNTIME_OPERATOR_MODE:
                return True
    return False


# ------------------------------ ANIMAL CARDS ---------------------------------

SEX_CARD = {"мужской": "male", "муж": "male", "женский": "female", "жен": "female"}
TYPE_CARD = {
    "кошка": "cat",
    "кот": "cat",
    "собака": "dog",
    "пес": "dog",
    "пёс": "dog",
    "котенок": "kitten",
    "котёнок": "kitten",
    "щенок": "puppy",
}
SIZE_CARD = {
    "маленький": "little",
    "маленькая": "little",
    "малый": "little",
    "средний": "medium",
    "средняя": "medium",
    "среднее": "medium",
    "большой": "big",
    "большая": "big",
    "крупный": "big",
}
FUR_CARD = {
    "короткошерстное": "shortHaired",
    "короткошерстный": "shortHaired",
    "средней длины": "mediumLength",
    "средняя длина": "mediumLength",
    "длинношерстное": "longHaired",
    "длинношерстный": "longHaired",
}
EARS_CARD = {
    "маленькие": "little",
    "большие": "big",
    "средние": "middle",
    "висячие": "lopEared",
    "вислоухие": "lopEared",
    "стоячие": "erect",
    "полустоячие": "semiErect",
    "полуcтоячие": "semiErect",
    "купированные": "cropped",
    "с клипсой": "clipped",
    "остроконечные": "pointed",
    "без ушей": "noEars",
    "одно ухо": "oneEar",
    "с фигурным вырезом (выщипом)": "plucked",
}
TAIL_CARD = {
    "большой": "big",
    "маленький": "little",
    "средний": "middle",
    "прямой": "straight",
    "короткий": "short",
    "кольцом": "ringShaped",
    "двойным кольцом": "doubleRingShaped",
    "пером": "featherShaped",
    "пушистый": "fluffy",
    "без хвоста": "noTail",
    "купированный": "cropped",
    "саблевидный": "saberShaped",
    "серповидный": "sickleShaped",
    "прутом": "rodShaped",
    "поленообразный": "logShaped",
}
ANIMAL_STATUS_CARD = {
    "находится в приюте": "atShelter",
    "в приюте": "atShelter",
    "выпущен в места обитания": "releasedToHabitat",
    "передан владельцу": "transferredToOwner",
    "пал (падеж)": "dead",
    "падеж": "dead",
}

# Aliases from Excel wording variants.
EARS_CARD.update(
    {
        "с фигурным вырезом (выщипом)": "plucked",
    }
)
ANIMAL_STATUS_CARD.update(
    {
        "находится в пункте временного содержания": "inThePointOfTemporaryContent",
        "в пункте временного содержания": "inThePointOfTemporaryContent",
        "выпущено в прежнюю среду обитания": "releasedToHabitat",
        "передано владельцу": "transferredToOwner",
        "проведено мероприятие по эвтаназии": "dead",
    }
)

ANIMAL_STATUS_CARD_CANONICAL_NAME = {
    "atShelter": "В приюте",
    "inThePointOfTemporaryContent": "В пункте временного содержания",
    "releasedToHabitat": "Выпущен в места обитания",
    "transferredToOwner": "Передан владельцу",
    "dead": "Падеж",
}


def map_code_card(dict_obj, value):
    name = as_string_or_null(value)
    if not name:
        return None
    code = _lookup_normalized(dict_obj, name)
    return {"code": code, "name": name} if code else {"name": name}


def sex_code_card(value):
    return _lookup_normalized(SEX_CARD, value)


def status_obj_card(value):
    name = as_string_or_null(value)
    if not name:
        return None
    code = _lookup_normalized(ANIMAL_STATUS_CARD, name)
    if not code:
        return {"name": name}
    return {"code": code, "name": ANIMAL_STATUS_CARD_CANONICAL_NAME.get(code) or name}


def strip_card_suffix(value):
    s = as_string_or_null(value)
    if not s:
        return None
    return re.sub(r"-[A-Za-z\u0410-\u044f\u0401\u0451]+$", "", s)


def split_fio(value):
    s = as_string_or_null(value)
    if not s:
        return {"last": None, "first": None, "patronymic": None}
    parts = re.sub(r"\s+", " ", s).strip().split(" ")
    return {
        "last": parts[0] if len(parts) > 0 else None,
        "first": parts[1] if len(parts) > 1 else None,
        "patronymic": " ".join(parts[2:]) if len(parts) > 2 else None,
    }


def format_phone_ru(value):
    raw = as_string_or_null(value)
    if not raw:
        return None
    digits = re.sub(r"\D", "", raw)
    if len(digits) == 10:
        digits = "7" + digits
    if len(digits) == 11 and digits.startswith("8"):
        digits = "7" + digits[1:]
    if len(digits) != 11 or not digits.startswith("7"):
        return raw
    return f"+7 ({digits[1:4]}) {digits[4:7]} {digits[7:9]} {digits[9:11]}"


def detect_animal_receiver_type(block):
    if not isinstance(block, dict):
        return None
    has_shelter = any(
        [
            as_string_or_null(block.get("shelterName")),
            as_string_or_null(block.get("shelterOGRN")),
            as_string_or_null(block.get("shelterINN")),
            as_string_or_null(block.get("shelterAddress")),
            block.get("shelterAddressObj"),
        ]
    )
    has_pvs = any(
        [
            as_string_or_null(block.get("pvsName")),
            as_string_or_null(block.get("pvsOGRN")),
            as_string_or_null(block.get("pvsINN")),
            as_string_or_null(block.get("pvsAddress")),
            block.get("pvsAddressObj"),
        ]
    )
    if has_shelter:
        return "shelter"
    if has_pvs:
        return "temporaryHoldingFacility"
    return None


def detect_receiver_type_last_wins(block):
    if not isinstance(block, dict):
        return None
    has_shelter = any(
        [
            as_string_or_null(block.get("shelterName")),
            as_string_or_null(block.get("shelterAddress")),
            as_string_or_null(block.get("shelterPhone")),
        ]
    )
    has_pvs = any(
        [
            as_string_or_null(block.get("pvsName")),
            as_string_or_null(block.get("pvsAddress")),
            as_string_or_null(block.get("pvsPhone")),
        ]
    )
    if has_pvs:
        return "temporaryHoldingFacility"
    if has_shelter:
        return "shelter"
    return None


def _list_or_empty(value):
    return value if isinstance(value, list) else []


def animal_snapshot_for_events(row, animal_obj):
    snap = {}
    if (animal_obj or {}).get("type"):
        snap["type"] = (animal_obj or {}).get("type")
    if (animal_obj or {}).get("size"):
        snap["size"] = (animal_obj or {}).get("size")
    if (animal_obj or {}).get("sex"):
        snap["sex"] = (animal_obj or {}).get("sex")
    if as_string_or_null(row.get("breed")):
        snap["breed"] = as_string_or_null(row.get("breed"))
    if as_string_or_null(row.get("age")):
        snap["approximateAge"] = as_string_or_null(row.get("age"))
    return snap


def build_animal_card(row):
    animal = {}
    if row.get("fur"):
        animal["fur"] = map_code_card(FUR_CARD, row.get("fur"))
    if row.get("ears"):
        animal["ears"] = map_code_card(EARS_CARD, row.get("ears"))
    if row.get("tail"):
        animal["tail"] = map_code_card(TAIL_CARD, row.get("tail"))
    if row.get("size"):
        animal["size"] = map_code_card(SIZE_CARD, row.get("size"))
    if row.get("type"):
        animal["type"] = map_code_card(TYPE_CARD, row.get("type"))

    sc = sex_code_card(row.get("sex"))
    if sc:
        animal["sex"] = {"code": sc, "name": as_string_or_null(row.get("sex"))}

    if as_string_or_null(row.get("breed")):
        animal["breed"] = as_string_or_null(row.get("breed"))
    if as_string_or_null(row.get("coloration")):
        animal["coloration"] = as_string_or_null(row.get("coloration"))
    if as_string_or_null(row.get("nickname")):
        animal["nameAtTimeOfAct"] = as_string_or_null(row.get("nickname"))

    weight = as_string_or_null(row.get("weight"))
    if weight:
        animal["weight"] = weight.replace(",", ".")
    temp = as_string_or_null(row.get("temperature"))
    if temp:
        animal["temperature"] = temp.replace(",", ".")

    if as_string_or_null(row.get("specialMarks")):
        animal["specialMarks"] = as_string_or_null(row.get("specialMarks"))
    if as_string_or_null(row.get("age")):
        animal["approximateAge"] = as_string_or_null(row.get("age"))
    if as_string_or_null(row.get("cageNumber")):
        animal["crateNumber"] = as_string_or_null(row.get("cageNumber"))
    if as_string_or_null(row.get("injuriesInfo")):
        animal["injuriesInflictedByAnimal"] = as_string_or_null(row.get("injuriesInfo"))
    if row.get("animalStatus"):
        animal["status"] = status_obj_card(row.get("animalStatus"))

    keeping = {}
    q_end = to_iso_z(row.get("quarantineUntilDate"))
    rel_s = to_iso_z(row.get("releaseFromShelterDate"))
    rel_p = to_iso_z(row.get("releaseFromPVSDate"))
    if q_end:
        keeping["quarantineEndDate"] = q_end
    if rel_s:
        keeping["releaseDateFromShelter"] = rel_s
    if rel_p:
        keeping["releaseDateFromTemporaryFacility"] = rel_p
    if keeping:
        animal["animalKeepingPeriod"] = keeping

    identity = row.get("identityMark") or {}
    if isinstance(identity, dict) and (identity.get("number") or identity.get("method") or identity.get("place")):
        animal["tagBlock"] = {}
        if as_string_or_null(identity.get("number")):
            animal["tagBlock"]["tagNumber"] = as_string_or_null(identity.get("number"))
        if as_string_or_null(identity.get("method")):
            animal["tagBlock"]["applicationMethod"] = as_string_or_null(identity.get("method"))
        if as_string_or_null(identity.get("place")):
            animal["tagBlock"]["applicationLocation"] = as_string_or_null(identity.get("place"))
    return animal


def build_events_card(row, animal_obj):
    events = []
    index_map = {}
    snap = animal_snapshot_for_events(row, animal_obj)

    dewormings = _list_or_empty(row.get("dewormings"))
    if dewormings:
        arr = []
        for d in dewormings:
            d = d or {}
            arr.append(
                {
                    **snap,
                    "dosage": as_string_or_null(d.get("dosage")),
                    "drugName": as_string_or_null(d.get("drugName")),
                    "executionDate": to_iso_z(d.get("date")),
                    "employeeFullName": as_string_or_null(d.get("employeeFIO")),
                    "employeePosition": as_string_or_null(d.get("employeePosition")),
                    "dewormingActNumber": as_string_or_null(d.get("actNumber")),
                }
            )
        index_map["deworming"] = len(events)
        events.append({"eventType": {"code": "deworming", "name": "Дегельминтизация"}, "deworming": arr})

    disinsections = _list_or_empty(row.get("disinsections"))
    if disinsections:
        arr = []
        for d in disinsections:
            d = d or {}
            arr.append(
                {
                    **snap,
                    "drugName": as_string_or_null(d.get("drugName")),
                    "executionDate": to_iso_z(d.get("date")),
                    "employeeFullName": as_string_or_null(d.get("employeeFIO")),
                    "employeePosition": as_string_or_null(d.get("employeePosition")),
                    "disinsectionActNumber": as_string_or_null(d.get("actNumber")),
                }
            )
        index_map["disinsection"] = len(events)
        events.append({"eventType": {"code": "disinsection", "name": "Дезинсекция"}, "disinsection": arr})

    vaccinations = _list_or_empty(row.get("vaccinations"))
    if vaccinations:
        arr = []
        for v in vaccinations:
            v = v or {}
            arr.append(
                {
                    **snap,
                    "dosage": as_string_or_null(v.get("dosage") or v.get("dose")),
                    "drugName": as_string_or_null(v.get("drugName")),
                    "executionDate": to_iso_z(v.get("date")),
                    "drugSeriesNumber": as_string_or_null(v.get("series")),
                    "employeeFullName": as_string_or_null(v.get("employeeFIO")),
                    "employeePosition": as_string_or_null(v.get("employeePosition")),
                    "vaccinationActNumber": as_string_or_null(v.get("actNumber")),
                }
            )
        index_map["vaccination"] = len(events)
        events.append({"eventType": {"code": "vaccination", "name": "Вакцинация"}, "vaccination": arr})

    sterilizations = _list_or_empty(row.get("sterilizations"))
    if sterilizations:
        s = sterilizations[0] or {}
        obj = {
            **snap,
            "executionDate": to_iso_z(s.get("date")),
            "employeeFullName": as_string_or_null(s.get("employeeFIO")),
            "employeePosition": as_string_or_null(s.get("employeePosition")),
            "sterilizationDrugName": as_string_or_null(s.get("drugName")),
            "sterilizationActNumber": as_string_or_null(s.get("actNumber")),
        }
        index_map["sterilization"] = len(events)
        events.append({"eventType": {"code": "sterilization", "name": "Стерилизация (кастрация)"}, "sterilization": obj})

    marking_events = _list_or_empty(row.get("markingEvents"))
    if marking_events:
        m = next(
            (
                x
                for x in marking_events
                if isinstance(x, dict)
                and any(
                    [
                        x.get("number"),
                        x.get("method"),
                        x.get("place"),
                        x.get("date"),
                        x.get("employeeFIO"),
                        x.get("employeePosition"),
                    ]
                )
            ),
            None,
        )
        if m:
            obj = {
                "tagNumber": as_string_or_null(m.get("number")),
                "applicationMethod": as_string_or_null(m.get("method")),
                "applicationLocation": as_string_or_null(m.get("place")),
                "executionDate": to_iso_z(m.get("date")),
                "employeeFullName": as_string_or_null(m.get("employeeFIO")),
                "employeePosition": as_string_or_null(m.get("employeePosition")),
            }
            index_map["identificationTagApplying"] = len(events)
            events.append(
                {
                    "eventType": {"code": "identificationTagApplying", "name": "Нанесение идентификационной метки"},
                    "identificationTagApplying": obj,
                }
            )

    examination = row.get("examination") or None
    if isinstance(examination, dict):
        member_keys = ["commissionMember257", "commissionMember258", "commissionMember259", "commissionMember260", "commissionMember261"]
        members = [as_string_or_null(examination.get(k)) for k in member_keys]
        members = [m for m in members if m]
        obj = {
            "actAuthor": as_string_or_null(examination.get("actAuthor")),
            "inspectionDate": to_iso_z(examination.get("date")),
            "commissionDecision": as_string_or_null(examination.get("commissionDecision")),
            "inspectionActNumber": as_string_or_null(examination.get("actNumber")),
            "reactionToFoodWithStranger": as_string_or_null(examination.get("foodReactionPresence")),
            "reactionToFoodOfferedByStranger": as_string_or_null(examination.get("foodReactionOffer")),
            "reactionToLoudSounds": as_string_or_null(examination.get("loudSoundReaction")),
            "commissionMemberFullName": ", ".join(members) if members else None,
        }
        if members:
            obj["commissionMemberFullNameBlock"] = [{"commissionMemberFullName": x} for x in members]
        index_map["inspection"] = len(events)
        events.append({"eventType": {"code": "inspection", "name": "Освидетельствование"}, "inspection": obj})

    euthanasia = row.get("euthanasia") or None
    if isinstance(euthanasia, dict):
        obj = {
            "dosage": as_string_or_null(euthanasia.get("dosage")),
            "drugName": as_string_or_null(euthanasia.get("drugName")),
            "euthanasiaDate": to_iso_z(euthanasia.get("date")),
            "euthanasiaTime": as_string_or_null(euthanasia.get("time")),
            "employeeFullName": as_string_or_null(euthanasia.get("employeeFIO")),
            "employeePosition": as_string_or_null(euthanasia.get("employeePosition")),
            "euthanasiaMethod": as_string_or_null(euthanasia.get("method")),
            "euthanasiaReason": as_string_or_null(euthanasia.get("reason")),
            "euthanasiaActNumber": as_string_or_null(euthanasia.get("actNumber")),
        }
        index_map["euthanasia"] = len(events)
        events.append({"eventType": {"code": "euthanasia", "name": "Эвтаназия"}, "euthanasia": obj})

    utilization = row.get("utilization") or None
    if isinstance(utilization, dict):
        obj = {
            "disposalDate": to_iso_z(utilization.get("date")),
            "disposalMethod": as_string_or_null(utilization.get("method")),
            "disposalReason": as_string_or_null(utilization.get("basis")),
            "employeeFullName": as_string_or_null(utilization.get("employeeFIO")),
            "employeePosition": as_string_or_null(utilization.get("employeePosition")),
            "disposalActNumber": as_string_or_null(utilization.get("actNumber")),
        }
        index_map["disposal"] = len(events)
        events.append({"eventType": {"code": "disposal", "name": "Утилизация"}, "disposal": obj})

    other_events = _list_or_empty(row.get("otherEvents"))
    if other_events:
        arr = []
        for o in other_events:
            o = o or {}
            arr.append(
                {
                    "executionDate": to_iso_z(o.get("date")),
                    "eventName": as_string_or_null(o.get("name")),
                    "eventDescription": as_string_or_null(o.get("description")),
                    "employeeFullName": as_string_or_null(o.get("employeeFIO")),
                    "employeePosition": as_string_or_null(o.get("employeePosition")),
                    "otherEventDocumentNumber": as_string_or_null(o.get("documentNumber")),
                }
            )
        index_map["other"] = len(events)
        events.append({"eventType": {"code": "other", "name": "Иное мероприятие"}, "other": arr})

    return {"events": events, "indexMap": index_map}


def build_release_info(row):
    r = row.get("releaseInfo") or None
    if not isinstance(r, dict):
        return None
    return {
        "actName": as_string_or_null(r.get("actName")),
        "actNumber": as_string_or_null(r.get("actNumber")),
        "actDate": to_iso_z(r.get("actDate")),
        "shelterName": as_string_or_null(r.get("shelterName")),
        "shelterAddress": as_string_or_null(r.get("shelterAddress")),
        "shelterINN": as_string_or_null(r.get("shelterINN")),
        "shelterOGRN": as_string_or_null(r.get("shelterOGRN")),
        "pvsName": as_string_or_null(r.get("pvsName")),
        "pvsAddress": as_string_or_null(r.get("pvsAddress")),
        "pvsINN": as_string_or_null(r.get("pvsINN")),
        "pvsOGRN": as_string_or_null(r.get("pvsOGRN")),
        "catcherFIO": as_string_or_null(r.get("catcherFIO")),
        "releaseAddress": as_string_or_null(r.get("releaseAddress")),
        "handoverActNumber": as_string_or_null(((row.get("handoverWithShelter") or {}).get("actNumber"))),
        "releaseActNumber": as_string_or_null(r.get("actNumber")),
        "trunsferActNumber": as_string_or_null(((row.get("transferToOwner") or {}).get("actNumber"))),
        "deathActNumber": as_string_or_null(((row.get("deathInfo") or {}).get("actNumber"))),
    }


def build_transfer_to_owner_block(row):
    t = row.get("transferToOwner") or None
    if not isinstance(t, dict):
        return None
    return {
        "actName": as_string_or_null(t.get("actName")),
        "actNumber": as_string_or_null(t.get("actNumber")),
        "transferDate": to_iso_z(t.get("transferDate")),
        "shelterName": as_string_or_null(t.get("shelterName")),
        "shelterAddress": as_string_or_null(t.get("shelterAddress")),
        "pvsName": as_string_or_null(t.get("pvsName")),
        "pvsAddress": as_string_or_null(t.get("pvsAddress")),
        "pvsINN": as_string_or_null(t.get("pvsINN")),
        "pvsOGRN": as_string_or_null(t.get("pvsOGRN")),
        "newOwnerFIO": as_string_or_null(t.get("newOwnerFIO")),
        "newOwnerAddress": as_string_or_null(t.get("newOwnerAddress")),
        "idSeries": as_string_or_null(t.get("idSeries")),
        "idNumber": as_string_or_null(t.get("idNumber")),
        "idDeptCode": as_string_or_null(t.get("idDeptCode")),
        "idIssueDate": to_iso_z(t.get("idIssueDate")),
        "idIssuedBy": as_string_or_null(t.get("idIssuedBy")),
    }


def build_death_info_block(row):
    d = row.get("deathInfo") or None
    if not isinstance(d, dict):
        return None
    return {
        "actName": as_string_or_null(d.get("actName")),
        "actNumber": as_string_or_null(d.get("actNumber")),
        "actDate": to_iso_z(d.get("actDate")),
        "deathDate": to_iso_z(d.get("deathDate")),
        "shelterName": as_string_or_null(d.get("shelterName")),
        "shelterAddress": as_string_or_null(d.get("shelterAddress")),
        "pvsName": as_string_or_null(d.get("pvsName")),
        "pvsAddress": as_string_or_null(d.get("pvsAddress")),
        "pvsINN": as_string_or_null(d.get("pvsINN")),
        "pvsOGRN": as_string_or_null(d.get("pvsOGRN")),
    }


def _append_upload(out, path, filename, base64_content, allow_external=False):
    fn = as_string_or_null(filename)
    if not fn:
        return
    b64 = as_string_or_null(base64_content)
    payload = {"path": path, "filename": fn, "allowExternal": bool(allow_external)}
    if b64:
        payload["base64"] = b64
    out.append(payload)


def collect_card_photo_uploads(row, logger):
    pending = []
    filename = row.get("photoFileName") or row.get("photoFilename") or row.get("photo")
    if as_string_or_null(filename):
        _append_upload(pending, "animal.photo", filename, row.get("photoBase64"), allow_external=False)
        return pending

    photos = _list_or_empty(row.get("photos"))
    valid = [x for x in photos if isinstance(x, dict) and (x.get("fileName") or x.get("filename"))]
    if valid:
        if len(valid) > 1:
            logger.warning("[CARD][PHOTO] photos[] has >1 file, animal.photo is object; first file will be used")
        first = valid[0]
        _append_upload(pending, "animal.photo", first.get("fileName") or first.get("filename"), first.get("base64"), allow_external=False)
    return pending


def collect_event_act_uploads(row, index_map):
    pending = []

    def push_files_from_item(item, *, event_index, event_key, item_index, target_field, single_prefix, plural_key):
        if event_index is None or not isinstance(item, dict):
            return

        files = []
        sb = item.get(f"{single_prefix}Base64")
        sf = item.get(f"{single_prefix}FileName") or item.get(f"{single_prefix}Filename") or item.get(single_prefix)
        if sf:
            files.append({"base64": sb, "filename": sf})

        lst = _list_or_empty(item.get(plural_key))
        for f in lst:
            if not isinstance(f, dict):
                continue
            b64 = f.get("base64")
            fn = f.get("fileName") or f.get("filename")
            if fn:
                files.append({"base64": b64, "filename": fn})

        for file_index, f in enumerate(files):
            _append_upload(
                pending,
                f"events[{event_index}].{event_key}[{item_index}].{target_field}[{file_index}]",
                f.get("filename"),
                f.get("base64"),
                allow_external=False,
            )

    def from_array(arr, map_key, single_prefix, plural_key, target_field):
        event_index = index_map.get(map_key)
        if event_index is None:
            return
        for idx, item in enumerate(_list_or_empty(arr)):
            push_files_from_item(
                item,
                event_index=event_index,
                event_key=map_key,
                item_index=idx,
                target_field=target_field,
                single_prefix=single_prefix,
                plural_key=plural_key,
            )

    from_array(row.get("dewormings"), "deworming", "dewormAct", "dewormActs", "dewormingActFile")
    from_array(row.get("disinsections"), "disinsection", "disinsectionAct", "disinsectionActs", "disinsectionActFile")
    from_array(row.get("vaccinations"), "vaccination", "vaccinationAct", "vaccinationActs", "vaccinationActFile")

    ster_index = index_map.get("sterilization")
    ster_list = _list_or_empty(row.get("sterilizations"))
    if ster_index is not None and ster_list:
        s = ster_list[0] or {}
        _append_upload(
            pending,
            f"events[{ster_index}].sterilization.sterilizationActFile[0]",
            s.get("sterilizationActFileName") or s.get("sterilizationActFilename") or s.get("sterilizationAct"),
            s.get("sterilizationActBase64"),
            allow_external=False,
        )
        for idx, f in enumerate(_list_or_empty(s.get("sterilizationActs"))):
            if not isinstance(f, dict):
                continue
            _append_upload(
                pending,
                f"events[{ster_index}].sterilization.sterilizationActFile[{idx}]",
                f.get("fileName") or f.get("filename"),
                f.get("base64"),
                allow_external=False,
            )

    def collect_act_from_container(container, event_key, target_field):
        event_index = index_map.get(event_key)
        if event_index is None or not isinstance(container, dict):
            return
        files = []
        single_name = container.get("actFileFileName") or container.get("actFileName") or container.get("actFile")
        if single_name:
            files.append({"base64": container.get("actFileBase64"), "filename": single_name})
        for f in _list_or_empty(container.get("actFiles")):
            if not isinstance(f, dict):
                continue
            if f.get("fileName") or f.get("filename"):
                files.append({"base64": f.get("base64"), "filename": f.get("fileName") or f.get("filename")})
        for idx, f in enumerate(files):
            _append_upload(
                pending,
                f"events[{event_index}].{event_key}.{target_field}[{idx}]",
                f.get("filename"),
                f.get("base64"),
                allow_external=False,
            )

    collect_act_from_container(row.get("examination"), "inspection", "inspectionActFile")
    collect_act_from_container(row.get("euthanasia"), "euthanasia", "euthanasiaActFile")
    collect_act_from_container(row.get("utilization"), "disposal", "disposalActFile")

    other_index = index_map.get("other")
    if other_index is not None:
        for j, e in enumerate(_list_or_empty(row.get("otherEvents"))):
            if not isinstance(e, dict):
                continue
            files = []
            single_name = e.get("otherEventDocumentFileName") or e.get("otherEventDocumentFilename") or e.get("otherEventDocument")
            if single_name:
                files.append(
                    {
                        "base64": e.get("otherEventDocumentBase64"),
                        "filename": single_name,
                    }
                )
            for f in _list_or_empty(e.get("otherEventDocuments")):
                if not isinstance(f, dict):
                    continue
                if f.get("fileName") or f.get("filename"):
                    files.append({"base64": f.get("base64"), "filename": f.get("fileName") or f.get("filename")})
            for k, f in enumerate(files):
                _append_upload(
                    pending,
                    f"events[{other_index}].other[{j}].otherEventDocumentFile[{k}]",
                    f.get("filename"),
                    f.get("base64"),
                    allow_external=False,
                )

    vet_files = []
    vet_file_name = row.get("vetInspectionActFileFileName") or row.get("vetInspectionActFileName") or row.get("vetInspectionActFile")
    if vet_file_name:
        vet_files.append(
            {
                "base64": row.get("vetInspectionActFileBase64"),
                "filename": vet_file_name,
            }
        )
    for f in _list_or_empty(row.get("vetInspectionActFiles")):
        if not isinstance(f, dict):
            continue
        if f.get("fileName") or f.get("filename"):
            vet_files.append({"base64": f.get("base64"), "filename": f.get("fileName") or f.get("filename")})
    for idx, f in enumerate(vet_files):
        _append_upload(pending, f"files[{idx}]", f.get("filename"), f.get("base64"), allow_external=False)

    return pending


def collect_release_act_uploads(row, release_doc):
    pending = []
    r = row.get("releaseInfo") or {}
    files = []
    single_name = r.get("actFileFileName") or r.get("actFileName") or r.get("actFile")
    if single_name:
        files.append({"base64": r.get("actFileBase64"), "filename": single_name})
    for f in _list_or_empty(r.get("actFiles")):
        if isinstance(f, dict) and (f.get("fileName") or f.get("filename")):
            files.append({"base64": f.get("base64"), "filename": f.get("fileName") or f.get("filename")})
    if not files:
        return pending

    has_shelter_data = bool((release_doc or {}).get("shelterData"))
    has_temp_data = bool((release_doc or {}).get("temporaryHoldingFacilityData"))
    for idx, f in enumerate(files):
        if has_shelter_data:
            _append_upload(pending, f"shelterData.releaseActFile[{idx}]", f.get("filename"), f.get("base64"), allow_external=False)
        if has_temp_data:
            _append_upload(pending, f"temporaryHoldingFacilityData.releaseActFile[{idx}]", f.get("filename"), f.get("base64"), allow_external=False)
    return pending


def collect_death_act_uploads(row):
    pending = []
    d = row.get("deathInfo") or {}
    files = []
    single_name = d.get("actFileFileName") or d.get("actFileName") or d.get("actFile")
    if single_name:
        files.append({"base64": d.get("actFileBase64"), "filename": single_name})
    for f in _list_or_empty(d.get("actFiles")):
        if isinstance(f, dict) and (f.get("fileName") or f.get("filename")):
            files.append({"base64": f.get("base64"), "filename": f.get("fileName") or f.get("filename")})
    for idx, f in enumerate(files):
        _append_upload(pending, f"actData.deathActFile[{idx}]", f.get("filename"), f.get("base64"), allow_external=False)
    return pending


def collect_transfer_act_uploads(row):
    pending = []
    t = row.get("transferToOwner") or {}
    files = []
    single_name = t.get("actFileFileName") or t.get("actFileName") or t.get("actFile")
    if single_name:
        files.append({"base64": t.get("actFileBase64"), "filename": single_name})
    for f in _list_or_empty(t.get("actFiles")):
        if isinstance(f, dict) and (f.get("fileName") or f.get("filename")):
            files.append({"base64": f.get("base64"), "filename": f.get("fileName") or f.get("filename")})
    for idx, f in enumerate(files):
        _append_upload(pending, f"actData.transferActFile[{idx}]", f.get("filename"), f.get("base64"), allow_external=False)
    return pending


def collect_handover_with_catcher_act_uploads(row):
    pending = []
    h = row.get("handoverWithCatcher") or {}
    files = []
    single_name = h.get("actFileFileName") or h.get("actFileName") or h.get("actFile")
    if single_name:
        files.append({"base64": h.get("actFileBase64"), "filename": single_name})
    for f in _list_or_empty(h.get("actFiles")):
        if isinstance(f, dict) and (f.get("fileName") or f.get("filename")):
            files.append({"base64": f.get("base64"), "filename": f.get("fileName") or f.get("filename")})
    for idx, f in enumerate(files):
        _append_upload(pending, f"transferActFile[{idx}]", f.get("filename"), f.get("base64"), allow_external=False)
    return pending


def collect_handover_with_shelter_act_uploads(row):
    pending = []
    h = row.get("handoverWithShelter") or {}
    files = []
    single_name = h.get("actFileFileName") or h.get("actFileName") or h.get("actFile")
    if single_name:
        files.append({"base64": h.get("actFileBase64"), "filename": single_name})
    for f in _list_or_empty(h.get("actFiles")):
        if isinstance(f, dict) and (f.get("fileName") or f.get("filename")):
            files.append({"base64": f.get("base64"), "filename": f.get("fileName") or f.get("filename")})
    for idx, f in enumerate(files):
        _append_upload(pending, f"transferActFile[{idx}]", f.get("filename"), f.get("base64"), allow_external=False)
    return pending


def validate_card_row_before_create(row):
    missing = []
    if not as_string_or_null(row.get("cardNumber")):
        missing.append("cardNumber")
    if not as_string_or_null(row.get("type")):
        missing.append("type")
    if missing:
        raise ValueError("VALIDATION: missing required fields: " + ", ".join(missing))


def build_card_record(row, resolved_orgs, logger):
    unit = resolved_orgs.get("unit")
    units = resolved_orgs.get("units") if isinstance(resolved_orgs.get("units"), list) else ([unit] if unit else [])

    animal = build_animal_card(row)
    events_bundle = build_events_card(row, animal)
    events = events_bundle.get("events") or []
    index_map = events_bundle.get("indexMap") or {}

    pending_uploads = []
    pending_uploads.extend(collect_card_photo_uploads(row, logger))
    pending_uploads.extend(collect_event_act_uploads(row, index_map))

    doc = {
        "guid": generate_guid(),
        "unit": unit,
        "units": units,
        "animal": animal,
        "events": events,
        "number": strip_card_suffix(row.get("cardNumber")),
        "parentEntries": CARD_COLLECTION,
        "municipality": as_string_or_null(row.get("municipality")),
        "inn": as_string_or_null(row.get("inn")),
        "ogrn": as_string_or_null(row.get("ogrn")),
    }

    release_info = build_release_info(row)
    transfer_to_owner = build_transfer_to_owner_block(row)
    death_info = build_death_info_block(row)
    if release_info:
        doc["releaseInfo"] = release_info
    if transfer_to_owner:
        doc["transferToOwner"] = transfer_to_owner
    if death_info:
        doc["deathInfo"] = death_info

    return {"record": doc, "pendingUploads": pending_uploads}


def build_release_act_record(row, resolved_orgs, card_record):
    r = row.get("releaseInfo") or None
    if not isinstance(r, dict):
        return None
    has_any = any(
        [
            as_string_or_null(r.get("actName")),
            as_string_or_null(r.get("actNumber")),
            as_string_or_null(r.get("actDate")),
            as_string_or_null(r.get("shelterName")),
            as_string_or_null(r.get("pvsName")),
        ]
    )
    if not has_any:
        return None

    unit = resolved_orgs.get("unit")
    units = resolved_orgs.get("units") if isinstance(resolved_orgs.get("units"), list) else ([unit] if unit else [])
    receiver_type = detect_animal_receiver_type(r)

    animal = dclone((card_record or {}).get("animal") or {})
    events = dclone((card_record or {}).get("events") or [])
    animal["events"] = events
    catch_address = build_minimal_address(r.get("releaseAddressObj") or r.get("releaseAddress"), row.get("region"))
    if catch_address:
        animal["catchAddress"] = catch_address

    doc = {
        "guid": generate_guid(),
        "parentEntries": RELEASE_COLLECTION,
        "unit": unit,
        "units": units,
        "number": {"number": strip_card_suffix(row.get("cardNumber"))},
        "releaseDate": to_iso_z(r.get("actDate")),
        "releaseActNumber": as_string_or_null(r.get("actNumber")),
        "animalReceiverType": receiver_type,
        "animalReleaseActType": {
            "code": "releaseActTitle",
            "name": as_string_or_null(r.get("actName")) or "Акт выпуска животного без владельца в места обитания",
        },
        "animal": animal,
    }

    shelter_has = any(
        [
            as_string_or_null(r.get("shelterName")),
            as_string_or_null(r.get("shelterOGRN")),
            as_string_or_null(r.get("shelterINN")),
            r.get("shelterAddressObj"),
            as_string_or_null(r.get("shelterAddress")),
        ]
    )
    pvs_has = any(
        [
            as_string_or_null(r.get("pvsName")),
            as_string_or_null(r.get("pvsOGRN")),
            as_string_or_null(r.get("pvsINN")),
            r.get("pvsAddressObj"),
            as_string_or_null(r.get("pvsAddress")),
        ]
    )
    exec_fio = as_string_or_null(r.get("catcherFIO"))

    if shelter_has:
        doc["shelterData"] = {
            "releaseExecutorFullName": exec_fio,
            "shelterName": as_string_or_null(r.get("shelterName")),
            "shelterInn": as_string_or_null(r.get("shelterINN")),
            "shelterOgrn": as_string_or_null(r.get("shelterOGRN")),
            "shelterAddress": build_minimal_address(r.get("shelterAddressObj") or r.get("shelterAddress"), row.get("region")),
            "releaseActFile": [],
        }
    if pvs_has:
        doc["temporaryHoldingFacilityData"] = {
            "releaseExecutorFullName": exec_fio,
            "shelterName": as_string_or_null(r.get("pvsName")),
            "shelterInn": as_string_or_null(r.get("pvsINN")),
            "shelterOgrn": as_string_or_null(r.get("pvsOGRN")),
            "shelterAddress": build_minimal_address(r.get("pvsAddressObj") or r.get("pvsAddress"), row.get("region")),
            "releaseActFile": [],
        }

    return {"record": doc, "pendingUploads": collect_release_act_uploads(row, doc)}


def build_death_act_record(row, resolved_orgs, card_record):
    d = row.get("deathInfo") or None
    if not isinstance(d, dict):
        return None
    has_any = any(
        [
            as_string_or_null(d.get("actName")),
            as_string_or_null(d.get("actNumber")),
            as_string_or_null(d.get("actDate")),
            as_string_or_null(d.get("deathDate")),
            as_string_or_null(d.get("shelterName")),
            as_string_or_null(d.get("pvsName")),
        ]
    )
    if not has_any:
        return None

    unit = resolved_orgs.get("unit")
    units = resolved_orgs.get("units") if isinstance(resolved_orgs.get("units"), list) else ([unit] if unit else [])
    receiver_type = detect_animal_receiver_type(d)

    animal = dclone((card_record or {}).get("animal") or {})
    events = dclone((card_record or {}).get("events") or [])
    animal["events"] = events
    catch_addr = build_minimal_address(
        d.get("shelterAddressObj") or d.get("shelterAddress") or d.get("pvsAddressObj") or d.get("pvsAddress"),
        row.get("region"),
    )
    if catch_addr:
        animal["catchAddress"] = catch_addr

    vet_fio = as_string_or_null(d.get("vetSpecialistFullName") or d.get("vetSpecialistFIO") or d.get("executorFIO") or row.get("vetSpecialistFullName"))
    vet_pos = as_string_or_null(d.get("vetSpecialistPosition") or d.get("vetSpecialistPos") or row.get("vetSpecialistPosition"))

    shelter_has = any(
        [
            as_string_or_null(d.get("shelterName")),
            as_string_or_null(d.get("shelterOGRN")),
            as_string_or_null(d.get("shelterINN")),
            d.get("shelterAddressObj"),
            as_string_or_null(d.get("shelterAddress")),
        ]
    )
    pvs_has = any(
        [
            as_string_or_null(d.get("pvsName")),
            as_string_or_null(d.get("pvsOGRN")),
            as_string_or_null(d.get("pvsINN")),
            d.get("pvsAddressObj"),
            as_string_or_null(d.get("pvsAddress")),
        ]
    )

    doc = {
        "guid": generate_guid(),
        "parentEntries": RELEASE_COLLECTION,
        "unit": unit,
        "units": units,
        "number": {"number": strip_card_suffix(row.get("cardNumber"))},
        "releaseDate": to_iso_z(d.get("actDate")),
        "releaseActNumber": as_string_or_null(d.get("actNumber")),
        "deathDate": to_iso_z(d.get("deathDate")),
        "animalReceiverType": receiver_type,
        "animalReleaseActType": {
            "code": "deathActTitle",
            "name": as_string_or_null(d.get("actName")) or "Акт падежа животного без владельца",
        },
        "animal": animal,
        "actData": {"deathActFile": []},
    }

    if shelter_has:
        doc["actData"]["shelterData"] = {
            "vetSpecialistFullName": vet_fio,
            "vetSpecialistPosition": vet_pos,
            "shelterName": as_string_or_null(d.get("shelterName")),
            "shelterNum": as_string_or_null(d.get("shelterNum")) or "",
            "shelterInn": as_string_or_null(d.get("shelterINN")),
            "shelterOgrn": as_string_or_null(d.get("shelterOGRN")),
            "shelterAddress": build_minimal_address(d.get("shelterAddressObj") or d.get("shelterAddress"), row.get("region")),
        }
    if pvs_has:
        doc["actData"]["temporaryHoldingFacilityData"] = {
            "vetSpecialistFullName": vet_fio,
            "vetSpecialistPosition": vet_pos,
            "shelterName": as_string_or_null(d.get("pvsName")),
            "shelterNum": as_string_or_null(d.get("pvsNum")) or "",
            "shelterInn": as_string_or_null(d.get("pvsINN")),
            "shelterOgrn": as_string_or_null(d.get("pvsOGRN")),
            "shelterAddress": build_minimal_address(d.get("pvsAddressObj") or d.get("pvsAddress"), row.get("region")),
        }

    return {"record": doc, "pendingUploads": collect_death_act_uploads(row)}


def build_transfer_owner_act_record(row, resolved_orgs, card_record):
    t = row.get("transferToOwner") or None
    if not isinstance(t, dict):
        return None
    has_any = any(
        [
            as_string_or_null(t.get("actName")),
            as_string_or_null(t.get("actNumber")),
            as_string_or_null(t.get("transferDate")),
            as_string_or_null(t.get("shelterName")),
            as_string_or_null(t.get("pvsName")),
            as_string_or_null(t.get("newOwnerFIO")),
            as_string_or_null(t.get("newOwnerAddress")),
        ]
    )
    if not has_any:
        return None

    unit = resolved_orgs.get("unit")
    units = resolved_orgs.get("units") if isinstance(resolved_orgs.get("units"), list) else ([unit] if unit else [])
    receiver_type = detect_animal_receiver_type(t)

    animal = dclone((card_record or {}).get("animal") or {})
    events = dclone((card_record or {}).get("events") or [])
    animal["events"] = events
    catch_addr = build_minimal_address(
        t.get("shelterAddressObj") or t.get("shelterAddress") or t.get("pvsAddressObj") or t.get("pvsAddress"),
        row.get("region"),
    )
    if catch_addr:
        animal["catchAddress"] = catch_addr

    vet_fio = as_string_or_null(t.get("vetSpecialistFullName") or t.get("vetSpecialistFIO") or row.get("vetSpecialistFullName"))
    vet_pos = as_string_or_null(t.get("vetSpecialistPosition") or t.get("vetSpecialistPos") or row.get("vetSpecialistPosition"))
    fio = split_fio(t.get("newOwnerFIO"))
    owner_address = build_minimal_address(t.get("newOwnerAddressObj") or t.get("newOwnerAddress"), row.get("region"))

    shelter_has = any(
        [
            as_string_or_null(t.get("shelterName")),
            as_string_or_null(t.get("shelterOGRN")),
            as_string_or_null(t.get("shelterINN")),
            t.get("shelterAddressObj"),
            as_string_or_null(t.get("shelterAddress")),
        ]
    )
    pvs_has = any(
        [
            as_string_or_null(t.get("pvsName")),
            as_string_or_null(t.get("pvsOGRN")),
            as_string_or_null(t.get("pvsINN")),
            t.get("pvsAddressObj"),
            as_string_or_null(t.get("pvsAddress")),
        ]
    )

    doc = {
        "guid": generate_guid(),
        "parentEntries": RELEASE_COLLECTION,
        "unit": unit,
        "units": units,
        "number": {"number": strip_card_suffix(row.get("cardNumber"))},
        "releaseDate": to_iso_z(t.get("transferDate")),
        "releaseActNumber": as_string_or_null(t.get("actNumber")),
        "animalReceiverType": receiver_type,
        "animalReleaseActType": {
            "code": "transferActTitle",
            "name": "Акт передачи животного без владельца прежнему или новому владельцу",
        },
        "animal": animal,
        "actData": {
            "transferActFile": [],
            "newOwnerData": {
                "newOwnerLastName": fio.get("last"),
                "newOwnerFirstName": fio.get("first"),
                "newOwnerPatronymic": fio.get("patronymic"),
                "newOwnerAddress": owner_address,
                "passportData": {
                    "IdentityDocumentSeries": as_string_or_null(t.get("idSeries")),
                    "IdentityDocumentNumber": as_string_or_null(t.get("idNumber")),
                    "IdentityDocumentSubdivisionCode": as_string_or_null(t.get("idDeptCode")),
                    "IdentityDocumentIssueDate": to_iso_z(t.get("idIssueDate")),
                    "IdentityDocumentIssuingAuthority": as_string_or_null(t.get("idIssuedBy")),
                },
            },
        },
    }

    if shelter_has:
        doc["actData"]["shelterData"] = {
            "vetSpecialistFullName": vet_fio,
            "vetSpecialistPosition": vet_pos,
            "shelterName": as_string_or_null(t.get("shelterName")),
            "shelterNum": as_string_or_null(t.get("shelterNum")) or "",
            "shelterInn": as_string_or_null(t.get("shelterINN")),
            "shelterOgrn": as_string_or_null(t.get("shelterOGRN")),
            "shelterAddress": build_minimal_address(t.get("shelterAddressObj") or t.get("shelterAddress"), row.get("region")),
        }
    if pvs_has:
        doc["actData"]["temporaryHoldingFacilityData"] = {
            "vetSpecialistFullName": vet_fio,
            "vetSpecialistPosition": vet_pos,
            "shelterName": as_string_or_null(t.get("pvsName")),
            "shelterNum": as_string_or_null(t.get("pvsNum")) or "",
            "shelterInn": as_string_or_null(t.get("pvsINN")),
            "shelterOgrn": as_string_or_null(t.get("pvsOGRN")),
            "shelterAddress": build_minimal_address(t.get("pvsAddressObj") or t.get("pvsAddress"), row.get("region")),
        }

    return {"record": doc, "pendingUploads": collect_transfer_act_uploads(row)}


def pick_animal_mini(card_animal, row):
    a = card_animal or {}
    out = {}
    if a.get("sex"):
        out["sex"] = a.get("sex")
    if a.get("type"):
        out["type"] = a.get("type")
    if a.get("breed") or as_string_or_null(row.get("breed")):
        out["breed"] = a.get("breed") or as_string_or_null(row.get("breed"))
    if a.get("coloration") or as_string_or_null(row.get("coloration")):
        out["coloration"] = a.get("coloration") or as_string_or_null(row.get("coloration"))
    if a.get("specialMarks") or as_string_or_null(row.get("specialMarks")):
        out["specialMarks"] = a.get("specialMarks") or as_string_or_null(row.get("specialMarks"))
    if a.get("approximateAge") or as_string_or_null(row.get("age")):
        out["approximateAge"] = a.get("approximateAge") or as_string_or_null(row.get("age"))
    return out


def build_animal_shelter_from_card(card_record):
    animal = dclone((card_record or {}).get("animal") or {})
    events = dclone((card_record or {}).get("events") or [])
    animal["events"] = events
    return animal


def build_handover_with_catcher_record(row, resolved_orgs, card_record):
    h = row.get("handoverWithCatcher") or None
    if not isinstance(h, dict):
        return None

    has_any = any(
        [
            as_string_or_null(h.get("actName")),
            as_string_or_null(h.get("actNumber")),
            as_string_or_null(h.get("orderNumber")),
            as_string_or_null(h.get("orderCreateDate")),
            as_string_or_null(h.get("catcherFIO")),
            as_string_or_null(h.get("shelterName")),
            as_string_or_null(h.get("pvsName")),
        ]
    )
    if not has_any:
        return None

    unit = resolved_orgs.get("unit")
    units = resolved_orgs.get("units") if isinstance(resolved_orgs.get("units"), list) else ([unit] if unit else [])
    receiver_code = detect_receiver_type_last_wins(h)
    date_act_mun = to_iso_z(h.get("actDate") or h.get("dateActMun") or h.get("actCreateDate") or h.get("orderCreateDate"))
    vet_fio = as_string_or_null(row.get("vetSpecialistFullName") or h.get("vetSpecialistFullName"))
    vet_pos = as_string_or_null(row.get("vetSpecialistPosition") or h.get("vetSpecialistPosition"))

    doc = {
        "guid": generate_guid(),
        "parentEntries": TRANSFER_ACT_COLLECTION,
        "unit": unit,
        "units": units,
        "animal": pick_animal_mini((card_record or {}).get("animal") or {}, row),
        "dataAppeal": {
            "dateActMun": date_act_mun,
            "requestNumber": as_string_or_null(h.get("orderNumber")),
            "requestCreationDate": to_iso_z(h.get("orderCreateDate")),
            "hunterFullName": as_string_or_null(h.get("catcherFIO")),
            "hunterPhoneNumber": format_phone_ru(h.get("catcherPhone")),
            "transferActNumber": as_string_or_null(h.get("actNumber")),
        },
        "transferActFile": [],
        "transferActType": {"code": "actWithHunter", "name": "Акт приёма-передачи с ловцом"},
    }
    if receiver_code:
        doc["animalReceiverType"] = {"code": receiver_code}

    shelter_has = any([as_string_or_null(h.get("shelterName")), as_string_or_null(h.get("shelterAddress")), as_string_or_null(h.get("shelterPhone"))])
    pvs_has = any([as_string_or_null(h.get("pvsName")), as_string_or_null(h.get("pvsAddress")), as_string_or_null(h.get("pvsPhone"))])

    if shelter_has:
        doc["dataAppeal"]["shelterData"] = {
            "shelterNum": "",
            "shelterName": as_string_or_null(h.get("shelterName")),
            "shelterInn": as_string_or_null(h.get("shelterINN")),
            "shelterOgrn": as_string_or_null(h.get("shelterOGRN")),
            "shelterAddress": build_minimal_address(h.get("shelterAddressObj") or h.get("shelterAddress"), row.get("region")),
            "vetSpecialistFullName": vet_fio,
            "vetSpecialistPosition": vet_pos,
        }
    if pvs_has:
        doc["dataAppeal"]["temporaryHoldingFacilityData"] = {
            "shelterNum": "",
            "shelterName": as_string_or_null(h.get("pvsName")),
            "shelterInn": as_string_or_null(h.get("pvsINN")),
            "shelterOgrn": as_string_or_null(h.get("pvsOGRN")),
            "shelterAddress": build_minimal_address(h.get("pvsAddressObj") or h.get("pvsAddress"), row.get("region")),
            "vetSpecialistFullName": vet_fio,
            "vetSpecialistPosition": vet_pos,
        }
    return {"record": doc, "pendingUploads": collect_handover_with_catcher_act_uploads(row)}


def build_handover_with_shelter_record(row, resolved_orgs, card_record):
    h = row.get("handoverWithShelter") or None
    if not isinstance(h, dict):
        return None

    has_any = any(
        [
            as_string_or_null(h.get("actName")),
            as_string_or_null(h.get("actNumber")),
            as_string_or_null(h.get("actDate")),
            as_string_or_null(h.get("shelterName")),
            as_string_or_null(h.get("pvsName")),
        ]
    )
    if not has_any:
        return None

    unit = resolved_orgs.get("unit")
    units = resolved_orgs.get("units") if isinstance(resolved_orgs.get("units"), list) else ([unit] if unit else [])
    thf_name = (
        as_string_or_null(h.get("pvsName"))
        or as_string_or_null((unit or {}).get("name"))
        or as_string_or_null(row.get("authorizedOrgName"))
        or as_string_or_null((DEFAULT_ORG or {}).get("name"))
    )
    trapping_org_name = as_string_or_null(h.get("shelterName"))

    doc = {
        "guid": generate_guid(),
        "parentEntries": TRANSFER_ACT_COLLECTION,
        "unit": unit,
        "units": units,
        "number": {"number": strip_card_suffix(row.get("cardNumber"))},
        "actData": {
            "thfName": thf_name,
            "transferActDate": to_iso_z(h.get("actDate")),
            "transferActNumber": as_string_or_null(h.get("actNumber")),
            "trappingOrganization": {
                "name": trapping_org_name,
                "shortName": trapping_org_name,
            },
        },
        "animalShelter": build_animal_shelter_from_card(card_record),
        "transferActFile": [],
        "transferActType": {"code": "actWithShelter", "name": "Акт приёма-передачи с приютом"},
    }
    return {"record": doc, "pendingUploads": collect_handover_with_shelter_act_uploads(row)}


def process_card_rows(session, logger, success_logger, fail_logger, rows, rollback_candidates, ok, fail):
    logger.info("[CARD] start rows=%s", len(rows))
    for i, row in enumerate(rows):
        row_num = int(row.get("__row_num", i + 1))
        card_number = row.get("cardNumber")
        try:
            validate_card_row_before_create(row)
            resolved = resolve_orgs_for_card_row(session, logger, row)
            if not resolved.get("authFound") and not resolved.get("shelterFound") and not DEFAULT_ORG_ENABLED:
                append_error(
                    fail_logger,
                    fail,
                    {
                        "registry": "card",
                        "index": row_num,
                        "cardNumber": card_number,
                        "stage": "org-lookup",
                        "error": "no-organizations-found",
                    },
                )
                continue

            built = build_card_record(row, resolved, logger)
            created = create_record(session, logger, CARD_COLLECTION, built["record"])
            main_id = created.get("_id")
            guid = created.get("guid") or built["record"]["guid"]
            if not main_id or not guid:
                raise RuntimeError("No _id/guid for created animal card")

            append_success(
                success_logger,
                rollback_candidates,
                ok,
                {
                    "registry": "card",
                    "index": row_num,
                    "cardNumber": card_number,
                    "_id": main_id,
                    "guid": guid,
                    "parentEntries": CARD_COLLECTION,
                },
            )

            card_doc = created
            if ENABLE_FILE_UPLOADS and built["pendingUploads"]:
                if DRY_RUN_LOG_UPLOAD_TARGETS:
                    for u in built["pendingUploads"]:
                        logger.warning("[CARD][DRY] upload planned: %s", {"path": u["path"], "filename": u["filename"]})
                else:
                    try:
                        card_doc = apply_uploads_to_doc(
                            session=session,
                            logger=logger,
                            collection=CARD_COLLECTION,
                            main_id=main_id,
                            guid=guid,
                            doc_state=card_doc,
                            pending_uploads=built["pendingUploads"],
                        )
                    except Exception as upload_exc:
                        append_error(
                            fail_logger,
                            fail,
                            {
                                "registry": "card",
                                "index": row_num,
                                "cardNumber": card_number,
                                "stage": "upload/link",
                                "error": serialize_exception(upload_exc),
                            },
                        )
                        log_processing_exception(logger, "[CARD] upload error:", row_num, upload_exc)
                        if _is_upload_error(upload_exc) and STOP_ON_FIRST_FATAL_UPLOAD and not RUNTIME_OPERATOR_MODE:
                            return True

            if row.get("handoverWithCatcher"):
                try:
                    built_h = build_handover_with_catcher_record(row, resolved, card_doc)
                    if built_h:
                        created_h = create_record(session, logger, TRANSFER_ACT_COLLECTION, built_h["record"])
                        h_id = created_h.get("_id")
                        h_guid = created_h.get("guid") or built_h["record"]["guid"]
                        if not h_id or not h_guid:
                            raise RuntimeError("No _id/guid for created handover-with-catcher act")
                        append_success(
                            success_logger,
                            rollback_candidates,
                            ok,
                            {
                                "registry": "handover-with-catcher",
                                "index": row_num,
                                "cardNumber": card_number,
                                "_id": h_id,
                                "guid": h_guid,
                                "parentEntries": TRANSFER_ACT_COLLECTION,
                            },
                        )
                        if ENABLE_FILE_UPLOADS and built_h["pendingUploads"]:
                            apply_uploads_to_doc(
                                session=session,
                                logger=logger,
                                collection=TRANSFER_ACT_COLLECTION,
                                main_id=h_id,
                                guid=h_guid,
                                doc_state=created_h,
                                pending_uploads=built_h["pendingUploads"],
                            )
                except Exception as sub_exc:
                    append_error(
                        fail_logger,
                        fail,
                        {
                            "registry": "card",
                            "index": row_num,
                            "cardNumber": card_number,
                            "stage": "handover-with-catcher",
                            "error": serialize_exception(sub_exc),
                        },
                    )
                    log_processing_exception(logger, "[CARD] handover-with-catcher error:", row_num, sub_exc)
                    if _is_upload_error(sub_exc) and STOP_ON_FIRST_FATAL_UPLOAD and not RUNTIME_OPERATOR_MODE:
                        return True

            if row.get("handoverWithShelter"):
                try:
                    built_hs = build_handover_with_shelter_record(row, resolved, card_doc)
                    if built_hs:
                        created_hs = create_record(session, logger, TRANSFER_ACT_COLLECTION, built_hs["record"])
                        hs_id = created_hs.get("_id")
                        hs_guid = created_hs.get("guid") or built_hs["record"]["guid"]
                        if not hs_id or not hs_guid:
                            raise RuntimeError("No _id/guid for created handover-with-shelter act")
                        append_success(
                            success_logger,
                            rollback_candidates,
                            ok,
                            {
                                "registry": "handover-with-shelter",
                                "index": row_num,
                                "cardNumber": card_number,
                                "_id": hs_id,
                                "guid": hs_guid,
                                "parentEntries": TRANSFER_ACT_COLLECTION,
                            },
                        )
                        if ENABLE_FILE_UPLOADS and built_hs["pendingUploads"]:
                            apply_uploads_to_doc(
                                session=session,
                                logger=logger,
                                collection=TRANSFER_ACT_COLLECTION,
                                main_id=hs_id,
                                guid=hs_guid,
                                doc_state=created_hs,
                                pending_uploads=built_hs["pendingUploads"],
                            )
                except Exception as sub_exc:
                    append_error(
                        fail_logger,
                        fail,
                        {
                            "registry": "card",
                            "index": row_num,
                            "cardNumber": card_number,
                            "stage": "handover-with-shelter",
                            "error": serialize_exception(sub_exc),
                        },
                    )
                    log_processing_exception(logger, "[CARD] handover-with-shelter error:", row_num, sub_exc)
                    if _is_upload_error(sub_exc) and STOP_ON_FIRST_FATAL_UPLOAD and not RUNTIME_OPERATOR_MODE:
                        return True

            if row.get("releaseInfo"):
                try:
                    built_release = build_release_act_record(row, resolved, card_doc)
                    if built_release:
                        created_release = create_record(session, logger, RELEASE_COLLECTION, built_release["record"])
                        release_id = created_release.get("_id")
                        release_guid = created_release.get("guid") or built_release["record"]["guid"]
                        if not release_id or not release_guid:
                            raise RuntimeError("No _id/guid for created release act")
                        append_success(
                            success_logger,
                            rollback_candidates,
                            ok,
                            {
                                "registry": "release-act",
                                "index": row_num,
                                "cardNumber": card_number,
                                "_id": release_id,
                                "guid": release_guid,
                                "parentEntries": RELEASE_COLLECTION,
                            },
                        )
                        if ENABLE_FILE_UPLOADS and built_release["pendingUploads"]:
                            apply_uploads_to_doc(
                                session=session,
                                logger=logger,
                                collection=RELEASE_COLLECTION,
                                main_id=release_id,
                                guid=release_guid,
                                doc_state=created_release,
                                pending_uploads=built_release["pendingUploads"],
                            )

                        refreshed_card = _fetch_record_by_id(session, logger, CARD_COLLECTION, main_id)
                        if isinstance(refreshed_card, dict):
                            card_doc = refreshed_card
                        card_doc.setdefault("releaseInfo", {})
                        card_doc["releaseInfo"]["releaseInfoRecordLink"] = ui_release_link(release_id)
                        card_doc = update_record(session, logger, CARD_COLLECTION, main_id, guid, card_doc)
                except Exception as sub_exc:
                    append_error(
                        fail_logger,
                        fail,
                        {
                            "registry": "card",
                            "index": row_num,
                            "cardNumber": card_number,
                            "stage": "release-act-create/link",
                            "error": serialize_exception(sub_exc),
                        },
                    )
                    log_processing_exception(logger, "[CARD] release-act error:", row_num, sub_exc)
                    if _is_upload_error(sub_exc) and STOP_ON_FIRST_FATAL_UPLOAD and not RUNTIME_OPERATOR_MODE:
                        return True

            if row.get("deathInfo"):
                try:
                    built_death = build_death_act_record(row, resolved, card_doc)
                    if built_death:
                        created_death = create_record(session, logger, RELEASE_COLLECTION, built_death["record"])
                        death_id = created_death.get("_id")
                        death_guid = created_death.get("guid") or built_death["record"]["guid"]
                        if not death_id or not death_guid:
                            raise RuntimeError("No _id/guid for created death act")
                        append_success(
                            success_logger,
                            rollback_candidates,
                            ok,
                            {
                                "registry": "death-act",
                                "index": row_num,
                                "cardNumber": card_number,
                                "_id": death_id,
                                "guid": death_guid,
                                "parentEntries": RELEASE_COLLECTION,
                            },
                        )
                        if ENABLE_FILE_UPLOADS and built_death["pendingUploads"]:
                            apply_uploads_to_doc(
                                session=session,
                                logger=logger,
                                collection=RELEASE_COLLECTION,
                                main_id=death_id,
                                guid=death_guid,
                                doc_state=created_death,
                                pending_uploads=built_death["pendingUploads"],
                            )
                except Exception as sub_exc:
                    append_error(
                        fail_logger,
                        fail,
                        {
                            "registry": "card",
                            "index": row_num,
                            "cardNumber": card_number,
                            "stage": "death-act-create",
                            "error": serialize_exception(sub_exc),
                        },
                    )
                    log_processing_exception(logger, "[CARD] death-act error:", row_num, sub_exc)
                    if _is_upload_error(sub_exc) and STOP_ON_FIRST_FATAL_UPLOAD and not RUNTIME_OPERATOR_MODE:
                        return True

            if row.get("transferToOwner"):
                try:
                    built_transfer = build_transfer_owner_act_record(row, resolved, card_doc)
                    if built_transfer:
                        created_transfer = create_record(session, logger, RELEASE_COLLECTION, built_transfer["record"])
                        transfer_id = created_transfer.get("_id")
                        transfer_guid = created_transfer.get("guid") or built_transfer["record"]["guid"]
                        if not transfer_id or not transfer_guid:
                            raise RuntimeError("No _id/guid for created transfer-to-owner act")
                        append_success(
                            success_logger,
                            rollback_candidates,
                            ok,
                            {
                                "registry": "transfer-owner-act",
                                "index": row_num,
                                "cardNumber": card_number,
                                "_id": transfer_id,
                                "guid": transfer_guid,
                                "parentEntries": RELEASE_COLLECTION,
                            },
                        )
                        if ENABLE_FILE_UPLOADS and built_transfer["pendingUploads"]:
                            apply_uploads_to_doc(
                                session=session,
                                logger=logger,
                                collection=RELEASE_COLLECTION,
                                main_id=transfer_id,
                                guid=transfer_guid,
                                doc_state=created_transfer,
                                pending_uploads=built_transfer["pendingUploads"],
                            )

                        refreshed_card = _fetch_record_by_id(session, logger, CARD_COLLECTION, main_id)
                        if isinstance(refreshed_card, dict):
                            card_doc = refreshed_card
                        card_doc.setdefault("transferToOwner", {})
                        card_doc["transferToOwner"]["transferToOwnerRecordLink"] = ui_release_link(transfer_id)
                        card_doc = update_record(session, logger, CARD_COLLECTION, main_id, guid, card_doc)
                except Exception as sub_exc:
                    append_error(
                        fail_logger,
                        fail,
                        {
                            "registry": "card",
                            "index": row_num,
                            "cardNumber": card_number,
                            "stage": "transfer-act-create/link",
                            "error": serialize_exception(sub_exc),
                        },
                    )
                    log_processing_exception(logger, "[CARD] transfer-to-owner act error:", row_num, sub_exc)
                    if _is_upload_error(sub_exc) and STOP_ON_FIRST_FATAL_UPLOAD and not RUNTIME_OPERATOR_MODE:
                        return True

        except Exception as exc:
            append_error(
                fail_logger,
                fail,
                {"registry": "card", "index": row_num, "cardNumber": card_number, "error": serialize_exception(exc)},
            )
            log_processing_exception(logger, "[CARD] row error:", row_num, exc)
            if _is_upload_error(exc):
                if STOP_ON_FIRST_FATAL_UPLOAD and not RUNTIME_OPERATOR_MODE:
                    return True
                continue
            if STOP_ON_FIRST_CREATE_ERROR and not RUNTIME_OPERATOR_MODE:
                return True
    return False


def _norm_path(base_dir: str, raw_path: str) -> str:
    p = str(raw_path or "").strip()
    if not p:
        return os.path.abspath(base_dir)
    if os.path.isabs(p):
        return os.path.abspath(p)
    return os.path.abspath(os.path.join(base_dir, p))


def _with_row_numbers(rows: List[Dict[str, Any]], start_row: int) -> List[Dict[str, Any]]:
    out = []
    for idx, row in enumerate(rows, start=start_row):
        item = dict(row or {})
        item["__row_num"] = idx
        out.append(item)
    return out


def _apply_limit(rows: List[Dict[str, Any]], limit: int) -> List[Dict[str, Any]]:
    if limit <= 0:
        return rows
    return rows[:limit]


def _append_unique_rollback(rollback_candidates: List[Dict[str, Any]], item: Dict[str, Any]):
    key = (str(item.get("_id") or ""), str(item.get("guid") or ""), str(item.get("parentEntries") or ""))
    if not all(key):
        return
    for existing in rollback_candidates:
        ex_key = (
            str(existing.get("_id") or ""),
            str(existing.get("guid") or ""),
            str(existing.get("parentEntries") or ""),
        )
        if ex_key == key:
            return
    rollback_candidates.append({"_id": key[0], "guid": key[1], "parentEntries": key[2]})


def _search_exists_by_id(session, logger, collection: str, main_id: str) -> bool:
    body = {
        "search": {
            "search": [
                {"andSubConditions": [{"field": "_id", "operator": "eq", "value": str(main_id)}]},
            ]
        }
    }
    data = search_collection(session, logger, collection, body)
    content = data.get("content") if isinstance(data, dict) else []
    return isinstance(content, list) and len(content) > 0


def _fetch_record_by_id(session, logger, collection: str, main_id: str):
    body = {
        "search": {
            "search": [
                {"andSubConditions": [{"field": "_id", "operator": "eq", "value": str(main_id)}]},
            ]
        }
    }
    data = search_collection(session, logger, collection, body)
    content = data.get("content") if isinstance(data, dict) else []
    if isinstance(content, list) and content:
        return content[0]
    return None


def _setup_runtime_profile(args) -> str:
    profile_name = str(args.profile or "custom").strip().lower()
    base_url = BASE_URL
    jwt_url = JWT_URL

    if profile_name in PROFILES and profile_name != "custom":
        profile = PROFILES[profile_name]
        base_url = profile.base_url
        jwt_url = profile.jwt_url

    if str(args.base_url or "").strip():
        base_url = str(args.base_url).strip()
    if str(args.jwt_url or "").strip():
        jwt_url = str(args.jwt_url).strip()

    if not jwt_url:
        jwt_url = base_url.rstrip("/") + "/jwt/"

    set_runtime_urls(base_url=base_url, jwt_url=jwt_url)

    normalized = str(base_url or "").strip().rstrip("/")
    for name, profile in PROFILES.items():
        if str(profile.base_url).strip().rstrip("/") == normalized:
            return name
    return "custom"


def _parse_args():
    parser = argparse.ArgumentParser(description="Pet migration runner (Excel/JSON -> PGS).")
    parser.add_argument("--profile", choices=["custom", "dev", "psi", "prod"], default="dev")
    parser.add_argument("--base-url", default="", help="Override stand base URL.")
    parser.add_argument("--jwt-url", default="", help="Override JWT page URL used for auto refresh.")
    parser.add_argument("--mode", choices=["auto", "single", "mass"], default="auto")
    parser.add_argument("--workbooks", default="", help="Explicit workbook list (separator ';' or new line).")
    parser.add_argument("--files-map", default="", help="Workbook->files folder map: 'book1.xlsm=one;book2.xlsm=two'.")
    parser.add_argument("--ask-files-always", action="store_true", help="Always ask files folder for each selected workbook.")
    parser.add_argument("--dry-run", action="store_true", help="Only parse input and log summaries, no create/update requests.")
    parser.add_argument("--auth-only", action="store_true", help="Validate auth and exit.")
    parser.add_argument("--skip-auth", action="store_true", help="Skip auth. Allowed only with --dry-run.")
    parser.add_argument("--no-prompt", action="store_true", help="Read cookie/token from files only.")
    parser.add_argument(
        "--no-interactive",
        action="store_true",
        help="Disable all interactive prompts (workbooks/files/operator decisions).",
    )
    parser.add_argument(
        "--operator-mode",
        action="store_true",
        help="Force operator decisions mode (retry/skip/abort prompts on row errors).",
    )
    parser.add_argument("--limit", type=int, default=0, help="Limit rows per registry job (0 = no limit).")

    resume_group = parser.add_mutually_exclusive_group()
    resume_group.add_argument("--resume", dest="resume", action="store_true", help="Force enable resume checkpoints.")
    resume_group.add_argument("--no-resume", dest="resume", action="store_false", help="Disable resume checkpoints.")
    parser.set_defaults(resume=None)

    parser.add_argument("--reset-state", action="store_true", help="Reset resume namespace before run.")
    parser.add_argument("--state-file", default="", help="Override checkpoint file path.")
    return parser.parse_args()


def _process_job_with_resume(
    *,
    session,
    logger,
    success_logger,
    fail_logger,
    process_fn,
    rows: List[Dict[str, Any]],
    created_items: List[Dict[str, Any]],
    errors: List[Dict[str, Any]],
    rollback_candidates: List[Dict[str, Any]],
    state: ResumeState,
    workbook_key: str,
    workbook_label: str,
    job_name: str,
    primary_registry: str,
    primary_collection: str,
    dry_run: bool,
    interactive: bool,
) -> bool:
    pending = []
    for row in rows:
        row_num = int(row.get("__row_num", 0) or 0)
        state_row = state.get(workbook_key, job_name, row_num)
        if state_row and not dry_run and VERIFY_CREATED and session is not None:
            state_collection = str(state_row.get("collection") or primary_collection)
            state_id = str(state_row.get("_id") or "")
            if not state_id:
                logger.warning("[RESUME] stale checkpoint without _id, reprocessing: %s %s row=%s", workbook_label, job_name, row_num)
                state.clear_row(workbook_key, job_name, row_num)
                state_row = None
            else:
                try:
                    exists = _search_exists_by_id(session, logger, state_collection, state_id)
                    if not exists:
                        logger.warning(
                            "[RESUME] checkpoint target missing, reprocessing: %s %s row=%s _id=%s",
                            workbook_label,
                            job_name,
                            row_num,
                            state_id,
                        )
                        state.clear_row(workbook_key, job_name, row_num)
                        state_row = None
                except Exception as exc:
                    logger.warning(
                        "[RESUME] pre-check failed, keep checkpoint: %s %s row=%s err=%s",
                        workbook_label,
                        job_name,
                        row_num,
                        exc,
                    )
        if state_row:
            resumed = {
                "workbook": workbook_label,
                "job": job_name,
                "index": row_num,
                "_id": state_row.get("_id"),
                "guid": state_row.get("guid"),
                "parentEntries": state_row.get("collection"),
                "registry": primary_registry,
                "resumed": True,
            }
            created_items.append(resumed)
            _append_unique_rollback(rollback_candidates, resumed)
        else:
            pending.append(row)

    if not pending:
        return False

    def _mark_new_successes(start_idx: int):
        if dry_run:
            return
        for item in created_items[start_idx:]:
            if str(item.get("registry") or "") != primary_registry:
                continue
            main_id = item.get("_id")
            guid = item.get("guid")
            row_num = item.get("index")
            if not main_id or not guid or row_num is None:
                continue
            try:
                row_idx_int = int(row_num)
            except Exception:
                continue
            state.mark_success(
                workbook_path=workbook_key,
                job_name=job_name,
                row_idx=row_idx_int,
                collection=primary_collection,
                main_id=str(main_id),
                guid=str(guid),
            )

    def _operator_row_action(row_num: int, exc_msg: str) -> str:
        while True:
            try:
                raw = input(
                    "[OPERATOR] %s/%s row=%s failed: %s\nActions: [R]etry / [S]kip / [A]bort: "
                    % (workbook_label, job_name, row_num, exc_msg)
                ).strip().lower()
            except EOFError:
                return "abort"
            if raw in {"r", "retry"}:
                return "retry"
            if raw in {"s", "skip"}:
                return "skip"
            if raw in {"a", "abort"}:
                return "abort"

    def _operator_row_post_error_action(row_num: int, exc_msg: str) -> str:
        while True:
            try:
                raw = input(
                    "[OPERATOR] %s/%s row=%s has non-fatal error: %s\nActions: [C]ontinue / [A]bort: "
                    % (workbook_label, job_name, row_num, exc_msg)
                ).strip().lower()
            except EOFError:
                return "abort"
            if raw in {"", "c", "continue"}:
                return "continue"
            if raw in {"a", "abort"}:
                return "abort"

    if RUNTIME_OPERATOR_MODE and interactive:
        for row in pending:
            row_num = int(row.get("__row_num", 0) or 0)
            while True:
                created_before = len(created_items)
                errors_before = len(errors)
                stopped = process_fn(
                    session=session,
                    logger=logger,
                    success_logger=success_logger,
                    fail_logger=fail_logger,
                    rows=[row],
                    rollback_candidates=rollback_candidates,
                    ok=created_items,
                    fail=errors,
                )
                _mark_new_successes(created_before)
                if stopped:
                    if len(errors) > errors_before:
                        last_err = errors[-1]
                        action = _operator_row_action(row_num, str(last_err.get("error") or "unknown-error"))
                        if action == "retry":
                            del errors[errors_before:]
                            continue
                        if action == "skip":
                            break
                        return True
                        return True

                row_success = False
                for item in created_items[created_before:]:
                    if str(item.get("registry") or "") != primary_registry:
                        continue
                    try:
                        if int(item.get("index")) == row_num:
                            row_success = True
                            break
                    except Exception:
                        continue

                if len(errors) > errors_before:
                    last_err = errors[-1]
                    if row_success:
                        action = _operator_row_post_error_action(row_num, str(last_err.get("error") or "unknown-error"))
                        if action == "abort":
                            return True
                        break
                    action = _operator_row_action(row_num, str(last_err.get("error") or "unknown-error"))
                    if action == "retry":
                        del errors[errors_before:]
                        continue
                    if action == "skip":
                        break
                    return True
                if row_success:
                    break
                break
        return False

    before_len = len(created_items)
    stopped = process_fn(
        session=session,
        logger=logger,
        success_logger=success_logger,
        fail_logger=fail_logger,
        rows=pending,
        rollback_candidates=rollback_candidates,
        ok=created_items,
        fail=errors,
    )
    _mark_new_successes(before_len)
    return stopped


def main():
    warnings.filterwarnings("ignore", category=InsecureRequestWarning)
    args = _parse_args()
    interactive = not args.no_interactive

    global RUNTIME_OPERATOR_MODE
    # Interactive runs default to operator workflow so the user can decide
    # retry/skip/abort without restarting migration.
    RUNTIME_OPERATOR_MODE = bool(args.operator_mode or interactive)

    logger = setup_logger()
    success_logger = setup_success_logger()
    fail_logger = setup_fail_logger()

    effective_profile = _setup_runtime_profile(args)
    logger.info(
        "Mode: mode=%s dry_run=%s auth_only=%s skip_auth=%s operator_mode=%s limit=%s profile=%s base_url=%s",
        args.mode,
        args.dry_run,
        args.auth_only,
        args.skip_auth,
        RUNTIME_OPERATOR_MODE,
        args.limit,
        effective_profile,
        get_runtime_base_url(),
    )

    if args.skip_auth and not args.dry_run:
        logger.error("--skip-auth can be used only with --dry-run")
        return 1
    if args.auth_only and args.skip_auth:
        logger.error("--auth-only cannot be used together with --skip-auth")
        return 1

    # `--no-prompt` affects only auth (cookie/token) input.
    # Workbook/files/operator prompts are controlled by `--no-interactive`.
    run_specs = []
    if USE_EXCEL_INPUT:
        run_specs = resolve_workbook_specs(
            mode=args.mode,
            workbook_paths_arg=args.workbooks,
            files_map_arg=args.files_map,
            interactive=interactive,
            ask_files_always=bool(args.ask_files_always),
        )
        logger.info("Selected workbooks: %s", [x.workbook_path for x in run_specs] or "-")
        for spec in run_specs:
            logger.info("Workbook files mapping: %s -> %s", os.path.basename(spec.workbook_path), spec.files_dir)

    state_enabled = False if args.dry_run else (RESUME_BY_DEFAULT if args.resume is None else bool(args.resume))
    state_path = _norm_path(SCRIPT_DIR, args.state_file or STATE_FILE)
    resume_namespace = "%s|%s|%s" % (effective_profile, get_runtime_base_url(), "universal-my-pet")
    state = ResumeState(path=Path(state_path), namespace=resume_namespace, enabled=state_enabled)
    if args.reset_state:
        state.reset_namespace()
        logger.info("State reset completed for namespace")
    logger.info("Resume: enabled=%s state_file=%s namespace=%s", state_enabled, state_path, resume_namespace)

    session = None
    if not args.skip_auth and not args.dry_run:
        session = setup_session(
            logger,
            no_prompt=args.no_prompt,
            operator_mode=RUNTIME_OPERATOR_MODE,
        )
        if not session:
            logger.error("Authorization failed, exiting")
            return 1
    elif args.skip_auth:
        logger.info("Auth skipped (--skip-auth)")

    if args.auth_only:
        logger.info("Auth-only mode completed")
        return 0

    created_items = []
    errors = []
    rollback_candidates = []
    stopped = False

    if args.dry_run:
        if run_specs:
            for spec in run_specs:
                set_active_files_dir(spec.files_dir)
                parsed = load_rows_from_excel(spec.workbook_path, logger)
                catch_rows = _apply_limit(_with_row_numbers(parsed.get("catch") or [], EXCEL_DATA_START_ROW), args.limit)
                stray_rows = _apply_limit(_with_row_numbers(parsed.get("stray") or [], EXCEL_DATA_START_ROW), args.limit)
                card_rows = _apply_limit(_with_row_numbers(parsed.get("card") or [], EXCEL_DATA_START_ROW), args.limit)
                logger.info(
                    "[DRY] workbook=%s rows: catch=%s stray=%s card=%s files_dir=%s",
                    os.path.basename(spec.workbook_path),
                    len(catch_rows),
                    len(stray_rows),
                    len(card_rows),
                    spec.files_dir,
                )
        else:
            catch_files = discover_input_files(CATCH_PART_GLOB)
            stray_files = discover_input_files(STRAY_PART_GLOB)
            card_files = discover_input_files(CARD_PART_GLOB)
            catch_rows = _apply_limit(_with_row_numbers(load_rows_from_files(catch_files, logger, "CATCH"), 1), args.limit) if catch_files else []
            stray_rows = _apply_limit(_with_row_numbers(load_rows_from_files(stray_files, logger, "STRAY"), 1), args.limit) if stray_files else []
            card_rows = _apply_limit(_with_row_numbers(load_rows_from_files(card_files, logger, "CARD"), 1), args.limit) if card_files else []
            logger.info("[DRY] json rows: catch=%s stray=%s card=%s", len(catch_rows), len(stray_rows), len(card_rows))
        logger.info("Dry-run completed")
        return 0

    if run_specs:
        for spec in run_specs:
            if stopped:
                break
            set_active_files_dir(spec.files_dir)
            logger.info("Processing workbook: %s", spec.workbook_path)
            parsed = load_rows_from_excel(spec.workbook_path, logger)
            catch_rows = _apply_limit(_with_row_numbers(parsed.get("catch") or [], EXCEL_DATA_START_ROW), args.limit)
            stray_rows = _apply_limit(_with_row_numbers(parsed.get("stray") or [], EXCEL_DATA_START_ROW), args.limit)
            card_rows = _apply_limit(_with_row_numbers(parsed.get("card") or [], EXCEL_DATA_START_ROW), args.limit)
            logger.info(
                "[INPUT] workbook=%s rows summary: catch=%s stray=%s card=%s",
                os.path.basename(spec.workbook_path),
                len(catch_rows),
                len(stray_rows),
                len(card_rows),
            )

            workbook_key = os.path.abspath(spec.workbook_path)
            workbook_label = os.path.basename(spec.workbook_path)

            if catch_rows and not stopped:
                stopped = _process_job_with_resume(
                    session=session,
                    logger=logger,
                    success_logger=success_logger,
                    fail_logger=fail_logger,
                    process_fn=process_order_rows,
                    rows=catch_rows,
                    created_items=created_items,
                    errors=errors,
                    rollback_candidates=rollback_candidates,
                    state=state,
                    workbook_key=workbook_key,
                    workbook_label=workbook_label,
                    job_name="catch_orders",
                    primary_registry="catch-order",
                    primary_collection=ORDER_COLLECTION,
                    dry_run=False,
                    interactive=interactive,
                )
            if stray_rows and not stopped:
                stopped = _process_job_with_resume(
                    session=session,
                    logger=logger,
                    success_logger=success_logger,
                    fail_logger=fail_logger,
                    process_fn=process_stray_rows,
                    rows=stray_rows,
                    created_items=created_items,
                    errors=errors,
                    rollback_candidates=rollback_candidates,
                    state=state,
                    workbook_key=workbook_key,
                    workbook_label=workbook_label,
                    job_name="stray_animals",
                    primary_registry="stray",
                    primary_collection=TARGET_COLLECTION,
                    dry_run=False,
                    interactive=interactive,
                )
            if card_rows and not stopped:
                stopped = _process_job_with_resume(
                    session=session,
                    logger=logger,
                    success_logger=success_logger,
                    fail_logger=fail_logger,
                    process_fn=process_card_rows,
                    rows=card_rows,
                    created_items=created_items,
                    errors=errors,
                    rollback_candidates=rollback_candidates,
                    state=state,
                    workbook_key=workbook_key,
                    workbook_label=workbook_label,
                    job_name="animal_cards",
                    primary_registry="card",
                    primary_collection=CARD_COLLECTION,
                    dry_run=False,
                    interactive=interactive,
                )
    else:
        set_active_files_dir(FILES_DIR)
        catch_files = discover_input_files(CATCH_PART_GLOB)
        stray_files = discover_input_files(STRAY_PART_GLOB)
        card_files = discover_input_files(CARD_PART_GLOB)
        logger.info("[INPUT] catch files: %s", [os.path.basename(x) for x in catch_files] or "-")
        logger.info("[INPUT] stray files: %s", [os.path.basename(x) for x in stray_files] or "-")
        logger.info("[INPUT] card files: %s", [os.path.basename(x) for x in card_files] or "-")

        if not catch_files and not stray_files and not card_files:
            logger.error(
                "No input source found. Excel not found by pattern '%s'; JSON patterns: %s, %s, %s",
                EXCEL_INPUT_GLOB,
                CATCH_PART_GLOB,
                STRAY_PART_GLOB,
                CARD_PART_GLOB,
            )
            return 1

        catch_rows = _apply_limit(_with_row_numbers(load_rows_from_files(catch_files, logger, "CATCH"), 1), args.limit) if catch_files else []
        stray_rows = _apply_limit(_with_row_numbers(load_rows_from_files(stray_files, logger, "STRAY"), 1), args.limit) if stray_files else []
        card_rows = _apply_limit(_with_row_numbers(load_rows_from_files(card_files, logger, "CARD"), 1), args.limit) if card_files else []

        if catch_rows and not stopped:
            stopped = _process_job_with_resume(
                session=session,
                logger=logger,
                success_logger=success_logger,
                fail_logger=fail_logger,
                process_fn=process_order_rows,
                rows=catch_rows,
                created_items=created_items,
                errors=errors,
                rollback_candidates=rollback_candidates,
                state=state,
                workbook_key="__json__",
                workbook_label="json",
                job_name="catch_orders",
                primary_registry="catch-order",
                primary_collection=ORDER_COLLECTION,
                dry_run=False,
                interactive=interactive,
            )
        if stray_rows and not stopped:
            stopped = _process_job_with_resume(
                session=session,
                logger=logger,
                success_logger=success_logger,
                fail_logger=fail_logger,
                process_fn=process_stray_rows,
                rows=stray_rows,
                created_items=created_items,
                errors=errors,
                rollback_candidates=rollback_candidates,
                state=state,
                workbook_key="__json__",
                workbook_label="json",
                job_name="stray_animals",
                primary_registry="stray",
                primary_collection=TARGET_COLLECTION,
                dry_run=False,
                interactive=interactive,
            )
        if card_rows and not stopped:
            stopped = _process_job_with_resume(
                session=session,
                logger=logger,
                success_logger=success_logger,
                fail_logger=fail_logger,
                process_fn=process_card_rows,
                rows=card_rows,
                created_items=created_items,
                errors=errors,
                rollback_candidates=rollback_candidates,
                state=state,
                workbook_key="__json__",
                workbook_label="json",
                job_name="animal_cards",
                primary_registry="card",
                primary_collection=CARD_COLLECTION,
                dry_run=False,
                interactive=interactive,
            )

    if session is not None and VERIFY_CREATED:
        verify_created_entries(session, logger, rollback_candidates)

    logger.info("===== DONE =====")
    logger.info("Created: %s", json.dumps(created_items, ensure_ascii=False, indent=2))
    logger.info("Errors: %s", json.dumps(errors, ensure_ascii=False, indent=2))
    print_rollback("final-summary", rollback_candidates, logger)

    if stopped:
        logger.warning("Processing stopped due to configured stop flags")
    return 0 if not stopped else 1


if __name__ == "__main__":
    sys.exit(main())
