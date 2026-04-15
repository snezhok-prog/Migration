from __future__ import annotations

import mimetypes
import re
from pathlib import Path
from typing import Any, Dict, Optional

import requests

from _config import (
    AUTH_TEST_COLLECTION,
    AUTO_JWT,
    BASE_URL,
    COOKIE_FILE,
    JWT_URL,
    REAUTH_ON_401,
    REAUTH_RETRIES,
    SAVE_AUTH,
    SCRIPT_DIR,
    TOKEN_FILE,
    UI_BASE_URL,
    VERIFY_SSL,
)
from _utils import (
    cookie_jar_to_string,
    extract_jwt,
    extract_jwt_from_html,
    parse_cookie_pairs,
    read_text_if_exists,
    write_text,
)


class ApiCallError(RuntimeError):
    pass


_RUNTIME_BASE_URL = str(BASE_URL).rstrip("/")
_RUNTIME_JWT_URL = str(JWT_URL or (str(BASE_URL).rstrip("/") + "/jwt/")).strip()
_RUNTIME_UI_BASE_URL = str(UI_BASE_URL or BASE_URL).rstrip("/")


def set_runtime_urls(base_url: Optional[str] = None, jwt_url: Optional[str] = None, ui_base_url: Optional[str] = None) -> None:
    global _RUNTIME_BASE_URL, _RUNTIME_JWT_URL, _RUNTIME_UI_BASE_URL
    if base_url:
        _RUNTIME_BASE_URL = str(base_url).strip().rstrip("/")
    if jwt_url:
        _RUNTIME_JWT_URL = str(jwt_url).strip()
    elif _RUNTIME_BASE_URL and not _RUNTIME_JWT_URL:
        _RUNTIME_JWT_URL = _RUNTIME_BASE_URL + "/jwt/"

    if ui_base_url:
        _RUNTIME_UI_BASE_URL = str(ui_base_url).strip().rstrip("/")
    elif base_url:
        _RUNTIME_UI_BASE_URL = _RUNTIME_BASE_URL


def get_runtime_base_url() -> str:
    return _RUNTIME_BASE_URL


def get_runtime_jwt_url() -> str:
    return _RUNTIME_JWT_URL


def get_runtime_ui_base_url() -> str:
    return _RUNTIME_UI_BASE_URL


def _build_url(path: str) -> str:
    return _RUNTIME_BASE_URL + "/" + str(path or "").lstrip("/")


def _safe_json(resp: requests.Response) -> Any:
    try:
        return resp.json()
    except Exception:
        return resp.text


def _apply_token_headers(session: requests.Session, token: str) -> None:
    clean = extract_jwt(token)
    if not clean:
        return
    session.headers["token"] = clean
    session.headers["Authorization"] = "Bearer " + clean


def _normalize_raw_cookie_header(raw_cookie: str) -> str:
    text = str(raw_cookie or "").replace("\ufeff", "").strip()
    if not text:
        return ""
    text = re.sub(r"^\s*cookie\s*:\s*", "", text, flags=re.IGNORECASE)
    return text.replace("\r", " ").replace("\n", " ").strip()


def _apply_cookie_pairs(session: requests.Session, cookie_raw: str, logger) -> int:
    raw_cookie_header = _normalize_raw_cookie_header(cookie_raw)
    pairs = parse_cookie_pairs(cookie_raw)
    host_match = re.match(r"^https?://([^/:]+)", _RUNTIME_BASE_URL.strip(), flags=re.IGNORECASE)
    domain = host_match.group(1) if host_match else ""
    parsed = 0
    for name, value in pairs:
        if domain:
            session.cookies.set(name, value, domain=domain, path="/")
        else:
            session.cookies.set(name, value, path="/")
        parsed += 1
    if raw_cookie_header:
        # Keep original cookie string for PSI compatibility (long PLATFORM_SESSION chains).
        session.headers["Cookie"] = raw_cookie_header
    elif parsed:
        session.headers["Cookie"] = "; ".join([f"{k}={v}" for k, v in pairs])
    logger.info("Parsed cookie pairs: %s", parsed)
    return parsed


def _auth_test(session: requests.Session, logger) -> bool:
    checks = [
        (
            _build_url(f"/api/v1/search/{AUTH_TEST_COLLECTION}"),
            {
                "search": {"search": [{"andSubConditions": [{"field": "_id", "operator": "notNull"}]}]},
                "limit": 1,
                "offset": 0,
            },
        ),
        (_build_url("/api/v1/search/subservices"), {}),
    ]
    for auth_url, body in checks:
        try:
            resp = session.post(auth_url, json=body, timeout=60)
        except Exception as exc:
            logger.warning("[AUTH TEST] request error: %s (%s)", exc, auth_url)
            continue
        content_type = str(resp.headers.get("content-type") or "").lower()
        logger.info("[AUTH TEST] POST %s -> %s | ct=%s", auth_url, resp.status_code, content_type)
        if resp.status_code == 200 and "application/json" in content_type:
            return True
        preview = (resp.text or "").replace("\r", " ").replace("\n", " ").strip()[:500]
        if preview:
            logger.warning("[AUTH TEST] non-success response preview: %s", preview)
    return False


def _drop_token_headers(session: requests.Session) -> None:
    session.headers.pop("token", None)
    session.headers.pop("Authorization", None)


def _save_auth_if_needed(session: requests.Session, token: str, logger) -> None:
    if not SAVE_AUTH:
        return
    try:
        token_path = Path(SCRIPT_DIR) / TOKEN_FILE
        cookie_path = Path(SCRIPT_DIR) / COOKIE_FILE
        if token:
            write_text(token_path, token)
        cookie_raw = cookie_jar_to_string(session.cookies)
        if cookie_raw:
            write_text(cookie_path, cookie_raw)
        logger.info("[AUTH] token/cookie files updated")
    except Exception as exc:
        logger.warning("[AUTH] failed to persist token/cookie: %s", exc)


def _refresh_token_from_jwt_page(session: requests.Session, logger) -> Optional[str]:
    candidates = []
    jwt_url = str(_RUNTIME_JWT_URL or "").strip()
    if jwt_url:
        candidates.extend([jwt_url.rstrip("/") + "/", jwt_url.rstrip("/") + "/?access=1"])
    candidates.extend([_build_url("/jwt/"), _build_url("/jwt/?access=1")])
    tried = set()

    for url in candidates:
        if url in tried:
            continue
        tried.add(url)
        try:
            resp = session.get(url, timeout=90, allow_redirects=True)
        except Exception as exc:
            logger.warning("[AUTH][JWT] GET %s failed: %s", url, exc)
            continue
        logger.info("[AUTH][JWT] GET %s -> %s", url, resp.status_code)
        if resp.status_code >= 400:
            continue
        token = extract_jwt_from_html(resp.text or "")
        if token:
            logger.info("[AUTH][JWT] token parsed from %s", url)
            return token
    return None


def _reauth_session(session: requests.Session, logger) -> bool:
    if not AUTO_JWT:
        return False
    token = _refresh_token_from_jwt_page(session, logger)
    if not token:
        logger.warning("[AUTH] auto-jwt failed: token not found")
        return False
    _apply_token_headers(session, token)
    _save_auth_if_needed(session, token, logger)
    return _auth_test(session, logger)


def setup_session(logger, no_prompt: bool = False) -> requests.Session:
    token_path = Path(SCRIPT_DIR) / TOKEN_FILE
    cookie_path = Path(SCRIPT_DIR) / COOKIE_FILE
    default_token = extract_jwt(read_text_if_exists(token_path))
    default_cookie = read_text_if_exists(cookie_path)

    print("\nCookie и token можно взять из файлов рядом со скриптом:")
    print(f"  cookie: {cookie_path}")
    print(f"  token:  {token_path}")

    if no_prompt:
        cookie_input = ""
        token_input = ""
    else:
        print("Нажмите Enter, чтобы использовать значения из файлов.")
        cookie_input = input("Cookie: ").replace("\ufeff", "").strip()
        token_input = input("JWT token: ").replace("\ufeff", "").strip()

    cookie_raw = cookie_input or default_cookie
    token = extract_jwt(token_input or default_token)
    if not cookie_raw and not token:
        raise RuntimeError("No cookie/token provided")

    session = requests.Session()
    session.trust_env = False
    session.verify = VERIFY_SSL
    session.headers.update(
        {
            "Accept": "application/json, text/plain, */*",
            "Origin": _RUNTIME_BASE_URL,
            "Referer": _RUNTIME_BASE_URL + "/",
            "User-Agent": "Mozilla/5.0",
        }
    )
    _apply_cookie_pairs(session, cookie_raw, logger)
    if token:
        _apply_token_headers(session, token)

    xsrf = session.cookies.get("XSRF-TOKEN") or session.cookies.get("XSRF_TOKEN")
    if xsrf:
        session.headers["X-XSRF-TOKEN"] = xsrf

    if _auth_test(session, logger):
        _save_auth_if_needed(session, token, logger)
        return session

    # Common PSI case: valid cookie + stale/foreign JWT from token.md.
    if token and cookie_raw:
        logger.warning("[AUTH] initial auth failed; retrying with cookie-only (without JWT headers)")
        _drop_token_headers(session)
        if _auth_test(session, logger):
            _save_auth_if_needed(session, "", logger)
            return session

    if _reauth_session(session, logger):
        return session

    raise RuntimeError("Auth failed. Update cookie/token and try again.")


def _request(session: requests.Session, logger, *, method: str, url: str, timeout: int = 120, **kwargs) -> requests.Response:
    resp = session.request(method=method, url=url, timeout=timeout, **kwargs)
    if resp.status_code in (401, 403, 413) and REAUTH_ON_401:
        for attempt in range(max(1, int(REAUTH_RETRIES))):
            logger.warning("[%s] %s -> %s (attempt reauth %s/%s)", method.upper(), url, resp.status_code, attempt + 1, REAUTH_RETRIES)
            if not _reauth_session(session, logger):
                break
            resp = session.request(method=method, url=url, timeout=timeout, **kwargs)
            if resp.status_code not in (401, 403, 413):
                break
    logger.info("[%s] %s -> %s", method.upper(), url, resp.status_code)
    return resp


def _request_json(session: requests.Session, logger, *, method: str, url: str, timeout: int = 120, **kwargs) -> Any:
    resp = _request(session, logger, method=method, url=url, timeout=timeout, **kwargs)
    payload = _safe_json(resp)
    if resp.status_code >= 400:
        raise ApiCallError(f"HTTP {resp.status_code}: {payload}")
    return payload


def create_record(session: requests.Session, logger, collection: str, body: Dict[str, Any]) -> Dict[str, Any]:
    url = _build_url(f"/api/v1/create/{collection}")
    data = _request_json(session, logger, method="POST", url=url, json=body)
    if isinstance(data, dict):
        return data
    raise ApiCallError(f"Unexpected create response type for {collection}: {type(data).__name__}")


def update_record(session: requests.Session, logger, *, collection: str, main_id: str, guid: str, body: Dict[str, Any]) -> Dict[str, Any]:
    url = _build_url(f"/api/v1/update/{collection}")
    data = _request_json(session, logger, method="PUT", url=url, params={"mainId": str(main_id), "guid": str(guid)}, json=body)
    if isinstance(data, dict):
        return data
    raise ApiCallError(f"Unexpected update response type for {collection}: {type(data).__name__}")


def delete_record(session: requests.Session, logger, *, collection: str, main_id: str, guid: str) -> int:
    url = _build_url(f"/api/v1/delete/{collection}")
    resp = _request(session, logger, method="DELETE", url=url, params={"mainId": str(main_id), "guid": str(guid or "")}, timeout=90)
    return resp.status_code


def delete_from_collection(session: requests.Session, logger, item: Dict[str, Any]) -> bool:
    status = delete_record(
        session,
        logger,
        collection=str(item.get("parentEntries") or ""),
        main_id=str(item.get("_id") or ""),
        guid=str(item.get("guid") or ""),
    )
    ok = status in (200, 202, 204)
    if ok:
        logger.info("Удалена запись: %s %s", item.get("parentEntries"), item.get("_id"))
    else:
        logger.error("Не удалось удалить запись: %s %s (HTTP %s)", item.get("parentEntries"), item.get("_id"), status)
    return ok


def upload_file(
    session: requests.Session,
    logger,
    *,
    entry_name: str,
    entry_id: str,
    entity_field_path: str,
    file_path: str,
    allow_external: bool = False,
) -> Dict[str, Any]:
    url = _build_url("/api/v1/storage/upload")
    data = {
        "entryName": entry_name,
        "entryId": entry_id,
        "entityFieldPath": entity_field_path or "",
        "allowExternal": "true" if allow_external else "false",
    }
    mime_type, _ = mimetypes.guess_type(file_path)
    content_type = mime_type or "application/octet-stream"
    with open(file_path, "rb") as fp:
        files = {"file": (Path(file_path).name, fp, content_type)}
        resp = _request(session, logger, method="POST", url=url, data=data, files=files, timeout=180)
    payload = _safe_json(resp)
    if resp.status_code >= 400:
        raise ApiCallError(f"Upload failed HTTP {resp.status_code}: {payload}")
    if isinstance(payload, dict):
        return payload
    raise ApiCallError(f"Unexpected upload response type: {type(payload).__name__}")


def build_file_meta(uploaded: Dict[str, Any], *, fallback_name: str, fallback_size: int, path: str, allow_external: bool) -> Dict[str, Any]:
    file_id = uploaded.get("_id") or uploaded.get("fileId") or uploaded.get("id") or ""
    original_name = uploaded.get("originalName") or uploaded.get("name") or fallback_name
    return {
        "_id": file_id,
        "size": uploaded.get("size") if uploaded.get("size") is not None else fallback_size,
        "isFile": True,
        "originalName": original_name,
        "allowExternal": bool(allow_external),
        "entityFieldPath": path,
    }


def search_org_by_ogrn(session: Optional[requests.Session], logger, ogrn: Any) -> Optional[Dict[str, Any]]:
    val = str(ogrn or "").strip()
    if not val or session is None:
        return None
    url = _build_url("/api/v1/search/organizations")
    body = {"search": {"search": [{"field": "ogrn", "operator": "eq", "value": val}]}, "size": 2}
    try:
        resp = session.post(url, json=body, timeout=60)
    except Exception as exc:
        logger.warning("[ORG SEARCH] request error for ogrn=%s: %s", val, exc)
        return None
    logger.info("[ORG SEARCH] %s -> %s", url, resp.status_code)
    if resp.status_code >= 400:
        return None
    payload = _safe_json(resp)
    content = payload.get("content") if isinstance(payload, dict) else []
    if isinstance(content, list):
        if len(content) == 1:
            return content[0]
        if len(content) > 1:
            first = content[0] if isinstance(content[0], dict) else None
            logger.warning(
                "[ORG SEARCH] multiple matches for ogrn=%s (count=%s), using first id=%s",
                val,
                len(content),
                (first or {}).get("_id"),
            )
            return first
    return None
