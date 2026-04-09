import json
import mimetypes
import os
import re
from pathlib import Path
from urllib.parse import quote, urlparse

import requests

from _config import (
    AUTH_TEST_COLLECTION,
    AUTO_JWT,
    BASE_URL,
    JWT_URL,
    REAUTH_ON_AUTH_ERROR,
    REAUTH_RETRIES,
    SAVE_AUTH,
    STORAGE_UPLOAD_FILE_PATH,
    STORAGE_UPLOAD_PATH,
    VERIFY_SSL,
)
from _utils import build_multipart_body, fix_mojibake_deep, jsonable, make_boundary, norm_ru, safe_json


class ApiCallError(RuntimeError):
    def __init__(self, message, code=None, data=None):
        super().__init__(message)
        self.code = code
        self.data = data


_RUNTIME_BASE_URL = str(BASE_URL).rstrip("/")
_RUNTIME_JWT_URL = str(JWT_URL or (str(BASE_URL).rstrip("/") + "/jwt/")).strip()


def set_runtime_urls(base_url=None, jwt_url=None):
    global _RUNTIME_BASE_URL, _RUNTIME_JWT_URL
    if base_url:
        _RUNTIME_BASE_URL = str(base_url).strip().rstrip("/")
    if jwt_url:
        _RUNTIME_JWT_URL = str(jwt_url).strip()
    elif _RUNTIME_BASE_URL and not _RUNTIME_JWT_URL:
        _RUNTIME_JWT_URL = _RUNTIME_BASE_URL + "/jwt/"


def get_runtime_base_url():
    return _RUNTIME_BASE_URL


def get_runtime_jwt_url():
    return _RUNTIME_JWT_URL


def _build_url(path):
    return _RUNTIME_BASE_URL + "/" + str(path or "").lstrip("/")


def _session_copy_into(target, source):
    target.cookies = source.cookies
    target.headers = source.headers
    target.verify = source.verify
    target.auth = source.auth


def _snapshot_file_stream_positions(files_obj):
    streams = []
    if not files_obj:
        return streams
    values = files_obj.values() if isinstance(files_obj, dict) else files_obj
    for item in values:
        candidates = item if isinstance(item, tuple) else (item,)
        for candidate in candidates:
            if hasattr(candidate, "read") and hasattr(candidate, "seek") and hasattr(candidate, "tell"):
                try:
                    streams.append((candidate, int(candidate.tell())))
                except Exception:
                    pass
                break
    return streams


def _rewind_file_stream_positions(streams):
    for handle, position in streams:
        try:
            handle.seek(position)
        except Exception:
            pass


def _read_text_if_exists(path):
    if not os.path.exists(path):
        return ""
    try:
        return open(path, "r", encoding="utf-8", errors="ignore").read().replace("\ufeff", "").strip()
    except Exception:
        return ""


def _write_text(path, text):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)


def _extract_jwt(text):
    raw = str(text or "").replace("\ufeff", "").strip()
    if not raw:
        return ""
    raw = re.sub(r"^\s*bearer\s+", "", raw, flags=re.IGNORECASE).strip()
    m = re.search(r"\b([A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+)\b", raw)
    return m.group(1) if m else raw


def _extract_jwt_from_html(text):
    html = text or ""
    patterns = [
        r'<textarea[^>]*name=["\']token["\'][^>]*>(.*?)</textarea>',
        r'<textarea[^>]*name=["\']jwt["\'][^>]*>(.*?)</textarea>',
    ]
    for p in patterns:
        m = re.search(p, html, flags=re.IGNORECASE | re.DOTALL)
        if not m:
            continue
        token = _extract_jwt(m.group(1))
        if token:
            return token
    return _extract_jwt(html)


def _parse_cookie_pairs(raw_cookie):
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
        if name and value and re.match(r"^[A-Za-z0-9_.\-]+$", name):
            out.append((name, value))
    return out


def _cookie_jar_to_string(cookie_jar):
    items = []
    for c in cookie_jar:
        name = getattr(c, "name", None)
        value = getattr(c, "value", None)
        if name:
            items.append("%s=%s" % (name, value or ""))
    return "; ".join(items)


def _normalize_raw_cookie_header(raw_cookie):
    text = str(raw_cookie or "").replace("\ufeff", "").strip()
    if not text:
        return ""
    text = re.sub(r"^\s*cookie\s*:\s*", "", text, flags=re.IGNORECASE)
    text = text.replace("\r", " ").replace("\n", " ").strip()
    return text


def _load_default_auth_from_files():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    token_md = os.path.join(script_dir, "token.md")
    cookie_md = os.path.join(script_dir, "cookie.md")

    token_raw = _read_text_if_exists(token_md)
    cookie_raw = _read_text_if_exists(cookie_md)

    return {
        "token": _extract_jwt(token_raw),
        "cookie": cookie_raw,
        "token_path": token_md,
        "cookie_path": cookie_md,
    }


def _save_auth_if_needed(session, logger):
    meta = getattr(session, "_auth_meta", None)
    if not isinstance(meta, dict) or not meta.get("save_auth"):
        return
    try:
        token = str(meta.get("token") or "").strip()
        if token:
            _write_text(meta.get("token_path"), token)
        cookie_raw = _cookie_jar_to_string(session.cookies)
        if cookie_raw:
            _write_text(meta.get("cookie_path"), cookie_raw)
        logger.info("[AUTH] token/cookie files updated")
    except Exception as exc:
        logger.warning("[AUTH] failed to persist token/cookie: %s", exc)


def _clean_token_for_headers(token):
    raw = str(token or "").replace("\ufeff", "").strip()
    if not raw:
        return ""
    # PSI expects raw JWT value in `token` and Bearer-form in `Authorization`.
    # Remove optional Bearer prefix first, then normalize through parser.
    bearer_clean = re.sub(r"^\s*bearer\s+", "", raw, flags=re.IGNORECASE).strip()
    return _extract_jwt(bearer_clean)


def _apply_token_headers(session, token):
    clean = _clean_token_for_headers(token)
    if not clean:
        return
    session.headers["token"] = clean
    session.headers["Authorization"] = "Bearer " + clean


def _ensure_auth_headers_from_meta(session):
    meta = getattr(session, "_auth_meta", None)
    if not isinstance(meta, dict):
        return
    token = str(meta.get("token") or "").strip()
    if token:
        _apply_token_headers(session, token)


def _apply_cookies(session, cookie_raw, logger):
    raw_cookie_header = _normalize_raw_cookie_header(cookie_raw)
    pairs = _parse_cookie_pairs(cookie_raw)
    host_match = re.match(r"^https?://([^/:]+)", _RUNTIME_BASE_URL, flags=re.IGNORECASE)
    domain = host_match.group(1) if host_match else ""
    parsed = 0
    for name, value in pairs:
        if domain:
            session.cookies.set(name, value, domain=domain, path="/")
        else:
            session.cookies.set(name, value, path="/")
        parsed += 1
    if raw_cookie_header:
        # Keep original cookie string for PSI compatibility (very long PLATFORM_SESSION chains).
        session.headers["Cookie"] = raw_cookie_header
    elif parsed:
        session.headers["Cookie"] = "; ".join(["%s=%s" % (k, v) for k, v in pairs])
    logger.info("Parsed cookie pairs: %s", parsed)
    return parsed


def _auth_test(session, logger):
    auth_url = _build_url("/api/v1/search/%s" % AUTH_TEST_COLLECTION)
    body = {
        "search": {"search": [{"andSubConditions": [{"field": "_id", "operator": "notNull"}]}]},
        "limit": 1,
        "offset": 0,
    }
    try:
        resp = session.post(auth_url, json=body, timeout=60)
    except Exception as exc:
        logger.warning("[AUTH TEST] request error: %s", exc)
        return False
    logger.info("[AUTH TEST] %s -> %s", auth_url, resp.status_code)
    return resp.status_code < 400


def _refresh_token_from_jwt_page(session, logger):
    candidates = []
    jwt_url = str(_RUNTIME_JWT_URL or "").strip()
    if jwt_url:
        candidates.extend([jwt_url.rstrip("/") + "/", jwt_url.rstrip("/") + "/?access=1"])
    candidates.extend([_build_url("/jwt/"), _build_url("/jwt/?access=1")])

    tried = set()
    for url in candidates:
        if not url or url in tried:
            continue
        tried.add(url)
        try:
            resp = session.get(url, timeout=90, allow_redirects=True)
        except Exception as exc:
            logger.warning("[AUTH][JWT] GET %s не выполнен: %s", url, exc)
            continue
        logger.info("[AUTH][JWT] GET %s -> %s", url, resp.status_code)
        if resp.status_code >= 400:
            continue
        token = _extract_jwt_from_html(resp.text or "")
        if token:
            logger.info("[AUTH][JWT] token parsed from %s", url)
            return token
    return ""


def _reauth_session(session, logger):
    meta = getattr(session, "_auth_meta", None)
    if not isinstance(meta, dict):
        return False
    if not meta.get("auto_jwt", True):
        return False

    token = _refresh_token_from_jwt_page(session, logger)
    if not token:
        logger.warning("[AUTH] auto-jwt не выполнен: токен не найден")
        return False

    _apply_token_headers(session, token)
    meta["token"] = _clean_token_for_headers(token)
    _save_auth_if_needed(session, logger)
    ok = _auth_test(session, logger)
    logger.info("[AUTH] re-auth result: %s", ok)
    return ok


def setup_session(
    logger,
    *,
    no_prompt=False,
    auto_jwt_override=None,
    save_auth_override=None,
    operator_mode=False,
):
    defaults = _load_default_auth_from_files()

    print("\nCookie и token можно взять из файлов рядом со скриптом:")
    print("  cookie: %s" % defaults["cookie_path"])
    print("  token:  %s" % defaults["token_path"])

    if no_prompt:
        print("Режим без запросов: используются значения только из файлов.")
        cookie_input = ""
        jwt_input = ""
    else:
        print("Нажмите Enter, чтобы использовать значения из файлов.")
        try:
            cookie_input = input("Cookie (или Enter): ").replace("\ufeff", "").strip()
        except EOFError:
            cookie_input = ""
        try:
            jwt_input = input("JWT token (или Enter): ").replace("\ufeff", "").strip()
        except EOFError:
            jwt_input = ""

    cookie_header = cookie_input or defaults["cookie"]
    jwt_token = _clean_token_for_headers(jwt_input or defaults["token"])

    if not cookie_header and not jwt_token:
        logger.error("Не переданы ни Cookie, ни JWT token")
        return None

    auto_jwt = AUTO_JWT if auto_jwt_override is None else bool(auto_jwt_override)
    save_auth = SAVE_AUTH if save_auth_override is None else bool(save_auth_override)

    session = requests.Session()
    session.trust_env = False
    session.verify = bool(VERIFY_SSL)
    session.headers.update(
        {
            "Accept": "application/json, text/plain, */*",
            "Origin": _RUNTIME_BASE_URL,
            "Referer": _RUNTIME_BASE_URL + "/",
            "User-Agent": "Mozilla/5.0",
        }
    )

    _apply_cookies(session, cookie_header, logger)
    if jwt_token:
        _apply_token_headers(session, jwt_token)

    xsrf = session.cookies.get("XSRF-TOKEN") or session.cookies.get("XSRF_TOKEN")
    if xsrf:
        session.headers["X-XSRF-TOKEN"] = xsrf

    meta = {
        "token_path": defaults["token_path"],
        "cookie_path": defaults["cookie_path"],
        "token": jwt_token,
        "auto_jwt": auto_jwt,
        "save_auth": save_auth,
        "operator_mode": bool(operator_mode),
    }
    setattr(session, "_auth_meta", meta)

    if _auth_test(session, logger):
        _save_auth_if_needed(session, logger)
        return session

    if auto_jwt and _reauth_session(session, logger):
        return session

    logger.error("Авторизация не прошла. Обновите cookie/token и повторите запуск.")
    return None


def api_request(session, logger, method, url, reauth_fn=None, max_retries=None, **kwargs):
    auth_codes = {401, 403}
    retries = max(1, int(max_retries if max_retries is not None else REAUTH_RETRIES))
    file_streams = _snapshot_file_stream_positions(kwargs.get("files"))

    response = None
    for _ in range(retries + 1):
        _ensure_auth_headers_from_meta(session)
        if file_streams:
            _rewind_file_stream_positions(file_streams)
        response = session.request(method=method.upper(), url=url, timeout=120, **kwargs)
        if not REAUTH_ON_AUTH_ERROR or response.status_code not in auth_codes:
            break
        logger.warning("HTTP %s от %s. Пробуем авто-переавторизацию", response.status_code, url)
        if not _reauth_session(session, logger):
            break

    if response is not None and response.status_code in auth_codes:
        meta = getattr(session, "_auth_meta", {}) or {}
        if bool(meta.get("operator_mode")):
            while True:
                try:
                    raw = input(
                        "[AUTH] Сессия истекла (HTTP %s). Действия: [В]ойти заново / [О]становить: "
                        % response.status_code
                    )
                except EOFError:
                    raw = "о"
                normalized = norm_ru(raw)
                if not normalized:
                    normalized = "в"
                if normalized in {"o", "о", "остановить", "abort", "a"}:
                    break
                if normalized not in {"в", "v", "login", "l", "войти", "логин"}:
                    continue
                refreshed = setup_session(
                    logger,
                    no_prompt=False,
                    auto_jwt_override=meta.get("auto_jwt", True),
                    save_auth_override=meta.get("save_auth", True),
                    operator_mode=True,
                )
                if not refreshed:
                    continue
                _session_copy_into(session, refreshed)
                if file_streams:
                    _rewind_file_stream_positions(file_streams)
                response = session.request(method=method.upper(), url=url, timeout=120, **kwargs)
                if response.status_code not in auth_codes:
                    break

    if response is None:
        raise RuntimeError("API request retries exhausted")
    return response


def _parse_response_data(response):
    content_type = str(response.headers.get("Content-Type") or "")
    text = response.text or ""
    if "application/json" in content_type:
        return safe_json(text)
    return text


def call_api(session, logger, method="GET", path="", body=None, extra_headers=None):
    url = _build_url(path)
    headers = {"Content-Type": "application/json"}
    if extra_headers:
        headers.update(extra_headers)

    kwargs = {"headers": headers}
    if body is not None:
        sanitized_body = fix_mojibake_deep(body)
        kwargs["data"] = json.dumps(jsonable(sanitized_body), ensure_ascii=False).encode("utf-8")

    response = api_request(session, logger, method, url, **kwargs)
    data = _parse_response_data(response)
    if 200 <= response.status_code < 300:
        return {"code": response.status_code, "data": data, "response": response}
    raise ApiCallError("HTTP %s %s" % (response.status_code, url), code=response.status_code, data=data)


def search_collection(session, logger, collection, body):
    path = "/api/v1/search/%s" % collection
    return call_api(session, logger, method="POST", path=path, body=body)["data"]


def create_record(session, logger, collection, body):
    path = "/api/v1/create/%s" % collection
    return call_api(session, logger, method="POST", path=path, body=body)["data"]


def update_record(session, logger, collection, main_id, guid, doc):
    q = "?mainId=%s&guid=%s" % (quote(str(main_id), safe=""), quote(str(guid), safe=""))
    path = "/api/v1/update/%s%s" % (collection, q)
    return call_api(session, logger, method="PUT", path=path, body=doc)["data"]


def upload_file_base64(session, logger, entry_name, entry_id, entity_field_path, filename, base64_content):
    boundary = make_boundary()
    fields = {
        "entryName": entry_name,
        "entryId": entry_id,
        "entityFieldPath": entity_field_path,
    }
    raw_body = build_multipart_body(
        boundary=boundary,
        filename=filename,
        fields=fields,
        base64_content=base64_content,
    )

    url = _build_url(STORAGE_UPLOAD_PATH)
    headers = {"Content-Type": "multipart/form-data; boundary=%s" % boundary}
    response = api_request(
        session=session,
        logger=logger,
        method="POST",
        url=url,
        headers=headers,
        data=raw_body.encode("utf-8"),
    )
    data = _parse_response_data(response)
    if 200 <= response.status_code < 300:
        return {"code": response.status_code, "data": data}
    raise ApiCallError("Ошибка загрузки: HTTP %s" % response.status_code, code=response.status_code, data=data)


def upload_file(
    session,
    logger,
    file_path,
    entry_name,
    entry_id,
    entity_field_path="",
    allow_external=False,
):
    if not os.path.isfile(file_path):
        raise ApiCallError("Local file not found: %s" % file_path)

    url = _build_url(STORAGE_UPLOAD_FILE_PATH)
    filename = os.path.basename(file_path)
    mime_type, _ = mimetypes.guess_type(file_path)
    content_type = mime_type or "application/octet-stream"
    data = {
        "entryName": entry_name,
        "entryId": str(entry_id),
        "entityFieldPath": entity_field_path or "",
        "allowExternal": "true" if allow_external else "false",
    }

    with open(file_path, "rb") as f:
        files = {"file": (filename, f, content_type)}
        response = api_request(
            session=session,
            logger=logger,
            method="POST",
            url=url,
            data=data,
            files=files,
        )

    resp_data = _parse_response_data(response)
    if 200 <= response.status_code < 300:
        return {"code": response.status_code, "data": resp_data}
    raise ApiCallError(
        "Ошибка загрузки файла: HTTP %s" % response.status_code,
        code=response.status_code,
        data=resp_data,
    )


def delete_from_collection(session, logger, data):
    main_id = data.get("_id")
    guid = data.get("guid")
    parent_entries = data.get("parentEntries")
    if not main_id or not guid or not parent_entries:
        logger.error("Cannot delete: _id, guid or parentEntries is missing")
        return False

    q = "?mainId=%s&guid=%s" % (quote(str(main_id), safe=""), quote(str(guid), safe=""))
    path = "/api/v1/delete/%s%s" % (parent_entries, q)
    url = _build_url(path)

    try:
        response = api_request(session, logger, "DELETE", url, max_retries=1)
    except Exception as exc:
        logger.error("Delete exception %s/%s: %s", parent_entries, main_id, exc)
        return False

    if response.status_code in (200, 202, 204, 404, 500):
        logger.info("Delete %s/%s finished with status %s", parent_entries, main_id, response.status_code)
        return True

    logger.error(
        "Delete error %s/%s: %s %s",
        parent_entries,
        main_id,
        response.status_code,
        (response.text or "")[:500],
    )
    return False
