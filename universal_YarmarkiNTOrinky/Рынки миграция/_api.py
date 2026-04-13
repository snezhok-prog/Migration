import copy
import json
import requests
import os
import mimetypes
import re

from pathlib import Path
from _config import BASE_URL, APPEAL_SETTINGS, STANDARD_CODES, SCRIPT_DIR, TOKEN_FILE, COOKIE_FILE
from _utils import (
    jsonable,
    generate_guid,
    to_iso_date,
    parse_date_to_birthday_obj,
    format_phone,
    format_multiple_phones,
    read_file_as_base64,
    make_boundary,
    build_multipart_body,
    find_file_in_dir,
    find_document_group_by_mnemonic
)


def get_standard_code(key):
    return STANDARD_CODES.get(key)


def upload_file(
    session,
    logger,
    file_path: str,
    entry_name: str,
    entry_id: str,
    entity_field_path: str = ""
):
    """
    РћР±СЏР·Р°С‚РµР»РµРЅ JWT С‚РѕРєРµРЅ, Р±РµР· РЅРµРіРѕ РЅРёС‡РµРіРѕ РЅРµ Р·Р°РіСЂСѓР·РёС‚СЃСЏ
    """
    file_name = os.path.basename(file_path)
    url = f"{BASE_URL}/api/v1/storage/upload"
    
    # рџ”Ґ Р’РђР–РќРћ: РІСЃРµ РїРѕР»СЏ С„РѕСЂРјС‹ вЂ” РІ data, Р° РЅРµ РІ params!
    data = {
        'entryName': entry_name,      # в†ђ Р±С‹Р»Рѕ РІ params, С‚РµРїРµСЂСЊ Р·РґРµСЃСЊ
        'entryId': entry_id,          # в†ђ Р±С‹Р»Рѕ РІ params, С‚РµРїРµСЂСЊ Р·РґРµСЃСЊ
        'entityFieldPath': entity_field_path,
        'allowExternal': 'false'
    }
    
    if not os.path.isfile(file_path):
        logger.error(f"Р¤Р°Р№Р» РЅРµ РЅР°Р№РґРµРЅ: {file_path}")
        return None
    
    if not file_name:#123
        file_name = Path(file_path).name
    mime_type, _ = mimetypes.guess_type(file_path)
    content_type = mime_type or "application/octet-stream"
    
    logger.info(f"РџРѕРґРіРѕС‚РѕРІРєР° Рє Р·Р°РіСЂСѓР·РєРµ: {file_name} ({content_type})")
    
    # РћР±РЅРѕРІР»СЏРµРј Referer РїРѕРґ РєРѕРЅРєСЂРµС‚РЅРѕРµ РґРµР»Рѕ
    session.headers["Referer"] = f"{BASE_URL}/AppRKN034/common-appeals/edit/{entry_id}"
    session.headers["Origin"] = BASE_URL
    
    try:
        with open(file_path, 'rb') as f:
            files = {'file': (file_name, f, content_type)}
            
            logger.debug(f"рџ”Ќ РћС‚РїСЂР°РІРєР°: url={url}, data={data}")
            logger.debug(f"рџ”Ќ Cookies: {list(session.cookies.keys())}")
            
            # рџ”Ґ РЈР±РёСЂР°РµРј params= вЂ” РІСЃРµ РґР°РЅРЅС‹Рµ РІ С‚РµР»Рµ С„РѕСЂРјС‹
            response = api_request(session, logger, "post", url, files=files, data=data)
            # response = session.post(
            #     url,
            #     files=files,
            #     data=data,  # в†ђ РІСЃРµ РїРѕР»СЏ С„РѕСЂРјС‹ Р·РґРµСЃСЊ
            #     timeout=120
            # )
    except Exception as e:
        logger.error(f"РћС€РёР±РєР° РїСЂРё Р·Р°РіСЂСѓР·РєРµ: {type(e).__name__}: {e}")
        return None

    logger.info(f"Р—Р°РїСЂРѕСЃ Рє {url}, СЃС‚Р°С‚СѓСЃ: {response.status_code}")
    
    if response.status_code not in (200, 201, 202):
        logger.error(f"РћС€РёР±РєР° HTTP {response.status_code}: {response.text}")
        return None

    try:
        result = response.json()
    except requests.exceptions.JSONDecodeError:
        if response.status_code in (200, 201, 202) and not response.text.strip():
            return {"status": "uploaded", "fileName": file_name}
        logger.error("РћС‚РІРµС‚ РЅРµ СЏРІР»СЏРµС‚СЃСЏ JSON")
        return None

    if isinstance(result, dict) and ("error" in result or result.get("success") is False):
        logger.error(f"API РІРµСЂРЅСѓР»Рѕ РѕС€РёР±РєСѓ: {result}")
        return None

    logger.info(f"вњ… Р¤Р°Р№Р» {file_name} Р·Р°РіСЂСѓР¶РµРЅ")
    return result


def delete_file_from_storage(session, logger, file_id: str):
    """
    РЈРґР°Р»СЏРµС‚ С„Р°Р№Р» РёР· С…СЂР°РЅРёР»РёС‰Р° РїРѕ fileId.
    
    Args:
        session (requests.Session): РђРІС‚РѕСЂРёР·РѕРІР°РЅРЅР°СЏ СЃРµСЃСЃРёСЏ (СЃ РєСѓРєР°РјРё Рё С‚РѕРєРµРЅРѕРј)
        logger (logging.Logger): Р›РѕРіРіРµСЂ
        file_id (str): РРґРµРЅС‚РёС„РёРєР°С‚РѕСЂ С„Р°Р№Р»Р° (РЅР°РїСЂРёРјРµСЂ, '6946daca2899a5480fe402dd')
    
    Returns:
        bool: True вЂ” РµСЃР»Рё СѓРґР°Р»РµРЅРёРµ РїСЂРѕС€Р»Рѕ СѓСЃРїРµС€РЅРѕ, False вЂ” РїСЂРё РѕС€РёР±РєРµ
    """
    # РЈР±РёСЂР°РµРј РІРѕР·РјРѕР¶РЅС‹Рµ РїСЂРѕР±РµР»С‹ РІ file_id
    file_id = file_id.strip()
    
    url = f"{BASE_URL}/api/v1/storage/remove"
    params = {"fileId": file_id}
    
    # РЇРІРЅРѕ СѓРєР°Р·С‹РІР°РµРј Р·Р°РіРѕР»РѕРІРєРё, РєР°Рє РІ fetch (С…РѕС‚СЏ session Рё С‚Р°Рє РѕС‚РїСЂР°РІРёС‚ РєСѓРєРё)
    headers = {
        "accept": "application/hal+json",
        "content-type": "application/json"
    }
    
    logger.info(f"РЈРґР°Р»РµРЅРёРµ С„Р°Р№Р»Р° РёР· С…СЂР°РЅРёР»РёС‰Р°: fileId={file_id}")
    
    try:
        response = session.delete(url, params=params, headers=headers)
        logger.info(f"РЎС‚Р°С‚СѓСЃ СѓРґР°Р»РµРЅРёСЏ С„Р°Р№Р»Р°: {response.status_code}")
        
        # РЈСЃРїРµС€РЅС‹Рµ СЃС‚Р°С‚СѓСЃС‹: 200, 204, РёРЅРѕРіРґР° 202
        if response.status_code in (200, 204, 202):
            logger.info("вњ… Р¤Р°Р№Р» СѓСЃРїРµС€РЅРѕ СѓРґР°Р»С‘РЅ РёР· С…СЂР°РЅРёР»РёС‰Р°")
            return True
        else:
            logger.error(f"вќЊ РћС€РёР±РєР° СѓРґР°Р»РµРЅРёСЏ С„Р°Р№Р»Р°: {response.status_code}")
            logger.error(f"РўРµР»Рѕ РѕС‚РІРµС‚Р°: {response.text[:500]}")
            return False
            
    except Exception as e:
        logger.error(f"рџ”Ґ РСЃРєР»СЋС‡РµРЅРёРµ РїСЂРё СѓРґР°Р»РµРЅРёРё С„Р°Р№Р»Р°: {e}")
        return False


def api_request(session, logger, method, url, reauth_fn=None, max_retries=3, **kwargs):
    """РЈРЅРёС„РёС†РёСЂРѕРІР°РЅРЅС‹Р№ Р·Р°РїСЂРѕСЃ Рє API СЃ РїРѕРІС‚РѕСЂРЅРѕР№ Р°РІС‚РѕСЂРёР·Р°С†РёРµР№ РЅР° 401/403/500."""
    if reauth_fn is None:
        reauth_fn = lambda log: setup_session(log, no_prompt=True)

    for attempt in range(max_retries + 1):
        try:
            request_fn = getattr(session, method.lower(), None)
            if request_fn is None:
                raise ValueError(f"РќРµРёР·РІРµСЃС‚РЅС‹Р№ HTTP РјРµС‚РѕРґ: {method}")

            response = request_fn(url, **kwargs)
            if response.status_code in (401, 403) or response.status_code == 500 and method.lower() != "delete":
                logger.warning(f"HTTP {response.status_code} РѕС‚ {url}. РџРѕРїС‹С‚РєР° СЂРµР°РІС‚РѕСЂРёР·Р°С†РёРё {attempt + 1}/{max_retries}")
                if attempt < max_retries:
                    new_session = reauth_fn(logger)
                    if new_session is not None:
                        session.cookies = new_session.cookies
                        session.headers = new_session.headers
                        session.verify = new_session.verify
                        session.auth = new_session.auth
                        continue
                return response
            
            return response
        except Exception as e:
            logger.error(f"РћС€РёР±РєР° РІ api_request ({method.upper()} {url}): {e}")
            if attempt < max_retries:
                continue
            raise

    raise RuntimeError("api_request: РёСЃС‡РµСЂРїР°РЅС‹ РїРѕРїС‹С‚РєРё")


def _read_text_if_exists(path: Path) -> str:
    if not path.exists():
        return ""
    return path.read_text(encoding="utf-8", errors="ignore").replace("\ufeff", "").strip()


def _extract_jwt(raw: str) -> str:
    text = str(raw or "").replace("\ufeff", "").strip()
    if not text:
        return ""
    if text.lower().startswith("bearer "):
        text = text[7:].strip()
    m = re.search(r"([A-Za-z0-9\-_]+\.[A-Za-z0-9\-_]+\.[A-Za-z0-9\-_]+)", text)
    if m:
        return m.group(1).strip()
    return text


def _parse_cookie_pairs(raw_cookie: str):
    text = str(raw_cookie or "").replace("\ufeff", "")
    if not text.strip():
        return []
    header_match = re.search(
        r"(?ims)\bcookie\b\s*[:=]?\s*(.+?)(?:\n\s*[A-Za-z][A-Za-z0-9\-]*\s*(?::|$)|$)",
        text,
    )
    if header_match:
        cookie_chunk = header_match.group(1)
    else:
        cookie_chunk = text
    cleaned = cookie_chunk.replace("\r", " ").replace("\n", " ").strip()
    out = []
    for part in cleaned.split(";"):
        part = part.strip()
        if not part or "=" not in part:
            continue
        name, value = part.split("=", 1)
        name = name.strip()
        value = value.strip()
        if name:
            out.append((name, value))
    return out


def setup_session(logger, no_prompt: bool = False):
    token_path = Path(SCRIPT_DIR) / TOKEN_FILE
    cookie_path = Path(SCRIPT_DIR) / COOKIE_FILE
    default_token = _extract_jwt(_read_text_if_exists(token_path))
    default_cookie = _read_text_if_exists(cookie_path)

    print("\nCookie и token можно взять из файлов рядом со скриптом:")
    print(f"  cookie: {cookie_path}")
    print(f"  token:  {token_path}")

    if no_prompt:
        cookie_input = ""
        token_input = ""
    else:
        print("Нажмите Enter, чтобы использовать значения из файлов.")
        cookie_input = input("Cookie (или Enter): ").replace("\ufeff", "").strip()
        token_input = input("JWT token (или Enter): ").replace("\ufeff", "").strip()

    cookie_header = cookie_input or default_cookie
    jwt_token = _extract_jwt(token_input or default_token)

    if not cookie_header:
        logger.error("Cookie не введён")
        return None

    session = requests.Session()
    session.verify = False
    session.headers.update(
        {
            "Accept": "application/json, text/plain, */*",
            "Origin": BASE_URL,
            "Referer": BASE_URL.rstrip("/") + "/",
            "User-Agent": "Mozilla/5.0",
        }
    )

    host_match = re.match(r"^https?://([^/:]+)", BASE_URL.strip(), flags=re.IGNORECASE)
    domain = host_match.group(1) if host_match else ""

    pairs = _parse_cookie_pairs(cookie_header)
    for name, value in pairs:
        if domain:
            session.cookies.set(name, value, domain=domain, path="/")
        else:
            session.cookies.set(name, value, path="/")
    if pairs:
        session.headers["Cookie"] = "; ".join([f"{k}={v}" for k, v in pairs])
    logger.info("Parsed cookie pairs: %s", len(pairs))

    if jwt_token:
        clean_token = jwt_token.replace("Bearer ", "").strip()
        session.headers["token"] = clean_token
        session.headers["Authorization"] = "Bearer " + clean_token

    xsrf = session.cookies.get("XSRF-TOKEN") or session.cookies.get("XSRF_TOKEN")
    if xsrf and "X-XSRF-TOKEN" not in session.headers:
        session.headers["X-XSRF-TOKEN"] = xsrf

    test_url = f"{BASE_URL}/api/v1/search/subservices"
    try:
        r = session.post(test_url, json={})
        logger.info(f"[AUTH TEST] POST {test_url} -> {r.status_code} | ct={r.headers.get('content-type')}")
        logger.debug(r.text[:500])

        if r.status_code == 200 and "application/json" in (r.headers.get("content-type") or ""):
            logger.info("✅ Авторизация для API выглядит рабочей")
            return session

        logger.error("❌ Авторизация для API НЕ рабочая (даже если / отдаёт 200).")
        logger.error(f"Ответ (первые 500): {r.text[:500]}")
        return None

    except Exception as e:
        logger.error(f"Ошибка при AUTH TEST: {e}")
        return None


def get_subservices(session, logger):
    search_data = {
        "search": {
            "search": [
                {
                    "field": "version",
                    "operator": "in",
                    "value": ["RKN012"]
                },
                {
                    "field": "notShowInList",
                    "operator": "neq",
                    "value": True
                }
            ]
        },
        "sort": "serviceCode,DESC"
    }
    url = f"{BASE_URL}/api/v1/search/subservices"
    response = api_request(session, logger, "post", url, json=search_data, max_retries=1)

    logger.info(f"Р—Р°РїСЂРѕСЃ Рє {url}, СЃС‚Р°С‚СѓСЃ: {response.status_code}")
    logger.debug(f"РўРµР»Рѕ РѕС‚РІРµС‚Р°: {response.text[:500]}")

    if response.status_code != 200:
        logger.error(f"РћС€РёР±РєР° HTTP {response.status_code}: {response.text}")
        return None

    try:
        result = response.json()
    except requests.exceptions.JSONDecodeError:
        logger.error("РћС‚РІРµС‚ РЅРµ СЏРІР»СЏРµС‚СЃСЏ JSON. Р’РѕР·РјРѕР¶РЅРѕ, РїСЂРѕР±Р»РµРјР° СЃ Р°РІС‚РѕСЂРёР·Р°С†РёРµР№ РёР»Рё URL.")
        logger.error(f"РўРµР»Рѕ РѕС‚РІРµС‚Р° (РїРµСЂРІС‹Рµ 500 СЃРёРјРІРѕР»РѕРІ): {response.text[:500]}")
        return None

    if "content" in result and len(result["content"]) > 0:
        return result["content"]

    logger.warning("РћС‚РІРµС‚ РЅРµ СЃРѕРґРµСЂР¶РёС‚ РґР°РЅРЅС‹С… (РїРѕР»Рµ 'content' РїСѓСЃС‚Рѕ РёР»Рё РѕС‚СЃСѓС‚СЃС‚РІСѓРµС‚)")
    return None


def get_unit(session, params, logger):
    search_org = {
        "page": 0,
        "size": 1,
        "search": {
            "search": []
        }
    }
    for k, v in params.items():
        search_org["search"]["search"].append({
            "field": k,
            "operator": "eq",
            "value": v
        })

    response = api_request(session, logger, "post", f"{BASE_URL}/api/v1/search/organizations", json=search_org, max_retries=1)
    if response.status_code != 200:
        return None

    result = response.json()
    if "content" in result and len(result["content"]) > 0:
        return result["content"][0]
    return None


def create_appeal_data(unit=None, data=None):
    unit_obj = APPEAL_SETTINGS["unit"] if unit is None else unit
    unit_id = APPEAL_SETTINGS["unit"]["id"] if unit is None else unit.get("id")
    number = data.get("number") if data is not None else None
    pin = data.get("pin") if data is not None else None
    executor = data.get("executor") if data is not None else None
    dateFinish = data.get("dateFinish") if data is not None else None

    return {
        "unitId": unit_id,
        "unit": unit_obj,
        "number": APPEAL_SETTINGS["number"] if number is None else number,
        "pin": APPEAL_SETTINGS["pin"] if pin is None else pin,
        "controlOperator": [],
        "events": [],
        "isCustomForm": True,
        "dataForExecuteAction": {},
        "status": APPEAL_SETTINGS["status"],
        "statusHistory": [APPEAL_SETTINGS["status"]],
        "isValid": False,
        "executor": APPEAL_SETTINGS["executor"] if executor is None else executor,
        "dateFinish": APPEAL_SETTINGS["dateFinish"] if dateFinish is None else dateFinish
    }


def create_subservice_data(subserviceTemplate, data=None):
    subservice = {
        "id": subserviceTemplate["_id"],
        "variant": None,
        "title": subserviceTemplate["titles"]["branch"][0]["title"],
        "titles": subserviceTemplate["titles"],
        "shortTitle": subserviceTemplate["titles"]["branch"][0]["shortTitle"],
        "serviceId": subserviceTemplate["serviceId"],
        "guid": generate_guid(),
        "subjects": [],
        "objects": [],
        "entities": [],
        "standardCode": subserviceTemplate["standardCode"],
        "version": subserviceTemplate["version"],
        "appealsCollection": APPEAL_SETTINGS["parentEntries"],
        "parentEntries": f"{APPEAL_SETTINGS['parentEntries']}.subservices",
        "responsibleOrganizations": None,
        "xsd": subserviceTemplate.get("xsd"),
        "mainElement": subserviceTemplate.get("mainElement"),
        "xsdData": {},
        "xsdRequired": True,
        "status": APPEAL_SETTINGS["status"],
        "statusHistory": [APPEAL_SETTINGS["status"]],
        "appealXsdDataValid": True,
        "xsdDataValid": False
    }
    if data:
        subservice.update(data)
    return subservice


def create_mainElement_data(data=None):
    operationType = data.get("operationType") if data else None
    xsdData = data.get("xsdData") if data else None
    xsd = data.get("xsd") if data else None

    return {
        "xsd": xsd,
        "titles": {
            "common": {},
            "object": {},
            "subject": {}
        },
        "xsdData": {} if xsdData is None else xsdData,
        "objectXsd": None,
        "subjectXsd": None,
        "objectMainXsd": None,
        "operationType": APPEAL_SETTINGS["operationType"] if operationType is None else operationType,
        "registryParams": {
            "structure": [],
            "useChecksTab": True,
            "useHistoryTab": True,
            "registersTabName": ""
        },
        "subjectMainXsd": None,
        "registryEntryType": APPEAL_SETTINGS["registryEntryType"]
    }


def create_subject_data(template, data=None):
    subject = copy.deepcopy(template)
    if data is not None:
        for k, v in data.items():
            subject[k] = v
    return subject


def create_appeal_with_entities(session, logger, appeal_data, subservice_data=None, subject_data=None, document_data=None, files_contents=None):
    appeal_url = f"{BASE_URL}/api/v1/create/{APPEAL_SETTINGS['parentEntries']}"

    try:
        logger.info("РћС‚РїСЂР°РІРєР° Р·Р°РїСЂРѕСЃР° РЅР° СЃРѕР·РґР°РЅРёРµ РѕР±СЂР°С‰РµРЅРёСЏ...")
        appeal_response = api_request(session, logger, "post", appeal_url, json=jsonable(appeal_data), max_retries=1)

        if appeal_response.status_code not in (200, 201):
            logger.error(f"РћС€РёР±РєР° СЃРѕР·РґР°РЅРёСЏ РѕР±СЂР°С‰РµРЅРёСЏ: {appeal_response.status_code}")
            logger.error(f"РўРµР»Рѕ РѕС‚РІРµС‚Р°: {appeal_response.text[:500]}")
            return False, None, None, None, None

        appeal = appeal_response.json()
        logger.info(f"вњ… РћР±СЂР°С‰РµРЅРёРµ СЃРѕР·РґР°РЅРѕ. ID: {appeal.get('_id')}, GUID: {appeal.get('guid')}")
    except Exception as e:
        logger.error(f"РСЃРєР»СЋС‡РµРЅРёРµ РїСЂРё СЃРѕР·РґР°РЅРёРё РѕР±СЂР°С‰РµРЅРёСЏ: {e}")
        return False, None, None, None, None

    appeal_id = appeal.get("_id")
    appeal_guid = appeal.get("guid")

    if not appeal_id or not appeal_guid:
        logger.error("РћС‚РІРµС‚ РѕС‚ СЃРѕР·РґР°РЅРёСЏ РѕР±СЂР°С‰РµРЅРёСЏ РЅРµ СЃРѕРґРµСЂР¶РёС‚ _id РёР»Рё guid")
        return False, appeal, None, None, None

    subservice = None
    subject = None
    document = None

    if subservice_data is not None:
        subservice_url = (
            f"{BASE_URL}/api/v1/create/{APPEAL_SETTINGS['parentEntries']}/subservices"
            f"?mainId={appeal_id}&parentGuid={appeal_guid}&parentEntries={APPEAL_SETTINGS['parentEntries']}.subservices"
        )

        try:
            logger.info("РћС‚РїСЂР°РІРєР° Р·Р°РїСЂРѕСЃР° РЅР° СЃРѕР·РґР°РЅРёРµ subservice...")
            subservice_response = api_request(session, logger, "post", subservice_url, json=jsonable(subservice_data), max_retries=1)
            if subservice_response.status_code not in (200, 201):
                logger.error(f"РћС€РёР±РєР° СЃРѕР·РґР°РЅРёСЏ subservice: {subservice_response.status_code}")
                logger.error(f"РўРµР»Рѕ РѕС‚РІРµС‚Р°: {subservice_response.text[:500]}")
                return False, appeal, None, None, None

            subservice = subservice_response.json()
            logger.info("вњ… Subservice СѓСЃРїРµС€РЅРѕ СЃРѕР·РґР°РЅ")
        except Exception as e:
            logger.error(f"РСЃРєР»СЋС‡РµРЅРёРµ РїСЂРё СЃРѕР·РґР°РЅРёРё subservice: {e}")
            return False, appeal, None, None, None

    if subject_data is not None:
        subject_url = (
            f"{BASE_URL}/api/v1/create/{APPEAL_SETTINGS['parentEntries']}/subjects"
            f"?mainId={appeal_id}&parentGuid={appeal_guid}&parentEntries={APPEAL_SETTINGS['parentEntries']}.subjects"
        )
        try:
            logger.info("РћС‚РїСЂР°РІРєР° Р·Р°РїСЂРѕСЃР° РЅР° СЃРѕР·РґР°РЅРёРµ subject...")
            subject_response = api_request(session, logger, "post", subject_url, json=jsonable(subject_data), max_retries=1)

            if subject_response.status_code not in (200, 201):
                logger.error(f"РћС€РёР±РєР° СЃРѕР·РґР°РЅРёСЏ subject: {subject_response.status_code}")
                logger.error(f"РўРµР»Рѕ РѕС‚РІРµС‚Р°: {subject_response.text[:500]}")
                return False, appeal, subservice, None, None

            subject = subject_response.json()
            logger.info("вњ… Subject СѓСЃРїРµС€РЅРѕ СЃРѕР·РґР°РЅ")
        except Exception as e:
            logger.error(f"РСЃРєР»СЋС‡РµРЅРёРµ РїСЂРё СЃРѕР·РґР°РЅРёРё subject: {e}")
            return False, appeal, subservice, None, None

    if document_data is not None:
        document_data["subserviceGuid"] = subservice.get("guid") if subservice else None
        document_url = (
            f"{BASE_URL}/api/v1/create/{APPEAL_SETTINGS['parentEntries']}/documents"
            f"?mainId={appeal_id}&parentGuid={appeal_guid}&parentEntries={APPEAL_SETTINGS['parentEntries']}.documents"
        )
        try:
            logger.info("РћС‚РїСЂР°РІРєР° Р·Р°РїСЂРѕСЃР° РЅР° СЃРѕР·РґР°РЅРёРµ document...")
            document_response = api_request(session, logger, "post", document_url, json=jsonable(document_data), max_retries=1)

            if document_response.status_code not in (200, 201):
                logger.error(f"РћС€РёР±РєР° СЃРѕР·РґР°РЅРёСЏ document: {document_response.status_code}")
                logger.error(f"РўРµР»Рѕ РѕС‚РІРµС‚Р°: {document_response.text[:500]}")
                return False, appeal, subservice, subject, None

            document = document_response.json()
            logger.info("вњ… Document СѓСЃРїРµС€РЅРѕ СЃРѕР·РґР°РЅ")
        except Exception as e:
            logger.error(f"РСЃРєР»СЋС‡РµРЅРёРµ РїСЂРё СЃРѕР·РґР°РЅРёРё document: {e}")
            return False, appeal, subservice, subject, None

    if files_contents is not None and document is not None:
        document_url = (
            f"{BASE_URL}/api/v1/update/{APPEAL_SETTINGS['parentEntries']}/documents"
            f"?mainId={appeal_id}&guid={document['guid']}&parentEntries={APPEAL_SETTINGS['parentEntries']}.documents"
        )
        try:
            file_metas = []
            file_upload_exception = False

            for b64, fileName in files_contents:
                file_meta = upload_file_to_stend(
                    session=session,
                    logger=logger,
                    filename=fileName,
                    base64_content=b64,
                    entry_id=appeal_id,
                    entity_field_path=""
                )
                if file_meta:
                    file_metas.append(file_meta)
                else:
                    file_upload_exception = True
                    break

            if file_upload_exception:
                for f in file_metas:
                    delete_file_from_storage(session, logger, f.get("_id"))
                return False, appeal, subservice, subject, document

            document["files"] = file_metas
            logger.info("РћС‚РїСЂР°РІРєР° Р·Р°РїСЂРѕСЃР° РЅР° РґРѕР±Р°РІР»РµРЅРёРµ РІ document files...")
            document_response = api_request(session, logger, "put", document_url, json=jsonable(document), max_retries=1)

            if document_response.status_code not in (200, 201):
                logger.error(f"РћС€РёР±РєР° РѕР±РЅРѕРІР»РµРЅРёСЏ document: {document_response.status_code}")
                logger.error(f"РўРµР»Рѕ РѕС‚РІРµС‚Р°: {document_response.text[:500]}")
                return False, appeal, subservice, subject, document

            document = document_response.json()
            logger.info("вњ… Р¤Р°Р№Р»С‹ СѓСЃРїРµС€РЅРѕ РґРѕР±Р°РІР»РµРЅС‹ РІ document")
        except Exception as e:
            logger.error(f"РСЃРєР»СЋС‡РµРЅРёРµ РїСЂРё РѕР±РЅРѕРІР»РµРЅРёРё document: {e}")
            return False, appeal, subservice, subject, document

    return True, appeal, subservice, subject, document


def delete_from_collection(session, logger, data):
    """
    РЈРЅРёРІРµСЂСЃР°Р»СЊРЅРѕРµ СѓРґР°Р»РµРЅРёРµ Р·Р°РїРёСЃРё РёР· РєРѕР»Р»РµРєС†РёРё
    data = {
        "_id": ...,
        "guid": ...,
        "parentEntries": ...,
        ...
    }
    """
    main_id = data.get("_id")
    guid = data.get("guid")
    parent_entries = data.get("parentEntries")

    if not main_id or not guid or not parent_entries:
        logger.error(f"вќЊ РќРµРІРѕР·РјРѕР¶РЅРѕ СѓРґР°Р»РёС‚СЊ: РѕС‚СЃСѓС‚СЃС‚РІСѓРµС‚ _id, guid РёР»Рё parent_entries РІ РґР°РЅРЅС‹С… РґР»СЏ РєРѕР»Р»РµРєС†РёРё")
        return False

    url = f"{BASE_URL}/api/v1/delete/{parent_entries}?mainId={main_id}&guid={guid}"
    try:
        logger.info(f"РћС‚РїСЂР°РІРєР° DELETE-Р·Р°РїСЂРѕСЃР° РґР»СЏ {parent_entries} вЂ” _id: {main_id}, guid: {guid}")
        response = api_request(session, logger, "delete", url, max_retries=1)

        if response.status_code in (200, 204, 202):
            logger.info(f"вњ… Р—Р°РїРёСЃСЊ СѓСЃРїРµС€РЅРѕ СѓРґР°Р»РµРЅР°: {parent_entries} вЂ” {main_id} ({guid})")
            return True
        if response.status_code == 404 or response.status_code == 500:
            logger.info(f"в„№пёЏ Р—Р°РїРёСЃСЊ РЅРµ РЅР°Р№РґРµРЅР° (РІРѕР·РјРѕР¶РЅРѕ, СѓР¶Рµ СѓРґР°Р»РµРЅР°): {parent_entries} вЂ” {main_id} ({guid})")
            return True

        logger.error(f"вќЊ РћС€РёР±РєР° СѓРґР°Р»РµРЅРёСЏ: СЃС‚Р°С‚СѓСЃ {response.status_code}")
        logger.error(f"РўРµР»Рѕ РѕС‚РІРµС‚Р°: {response.text[:500]}")
        return False

    except Exception as e:
        logger.error(f"вќЊ РСЃРєР»СЋС‡РµРЅРёРµ РїСЂРё СѓРґР°Р»РµРЅРёРё РёР· {parent_entries}: {e}")
        return False

