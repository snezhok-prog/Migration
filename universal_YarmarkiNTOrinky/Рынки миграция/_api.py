import copy
import json
import os
import mimetypes
import re
from typing import Optional

import requests

from pathlib import Path
from _config import (
    APPEAL_SETTINGS,
    AUTH_TEST_COLLECTION,
    BASE_URL,
    COOKIE_FILE,
    JWT_URL,
    REAUTH_ON_401,
    REAUTH_RETRIES,
    SAVE_AUTH,
    SCRIPT_DIR,
    STANDARD_CODES,
    TOKEN_FILE,
    UI_BASE_URL,
    VERIFY_SSL,
)
from _utils import jsonable, generate_guid


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
        _RUNTIME_JWT_URL = _RUNTIME_BASE_URL.rstrip("/") + "/jwt/"
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


def _safe_json(resp: requests.Response):
    try:
        return resp.json()
    except Exception:
        return resp.text


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
    Обязателен JWT токен, без него ничего не загрузится
    """
    file_name = os.path.basename(file_path)
    url = _build_url("/api/v1/storage/upload")
    
    # 🔥 ВАЖНО: все поля формы — в data, а не в params!
    data = {
        'entryName': entry_name,      # ← было в params, теперь здесь
        'entryId': entry_id,          # ← было в params, теперь здесь
        'entityFieldPath': entity_field_path,
        'allowExternal': 'false'
    }
    
    if not os.path.isfile(file_path):
        logger.error(f"Файл не найден: {file_path}")
        return None
    
    if not file_name:#123
        file_name = Path(file_path).name
    mime_type, _ = mimetypes.guess_type(file_path)
    content_type = mime_type or "application/octet-stream"
    
    logger.info(f"Подготовка к загрузке: {file_name} ({content_type})")
    
    # Обновляем Referer под конкретное дело
    session.headers["Referer"] = _RUNTIME_BASE_URL + f"/AppRKN034/common-appeals/edit/{entry_id}"
    session.headers["Origin"] = _RUNTIME_BASE_URL
    
    try:
        with open(file_path, 'rb') as f:
            files = {'file': (file_name, f, content_type)}
            
            logger.debug(f"🔍 Отправка: url={url}, data={data}")
            logger.debug(f"🔍 Cookies: {list(session.cookies.keys())}")
            
            # 🔥 Убираем params= — все данные в теле формы
            response = api_request(session, logger, "post", url, files=files, data=data)
            # response = session.post(
            #     url,
            #     files=files,
            #     data=data,  # ← все поля формы здесь
            #     timeout=120
            # )
    except Exception as e:
        logger.error(f"Ошибка при загрузке: {type(e).__name__}: {e}")
        return None

    logger.info(f"Запрос к {url}, статус: {response.status_code}")
    
    if response.status_code not in (200, 201, 202):
        logger.error(f"Ошибка HTTP {response.status_code}: {response.text}")
        return None

    try:
        result = response.json()
    except requests.exceptions.JSONDecodeError:
        if response.status_code in (200, 201, 202) and not response.text.strip():
            return {"status": "uploaded", "fileName": file_name}
        logger.error("Ответ не является JSON")
        return None

    if isinstance(result, dict) and ("error" in result or result.get("success") is False):
        logger.error(f"API вернуло ошибку: {result}")
        return None

    logger.info(f"✅ Файл {file_name} загружен")
    return result


def delete_file_from_storage(session, logger, file_id: str):
    """
    Удаляет файл из хранилища по fileId.
    
    Args:
        session (requests.Session): Авторизованная сессия (с куками и токеном)
        logger (logging.Logger): Логгер
        file_id (str): Идентификатор файла (например, '6946daca2899a5480fe402dd')
    
    Returns:
        bool: True — если удаление прошло успешно, False — при ошибке
    """
    # Убираем возможные пробелы в file_id
    file_id = file_id.strip()
    
    url = _build_url("/api/v1/storage/remove")
    params = {"fileId": file_id}
    
    # Явно указываем заголовки, как в fetch (хотя session и так отправит куки)
    headers = {
        "accept": "application/hal+json",
        "content-type": "application/json"
    }
    
    logger.info(f"Удаление файла из хранилища: fileId={file_id}")
    
    try:
        response = session.delete(url, params=params, headers=headers)
        logger.info(f"Статус удаления файла: {response.status_code}")
        
        # Успешные статусы: 200, 204, иногда 202
        if response.status_code in (200, 204, 202):
            logger.info("✅ Файл успешно удалён из хранилища")
            return True
        else:
            logger.error(f"❌ Ошибка удаления файла: {response.status_code}")
            logger.error(f"Тело ответа: {response.text[:500]}")
            return False
            
    except Exception as e:
        logger.error(f"🔥 Исключение при удалении файла: {e}")
        return False


def api_request(session, logger, method, url, reauth_fn=None, max_retries=3, **kwargs):
    """Унифицированный запрос к API с повторной авторизацией на 401/403/500."""
    if reauth_fn is None:
        reauth_fn = lambda log: setup_session(log, no_prompt=True)

    for attempt in range(max_retries + 1):
        try:
            request_fn = getattr(session, method.lower(), None)
            if request_fn is None:
                raise ValueError(f"Неизвестный HTTP метод: {method}")

            response = request_fn(url, **kwargs)
            if response.status_code in (401, 403) or response.status_code == 500 and method.lower() != "delete":
                logger.warning(f"HTTP {response.status_code} от {url}. Попытка реавторизации {attempt + 1}/{max_retries}")
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
            logger.error(f"Ошибка в api_request ({method.upper()} {url}): {e}")
            if attempt < max_retries:
                continue
            raise

    raise RuntimeError("api_request: исчерпаны попытки")


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
    session.verify = VERIFY_SSL
    session.headers.update(
        {
            "Accept": "application/json, text/plain, */*",
            "Origin": _RUNTIME_BASE_URL,
            "Referer": _RUNTIME_BASE_URL.rstrip("/") + "/",
            "User-Agent": "Mozilla/5.0",
        }
    )

    host_match = re.match(r"^https?://([^/:]+)", _RUNTIME_BASE_URL.strip(), flags=re.IGNORECASE)
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

    test_url = _build_url("/api/v1/search/subservices")
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
    url = _build_url("/api/v1/search/subservices")
    response = api_request(session, logger, "post", url, json=search_data, max_retries=1)

    logger.info(f"Запрос к {url}, статус: {response.status_code}")
    logger.debug(f"Тело ответа: {response.text[:500]}")

    if response.status_code != 200:
        logger.error(f"Ошибка HTTP {response.status_code}: {response.text}")
        return None

    try:
        result = response.json()
    except requests.exceptions.JSONDecodeError:
        logger.error("Ответ не является JSON. Возможно, проблема с авторизацией или URL.")
        logger.error(f"Тело ответа (первые 500 символов): {response.text[:500]}")
        return None

    if "content" in result and len(result["content"]) > 0:
        return result["content"]

    logger.warning("Ответ не содержит данных (поле 'content' пусто или отсутствует)")
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

    response = api_request(session, logger, "post", _build_url("/api/v1/search/organizations"), json=search_org, max_retries=1)
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
    appeal_url = _build_url(f"/api/v1/create/{APPEAL_SETTINGS['parentEntries']}")

    try:
        logger.info("Отправка запроса на создание обращения...")
        appeal_response = api_request(session, logger, "post", appeal_url, json=jsonable(appeal_data), max_retries=1)

        if appeal_response.status_code not in (200, 201):
            logger.error(f"Ошибка создания обращения: {appeal_response.status_code}")
            logger.error(f"Тело ответа: {appeal_response.text[:500]}")
            return False, None, None, None, None

        appeal = appeal_response.json()
        logger.info(f"✅ Обращение создано. ID: {appeal.get('_id')}, GUID: {appeal.get('guid')}")
    except Exception as e:
        logger.error(f"Исключение при создании обращения: {e}")
        return False, None, None, None, None

    appeal_id = appeal.get("_id")
    appeal_guid = appeal.get("guid")

    if not appeal_id or not appeal_guid:
        logger.error("Ответ от создания обращения не содержит _id или guid")
        return False, appeal, None, None, None

    subservice = None
    subject = None
    document = None

    if subservice_data is not None:
        subservice_url = (
            _build_url(f"/api/v1/create/{APPEAL_SETTINGS['parentEntries']}/subservices")
            + 
            f"?mainId={appeal_id}&parentGuid={appeal_guid}&parentEntries={APPEAL_SETTINGS['parentEntries']}.subservices"
        )

        try:
            logger.info("Отправка запроса на создание subservice...")
            subservice_response = api_request(session, logger, "post", subservice_url, json=jsonable(subservice_data), max_retries=1)
            if subservice_response.status_code not in (200, 201):
                logger.error(f"Ошибка создания subservice: {subservice_response.status_code}")
                logger.error(f"Тело ответа: {subservice_response.text[:500]}")
                return False, appeal, None, None, None

            subservice = subservice_response.json()
            logger.info("✅ Subservice успешно создан")
        except Exception as e:
            logger.error(f"Исключение при создании subservice: {e}")
            return False, appeal, None, None, None

    if subject_data is not None:
        subject_url = (
            _build_url(f"/api/v1/create/{APPEAL_SETTINGS['parentEntries']}/subjects")
            +
            f"?mainId={appeal_id}&parentGuid={appeal_guid}&parentEntries={APPEAL_SETTINGS['parentEntries']}.subjects"
        )
        try:
            logger.info("Отправка запроса на создание subject...")
            subject_response = api_request(session, logger, "post", subject_url, json=jsonable(subject_data), max_retries=1)

            if subject_response.status_code not in (200, 201):
                logger.error(f"Ошибка создания subject: {subject_response.status_code}")
                logger.error(f"Тело ответа: {subject_response.text[:500]}")
                return False, appeal, subservice, None, None

            subject = subject_response.json()
            logger.info("✅ Subject успешно создан")
        except Exception as e:
            logger.error(f"Исключение при создании subject: {e}")
            return False, appeal, subservice, None, None

    if document_data is not None:
        document_data["subserviceGuid"] = subservice.get("guid") if subservice else None
        document_url = (
            _build_url(f"/api/v1/create/{APPEAL_SETTINGS['parentEntries']}/documents")
            +
            f"?mainId={appeal_id}&parentGuid={appeal_guid}&parentEntries={APPEAL_SETTINGS['parentEntries']}.documents"
        )
        try:
            logger.info("Отправка запроса на создание document...")
            document_response = api_request(session, logger, "post", document_url, json=jsonable(document_data), max_retries=1)

            if document_response.status_code not in (200, 201):
                logger.error(f"Ошибка создания document: {document_response.status_code}")
                logger.error(f"Тело ответа: {document_response.text[:500]}")
                return False, appeal, subservice, subject, None

            document = document_response.json()
            logger.info("✅ Document успешно создан")
        except Exception as e:
            logger.error(f"Исключение при создании document: {e}")
            return False, appeal, subservice, subject, None

    if files_contents is not None and document is not None:
        document_url = (
            _build_url(f"/api/v1/update/{APPEAL_SETTINGS['parentEntries']}/documents")
            +
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
            logger.info("Отправка запроса на добавление в document files...")
            document_response = api_request(session, logger, "put", document_url, json=jsonable(document), max_retries=1)

            if document_response.status_code not in (200, 201):
                logger.error(f"Ошибка обновления document: {document_response.status_code}")
                logger.error(f"Тело ответа: {document_response.text[:500]}")
                return False, appeal, subservice, subject, document

            document = document_response.json()
            logger.info("✅ Файлы успешно добавлены в document")
        except Exception as e:
            logger.error(f"Исключение при обновлении document: {e}")
            return False, appeal, subservice, subject, document

    return True, appeal, subservice, subject, document


def delete_from_collection(session, logger, data):
    """
    Универсальное удаление записи из коллекции
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
        logger.error(f"❌ Невозможно удалить: отсутствует _id, guid или parent_entries в данных для коллекции")
        return False

    url = _build_url(f"/api/v1/delete/{parent_entries}?mainId={main_id}&guid={guid}")
    try:
        logger.info(f"Отправка DELETE-запроса для {parent_entries} — _id: {main_id}, guid: {guid}")
        response = api_request(session, logger, "delete", url, max_retries=1)

        if response.status_code in (200, 204, 202):
            logger.info(f"✅ Запись успешно удалена: {parent_entries} — {main_id} ({guid})")
            return True
        if response.status_code == 404 or response.status_code == 500:
            logger.info(f"ℹ️ Запись не найдена (возможно, уже удалена): {parent_entries} — {main_id} ({guid})")
            return True

        logger.error(f"❌ Ошибка удаления: статус {response.status_code}")
        logger.error(f"Тело ответа: {response.text[:500]}")
        return False

    except Exception as e:
        logger.error(f"❌ Исключение при удалении из {parent_entries}: {e}")
        return False

