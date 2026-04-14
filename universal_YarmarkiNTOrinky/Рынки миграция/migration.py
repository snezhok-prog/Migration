from __future__ import annotations

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
import requests
from urllib3.exceptions import InsecureRequestWarning

from _api import (
    api_request,
    create_appeal_data,
    create_appeal_with_entities,
    create_mainElement_data,
    create_subject_data,
    create_subservice_data,
    delete_file_from_storage,
    delete_from_collection,
    get_runtime_base_url,
    get_runtime_ui_base_url,
    get_standard_code,
    get_subservices,
    get_unit,
    set_runtime_urls,
    setup_session,
    upload_file,
)
from _config import (
    BASE_URL,
    EXCEL_FILE_NAME,
    EXCEL_INPUT_GLOB,
    EXCEL_LISTS,
    FILES_DIR,
    JWT_URL,
    LICENSES_COLLECTION,
    RECORDS_COLLECTION,
    RECORDS_TEMPLATES,
    RESUME_BY_DEFAULT,
    SCRIPT_DIR,
    STANDARD_CODES,
    STATE_FILE,
    TEST,
    UI_BASE_URL,
    UNIT,
)
from _excel_input import discover_excel_files
from _logger import setup_fail_logger, setup_logger, setup_success_logger, setup_user_logger
from _profiles import PROFILES
from _state import ResumeState
from _templates import SUBJECT_IP, SUBJECT_UL
from _utils import (
    find_document_group_by_mnemonic,
    find_file_in_dir,
    format_multiple_phones,
    format_phone,
    generate_guid,
    jsonable,
    nz,
    parse_date_to_birthday_obj,
    parse_key_value_mapping,
    parse_path_list,
    read_excel,
    read_file_as_base64,
    split_sc,
    to_iso_date,
)

NSI_LOCAL_OBJECT_MARKET_COLLECTION = "nsiLocalObjectMarket"

def normStr(s):
    if s is None:
        return None
    s = str(s).strip()
    return s if s else None

def split_postal_address(s):
    raw = str(s).strip()
    if not raw:
        return {"postalCode": None, "fullAddress": None}

    # "236006, Калининград ..." -> postalCode=236006, fullAddress="Калининград ..."
    m = re.match(r"^(\d{6})(?:,\s*)?(.*)$", raw)
    if not m:
        return {"postalCode": None, "fullAddress": raw}

    postal_code = m.group(1)
    rest = (m.group(2) or "").strip()

    if not rest:
        rest = None

    return {
        "postalCode": postal_code,
        "fullAddress": rest
    }

def norm_key(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())

def dash_str(value):
    value = normStr(value)
    return value if value is not None else "-"

def format_date_to_dmy(value):
    text = normStr(value)
    if not text:
        return None
    try:
        dt = pd.to_datetime(text, errors="coerce", dayfirst=True)
        if pd.isna(dt):
            return None
        return dt.strftime("%d.%m.%Y")
    except Exception:
        return None

def collect_non_empty_values(row, column_names):
    values = []
    for column_name in column_names:
        value = normStr(row.get(column_name))
        if value is not None:
            values.append(value)
    return values

def build_nsi_local_object_market_payload(row, unit, guid):
    market_specialization = normStr(row.get("Специализация рынка"))
    market_other_specialization = normStr(row.get("Иная специализация рынка"))
    if norm_key(market_specialization) == "иная":
        market_specialization_value = market_other_specialization or market_specialization
    else:
        market_specialization_value = market_specialization or market_other_specialization

    full_address = "; ".join(collect_non_empty_values(
        row,
        [
            "1. Местоположение объекта или объектов недвижимости, где предполагается организовать рынок",
            "2. Местоположение объекта или объектов недвижимости, где предполагается организовать рынок",
            "3. Местоположение объекта или объектов недвижимости, где предполагается организовать рынок"
        ]
    ))
    cad_number = "; ".join(collect_non_empty_values(
        row,
        [
            "1. Кадастровый номер объекта недвижимости",
            "2. Кадастровый номер объекта недвижимости",
            "3. Кадастровый номер объекта недвижимости"
        ]
    ))

    unit_id = None
    if isinstance(unit, dict):
        unit_id = unit.get("id") or unit.get("_id")

    return {
        "guid": dash_str(guid),
        "code": dash_str(guid),
        "autokey": dash_str(guid),
        "ObjectID": dash_str(guid),
        "parentEntries": NSI_LOCAL_OBJECT_MARKET_COLLECTION,
        "LayerId": "2",
        "Layer": "Розничные рынки",
        "Subject": dash_str(row.get("Субъект РФ")),
        "Disctrict": dash_str(row.get("Муниципальный район/округ, городской округ или внутригородская территория")),
        "FullAddress": dash_str(full_address),
        "GeoCoordinates": dash_str(row.get("Геокоординаты точки, на которой расположен рынок")),
        "CadNumber": dash_str(cad_number),
        "PermissionNumber": dash_str(row.get("Номер разрешения")),
        "PermissionStartDate": dash_str(format_date_to_dmy(row.get("Дата выдачи разрешения"))),
        "PermissionEndDate": dash_str(format_date_to_dmy(row.get("Дата завершения действия разрешения"))),
        "PermissionStatus": dash_str(row.get("Статус разрешения")),
        "TitleMarket": dash_str(row.get("Наименование рынка")),
        "TypeMarket": dash_str(row.get("Тип рынка")),
        "MarketSpecialization": dash_str(market_specialization_value),
        "OperatingPeriod": dash_str(row.get("Период функционирования рынка")),
        "StartTimeMarket": dash_str(row.get("Основное время начала работы рынка")),
        "EndTimeMarket": dash_str(row.get("Основное время окончания работы рынка")),
        "PlaceNumber": dash_str(row.get("Число торговых мест, шт.")),
        "OperatorName": dash_str(row.get("Наименование оператора")),
        "OperatorINN": dash_str(row.get("ИНН оператора")),
        "OperatorOGRN": dash_str(row.get("ОГРН оператора")),
        "OperatorNumber": dash_str(row.get("Контактный номер оператора")),
        "OperatorEmail": dash_str(row.get("Адрес электронной почты оператора")),
        "dictionaryType": "local",
        "dictionaryUnitId": dash_str(unit_id)
    }

def log_success(success_logger, record):
    success_logger.info(json.dumps(record, ensure_ascii=False))


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
    action = _ask_error_action()
    if action == "retry":
        logger.warning("Retry для данного типа ошибки не поддержан, строка будет пропущена")
    return action


def _selected_lists(sheet_mode: str) -> List[str]:
    mapping = {
        "permits": "2. Реестр разрешений",
        "markets": "3. Реестр рынков",
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
    workbook_abs = str(Path(workbook_path).resolve())
    workbook_name = Path(workbook_abs).name
    workbook_stem = Path(workbook_abs).stem
    mapped_raw = files_map.get(workbook_abs) or files_map.get(workbook_name) or files_map.get(workbook_stem)

    default_dir = str(Path(files_root).resolve())
    if mapped_raw:
        mapped_path = Path(str(mapped_raw).strip())
        if not mapped_path.is_absolute():
            mapped_path = (Path(SCRIPT_DIR) / mapped_path).resolve()
        default_dir = str(mapped_path)
    else:
        root = Path(default_dir)
        if root.exists() and root.is_dir():
            for hint in _numeric_hints(workbook_name):
                for folder_name in (
                    hint,
                    f"part{hint}",
                    f"book{hint}",
                    "one" if hint == "1" else "",
                    "two" if hint == "2" else "",
                    "three" if hint == "3" else "",
                ):
                    if not folder_name:
                        continue
                    candidate = root / folder_name
                    if candidate.exists() and candidate.is_dir():
                        default_dir = str(candidate.resolve())
                        break

    if interactive and (prompt_always or not mapped_raw):
        chosen = _prompt_with_default(f"Files directory for {workbook_name}", default_dir, True)
        return str(Path(chosen).expanduser().resolve())
    return str(Path(default_dir).expanduser().resolve())


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

    specs: List[WorkbookRunSpec] = []
    for workbook_path in selected:
        workbook_abs = str(Path(workbook_path).expanduser().resolve())
        files_dir = _infer_files_dir_for_workbook(
            workbook_path=workbook_abs,
            files_root=files_root,
            files_map=files_map,
            interactive=interactive,
            prompt_always=bool(args.ask_files_always),
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
        choice = normStr(raw).lower()
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
        "===== СТАРТ МИГРАЦИИ РЫНКОВ =====",
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
    parser = argparse.ArgumentParser(description="Миграция данных реестров рынков")
    parser.add_argument("--profile", choices=["custom", "dev", "psi", "prod"], default="dev")
    parser.add_argument("--base-url", default="")
    parser.add_argument("--jwt-url", default="")
    parser.add_argument("--ui-base-url", default="")

    parser.add_argument("--mode", choices=["auto", "single", "mass"], default="auto")
    parser.add_argument("--workbook", default=str(Path(SCRIPT_DIR) / EXCEL_FILE_NAME), help="Путь к книге Excel (совместимость)")
    parser.add_argument("--workbooks", default="", help="Явный список книг (разделитель ';' или новая строка)")
    parser.add_argument("--files-dir", default=FILES_DIR, help="Папка с файлами для загрузки")
    parser.add_argument("--files-map", default="", help="Связка книга->папка файлов: 'book1.xlsm=dir1;book2.xlsm=dir2'")
    parser.add_argument("--ask-files-always", action="store_true", help="Всегда спрашивать папку файлов для каждой книги")

    parser.add_argument("--sheet", choices=["all", "permits", "markets"], default="all", help="Какие листы запускать")
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


def main():
    args = parse_args()
    warnings.filterwarnings("ignore", category=InsecureRequestWarning)

    logger = setup_logger()
    successLogger = setup_success_logger()
    failLogger = setup_fail_logger()
    userLogger = setup_user_logger()

    interactive = not args.no_interactive and not args.no_prompt
    try:
        profile_name, runtime_base_url, runtime_jwt_url, runtime_ui_url = _resolve_runtime(args)
    except Exception as exc:
        logger.error("Ошибка runtime конфигурации: %s", exc)
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
        logger.error("Ошибка выбора книг/папок: %s", exc)
        return 1

    state_path = Path(args.state_file).expanduser().resolve()
    state_namespace = f"markets_migration:{profile_name}"
    state = ResumeState(path=state_path, namespace=state_namespace, enabled=True)
    if args.reset_state:
        state.reset_namespace()
        logger.info("State reset: %s", state_path)

    resume_enabled = bool(args.resume)
    if resume_enabled:
        strategy = _choose_resume_strategy(state=state, logger=logger, user_logger=userLogger, interactive=interactive)
        if strategy == "abort":
            logger.warning("Остановка по выбору оператора перед стартом")
            return 1
        if strategy == "reset":
            state.clear_rows()
            resume_enabled = False
            logger.info("Checkpoints очищены перед стартом")

    logger.info(
        "%s",
        _console_block(
            "Markets migration start",
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
        user_logger=userLogger,
        profile=profile_name,
        base_url=get_runtime_base_url(),
        mode=resolved_mode,
        interactive=interactive,
        operator_mode=bool(args.operator_mode),
        state_file_path=str(state_path),
        success_log_path=getattr(successLogger, "log_path", ""),
        fail_log_path=getattr(failLogger, "log_path", ""),
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
    processed_rows = 0
    created_rows = 0
    resumed_skips = 0
    failed_rows = 0

    try:
        session = None
        if not skip_auth:
            session = setup_session(logger, no_prompt=(args.no_prompt or args.no_interactive))
            if session is None:
                raise RuntimeError("Авторизация не выполнена")
            logger.info("Авторизация выполнена успешно")

        if args.auth_only:
            logger.info("--auth-only: авторизация проверена, миграция не запускается")
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
            excel_path = str(workbook_path)
            if not workbook_path.exists():
                raise FileNotFoundError(f"Книга не найдена: {workbook_path}")

            logger.info("%s", _console_block("Workbook", [f"Path: {workbook_path}", f"Files: {files_dir_active}"]))
            userLogger.info(
                "WORKBOOK_START | workbook=%s | files=%s | sheet=%s | resume=%s",
                excel_path,
                files_dir_active,
                args.sheet,
                resume_enabled,
            )

            for list_name in _selected_lists(args.sheet):
                logger.info("Обработка листа: %s | книга: %s", list_name, workbook_path.name)

                excel = read_excel(excel_path, skiprows=3, sheet_name=list_name)
                if excel is None:
                    raise RuntimeError(f"Не удалось прочитать лист '{list_name}' из книги {excel_path}")
                excel = excel.iloc[1:].reset_index(drop=True)  # Удаляем строку с подписями и сбрасываем индекс

                logger.info("Загружено строк: %s", len(excel))
                excel.columns = [c.strip() for c in excel.columns]
                rows_total = len(excel)
                for i, row in enumerate(excel.to_dict("records"), start=1):
                    logger.info("%s/%s", i, rows_total)
                    if args.limit and i > int(args.limit):
                        logger.info("Достигнут лимит строк (%s) для листа %s", args.limit, list_name)
                        break

                    if resume_enabled:
                        checkpoint = state.get(str(workbook_path), list_name, i)
                        if isinstance(checkpoint, dict):
                            resumed_skips += 1
                            logger.info(
                                "[RESUME][SKIP] workbook=%s sheet=%s row=%s _id=%s",
                                workbook_path.name,
                                list_name,
                                i,
                                checkpoint.get("_id"),
                            )
                            continue

                    processed_rows += 1
                    unit = UNIT

                    if TEST:
                        print(row)
                    MAP_PERMISSION_STATUS = {
                        "действует":      { "code": "Working",     "name": "Действует" },
                        "приостановлено": { "code": "Stop",        "name": "Приостановлено" },
                        "аннулировано":   { "code": "Annul",       "name": "Аннулировано" },
                        "не действует":   { "code": "doesNotWork", "name": "Не действует" },
                        "черновик":       { "code": "Draft",       "name": "Черновик" }
                    }
                    MAP_MARKET_TYPE = {
                        "специализированный": { "code": "Specialized", "_id": "68ed17bb27eea1af1d5473f7", "name": "Специализированный" },
                        "универсальный":      { "code": "Universal",   "_id": "68ed1979b12469d98995f24e", "name": "Универсальный" }
                    }
                    MAP_SPECIALIZATION = {
                        "сельскохозяйственный":                       { "code":"Agricultural", "parentId":"68ed17bb27eea1af1d5473f7", "name":"Сельскохозяйственный" },
                        "сельскохозяйственный кооперативный":        { "code":"AgriculturalCooperative", "parentId":"68ed17bb27eea1af1d5473f7", "name":"Сельскохозяйственный кооперативный" },
                        "вещевой":                                   { "code":"Clothing", "parentId":"68ed17bb27eea1af1d5473f7", "name":"Вещевой" },
                        "по продаже радио- и электробытовой техники": { "code":"For the sale of radio and electrical appliances", "parentId":"68ed17bb27eea1af1d5473f7", "name":"По продаже радио- и электробытовой техники" },
                        "иная":                                      { "code":"Other", "parentId":"68ed17bb27eea1af1d5473f7", "name":"Иная" },
                        "иное":                                      { "code":"Other", "parentId":"68ed17bb27eea1af1d5473f7", "name":"Иная" },
                        "по продаже строительных материалов":         { "code":"SaleBuildingMaterials", "parentId":"68ed17bb27eea1af1d5473f7", "name":"По продаже строительных материалов" },
                        "по продаже продуктов питания":               { "code":"SaleProducts", "parentId":"68ed17bb27eea1af1d5473f7", "name":"По продаже продуктов питания" }
                    }
                    MAP_OPERATION_PERIOD = {
                        "постоянный": { "code":"PermanentType", "name":"Постоянный" },
                        "временный":  { "code":"TemporaryType", "name":"Временный" }
                    }
                    MAP_ORG_STATE_FORM = {
                        "негосударственные пенсионные фонды": { "code": "70402", "name": "Негосударственные пенсионные фонды" },
                        "районные суды, городские суды, межрайонные суды (районные суды)": { "code": "30008", "name": "Районные суды, городские суды, межрайонные суды (районные суды)" },
                        "товарищества собственников недвижимости": { "code": "20700", "name": "Товарищества собственников недвижимости" },
                        "обособленные подразделения юридических лиц": { "code": "30003", "name": "Обособленные подразделения юридических лиц" },
                        "адвокаты, учредившие адвокатский кабинет": { "code": "50201", "name": "Адвокаты, учредившие адвокатский кабинет" },
                        "сельскохозяйственные потребительские животноводческие кооперативы": { "code": "20115", "name": "Сельскохозяйственные потребительские животноводческие кооперативы" },
                        "садоводческие или огороднические некоммерческие товарищества": { "code": "20702", "name": "Садоводческие или огороднические некоммерческие товарищества" },
                        "жилищные или жилищно-строительные кооперативы": { "code": "20102", "name": "Жилищные или жилищно-строительные кооперативы" },
                        "казачьи общества, внесенные в государственный реестр казачьих обществ в российской федерации": { "code": "21100", "name": "Казачьи общества, внесенные в государственный реестр казачьих обществ в Российской Федерации" },
                        "сельскохозяйственные производственные кооперативы": { "code": "14100", "name": "Сельскохозяйственные производственные кооперативы" },
                        "нотариусы, занимающиеся частной практикой": { "code": "50202", "name": "Нотариусы, занимающиеся частной практикой" },
                        "крестьянские (фермерские) хозяйства": { "code": "15300", "name": "Крестьянские (фермерские) хозяйства" },
                        "благотворительные учреждения": { "code": "75502", "name": "Благотворительные учреждения" },
                        "прочие юридические лица, являющиеся коммерческими организациями": { "code": "19000", "name": "Прочие юридические лица, являющиеся коммерческими организациями" },
                        "объединения фермерских хозяйств": { "code": "20613", "name": "Объединения фермерских хозяйств" },
                        "нотариальные палаты": { "code": "20610", "name": "Нотариальные палаты" },
                        "некоммерческие партнерства": { "code": "20614", "name": "Некоммерческие партнерства" },
                        "общественные фонды": { "code": "70403", "name": "Общественные фонды" },
                        "федеральные государственные автономные учреждения": { "code": "75101", "name": "Федеральные государственные автономные учреждения" },
                        "сельскохозяйственные потребительские обслуживающие кооперативы": { "code": "20111", "name": "Сельскохозяйственные потребительские обслуживающие кооперативы" },
                        "общественные организации": { "code": "20200", "name": "Общественные  организации" },
                        "государственные корпорации": { "code": "71601", "name": "Государственные корпорации" },
                        "главы крестьянских (фермерских) хозяйств": { "code": "50101", "name": "Главы крестьянских (фермерских) хозяйств" },
                        "автономные некоммерческие организации": { "code": "71400", "name": "Автономные некоммерческие организации" },
                        "учреждения": { "code": "75000", "name": "Учреждения" },
                        "государственные автономные учреждения субъектов российской федерации": { "code": "75201", "name": "Государственные автономные учреждения субъектов Российской Федерации" },
                        "объединения (ассоциации и союзы) благотворительных организаций": { "code": "20620", "name": "Объединения (ассоциации и союзы) благотворительных организаций" },
                        "публичные акционерные общества": { "code": "12247", "name": "Публичные акционерные общества" },
                        "государственные академии наук": { "code": "75300", "name": "Государственные академии наук" },
                        "государственные компании": { "code": "71602", "name": "Государственные компании" },
                        "представительства юридических лиц": { "code": "30001", "name": "Представительства юридических лиц" },
                        "государственные бюджетные учреждения субъектов российской федерации": { "code": "75203", "name": "Государственные бюджетные учреждения субъектов Российской Федерации" },
                        "союзы (ассоциации) кредитных кооперативов": { "code": "20604", "name": "Союзы (ассоциации) кредитных кооперативов" },
                        "межправительственные международные организации": { "code": "40001", "name": "Межправительственные международные организации" },
                        "муниципальные бюджетные учреждения": { "code": "75403", "name": "Муниципальные бюджетные учреждения" },
                        "государственные унитарные предприятия субъектов российской федерации": { "code": "65242", "name": "Государственные унитарные предприятия субъектов Российской Федерации" },
                        "полные товарищества": { "code": "11051", "name": "Полные товарищества" },
                        "общественные движения": { "code": "20210", "name": "Общественные движения" },
                        "рыболовецкие артели (колхозы)": { "code": "14154", "name": "Рыболовецкие артели (колхозы)" },
                        "потребительские общества": { "code": "20107", "name": "Потребительские общества" },
                        "союзы (ассоциации) общин малочисленных народов": { "code": "20607", "name": "Союзы (ассоциации) общин малочисленных народов" },
                        "сельскохозяйственные потребительские перерабатывающие кооперативы": { "code": "20109", "name": "Сельскохозяйственные потребительские перерабатывающие  кооперативы" },
                        "учреждения, созданные российской федерацией": { "code": "75100", "name": "Учреждения, созданные Российской Федерацией" },
                        "производственные кооперативы (кроме сельскохозяйственных производственных кооперативов)": { "code": "14200", "name": "Производственные кооперативы (кроме сельскохозяйственных производственных кооперативов)" },
                        "учреждения, созданные субъектом российской федерации": { "code": "75200", "name": "Учреждения, созданные субъектом Российской Федерации" },
                        "государственные казенные учреждения субъектов российской федерации": { "code": "75204", "name": "Государственные казенные учреждения субъектов Российской Федерации" },
                        "федеральные государственные унитарные предприятия": { "code": "65241", "name": "Федеральные государственные унитарные предприятия" },
                        "саморегулируемые организации": { "code": "20619", "name": "Саморегулируемые организации" },
                        "территориальные общественные самоуправления": { "code": "20217", "name": "Территориальные общественные самоуправления" },
                        "акционерные общества": { "code": "12200", "name": "Акционерные общества" },
                        "кредитные потребительские кооперативы граждан": { "code": "20105", "name": "Кредитные потребительские  кооперативы граждан" },
                        "казенные предприятия субъектов российской федерации": { "code": "65142", "name": "Казенные предприятия субъектов Российской Федерации" },
                        "советы муниципальных образований субъектов российской федерации": { "code": "20603", "name": "Советы муниципальных образований субъектов Российской Федерации" },
                        "сельскохозяйственные потребительские снабженческие кооперативы": { "code": "20112", "name": "Сельскохозяйственные потребительские снабженческие кооперативы" },
                        "ассоциации (союзы)": { "code": "20600", "name": "Ассоциации (союзы)" },
                        "филиалы юридических лиц": { "code": "30002", "name": "Филиалы юридических лиц" },
                        "муниципальные казенные предприятия": { "code": "65143", "name": "Муниципальные казенные предприятия" },
                        "жилищные накопительные кооперативы": { "code": "20103", "name": "Жилищные накопительные кооперативы" },
                        "органы общественной самодеятельности": { "code": "20211", "name": "Органы общественной самодеятельности" },
                        "религиозные организации": { "code": "71500", "name": "Религиозные организации" },
                        "благотворительные фонды": { "code": "70401", "name": "Благотворительные фонды" },
                        "федеральные государственные казенные учреждения": { "code": "75104", "name": "Федеральные государственные казенные учреждения" },
                        "учреждения, созданные муниципальным образованием (муниципальные учреждения)": { "code": "75400", "name": "Учреждения, созданные муниципальным образованием (муниципальные учреждения)" },
                        "общественные учреждения": { "code": "75505", "name": "Общественные учреждения" },
                        "производственные кооперативы (артели)": { "code": "14000", "name": "Производственные кооперативы (артели)" },
                        "муниципальные автономные учреждения": { "code": "75401", "name": "Муниципальные автономные учреждения" },
                        "хозяйственные общества": { "code": "12000", "name": "Хозяйственные общества" },
                        "адвокатские палаты": { "code": "20609", "name": "Адвокатские палаты" },
                        "общества взаимного страхования": { "code": "20108", "name": "Общества взаимного страхования" },
                        "союзы (ассоциации) общественных объединений": { "code": "20606", "name": "Союзы (ассоциации) общественных объединений" },
                        "общества с ограниченной ответственностью": { "code": "12300", "name": "Общества с ограниченной ответственностью" },
                        "хозяйственные партнерства": { "code": "13000", "name": "Хозяйственные партнерства" },
                        "структурные подразделения обособленных подразделений юридических лиц": { "code": "30004", "name": "Структурные подразделения обособленных подразделений юридических лиц" },
                        "простые товарищества": { "code": "30006", "name": "Простые товарищества" },
                        "коллегии адвокатов": { "code": "20616", "name": "Коллегии адвокатов" },
                        "торгово-промышленные палаты": { "code": "20611", "name": "Торгово-промышленные палаты" },
                        "индивидуальные предприниматели": { "code": "50102", "name": "Индивидуальные предприниматели" },
                        "отделения иностранных некоммерческих неправительственных организаций": { "code": "71610", "name": "Отделения иностранных некоммерческих неправительственных организаций" },
                        "гаражные и гаражно-строительные кооперативы": { "code": "20101", "name": "Гаражные и гаражно-строительные кооперативы" },
                        "частные учреждения": { "code": "75500", "name": "Частные учреждения" },
                        "экологические фонды": { "code": "70404", "name": "Экологические фонды" },
                        "неправительственные международные организации": { "code": "40002", "name": "Неправительственные международные организации" },
                        "союзы потребительских обществ": { "code": "20608", "name": "Союзы потребительских обществ" },
                        "федеральные казенные предприятия": { "code": "65141", "name": "Федеральные казенные предприятия" },
                        "потребительские кооперативы": { "code": "20100", "name": "Потребительские кооперативы" },
                        "фонды проката": { "code": "20121", "name": "Фонды проката" },
                        "публично-правовые компании": { "code": "71600", "name": "Публично-правовые компании" },
                        "фонды": { "code": "70400", "name": "Фонды" },
                        "федеральные государственные бюджетные учреждения": { "code": "75103", "name": "Федеральные государственные бюджетные учреждения" },
                        "товарищества на вере (коммандитные товарищества)": { "code": "11064", "name": "Товарищества на вере (коммандитные товарищества)" },
                        "кредитные потребительские кооперативы": { "code": "20104", "name": "Кредитные потребительские кооперативы" },
                        "товарищества собственников жилья": { "code": "20716", "name": "Товарищества собственников жилья" },
                        "общины коренных малочисленных народов российской федерации": { "code": "21200", "name": "Общины коренных малочисленных народов Российской Федерации" },
                        "хозяйственные товарищества": { "code": "11000", "name": "Хозяйственные товарищества" },
                        "паевые инвестиционные фонды": { "code": "30005", "name": "Паевые инвестиционные фонды" },
                        "политические партии": { "code": "20201", "name": "Политические партии" },
                        "объединения работодателей": { "code": "20612", "name": "Объединения работодателей" },
                        "сельскохозяйственные потребительские сбытовые (торговые) кооперативы": { "code": "20110", "name": "Сельскохозяйственные потребительские сбытовые (торговые) кооперативы" },
                        "непубличные акционерные общества": { "code": "12267", "name": "Непубличные акционерные общества" },
                        "союзы (ассоциации) кооперативов": { "code": "20605", "name": "Союзы (ассоциации) кооперативов" },
                        "муниципальные казенные учреждения": { "code": "75404", "name": "Муниципальные казенные учреждения" },
                        "муниципальные унитарные предприятия": { "code": "65243", "name": "Муниципальные унитарные предприятия" },
                        "адвокатские бюро": { "code": "20615", "name": "Адвокатские бюро" },
                        "сельскохозяйственные артели (колхозы)": { "code": "14153", "name": "Сельскохозяйственные артели (колхозы)" },
                    }
                    recData = copy.deepcopy(RECORDS_TEMPLATES.get(list_name))
                    if list_name == "2. Реестр разрешений":
                        
                        if not TEST and not args.dry_run and session is not None:
                            orgOGRN = row.get("ОГРН уполномоченного органа")
                            if pd.notna(orgOGRN) and orgOGRN.strip():
                                orgSearchParams = {"ogrn": str(orgOGRN.strip())}
                                unitSearch = get_unit(session, orgSearchParams, logger)
                                if unitSearch is not None:
                                    unit = unitSearch.copy()
                                    unit["id"] = unit.pop("_id")
                        recData["guid"] = generate_guid()
                        recData["parentEntries"] = "reestrpermitsReestr"
                        recData["unit"] = unit
                        recData["generalInformation"] = {
                            "Subject": row.get("Субъект РФ"),
                            "Disctrict": row.get("Муниципальный район/округ, городской округ или внутригородская территория")
                        }
                        recData["permission"] = {
                            "PermissionStatus": MAP_PERMISSION_STATUS[row.get("Статус разрешения").lower()],
                            "PermissionNumber": row.get("Номер разрешения"),
                            "PermissionStartDate": row.get("Дата выдачи разрешения"),
                            "administrationName": None,
                            "permissionEffectiveStartDate": row.get("Дата начала действия разрешения"),
                            "PermissionEndDate": row.get("Дата завершения действия разрешения"),
                            "permissionExtensionDate": row.get("Дата, до которой продлено действие разрешения"),
                            "reissuePermissionFile": None
                        }
                    elif list_name == "3. Реестр рынков":
                        if not TEST and not args.dry_run and session is not None:
                            orgOGRN = row.get("ОГРН уполномоченного органа")
                            if pd.notna(orgOGRN) and orgOGRN.strip():
                                orgSearchParams = {"ogrn": str(orgOGRN.strip())}
                                unitSearch = get_unit(session, orgSearchParams, logger)
                                if unitSearch is not None:
                                    unit = unitSearch.copy()
                                    unit["id"] = unit.pop("_id")
                        recData["guid"] = generate_guid()
                        recData["parentEntries"] = "reestrmarketReestr"
                        recData["unit"] = unit
                        recData["TotalInfo"] = {
                            "Subject":   row.get("Субъект РФ"),
                            "Disctrict": row.get("Муниципальный район/округ, городской округ или внутригородская территория")
                        }
                        recData["InfoRetailMarket"] = {
                            "RetailName": row.get("Наименование рынка"),
                            "PermissionStatus": MAP_PERMISSION_STATUS[row.get("Статус разрешения").lower()] if pd.notna(row.get("Статус разрешения")) else None,
                            "PermissionNumber": row.get("Номер разрешения") if pd.notna(row.get("Номер разрешения")) else None,
                            "PermissionStartDate": to_iso_date(row.get("Дата выдачи разрешения")) if pd.notna(row.get("Дата выдачи разрешения")) else None,
                            "PermissionEndDate": to_iso_date(row.get("Дата завершения действия разрешения")) if pd.notna(row.get("Дата завершения действия разрешения")) else None,
                            "marketType": MAP_MARKET_TYPE.get(row.get("Тип рынка").lower(), {}).get("code") if pd.notna(row.get("Тип рынка")) else None,
                            "marketSpecialization":    MAP_SPECIALIZATION.get(row.get("Специализация рынка").lower(), {}).get("code") if pd.notna(row.get("Специализация рынка")) else None,
                            "marketOtherSpecialization": row.get("Другая специализация рынка") if pd.notna(row.get("Другая специализация рынка")) else None,
                            "constituentFiles": None,
                            "GeoCoordinates": row.get("Геокоординаты точки, на которой расположен рынок") if pd.notna(row.get("Геокоординаты точки, на которой расположен рынок")) else None,
                            "marketArea": float(row.get("Площадь рынка, кв. м.")) if pd.notna(row.get("Площадь рынка, кв. м.")) else None,
                            "PlaceNumber": int(row.get("Число торговых мест, шт.")) if pd.notna(row.get("Число торговых мест, шт.")) else None,
                            "operationPeriod": MAP_OPERATION_PERIOD.get(row.get("Период действия разрешения").lower(), {"name": row.get("Период действия разрешения") }) if pd.notna(row.get("Период действия разрешения")) else None,
                            "marketOpeningTime": row.get("Основное время начала работы рынка") if pd.notna(row.get("Основное время начала работы рынка")) else None,
                            "marketClosingTime": row.get("Основное время окончания работы рынка") if pd.notna(row.get("Основное время окончания работы рынка")) else None,
                            "sanitaryDayOfMonth": row.get("Санитарный день месяца") if pd.notna(row.get("Санитарный день месяца")) else None,
                            "blockMarketAddress": [],
                            "blockCadNumber": [],
                            "cadsObjects": [],
                            "dayOnBlock": [],
                            "BlockDayOff": []
                        }
                        paUL = split_postal_address(row.get("Юридический адрес")) if pd.notna(row.get("Юридический адрес")) else {"postalCode": None, "fullAddress": None}
                        paFA = split_postal_address(row.get("Фактический адрес")) if pd.notna(row.get("Фактический адрес")) else {"postalCode": None, "fullAddress": None}
                        recData["InfoCompanyManagerMarket"] = {
                            "OperatorName":   row.get("Наименование оператора") if pd.notna(row.get("Наименование оператора")) else None,
                            "ShortNameUL":    row.get("Краткое наименование ЮЛ") if pd.notna(row.get("Краткое наименование ЮЛ")) else None,
                            "AddressUL":     { "postalCode": paUL["postalCode"], "fullAddress": paUL["fullAddress"] },
                            "AddressActual": { "postalCode": paFA["postalCode"], "fullAddress" : paFA["fullAddress"] },
                            "OperatorINN":    normStr(row.get("ИНН оператора")) if pd.notna(row.get("ИНН оператора")) else None,
                            "OperatorOGRN":   normStr(row.get("ОГРН оператора")) if pd.notna(row.get("ОГРН оператора")) else None,
                            "RykFIO":         normStr(row.get("ФИО руководителя")) if pd.notna(row.get("ФИО руководителя")) else None,
                            "OperatorNumber": normStr(row.get("Контактный номер оператора")) if pd.notna(row.get("Контактный номер оператора")) else None,
                            "OperatorEmail":  normStr(row.get("Адрес электронной почты оператора")) if pd.notna(row.get("Адрес электронной почты оператора")) else None,
                            "OrgStateForm": MAP_ORG_STATE_FORM.get(row.get("Организационно-правовая форма").lower(), {"name": row.get("Организационно-правовая форма") }) if pd.notna(row.get("Организационно-правовая форма")) else None
                        }

                    if TEST or args.dry_run:
                        logger.info(f"TEST MODE: Структура для строки {i} | {json.dumps(recData, ensure_ascii=False)}")
                        if list_name == "3. Реестр рынков":
                            logger.info(
                                f"TEST MODE NSI {list_name}: "
                                f"{json.dumps(build_nsi_local_object_market_payload(row, unit, recData['guid']), ensure_ascii=False)}"
                            )
                        continue

                    recordURL = f"{get_runtime_base_url()}/api/v1/create/{recData['parentEntries']}"
                    recordRes = api_request(session, logger, "post", recordURL, json=jsonable(recData))
                    if not recordRes.ok:
                        logger.error(f"Ошибка при создании записи")
                        failLogger.info(i)
                        failed_rows += 1
                        action = _operator_action_on_row_error(
                            operator_mode=bool(args.operator_mode),
                            interactive=interactive,
                            logger=logger,
                            context=f"workbook={workbook_path.name} sheet={list_name} row={i}: создание записи",
                        )
                        if action == "abort":
                            stopped = True
                            break
                        continue
                    recordResJSON = recordRes.json()
                    record_id = recordResJSON["_id"]
                    record_guid = recordResJSON["guid"]
                    if not TEST and not args.dry_run and session is not None:
                        # Files pathes
                        files_pathes = []
                        fileExeption = False
                        if pd.notna(row.get("Разрешение на организацию розничного рынка")) and isinstance(row.get("Разрешение на организацию розничного рынка"), str) and row.get("Разрешение на организацию розничного рынка") != "":
                            fileIds = row.get("Разрешение на организацию розничного рынка").split(";")
                            for fileId in fileIds:
                                fileId_clean = fileId.replace("\n", "").replace("\r", "")
                                file_path = find_file_in_dir(files_dir_active, fileId_clean)
                                if file_path:
                                    logger.info(f"Найден файл: {os.path.basename(file_path)}")
                                    files_pathes.append(file_path)
                                else:
                                    logger.error(f"Файл не найден по шаблону: {fileId_clean}")
                                    fileExeption = True
                                    break
                        if fileExeption:
                            failLogger.info(i)
                            failed_rows += 1
                            delete_from_collection(session, logger, recordResJSON)
                            logger.info(f"Ошибка при получении файла {i}")
                            action = _operator_action_on_row_error(
                                operator_mode=bool(args.operator_mode),
                                interactive=interactive,
                                logger=logger,
                                context=f"workbook={workbook_path.name} sheet={list_name} row={i}: файл не найден",
                            )
                            if action == "abort":
                                stopped = True
                                break
                            continue
                        # Files upload
                        file_objects = []
                        exception_f_o = False
                        for file_p in files_pathes:
                            logger.info(f"Загрузка файла {file_p}")
                            file_object = upload_file(session, logger, file_p, recData["parentEntries"], record_id, entity_field_path="")
                            if file_object is None:
                                logger.error(f"Ошибка при загрузке файла {file_p}")
                                failLogger.info(i)
                                failed_rows += 1
                                exception_f_o = True
                                break
                            file_objects.append(file_object)
                        if exception_f_o:
                            logger.info(f"Удаление данных по {i}")
                            for file_o in file_objects:
                                delete_file_from_storage(session, logger, file_o._id)
                                delete_from_collection(session, logger, recordResJSON)
                            failLogger.info(i)
                            failed_rows += 1
                            logger.info(f"Завершено удаление данных по {i}")
                            action = _operator_action_on_row_error(
                                operator_mode=bool(args.operator_mode),
                                interactive=interactive,
                                logger=logger,
                                context=f"workbook={workbook_path.name} sheet={list_name} row={i}: ошибка загрузки файла",
                            )
                            if action == "abort":
                                stopped = True
                                break
                            continue
                        if len(file_objects) > 0:
                            updRecData = {
                                "_id": record_id,
                                "guid": record_guid,
                                "parentEntries": recData['parentEntries'],
                                "permission": {
                                    "PermissionStatus": MAP_PERMISSION_STATUS[row.get("Статус разрешения").lower()],
                                    "PermissionNumber": row.get("Номер разрешения"),
                                    "PermissionStartDate": row.get("Дата выдачи разрешения"),
                                    "administrationName": None,
                                    "permissionEffectiveStartDate": row.get("Дата начала действия разрешения"),
                                    "PermissionEndDate": row.get("Дата завершения действия разрешения"),
                                    "permissionExtensionDate": row.get("Дата, до которой продлено действие разрешения"),
                                    "reissuePermissionFile": file_objects[0]
                                }
                            }
                            recordUpdURL = f"{get_runtime_base_url()}/api/v1/update/{recData['parentEntries']}?mainId={record_id}&guid={recordResJSON['guid']}"
                            recUpdRes = api_request(session, logger, "put", recordUpdURL, json=jsonable(updRecData))
                            if recUpdRes.status_code != requests.codes.ok:
                                logger.info(f"Удаление данных по {i}")
                                for file_o in file_objects:
                                    delete_file_from_storage(session, logger, file_o["_id"])
                                delete_from_collection(session, logger, recordResJSON)
                                failLogger.info(i)
                                failed_rows += 1
                                logger.info(f"Завершено удаление данных по {i}")
                                action = _operator_action_on_row_error(
                                    operator_mode=bool(args.operator_mode),
                                    interactive=interactive,
                                    logger=logger,
                                    context=f"workbook={workbook_path.name} sheet={list_name} row={i}: ошибка обновления записи",
                                )
                                if action == "abort":
                                    stopped = True
                                    break
                                continue
                    # Final log
                    logger.info(f"Создана структура для строки {i} | _id записи - {record_id}")
                    log_success(successLogger, {"_id": record_id, "guid": record_guid, "parentEntries": recData["parentEntries"]})

                    if list_name == "3. Реестр рынков":
                        local_payload = build_nsi_local_object_market_payload(row, unit, record_guid)
                        localRecordURL = f"{get_runtime_base_url()}/api/v1/create/{NSI_LOCAL_OBJECT_MARKET_COLLECTION}"
                        localRecordRes = api_request(session, logger, "post", localRecordURL, json=jsonable(local_payload))
                        if not localRecordRes.ok:
                            logger.error("Ошибка при создании записи в nsiLocalObjectMarket")
                            failLogger.info(f"{list_name}:{i}:nsi")
                            failed_rows += 1
                            action = _operator_action_on_row_error(
                                operator_mode=bool(args.operator_mode),
                                interactive=interactive,
                                logger=logger,
                                context=f"workbook={workbook_path.name} sheet={list_name} row={i}: nsiLocalObjectMarket",
                            )
                            if action == "abort":
                                stopped = True
                                break
                            continue
                        localRecordJSON = localRecordRes.json()
                        logger.info(
                            f"Создана запись в {NSI_LOCAL_OBJECT_MARKET_COLLECTION} | "
                            f"_id записи - {localRecordJSON.get('_id')}"
                        )
                        log_success(
                            successLogger,
                            {
                                "_id": localRecordJSON.get("_id"),
                                "guid": localRecordJSON.get("guid", local_payload["guid"]),
                                "parentEntries": NSI_LOCAL_OBJECT_MARKET_COLLECTION
                            }
                        )
                    state.mark_success(
                        workbook_path=str(workbook_path),
                        job_name=list_name,
                        row_idx=i,
                        collection=str(recData["parentEntries"]),
                        main_id=str(record_id),
                        guid=str(record_guid),
                        had_errors=False,
                        error_count=0,
                    )
                    created_rows += 1
                    if stopped:
                        logger.warning("Migration stopped on workbook=%s sheet=%s row=%s", workbook_path.name, list_name, i)
                        break
            logger.info(f"Завершена обработка листа: {list_name}")
            if stopped:
                break
            logger.info("Обработка файла завершена")
            if stopped:
                break

    except KeyboardInterrupt:
        stopped = True
        logger.warning("Остановка по Ctrl+C")
    except Exception as e:
        fatal_error = True
        logger.error("Критическая ошибка выполнения: %s", e)
        logger.debug(traceback.format_exc())
    finally:
        summary = {
            "processedRows": processed_rows,
            "createdRows": created_rows,
            "resumeSkips": resumed_skips,
            "failedRows": failed_rows,
            "workbooks": len(workbook_specs),
            "profile": profile_name,
            "baseUrl": get_runtime_base_url(),
            "mode": resolved_mode,
            "dryRun": bool(args.dry_run),
            "skipAuth": bool(skip_auth),
        }
        if stopped:
            status = "stopped"
            state.finish_run(status="stopped", summary=summary, clear_rows=False)
        elif fatal_error or failed_rows > 0:
            status = "failed"
            state.finish_run(status="failed", summary=summary, clear_rows=False)
        else:
            status = "completed"
            state.finish_run(status="completed", summary=summary, clear_rows=True)
        userLogger.info("FINISH | status=%s | summary=%s", status, json.dumps(summary, ensure_ascii=False))
        logger.info(
            "Итог: status=%s created=%s failed=%s resumed_skips=%s",
            status,
            created_rows,
            failed_rows,
            resumed_skips,
        )

    return 0 if not stopped and not fatal_error and failed_rows == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())

