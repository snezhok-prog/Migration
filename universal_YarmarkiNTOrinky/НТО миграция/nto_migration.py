from __future__ import annotations

import argparse
import copy
import json
import os
import traceback
import warnings
from pathlib import Path
from typing import Any, Dict, List, Sequence, Tuple

from openpyxl import load_workbook
from urllib3.exceptions import InsecureRequestWarning

from _api import build_file_meta, create_record, setup_session, update_record, upload_file
from _config import (
    EXCEL_FILE_NAME,
    FILES_DIR,
    NTO_MESTO_COLLECTION,
    NTO_NSI_COLLECTION,
    RESUME_BY_DEFAULT,
    ROLLBACK_BODY_FILE,
    SCRIPT_DIR,
    SHEET_MESTO,
    SHEET_TORGI,
    STATE_FILE,
    TORGI_COLLECTION,
)
from _logger import setup_fail_logger, setup_logger, setup_success_logger, setup_user_logger
from _state import ResumeState
from _utils import dump_json, resolve_local_file_path, set_by_path
from _vba_nto import (
    PendingUpload,
    build_nsi_local_object_nto_payload,
    iter_vba_rows,
    norm_str,
    transform_row_to_bidding,
    transform_row_to_registry,
)


def _build_file_placeholder(*, filename: str, size: int, entity_field_path: str, allow_external: bool) -> Dict[str, Any]:
    return {
        "_id": "",
        "originalName": filename,
        "size": int(size or 0),
        "isFile": True,
        "entityFieldPath": entity_field_path,
        "allowExternal": bool(allow_external),
    }


def _rollback_payload(rollback_items: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    return [
        {
            "_id": item.get("_id"),
            "guid": item.get("guid"),
            "parentEntries": item.get("parentEntries"),
        }
        for item in rollback_items
    ]


def _append_unique_rollback(target: List[Dict[str, Any]], entry: Dict[str, Any]) -> None:
    key = (str(entry.get("parentEntries")), str(entry.get("_id")))
    for existing in target:
        if (str(existing.get("parentEntries")), str(existing.get("_id"))) == key:
            return
    target.append(entry)


def _write_rollback_body(rollback_items: Sequence[Dict[str, Any]]) -> None:
    dump_json(Path(ROLLBACK_BODY_FILE), _rollback_payload(list(rollback_items)))


def _resolve_uploads(
    uploads: Sequence[PendingUpload],
    *,
    files_dir: Path,
    base_dir: Path,
) -> List[Dict[str, Any]]:
    resolved: List[Dict[str, Any]] = []
    for task in uploads:
        local_path = resolve_local_file_path(task.filename, files_dir=files_dir, base_dir=base_dir)
        if local_path:
            filename = os.path.basename(local_path)
            size = os.path.getsize(local_path)
        elif task.allow_external:
            filename = os.path.basename(str(task.filename or "")) or str(task.filename or "")
            size = 0
        else:
            raise FileNotFoundError(f"Файл для загрузки не найден: {task.filename}")
        resolved.append(
            {
                "task": task,
                "local_path": local_path,
                "filename": filename,
                "size": size,
            }
        )
    return resolved


def _apply_pending_uploads(
    *,
    session,
    logger,
    collection: str,
    main_id: str,
    guid: str,
    doc_state: Dict[str, Any],
    uploads: Sequence[PendingUpload],
    files_dir: Path,
    base_dir: Path,
    continue_on_upload_error: bool = True,
) -> Dict[str, Any]:
    if not uploads:
        return doc_state

    resolved = _resolve_uploads(uploads, files_dir=files_dir, base_dir=base_dir)

    for item in resolved:
        task = item["task"]
        set_by_path(
            doc_state,
            task.target_path,
            _build_file_placeholder(
                filename=item["filename"],
                size=item["size"],
                entity_field_path=task.target_path,
                allow_external=task.allow_external,
            ),
        )

    update_record(session, logger, collection=collection, main_id=main_id, guid=guid, body=doc_state)

    for item in resolved:
        task = item["task"]
        local_path = item["local_path"]
        if not local_path:
            logger.info(
                "[UPLOAD][SKIP-EXTERNAL] collection=%s id=%s field=%s file=%s",
                collection,
                main_id,
                task.target_path,
                task.filename,
            )
            continue
        try:
            uploaded = upload_file(
                session,
                logger,
                entry_name=collection,
                entry_id=main_id,
                entity_field_path=task.target_path,
                file_path=local_path,
                allow_external=task.allow_external,
            )
            meta = build_file_meta(
                uploaded,
                fallback_name=item["filename"],
                fallback_size=item["size"],
                path=task.target_path,
                allow_external=task.allow_external,
            )
            set_by_path(doc_state, task.target_path, meta)
        except Exception:
            logger.exception(
                "[UPLOAD][FAIL] collection=%s id=%s field=%s file=%s",
                collection,
                main_id,
                task.target_path,
                local_path,
            )
            if not continue_on_upload_error:
                raise

    updated = update_record(session, logger, collection=collection, main_id=main_id, guid=guid, body=doc_state)
    if isinstance(updated, dict):
        return updated
    return doc_state


def _process_mesto_row(
    *,
    row_idx: int,
    row_data: Dict[str, Any],
    session,
    logger,
    success_logger,
    rollback_items: List[Dict[str, Any]],
    files_dir: Path,
    base_dir: Path,
    dry_run: bool,
) -> Dict[str, Any]:
    payload, pending_uploads = transform_row_to_registry(row_data, session, logger)
    if dry_run:
        logger.info("[DRY][%s][ROW:%s] payload keys=%s uploads=%s", SHEET_MESTO, row_idx, sorted(payload.keys()), len(pending_uploads))
        return {"collection": NTO_MESTO_COLLECTION, "_id": "", "guid": str(payload.get("guid") or "")}

    created = create_record(session, logger, NTO_MESTO_COLLECTION, payload)
    main_id = str(created.get("_id") or "")
    final_guid = str(created.get("guid") or payload.get("guid") or "")
    if not main_id or not final_guid:
        raise RuntimeError("Нет _id/guid у созданной записи NTOmesto")

    payload["guid"] = final_guid
    gos_link = f"https://www.gosuslugi.ru/trade/{final_guid}"
    payload.setdefault("ntoInformation", {})
    payload["ntoInformation"]["GosuslugiData"] = gos_link

    doc_state = {"_id": main_id, "guid": final_guid, **copy.deepcopy(payload)}
    doc_state["guid"] = final_guid
    doc_state.setdefault("ntoInformation", {})
    doc_state["ntoInformation"]["GosuslugiData"] = gos_link

    project_status_name = norm_str(((payload.get("ntoInformation") or {}).get("ProjectStatus") or {}).get("name")).lower()
    if project_status_name == "утверждено":
        try:
            nsi_payload = build_nsi_local_object_nto_payload({**payload, "guid": final_guid})
            created_nsi = create_record(session, logger, NTO_NSI_COLLECTION, nsi_payload)
            nsi_entry = {
                "sheet": SHEET_MESTO,
                "row": row_idx,
                "_id": created_nsi.get("_id"),
                "guid": created_nsi.get("guid") or nsi_payload.get("guid"),
                "parentEntries": NTO_NSI_COLLECTION,
            }
            success_logger.info(json.dumps(nsi_entry, ensure_ascii=False))
            _append_unique_rollback(rollback_items, nsi_entry)
            _write_rollback_body(rollback_items)
        except Exception:
            logger.exception("[NSI][FAIL][ROW:%s]", row_idx)

    if any(task.target_path == "blockDopSogl[0].SupplementalAgreementIfAny" for task in pending_uploads):
        if not isinstance(doc_state.get("blockDopSogl"), list):
            doc_state["blockDopSogl"] = []
        while len(doc_state["blockDopSogl"]) == 0:
            doc_state["blockDopSogl"].append({})

    if pending_uploads:
        _apply_pending_uploads(
            session=session,
            logger=logger,
            collection=NTO_MESTO_COLLECTION,
            main_id=main_id,
            guid=final_guid,
            doc_state=doc_state,
            uploads=pending_uploads,
            files_dir=files_dir,
            base_dir=base_dir,
            continue_on_upload_error=True,
        )

    entry = {
        "sheet": SHEET_MESTO,
        "row": row_idx,
        "_id": main_id,
        "guid": final_guid,
        "parentEntries": NTO_MESTO_COLLECTION,
        "uploads": len(pending_uploads),
    }
    success_logger.info(json.dumps(entry, ensure_ascii=False))
    _append_unique_rollback(rollback_items, entry)
    _write_rollback_body(rollback_items)
    return {"collection": NTO_MESTO_COLLECTION, "_id": main_id, "guid": final_guid}


def _process_torgi_row(
    *,
    row_idx: int,
    row_data: Dict[str, Any],
    session,
    logger,
    success_logger,
    rollback_items: List[Dict[str, Any]],
    files_dir: Path,
    base_dir: Path,
    dry_run: bool,
) -> Dict[str, Any]:
    payload, pending_uploads = transform_row_to_bidding(row_data, session, logger)
    if dry_run:
        logger.info("[DRY][%s][ROW:%s] payload keys=%s uploads=%s", SHEET_TORGI, row_idx, sorted(payload.keys()), len(pending_uploads))
        return {"collection": TORGI_COLLECTION, "_id": "", "guid": str(payload.get("guid") or "")}

    created = create_record(session, logger, TORGI_COLLECTION, payload)
    main_id = str(created.get("_id") or "")
    guid = str(created.get("guid") or payload.get("guid") or "")
    if not main_id or not guid:
        raise RuntimeError("Нет _id/guid у созданной записи reestrbiddingReestr")

    auid = created.get("auid")
    payload["guid"] = guid
    doc_state = {"_id": main_id, "guid": guid, **({"auid": auid} if auid else {}), **copy.deepcopy(payload)}

    if pending_uploads:
        _apply_pending_uploads(
            session=session,
            logger=logger,
            collection=TORGI_COLLECTION,
            main_id=main_id,
            guid=guid,
            doc_state=doc_state,
            uploads=pending_uploads,
            files_dir=files_dir,
            base_dir=base_dir,
            continue_on_upload_error=True,
        )

    entry = {
        "sheet": SHEET_TORGI,
        "row": row_idx,
        "_id": main_id,
        "guid": guid,
        "parentEntries": TORGI_COLLECTION,
        "uploads": len(pending_uploads),
    }
    success_logger.info(json.dumps(entry, ensure_ascii=False))
    _append_unique_rollback(rollback_items, entry)
    _write_rollback_body(rollback_items)
    return {"collection": TORGI_COLLECTION, "_id": main_id, "guid": guid}


def _iter_selected_sheets(selected: str) -> List[Tuple[str, str]]:
    mapping = []
    if selected in {"all", "mesto"}:
        mapping.append((SHEET_MESTO, "mesto"))
    if selected in {"all", "torgi"}:
        mapping.append((SHEET_TORGI, "torgi"))
    return mapping


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


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Миграция НТО из Excel напрямую в API")
    parser.add_argument("--workbook", default=str(Path(SCRIPT_DIR) / EXCEL_FILE_NAME), help="Путь к книге Excel")
    parser.add_argument("--files-dir", default=FILES_DIR, help="Папка с файлами для загрузки")
    parser.add_argument("--sheet", choices=["all", "mesto", "torgi"], default="all", help="Какие листы запускать")
    parser.add_argument("--limit", type=int, default=0, help="Ограничение по числу строк на лист")
    parser.add_argument("--dry-run", action="store_true", help="Только собрать payload без записи в API")
    parser.add_argument("--no-auth", action="store_true", help="Не логиниться в API")
    parser.add_argument("--operator-mode", action="store_true", help="На ошибке строки: retry/skip/abort")
    parser.add_argument("--no-prompt", action="store_true", help="Не запрашивать input, использовать значения из файлов")
    parser.add_argument("--no-interactive", action="store_true", help="Отключить интерактивный режим")
    parser.add_argument("--state-file", default=str(STATE_FILE), help="Путь к checkpoints JSON")
    parser.add_argument("--reset-state", action="store_true", help="Очистить checkpoints перед запуском")
    parser.add_argument("--resume", dest="resume", action="store_true", default=RESUME_BY_DEFAULT, help="Продолжать с checkpoints")
    parser.add_argument("--no-resume", dest="resume", action="store_false", help="Игнорировать checkpoints")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    warnings.filterwarnings("ignore", category=InsecureRequestWarning)

    logger = setup_logger()
    success_logger = setup_success_logger()
    fail_logger = setup_fail_logger()
    user_logger = setup_user_logger()

    interactive = not args.no_interactive and not args.no_prompt

    workbook_default = str(Path(args.workbook).resolve())
    files_default = str(Path(args.files_dir).resolve())

    workbook_path = Path(_prompt_with_default("Excel workbook", workbook_default, interactive)).expanduser().resolve()
    files_dir = Path(_prompt_with_default("Files directory", files_default, interactive)).expanduser().resolve()
    base_dir = Path(SCRIPT_DIR).resolve()

    logger.info("=" * 92)
    logger.info("NTO migration start")
    logger.info("Workbook: %s", workbook_path)
    logger.info("Files dir: %s", files_dir)
    logger.info("Sheet mode: %s", args.sheet)
    logger.info("Dry run: %s", args.dry_run)
    logger.info("Resume: %s", args.resume)
    logger.info("State file: %s", args.state_file)
    logger.info("Interactive: %s", interactive)
    logger.info("Operator mode: %s", args.operator_mode)
    logger.info("=" * 92)

    user_logger.info("START | workbook=%s | files=%s | sheet=%s | dry_run=%s | resume=%s", workbook_path, files_dir, args.sheet, args.dry_run, args.resume)

    if not workbook_path.exists():
        logger.error("Workbook not found: %s", workbook_path)
        return 1

    state_path = Path(args.state_file).expanduser().resolve()
    state = ResumeState(path=state_path, namespace="nto_migration", enabled=True)
    if args.reset_state:
        state.reset_namespace()
        logger.info("State reset: %s", state_path)

    resume_enabled = bool(args.resume)
    existing_rows = state.rows_count()
    run_info = state.get_run_info()

    if resume_enabled and existing_rows > 0:
        status = str(run_info.get("status") or "unknown")
        started_at = run_info.get("startedAt") or "-"
        logger.info("Found previous checkpoints: rows=%s status=%s startedAt=%s", existing_rows, status, started_at)
        if interactive:
            should_continue = _ask_yes_no("Продолжить с предыдущих checkpoints?", default_yes=True)
            if not should_continue:
                state.clear_rows()
                resume_enabled = False
                logger.info("Checkpoints cleared by operator")

    state.begin_run(
        workbook=str(workbook_path),
        filesDir=str(files_dir),
        sheet=args.sheet,
        dryRun=bool(args.dry_run),
        resume=resume_enabled,
        operatorMode=bool(args.operator_mode),
    )

    session = None
    if not args.dry_run and not args.no_auth:
        session = setup_session(logger, no_prompt=(args.no_prompt or args.no_interactive))
        if session is None:
            state.finish_run(status="failed", summary={"reason": "auth_failed"}, clear_rows=False)
            return 1

    workbook = load_workbook(workbook_path, data_only=True)
    rollback_items: List[Dict[str, Any]] = []

    processed_rows = 0
    created_rows = 0
    resumed_skips = 0
    failed_rows = 0
    stopped = False

    try:
        for sheet_name, kind in _iter_selected_sheets(args.sheet):
            if sheet_name not in workbook.sheetnames:
                raise ValueError(f"Лист не найден: {sheet_name}")
            ws = workbook[sheet_name]
            logger.info("Обработка листа: %s", sheet_name)

            for row_idx, row_data in iter_vba_rows(ws, row_limit=max(0, int(args.limit or 0))):
                row_key = state.get(str(workbook_path), sheet_name, row_idx)
                if resume_enabled and isinstance(row_key, dict):
                    resumed_skips += 1
                    logger.info("[RESUME][SKIP] sheet=%s row=%s _id=%s", sheet_name, row_idx, row_key.get("_id"))
                    continue

                processed_rows += 1
                while True:
                    try:
                        if kind == "mesto":
                            result = _process_mesto_row(
                                row_idx=row_idx,
                                row_data=row_data,
                                session=session,
                                logger=logger,
                                success_logger=success_logger,
                                rollback_items=rollback_items,
                                files_dir=files_dir,
                                base_dir=base_dir,
                                dry_run=args.dry_run,
                            )
                        else:
                            result = _process_torgi_row(
                                row_idx=row_idx,
                                row_data=row_data,
                                session=session,
                                logger=logger,
                                success_logger=success_logger,
                                rollback_items=rollback_items,
                                files_dir=files_dir,
                                base_dir=base_dir,
                                dry_run=args.dry_run,
                            )

                        if not args.dry_run and isinstance(result, dict) and result.get("_id"):
                            created_rows += 1
                            state.mark_success(
                                workbook_path=str(workbook_path),
                                job_name=sheet_name,
                                row_idx=row_idx,
                                collection=str(result.get("collection") or ""),
                                main_id=str(result.get("_id") or ""),
                                guid=str(result.get("guid") or ""),
                                had_errors=False,
                                error_count=0,
                            )
                        break
                    except Exception as exc:
                        failed_rows += 1
                        error_payload = {
                            "sheet": sheet_name,
                            "row": row_idx,
                            "error": str(exc),
                            "traceback": traceback.format_exc(limit=30),
                        }
                        fail_logger.info(json.dumps(error_payload, ensure_ascii=False))
                        logger.exception("[FAIL][%s][ROW:%s]", sheet_name, row_idx)
                        user_logger.info("ROW_ERROR | sheet=%s | row=%s | error=%s", sheet_name, row_idx, str(exc))

                        if args.operator_mode and interactive:
                            action = _ask_error_action()
                            if action == "retry":
                                continue
                            if action == "skip":
                                logger.warning("[ROW][SKIP] sheet=%s row=%s", sheet_name, row_idx)
                                break
                            stopped = True
                            break

                        stopped = True
                        break

                if stopped:
                    logger.warning("Migration stopped on sheet=%s row=%s", sheet_name, row_idx)
                    break

            if stopped:
                break

    except KeyboardInterrupt:
        stopped = True
        logger.warning("Остановка по Ctrl+C")
    except Exception as exc:
        stopped = True
        logger.error("Критическая ошибка выполнения: %s", exc)
        logger.debug(traceback.format_exc())
    finally:
        summary = {
            "processedRows": processed_rows,
            "createdRows": created_rows,
            "resumeSkips": resumed_skips,
            "failedRows": failed_rows,
            "rollbackItems": len(rollback_items),
        }
        if stopped:
            state.finish_run(status="stopped", summary=summary, clear_rows=False)
        elif failed_rows > 0:
            state.finish_run(status="failed", summary=summary, clear_rows=False)
        else:
            state.finish_run(status="completed", summary=summary, clear_rows=True)

        user_logger.info("FINISH | status=%s | summary=%s", "stopped" if stopped else ("failed" if failed_rows > 0 else "completed"), json.dumps(summary, ensure_ascii=False))

    if rollback_items:
        _write_rollback_body(rollback_items)

    logger.info("Миграция завершена. Создано: %s, пропущено по resume: %s, ошибок: %s", created_rows, resumed_skips, failed_rows)

    if stopped:
        return 1
    if failed_rows > 0:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
