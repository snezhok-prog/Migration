import argparse
import glob
import json
import os
import warnings

from urllib3.exceptions import InsecureRequestWarning

from _api import delete_from_collection, setup_session
from _config import ROLLBACK_BODY_PATH, SUCCESS_LOG_PATTERN
from _logger import setup_rollback_logger


def _is_valid_rollback_item(data):
    return isinstance(data, dict) and all(k in data for k in ("_id", "guid", "parentEntries"))


def _matches_filter(data, collections_filter):
    if not collections_filter:
        return True
    return str(data.get("parentEntries")) in collections_filter


def iter_log_records(logger, collections_filter=None):
    log_files = glob.glob(SUCCESS_LOG_PATTERN)
    if not log_files:
        logger.warning("Не найдено success-логов по шаблону: %s", SUCCESS_LOG_PATTERN)
        return

    logger.info("Найдено success-логов: %s", len(log_files))
    for log_file in log_files:
        logger.info("Чтение success-лога: %s", os.path.basename(log_file))
        with open(log_file, "r", encoding="utf-8") as f:
            for line_num, line in enumerate(f, 1):
                line = line.strip()
                if not line:
                    continue
                try:
                    data = json.loads(line)
                except json.JSONDecodeError as exc:
                    logger.error("Ошибка JSON в %s:%s -> %s", log_file, line_num, exc)
                    continue
                if not _is_valid_rollback_item(data):
                    logger.warning("Некорректная строка в %s:%s", log_file, line_num)
                    continue
                if _matches_filter(data, collections_filter):
                    yield data


def iter_rollback_body(logger, collections_filter=None):
    if not os.path.exists(ROLLBACK_BODY_PATH):
        return
    try:
        raw = open(ROLLBACK_BODY_PATH, "r", encoding="utf-8").read().strip()
    except Exception as exc:
        logger.error("Не удалось прочитать %s: %s", ROLLBACK_BODY_PATH, exc)
        return
    if not raw:
        return
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        logger.error("Некорректный JSON в %s: %s", ROLLBACK_BODY_PATH, exc)
        return
    if not isinstance(data, list):
        logger.error("Ожидался массив в %s", ROLLBACK_BODY_PATH)
        return
    logger.info("Найден ROLLBACK_BODY: %s записей", len(data))
    for item in data:
        if _is_valid_rollback_item(item) and _matches_filter(item, collections_filter):
            yield item


def run_rollback(collections_filter=None):
    warnings.filterwarnings("ignore", category=InsecureRequestWarning)
    logger = setup_rollback_logger()
    logger.info("Запуск rollback-скрипта")
    if collections_filter:
        logger.info("Фильтр коллекций: %s", ", ".join(sorted(collections_filter)))

    session = setup_session(logger)
    if not session:
        logger.error("Не удалось авторизоваться. Завершение.")
        return 1

    total_deleted = 0
    total_failed = 0
    seen = set()

    for data in list(iter_rollback_body(logger, collections_filter)) + list(iter_log_records(logger, collections_filter)):
        key = (str(data.get("_id")), str(data.get("guid")), str(data.get("parentEntries")))
        if key in seen:
            continue
        seen.add(key)

        success = delete_from_collection(session, logger, data)
        if success:
            total_deleted += 1
        else:
            total_failed += 1

    logger.info("Rollback завершен. Удалено: %s, Ошибок: %s", total_deleted, total_failed)
    return 0 if total_failed == 0 else 1


def _parse_collections_arg(raw):
    if not raw:
        return None
    parts = [x.strip() for x in str(raw).split(",")]
    items = [x for x in parts if x]
    return set(items) if items else None


def main():
    parser = argparse.ArgumentParser(description="Rollback created records from ROLLBACK_BODY and success logs")
    parser.add_argument(
        "--collections",
        default="",
        help="Comma-separated collection names to rollback, e.g. animalsRecordsCollectionTwo,animalCatchActRegistryCollection",
    )
    args = parser.parse_args()
    return run_rollback(collections_filter=_parse_collections_arg(args.collections))


if __name__ == "__main__":
    raise SystemExit(main())
