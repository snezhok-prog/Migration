from __future__ import annotations

import glob
import json
import warnings

from urllib3.exceptions import InsecureRequestWarning

from _api import delete_from_collection, setup_session
from _config import SUCCESS_LOG_PATTERN
from _logger import setup_rollback_logger


def main() -> int:
    warnings.filterwarnings("ignore", category=InsecureRequestWarning)
    logger = setup_rollback_logger()
    session = setup_session(logger)

    log_files = glob.glob(SUCCESS_LOG_PATTERN)
    if not log_files:
        logger.warning("Не найдены success-логи по шаблону: %s", SUCCESS_LOG_PATTERN)
        return 0

    deleted = 0
    failed = 0
    for log_file in log_files:
        logger.info("Обработка success-лога: %s", log_file)
        with open(log_file, "r", encoding="utf-8") as fh:
            for line_no, line in enumerate(fh, start=1):
                raw = line.strip()
                if not raw:
                    continue
                try:
                    data = json.loads(raw)
                except json.JSONDecodeError as exc:
                    logger.error("Ошибка JSON в %s:%s: %s", log_file, line_no, exc)
                    failed += 1
                    continue
                if not isinstance(data, dict) or "_id" not in data or "guid" not in data or "parentEntries" not in data:
                    logger.warning("Пропуск строки %s:%s, нет обязательных полей", log_file, line_no)
                    continue
                if delete_from_collection(session, logger, data):
                    deleted += 1
                else:
                    failed += 1

    logger.info("Rollback завершён. Удалено: %s, ошибок: %s", deleted, failed)
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
