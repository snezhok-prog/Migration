from __future__ import annotations

import argparse
import glob
import json
import warnings

from urllib3.exceptions import InsecureRequestWarning

from _api import delete_from_collection, set_runtime_urls, setup_session
from _config import BASE_URL, JWT_URL, SUCCESS_LOG_PATTERN, UI_BASE_URL
from _logger import setup_rollback_logger
from _profiles import PROFILES


def _resolve_runtime(args):
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
    return base_url, jwt_url, ui_base_url


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Rollback по success логам для миграции ярмарок")
    parser.add_argument("--profile", choices=["custom", "dev", "psi", "prod"], default="dev")
    parser.add_argument("--base-url", default="")
    parser.add_argument("--jwt-url", default="")
    parser.add_argument("--ui-base-url", default="")
    parser.add_argument("--success-log-glob", default=SUCCESS_LOG_PATTERN, help="Glob-шаблон success логов")
    parser.add_argument("--no-prompt", action="store_true", help="Не спрашивать cookie/token, читать из token.md/cookie.md")
    return parser.parse_args()


def main() -> int:
    warnings.filterwarnings("ignore", category=InsecureRequestWarning)
    args = parse_args()
    logger = setup_rollback_logger()

    try:
        base_url, jwt_url, ui_base_url = _resolve_runtime(args)
    except Exception as exc:
        logger.error("Ошибка runtime конфигурации: %s", exc)
        return 1

    set_runtime_urls(base_url=base_url, jwt_url=jwt_url, ui_base_url=ui_base_url)
    logger.info("Rollback profile=%s base_url=%s", args.profile, base_url)

    try:
        session = setup_session(logger, no_prompt=bool(args.no_prompt))
    except Exception as exc:
        logger.error("Авторизация не выполнена: %s", exc)
        return 1
    if session is None:
        logger.error("Авторизация не выполнена: setup_session вернул None")
        return 1

    log_files = glob.glob(args.success_log_glob)
    if not log_files:
        logger.warning("Не найдены success-логи по шаблону: %s", args.success_log_glob)
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
