import logging
import os
from datetime import datetime


def _ensure_logs_dir():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    logs_dir = os.path.join(script_dir, "logs")
    os.makedirs(logs_dir, exist_ok=True)
    return logs_dir


def _reset_logger_handlers(logger):
    for handler in list(logger.handlers):
        logger.removeHandler(handler)
        handler.close()


def setup_logger():
    logs_dir = _ensure_logs_dir()
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_path = os.path.join(logs_dir, f"script_creation_log-{timestamp}.txt")

    logger = logging.getLogger("migration_main")
    logger.setLevel(logging.INFO)
    logger.propagate = False
    _reset_logger_handlers(logger)

    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

    fh = logging.FileHandler(log_path, mode="a", encoding="utf-8")
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    sh = logging.StreamHandler()
    sh.setFormatter(formatter)
    logger.addHandler(sh)

    logger.info("Запуск скрипта. Лог: %s", log_path)
    return logger


def setup_success_logger():
    logs_dir = _ensure_logs_dir()
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_path = os.path.join(logs_dir, f"success_log-{timestamp}.txt")

    logger = logging.getLogger("migration_success")
    logger.setLevel(logging.INFO)
    logger.propagate = False
    _reset_logger_handlers(logger)

    fh = logging.FileHandler(log_path, mode="a", encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(message)s"))
    logger.addHandler(fh)
    return logger


def setup_fail_logger():
    logs_dir = _ensure_logs_dir()
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_path = os.path.join(logs_dir, f"fail_log-{timestamp}.txt")

    logger = logging.getLogger("migration_fail")
    logger.setLevel(logging.INFO)
    logger.propagate = False
    _reset_logger_handlers(logger)

    fh = logging.FileHandler(log_path, mode="a", encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(message)s"))
    logger.addHandler(fh)
    return logger


def setup_rollback_logger():
    logs_dir = _ensure_logs_dir()
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_path = os.path.join(logs_dir, f"script_rollback_log-{timestamp}.txt")

    logger = logging.getLogger("migration_rollback")
    logger.setLevel(logging.INFO)
    logger.propagate = False
    _reset_logger_handlers(logger)

    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

    fh = logging.FileHandler(log_path, mode="a", encoding="utf-8")
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    sh = logging.StreamHandler()
    sh.setFormatter(formatter)
    logger.addHandler(sh)
    return logger

