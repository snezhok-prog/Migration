"""
Скрипт отката ищет в logs все success*.txt файлы
Удаляет построчно в каждом из них запись вида 
{ "_id": ..., "guid": ..., "parentEntries": ..., ... }
"""

import json
import glob
import os
import warnings

from urllib3.exceptions import InsecureRequestWarning
from _api import (
    setup_session,
    delete_from_collection
)
from _logger import setup_rollback_logger
from _config import SUCCESS_LOG_PATTERN

def main():
    warnings.filterwarnings("ignore", category=InsecureRequestWarning)
    logger = setup_rollback_logger()
    logger.info("Запуск скрипта отката")

    session = setup_session(logger)
    if not session:
        logger.error("Не удалось авторизоваться. Завершение.")
        return
    
    # Находим все success-логи
    log_files = glob.glob(SUCCESS_LOG_PATTERN)
    if not log_files:
        logger.warning(f"Не найдено файлов по шаблону: {SUCCESS_LOG_PATTERN}")
        return

    logger.info(f"Найдено файлов для обработки: {len(log_files)}")
    total_deleted = 0
    total_failed = 0

    for log_file in log_files:
        logger.info(f"Обработка файла: {os.path.basename(log_file)}")
        with open(log_file, "r", encoding="utf-8") as f:
            for line_num, line in enumerate(f, 1):
                line = line.strip()
                if not line:
                    continue
                try:
                    data = json.loads(line)
                    if isinstance(data, dict) and "_id" in data and "guid" in data and "parentEntries" in data:
                        success = delete_from_collection(session, logger, data)
                        if success:
                            total_deleted += 1
                        else:
                            total_failed += 1
                    else:
                        logger.warning(f"Некорректный формат в {log_file}:{line_num} — {line[:100]}...")
                except json.JSONDecodeError as e:
                    logger.error(f"Ошибка JSON в {log_file}:{line_num} — {e}")
                    total_failed += 1

    logger.info(f"Обработка завершена. Удалено: {total_deleted}, Ошибок: {total_failed}")

    
if __name__ == "__main__":
    main()