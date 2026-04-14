# Integration Guide: Рынки миграция

Документ фиксирует целевую архитектуру скрипта `migration.py` после унификации с подходом `universal_myPet`.

## 1. Назначение

Проект переносит данные из XLSM в два реестра:
- `2. Реестр разрешений`
- `3. Реестр рынков`

Дополнительно для строк листа рынков создается запись в `nsiLocalObjectMarket`.

## 2. Ключевые файлы

- `migration.py` — основной раннер миграции
- `_api.py` — API-слой, авторизация, runtime URL профилей
- `_config.py` — базовые константы и шаблоны
- `_profiles.py` — профили стендов (`dev|psi|prod|custom`)
- `_excel_input.py` — выбор книг в режимах `single/mass`
- `_state.py` — checkpoints/resume
- `_logger.py` — script/success/fail/user/rollback логирование
- `_utils.py` — утилиты парсинга/форматирования/файлов
- `rollback.py` — откат по success-логам

## 3. Поток выполнения

1. Разбор CLI-флагов (`parse_args`).
2. Разрешение стенда (`_resolve_runtime`) и установка runtime URL (`set_runtime_urls`).
3. Выбор входных книг/папок файлов (`_resolve_workbook_specs`).
4. Инициализация checkpoints (`ResumeState`) и стратегия resume (`_choose_resume_strategy`).
5. Авторизация (`setup_session`) либо `--skip-auth` для `--dry-run`.
6. Построчная обработка листов:
   - формирование `recData` (без изменения маппинга);
   - создание записи в целевой коллекции;
   - загрузка файлов (если есть);
   - обновление записи после загрузки;
   - создание NSI-записи для листа рынков;
   - запись в success/fail логи и checkpoints.
7. Завершение запуска с финальным `state.finish_run(...)`.

## 4. CLI-режимы

- `--profile dev|psi|prod|custom`
- `--mode auto|single|mass`
- `--workbook` / `--workbooks`
- `--files-dir` / `--files-map` / `--ask-files-always`
- `--sheet all|permits|markets`
- `--limit`
- `--auth-only`
- `--dry-run`
- `--operator-mode`
- `--resume` / `--no-resume`
- `--reset-state`
- `--state-file`
- `--no-prompt`, `--no-interactive`

## 5. Checkpoints и восстановление

- Namespace checkpoints: `markets_migration:{profile}`.
- На успешном полном завершении rows очищаются (`clear_rows=True`).
- При `failed/stopped` rows сохраняются для продолжения.
- В non-interactive режиме при незавершенном запуске выбирается автоматическое продолжение.

## 6. Operator mode

`--operator-mode` включает реакцию на ошибку строки:
- `retry`
- `skip`
- `abort`

Для веток, где повтор не безопасен без перестройки шага, `retry` логируется и строка пропускается.

## 7. Логи

- `logs/script_creation_log-*.txt`
- `logs/user_log-*.txt`
- `logs/success_log-*.txt`
- `logs/fail_log-*.txt`
- `logs/script_rollback_log-*.txt`

## 8. Откат

`rollback.py` читает success-логи и удаляет записи по `_id/guid/parentEntries` с учетом выбранного профиля.

## 9. Границы изменений

При дальнейших доработках:
- не изменять маппинг полей Excel -> payload без отдельной задачи;
- менять только orchestration/ops слой (CLI, checkpoints, логирование, runtime, rollback);
- подтверждать изменения через `py_compile` и dry-run сценарии.
