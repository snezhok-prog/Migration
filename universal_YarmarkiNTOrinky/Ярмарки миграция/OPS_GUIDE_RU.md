# OPS Guide (Ярмарки)

## 1. Стандартный запуск

```bash
python migration.py --profile dev
```

## 2. Проверка авторизации перед миграцией

```bash
python migration.py --profile dev --auth-only
```

## 3. Миграция в operator-mode

```bash
python migration.py --profile dev --operator-mode
```

На ошибке строки доступны действия:
- `retry` — повторить текущую строку
- `skip` — пропустить строку и продолжить
- `abort` — остановить миграцию

## 4. Продолжение после остановки

Повторный запуск:

```bash
python migration.py --profile dev
```

Если `state/checkpoints.json` содержит незавершенную миграцию, скрипт предложит продолжить с checkpoint.

## 5. Сброс checkpoints

```bash
python migration.py --profile dev --reset-state --no-resume
```

## 6. Dry-run

```bash
python migration.py --profile dev --dry-run --skip-auth
```

## 7. Mass-режим

```bash
python migration.py --profile dev --mode mass
```

или c явным списком:

```bash
python migration.py --profile dev --mode mass --workbooks "book1.xlsm;book2.xlsm"
```

## 8. Привязка папок файлов

```bash
python migration.py --profile dev --files-map "book1.xlsm=files/one;book2.xlsm=files/two"
```

## 9. Откат

```bash
python rollback.py --profile dev
```

## 10. Логи

- `logs/script_creation_log-*.txt` — технический лог
- `logs/user_log-*.txt` — лог для оператора
- `logs/success_log-*.txt` — успешно созданные записи
- `logs/fail_log-*.txt` — ошибки строк
