# Integration Guide (Ярмарки)

Документ для команды интеграции: как использовать унифицированный скрипт миграции ярмарок без изменения маппинга данных.

## 1. Точка входа

- Основной запуск: `python migration.py --profile dev`
- `migration.py` — обёртка, которая вызывает `fair_migration.main()`

## 2. Что унифицировано

- Профили стендов `dev|psi|prod|custom`
- Runtime URL переключение (`base/jwt/ui`)
- `single/mass` режим выбора книг
- Привязка книг к разным папкам файлов (`--files-map`)
- Checkpoints/resume (`state/checkpoints.json`)
- Operator mode (`retry/skip/abort`)
- User/script/success/fail/rollback логирование

## 3. Что оставлено без изменений

- Маппинг данных и трансформация строк Excel в payload
- Внутренняя структура целевых записей и коллекций

## 4. Основные CLI сценарии

```bash
python migration.py --profile dev
python migration.py --profile dev --auth-only
python migration.py --profile dev --dry-run --skip-auth
python migration.py --profile dev --operator-mode
python migration.py --profile dev --mode mass
python rollback.py --profile dev
```

## 5. Файлы окружения

- `cookie.md` — cookies для API
- `token.md` — JWT
- `state/checkpoints.json` — чекпоинты запуска
- `ROLLBACK_BODY.json` — тело для сценариев отката
