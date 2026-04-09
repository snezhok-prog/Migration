# Operations Handoff

Checklist for handing migration package to operations/infra team.

## 1) Required package content

Must include:
- `migration.py`
- `requirements.txt`
- `_api.py`, `_config.py`, `_excel_input.py`, `_logger.py`, `_profiles.py`, `_state.py`, `_utils.py`
- `clear_collections.py`, `rollback*.py`
- active `.xlsm` workbook
- `files/` attachments
- `README.md`

## 2) Before packaging

- Confirm target profile (`psi` when preparing PSI handoff).
- Ensure there are no absolute local-PC paths in code/config.
- Remove runtime artifacts: `logs/*`, `__pycache__`, temp workbook copies.
- Ensure token/cookie files are handled per handoff policy.

## 3) Minimum final checks

```bash
python -m py_compile migration.py
python migration.py --dry-run --skip-auth --no-interactive --limit 1
python migration.py --profile psi --auth-only --no-interactive
```

## 4) Runtime command for operations

```bash
python migration.py --profile psi
```

Non-interactive variant:

```bash
python migration.py --profile psi --no-prompt --no-interactive
```

## 5) Include with handoff

- short release note (what changed),
- known limitations (example: 413 for oversized uploads),
- rollback commands:
  - `python rollback.py`
  - `python rollback_orders.py`
  - `python rollback_stray.py`
  - `python rollback_cards.py`
