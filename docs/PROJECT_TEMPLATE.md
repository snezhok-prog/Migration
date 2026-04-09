# Migration Project Template

Use this template to bootstrap the next migration project.

## 1) Target structure

```text
new_migration_project/
  migration.py
  _api.py
  _excel_input.py
  _profiles.py
  _state.py
  _logger.py
  _utils.py
  _config.py
  clear_collections.py
  rollback.py
  rollback_*.py
  README.md
  requirements.txt
  files/
  logs/
  state/
```

## 2) Mandatory capabilities

- automated request generation and send
- file upload pipeline (multipart/base64 fallback if needed)
- detailed success/fail logs
- rollback payload generation
- resume via checkpoints
- operator mode (retry/relogin/skip/abort)
- single/mass workbook modes

## 3) Minimum CLI flags

- `--profile`
- `--dry-run`
- `--skip-auth`
- `--auth-only`
- `--no-interactive`
- `--operator-mode`
- `--reset-state`
- `--no-resume`
- `--mode single|mass`
- `--workbooks`
- `--files-map`

## 4) Quality gate before release

1. `py_compile` passes
2. `dry-run` on small limit
3. real DEV smoke run
4. logs + rollback artifacts reviewed
5. resume flow validated via controlled stop/restart

## 5) Required docs in each new project

- `README.md` with quick start and commands
- `MAPPING_PARITY.md` (legacy vs new mapping matrix)
- operations handoff checklist
- reusable prompt for the next AI agent
