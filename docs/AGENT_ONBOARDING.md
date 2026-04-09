# Agent Onboarding

This guide helps a new engineer or AI agent enter the project quickly.

## 1) Where production code lives

Active migration project: `universal_myPet/`.

Key files:
- `universal_myPet/migration.py` - main migration flow
- `universal_myPet/_excel_input.py` - Excel parsing
- `universal_myPet/_api.py` - HTTP/API client calls
- `universal_myPet/_profiles.py` - environment profiles (`dev/psi/prod`)
- `universal_myPet/_state.py` - resume/checkpoints

## 2) Baseline commands

```bash
python universal_myPet/migration.py --profile dev --auth-only
python universal_myPet/migration.py --dry-run --skip-auth --no-interactive --limit 1
python universal_myPet/migration.py --profile dev --no-prompt --no-interactive --limit 1
```

## 3) Required checks after changes

1. Syntax:
```bash
python -m py_compile universal_myPet/migration.py
```

2. Dry-run:
```bash
python universal_myPet/migration.py --dry-run --skip-auth --no-interactive --limit 1
```

3. DEV smoke (if API/operator/resume changed):
```bash
python universal_myPet/migration.py --profile dev --no-prompt --no-interactive --limit 1
```

## 4) Never commit these runtime files

- `token.md`, `cookie.md` with live session values
- `state/checkpoints.json` from real runs
- `ROLLBACK_BODY.json` from real runs
- runtime logs and `__pycache__`

## 5) Definition of Done

A task is done only when:
- code changes are implemented,
- behavior is verified by relevant commands,
- docs are updated if needed,
- report clearly states what was verified and what was not.
