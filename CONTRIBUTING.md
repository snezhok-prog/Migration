# Contributing

## Branching

- Base branch: `main`
- Feature branch: `feature/<short-topic>`
- Fix branch: `fix/<short-topic>`
- Release branch: `release/<yyyy-mm-dd>`

## Commit style

Use:

```text
<type>: <short summary>
```

Types:
- `feat` - new functionality
- `fix` - bug fix
- `docs` - documentation only
- `refactor` - internal code cleanup without behavior change
- `test` - tests/check improvements
- `build` - packaging/release pipeline changes

## Verification before PR

Minimum for `universal_myPet` changes:

```bash
python -m py_compile universal_myPet/migration.py
python universal_myPet/migration.py --dry-run --skip-auth --no-interactive --limit 1
```

If API/resume/operator behavior changed, run DEV smoke too:

```bash
python universal_myPet/migration.py --profile dev --no-prompt --no-interactive --limit 1
```

## PR checklist

- [ ] Change was validated with relevant checks
- [ ] Docs were updated if behavior changed
- [ ] No secrets/tokens/cookies were committed
- [ ] No runtime artifacts were committed (`logs`, `checkpoints`, `__pycache__`)
- [ ] Rollback plan is clear for risky changes

## Release/handoff checklist

- [ ] Clean transfer package was built
- [ ] Package has no local secrets
- [ ] PSI command is documented and tested in dry/smoke mode
- [ ] Rollback commands are included
