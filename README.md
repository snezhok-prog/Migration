# Migration Repository Playbook

This repository hosts production data-migration tooling.

Active project right now: [`universal_myPet`](./universal_myPet/README.md).

## Repository layout

- `universal_myPet/` - production migration script for the "My Pet" scenario (Python).
- `docs/` - onboarding, runbooks, and reusable templates.
- `.github/` - issue and pull request templates.
- Root-level legacy assets are intentionally ignored and kept outside active workflow.

## Quick start

```bash
cd universal_myPet
pip install -r requirements.txt
python migration.py --profile dev --auth-only
python migration.py --profile dev --no-prompt --no-interactive --limit 1
```

## Key documents

- Current migration engine docs: [`universal_myPet/README.md`](./universal_myPet/README.md)
- Operations handoff runbook: [`docs/OPERATIONS_HANDOFF.md`](./docs/OPERATIONS_HANDOFF.md)
- Template for next migration projects: [`docs/PROJECT_TEMPLATE.md`](./docs/PROJECT_TEMPLATE.md)
- AI/engineer onboarding: [`docs/AGENT_ONBOARDING.md`](./docs/AGENT_ONBOARDING.md)
- Wiki drafts for GitHub Wiki: [`docs/WIKI_HOME.md`](./docs/WIKI_HOME.md), [`docs/WIKI_PAGES_TEMPLATE.md`](./docs/WIKI_PAGES_TEMPLATE.md)
- Reusable AI prompt: [`AI_AGENT_PROMPT.md`](./AI_AGENT_PROMPT.md)

## Standard workflow

1. Create a branch from `main`: `feature/<short-topic>`.
2. Keep changes small and atomic.
3. Run relevant checks (at minimum: `py_compile` + smoke run).
4. Open a PR using the provided template.

Details: [`CONTRIBUTING.md`](./CONTRIBUTING.md).

## Data safety

- Do not commit live `token.md` / `cookie.md` values.
- Do not commit runtime logs/checkpoints from real runs.
- Build handoff archives without local secrets.
