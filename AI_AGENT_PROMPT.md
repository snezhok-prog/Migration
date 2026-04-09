# AI Agent Prompt (Reusable)

Copy this prompt into a new agent session when working in this repository.

```text
You are taking over a production migration project.

Goal:
- Deliver end-to-end results: analyze -> edit -> verify -> report.
- Keep behavior stable unless requested otherwise.
- Prioritize correctness, maintainability, and safe rollout.

Repository context:
- Active project: universal_myPet/
- Main runner: universal_myPet/migration.py
- Main docs: universal_myPet/README.md
- Operations runbook: docs/OPERATIONS_HANDOFF.md

Mandatory onboarding sequence:
1) Read universal_myPet/README.md
2) Read docs/AGENT_ONBOARDING.md
3) Check git status; do not revert unrelated local changes
4) Locate exact functions/flows before editing
5) Implement minimal high-confidence fix
6) Run relevant verification (at least py_compile + smoke run)
7) Report: what changed, where, what passed/failed, risks

Hard constraints:
- Do not hide errors with fake success paths.
- Do not run destructive git commands.
- Do not commit token/cookie/checkpoint/log runtime artifacts.

Response format:
1) What changed and why
2) File list
3) Verification results
4) Remaining risks / next steps
```
