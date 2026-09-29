# Agent Challenge Observations

Sanitized intake for `agent-challenges/` probe runs. This file is tracked, so
the `AGENTS.md` rule applies here: private documents stay under ignored
`private-cases/`, and no private document text, filenames, screenshots, or
document contents may appear below.

## What to record per run

- Challenge id (one of `challenges/*.json`) and run date.
- Outcome: completed, partial, or failed, in the runner final-response terms.
- Command classes used (for example `read`, `dump`, `check`, `apply`).
- Diagnostic codes encountered (for example `E6002`).
- Target-ID classes touched (for example `M.P0001`-style paragraph IDs).
- Change/run markup types observed (for example inserted-run, deleted-run).
- Aggregate counts only (commands run, retries, diagnostics seen).
- Sanitized summary path (under `artifacts/agent-challenges/`) and ignored run directory (under `private-cases/_runs/`).
- Tooling gaps only in notes; never document content.

## What never to record

Private text, private filenames, screenshots, raw OOXML, excerpts from the
agent transcript, or anything that identifies the private input document.

## Entry template

Copy the template for each run and keep every field sanitized.

```text
### YYYY-MM-DD — <challenge-id> — <completed|partial|failed>

- Commands:
- Diagnostics:
- Target IDs:
- Markup types:
- Counts:
- Artifacts:
- Notes:
```

## Runs

### 2026-09-29 — tracked-replace-probe (synthetic self-run) — completed

- Commands: read, capabilities, help, lint, check, apply, changes
- Diagnostics: E3201 on a deliberate guard-mismatch probe; none on the success path
- Target IDs: main-story paragraph IDs; one hyperlink ID observed during discovery and avoided for the tracked edit
- Markup types: deleted-run, inserted-run
- Counts: 9 commands, 1 deliberate failing check, 0 retries on the success path
- Artifacts: artifacts/agent-challenges/synthetic-2026-09-29.json
- Notes: synthetic 3-paragraph document with a hyperlink in the middle paragraph. Operator is repo-aware, so discovery ease is not a fresh-agent signal. Capabilities verdicts identified the safe plain-text paragraph. Check reported the resolved paragraph; apply recorded the requested author and timestamp; changes showed one deleted-run and one inserted-run pair. The deliberate guard mismatch failed with code, target, operation, line, column, and help pointer.

### 2026-09-29 — table-guarded-edit (synthetic self-run) — completed

- Commands: read, help, lint, check, apply, changes
- Diagnostics: none
- Target IDs: table, row, and cell IDs
- Markup types: none (untracked edit)
- Counts: 6 commands, 0 retries
- Artifacts: artifacts/agent-challenges/synthetic-2026-09-29.json
- Notes: synthetic 2-row 2-column table document. The help starter pattern transferred directly with added row-count and column-count guards. Check reported the resolved cell with row, column, and before/after counts. No tailored examples were needed.

### 2026-09-29 — preserve-edit (synthetic self-run) — completed

- Commands: check, apply, read, changes
- Diagnostics: none
- Target IDs: main-story paragraph IDs
- Markup types: none asserted (preserve mode)
- Counts: 4 commands, 0 retries
- Artifacts: artifacts/agent-challenges/synthetic-2026-09-29.json
- Notes: insert-after on the first paragraph with preserve mode. Check reported the created paragraph ID; readback showed the hyperlink record moved with its paragraph as positions shifted. Input document still validates.