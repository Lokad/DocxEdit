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
### 2026-09-29 — alias-compose-probe (synthetic self-run) — completed

- Commands: read, capabilities, styles, lint, check, apply, validate, changes
- Diagnostics: E7101 on the first check from a style display-name case mismatch; none on the success path
- Target IDs: main-story paragraph IDs; the insert created M.P0003 and alias @sec1 resolved to M.P0003 in both follow-up operations
- Markup types: none (untracked edit)
- Counts: 13 commands, 1 failing check, 1 CLI usage error on capabilities arguments, 1 recovery retry
- Artifacts: artifacts/agent-challenges/synthetic-2026-09-29.json
- Notes: public Word fixture input with synthetic patch text. Operator is repo-aware, so discovery ease is not a fresh-agent signal. The alias flow worked end to end: check reported created-target-ids M.P0003 plus affected targets, and apply agreed with identical reporting. The first check failed because the patch used display name Heading 2 while the fixture names it heading 2; the E7101 diagnostic carried code, target, operation, line, column, and help pointer, and the styles listing gave the exact styleId for the retry. Gap found during the run: help patch lists as per operation but had no Result-aliases concept section, so the @name mechanism was only documented in patch-format.md and per-operation help; a help overview addition is committed alongside this run.
### 2026-09-29 — tracked-review-workflow (synthetic self-run) — completed

- Commands: check, apply, read, changes, validate
- Diagnostics: none
- Target IDs: main-story paragraph IDs; the comment was created as comment:0
- Markup types: deleted-run, inserted-run, comment with range and reference markup
- Counts: 6 commands, 0 failing checks, 1 CLI usage error on a wrong changes flag, 0 retries
- Artifacts: artifacts/agent-challenges/synthetic-2026-09-29.json
- Notes: public Word fixture input with synthetic patch text. Operator is repo-aware, so workflow ease is not a fresh-agent signal. One patch carried a guarded text replacement plus a comment under require policy: check passed, apply recorded generated revision IDs with created-target-ids comment:0, markup readback showed the revision pair, changes showed the runs plus the comment under separate authors, and structural validation passed. The D15 atomic review workflow behaves as documented, and a positive require-mode edit case now locks the behavior in.
### 2026-09-29 — semantic-anchor-edit (synthetic self-run) — completed

- Commands: read, check, apply, capabilities, dump, validate
- Diagnostics: E1201 from capabilities and dump against a bookmark ID; none on the patch path
- Target IDs: main-story paragraph IDs; created bookmark M.B0001; semantic selector bookmark:ReviewAnchor for the edit
- Markup types: none (untracked edit); bookmark markers preserved complete around the replacement
- Counts: 9 commands, 0 failing checks, 2 inspection refusals, 0 retries
- Artifacts: artifacts/agent-challenges/synthetic-2026-09-29.json
- Notes: public Word fixture input with synthetic patch text. Operator is repo-aware, so discovery ease is not a fresh-agent signal. The fixture had no semantic anchors, so the run bootstrapped one with a guarded add-bookmark, then replaced its text through bookmark:ReviewAnchor with an expect-text guard; readback kept the markers complete and structural validation passed. Gap found: capabilities and dump refuse bookmark IDs, so semantic anchors are discoverable only through read; bookmark inspection and capability verdicts would close the D17 loop. A guarded selector edit case now locks the supported path in.