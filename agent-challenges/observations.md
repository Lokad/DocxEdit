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

No runs recorded yet.

