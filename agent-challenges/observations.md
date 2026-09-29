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
### 2026-09-29 — ambiguity-recovery (synthetic self-run) — completed

- Commands: read, apply, check, validate
- Diagnostics: E1202 on the selector patch and on an occurrence-based retry; none on the explicit-ID path
- Target IDs: main-story paragraph IDs; selector matched M.P0002 and M.P0003 with matches=2; recovery targeted M.P0003 with a guard
- Markup types: none (untracked edit)
- Counts: 8 commands, 2 failing checks, 0 retries after the working recovery
- Artifacts: artifacts/agent-challenges/synthetic-2026-09-29.json
- Notes: public Word fixture input with synthetic patch text. Operator is repo-aware, so recovery ease is not a fresh-agent signal. Duplicating a paragraph set up the ambiguity; the E1202 diagnostic named both candidates with the match count, help pointer, and line and column. A first recovery wrongly used occurrence, which selects find matches rather than selector targets, so E1202 persisted; the documented explicit-ID recovery with expect-text then edited only M.P0003 while M.P0002 stayed byte-identical in readback, and structural validation passed. The occurrence scope trap is worth one line in patch-format.md. A negative E1202 edit case now locks the refusal in.
### 2026-09-29 — unsupported-recovery (synthetic self-run) — completed

- Commands: read, check, capabilities, help, apply, validate
- Diagnostics: E4316 on the delete-column check; none on the recovery path
- Target IDs: table, row, and cell IDs; recovery cleared M.T0001.R01.C02 and M.T0001.R02.C02 with empty text
- Markup types: none (untracked edit)
- Counts: 7 commands, 1 failing check, 0 retries after the working recovery
- Artifacts: artifacts/agent-challenges/synthetic-2026-09-29.json
- Notes: synthetic 2-row 2-column document built with archive tooling since docxedit has no create command. Operator is repo-aware, so recovery ease is not a fresh-agent signal. The E4316 refusal carried reason, target, operation, line, column, and help pointer, and help delete-column is honest about the limitation. Capabilities refuse table IDs, so the recovery was found at cell level, where set-cell verdicts are supported. Clearing with empty text kept the grid shape, which is honestly different from structural deletion; readback and structural validation confirmed the result. A negative E4316 edit case now locks the refusal in.
### 2026-09-29 — iterative-tracked-edit (synthetic self-run) — completed

- Commands: apply, read, check, changes, validate
- Diagnostics: E6002 on the overlapping third patch; none on either disjoint edit
- Target IDs: main-story paragraph IDs; revisions allocated 1 and 2, then 3 and 4
- Markup types: deleted-run and inserted-run pairs, two pairs side by side
- Counts: 7 commands, 1 failing check, 0 retries
- Artifacts: artifacts/agent-challenges/synthetic-2026-09-29.json
- Notes: public Word fixture input with synthetic patch text. Operator is repo-aware, so workflow ease is not a fresh-agent signal. A first require-mode replacement created one revision pair; a second patch on untouched text in the same paragraph created a second pair with continuing IDs while markup readback kept both pairs coherent, and structural validation passed. A third patch overlapping the existing revision failed safely with E6002 carrying reason, target, operation, line, column, and help pointer. The D05 cross-patch workflow behaves as documented, and a two-patch unit test now locks it in.
### 2026-09-29 — tracked-dump-verification (synthetic self-run) — completed

- Commands: changes, dump, context, validate
- Diagnostics: none
- Target IDs: M.P0002; run IDs M.P0002.R0001 through R0006; change IDs M.CH0001 through M.CH0004
- Markup types: deleted-run and inserted-run, two pairs with distinct timestamps
- Counts: 4 commands, 0 failing checks, 0 retries
- Artifacts: artifacts/agent-challenges/synthetic-2026-09-29.json
- Notes: verification half runs on the iterative-tracked-edit output, which already covers the edit half. Operator is repo-aware, so discovery ease is not a fresh-agent signal. Dump runs expose run IDs, markup types, revision IDs, authors, and timestamps per run plus a changes section, so both revision pairs verify without raw OOXML. Change IDs live in a separate M.CH namespace from run IDs and cross-link through revision IDs and parent target IDs. No fallback or retry was needed.

### 2026-09-29 — markup-targeting (synthetic self-run) — completed

- Commands: changes, context
- Diagnostics: none
- Target IDs: change M.CH0001 points at parent target M.P0002; context shows its neighbors
- Markup types: none inspected beyond the inventory listing
- Counts: 2 commands, 0 failing checks, 0 retries
- Artifacts: artifacts/agent-challenges/synthetic-2026-09-29.json
- Notes: read-only navigation on the same reviewed document. Operator is repo-aware, so navigation ease is not a fresh-agent signal. The changes inventory bridges directly to editable targets through parent target IDs, and context renders the neighborhood metadata-only by default without forcing text exposure. No private text was needed at any step.
### 2026-09-29 — whitespace-selector (synthetic self-run) — completed

- Commands: apply, read, check, validate
- Diagnostics: E3201 on the deliberately normalized guard; none elsewhere
- Target IDs: main-story paragraph IDs; normalized text selector resolved to M.P0002
- Markup types: none (untracked edit)
- Counts: 7 commands, 1 failing check, 0 retries
- Artifacts: artifacts/agent-challenges/synthetic-2026-09-29.json
- Notes: public Word fixture input with synthetic patch text. Operator is repo-aware, so workflow ease is not a fresh-agent signal. A paragraph with a doubled space was addressed through a single-spaced text selector and resolved correctly; the single-spaced expect-text guard failed with E3201 while the exact-spaced guard passed, and the applied edit replaced only the find span with spacing preserved. The D09 normalized-match versus exact-guard contract from patch-format.md behaves as documented, matching existing unit coverage.
### 2026-09-29 — markup-inventory (synthetic self-run) — completed

- Commands: apply, changes, read
- Diagnostics: none
- Target IDs: main-story paragraph IDs; one comment anchored to M.P0002
- Markup types: deleted-run and inserted-run pairs plus comment range markup
- Counts: 5 commands, 0 failing checks, 0 retries
- Artifacts: artifacts/agent-challenges/synthetic-2026-09-29.json
- Notes: synthetic document with two revision pairs plus one anchored comment. Operator is repo-aware, so discovery ease is not a fresh-agent signal. Default changes output inventories markup by type, author, part, and story with no document text; a byte scan of the 40-line output found no revision or comment text. Read summary gives shape counts without text. No usability gap observed on this surface.
### 2026-09-29 — story-scope-selector (synthetic self-run) — completed

- Commands: read, find, check, apply, validate
- Diagnostics: E1201 on the main-story text selector against header-only text; none on the explicit-ID path
- Target IDs: H001.P0001 discovered through headers-footers coverage; main-story suggestion pointed at M.P0001
- Markup types: none (untracked edit)
- Counts: 7 commands, 1 failing check, 0 retries after the working recovery
- Artifacts: artifacts/agent-challenges/synthetic-2026-09-29.json
- Notes: synthetic main-plus-header document built with archive tooling since docxedit has no create command. Operator is repo-aware, so discovery ease is not a fresh-agent signal. Semantic text selectors search only the main story while find needs its headers-footers flag for the same text; the E1201 nearby-paragraph suggestion stayed main-story scoped. Recovery through the explicit H001 ID with check before apply changed only the header paragraph, and structural validation passed. The story scope is now documented in patch-format.md and help patch.

### 2026-09-29 — image-alt-edit (synthetic self-run) — completed

- Commands: read, help, check, apply, read, validate
- Diagnostics: E3201 on a deliberate guard-mismatch probe; none on the success path
- Target IDs: image ID M.I0001 discovered through read; containing paragraph M.P0001 observed
- Markup types: none (untracked metadata edit)
- Counts: 7 commands, 1 deliberate failing check, 0 retries
- Artifacts: artifacts/agent-challenges/synthetic-2026-09-29.json
- Notes: synthetic single-image document built with archive tooling since docxedit has no create command. Operator is repo-aware, so discovery ease is not a fresh-agent signal. Focused help gave minimal plus guarded examples that transferred verbatim; the mismatch diagnostic carried code, both values, target, operation, line, column, and help pointer. Check with a preview budget showed before/after identically through check and apply; default reports stay metadata-only unless the budget is set. Readback plus structural validation confirmed the result.

### 2026-09-29 — markup-inventory (live codex run) — completed

- Commands: help, changes, read, validate
- Diagnostics: none
- Target IDs: none present; no markup IDs available
- Markup types: none present (zero tracked-change and comment records)
- Counts: 10 commands, 0 failing checks, 0 retries; about 108k input tokens with about 87k cached
- Artifacts: artifacts/agent-challenges/markup-inventory/20260929T133603Z/summary.json
- Notes: first live codex-exec run in this thread (ephemeral, prompt-only, codex-cli 0.155.1) using the public fixture stand-in, whose hash was identical before and after. The agent consulted focused help before each new command class, stayed within docxedit with no forbidden inspection, and disclosed no text. Finding: empty changes and read results lack explicit zero totals and scanned-story coverage, forcing interpretation of empty arrays. Caveat: the fixture carries no markup, so populated-inventory usability remains unassessed; model, usage, and thread are recorded in the summary.

### 2026-09-29 — tracked-replace-probe (live codex run) — completed

- Commands: help, read, changes, dump, check, apply, changes, dump
- Diagnostics: none
- Target IDs: main-story paragraph IDs; edited M.P0002
- Markup types: deleted-run, inserted-run (one pair with requested author and timestamp)
- Counts: 11 commands, 0 failing checks, 0 retries; about 180k input tokens with about 151k cached
- Artifacts: artifacts/agent-challenges/tracked-replace-probe/20260929T133836Z/summary.json
- Notes: second live codex-exec run (ephemeral, prompt-only, codex-cli 0.155.1) on the public fixture stand-in, hash-identical before and after. The agent led with help patch and apply, discovered the target through read plus dump runs, wrote a fully guarded tracked replacement with explicit author and timestamp under require, and verified through changes (0 to 2 records), dump, and input-hash comparison. Post-checks confirm the output reads and changes cleanly. No agent-reported weaknesses; no text disclosed. Caveat: single clean-path run on a simple paragraph; ambiguity and overlap recovery remain assessed only in synthetic self-runs.

### 2026-09-29 — ambiguity-recovery (live codex run) — failed

- Commands: help patch attempted with no result; no docxedit edit commands ran
- Diagnostics: none observed (E1202 recovery untested)
- Target IDs: none discovered
- Markup types: none observed
- Counts: 0 effective commands; run abandoned after several minutes of unresponsive tool execution
- Artifacts: private run directory private-cases/_runs/ambiguity-recovery/20260929T134140Z (no sanitized summary; runner had not finalized)
- Notes: third live codex-exec run (ephemeral, prompt-only) on the public fixture stand-in. The agent reported shell execution unresponsive including basic directory checks, stopped its pending requests, made no edits, and disclosed no text. No product signal: the failure is in agent tool execution, not docxedit behavior, and the two prior live runs on the same substrate succeeded. Possible confounds: observer polled the run directory aggressively during the run, and the agent process burned about 10 minutes of CPU while producing only its final response. Do not read this as evidence about ambiguity recovery itself, which remains assessed only in synthetic self-runs; a rerun with wider observer polling gaps would test flakiness versus a systematic problem.

### 2026-09-29 — ambiguity-recovery (live codex run) — completed

- Commands: help, read, check, apply, read, validate
- Diagnostics: E1202 on the ambiguous selector; none on the recovery path
- Target IDs: selector matched M.P0001 and M.P0003 with matches 2; recovery targeted M.P0003 with a guard
- Markup types: none (untracked edit)
- Counts: 16 commands, 1 failing check, 0 retries after the working recovery; about 302k input tokens with about 270k cached
- Artifacts: artifacts/agent-challenges/ambiguity-recovery/20260929T135222Z/summary.json
- Notes: rerun after the earlier tooling-blocked attempt on the same challenge; with wider observer polling gaps it completed cleanly, supporting flake over systematic for the prior failure. The agent manufactured ambiguity in a separate working copy with the input hash identical throughout, hit E1202 with both candidates named, and recovered through the explicit ID plus expect-text without touching occurrence. Verified only the chosen paragraph changed; structural validation clean. No agent-reported weaknesses; no text disclosed.

### 2026-09-29 — image-alt-edit (live codex run) — partial

- Commands: help, read, media, validate, check
- Diagnostics: E1201 on the diagnostic probe (M.I0001 matched zero targets); nothing else
- Target IDs: none (zero images; the M.I0001 placeholder never resolved)
- Markup types: not applicable
- Counts: 11 commands, 1 failing check, 0 applies, 0 retries; about 109k input tokens with about 83k cached
- Artifacts: artifacts/agent-challenges/image-alt-edit/20260929T135728Z/summary.json
- Notes: first live run of the authored image-alt-edit challenge, against the imageless public fixture stand-in. The agent confirmed zero images through read plus the media command, wrote an explicitly non-appliable diagnostic probe to capture the E1201 failure mode, and stopped without applying or creating output. Runner reports harness success; challenge outcome partial by design since no edit was possible. Validates the challenge failure branch, E1201 discoverability, and fresh-agent use of the media command; no text disclosed. A follow-up on an image-bearing input would exercise the edit branch.

### 2026-09-29 — iterative-tracked-edit (live codex run) — completed

- Commands: help, read, dump, changes, check, apply, validate
- Diagnostics: E6002 on the deliberate overlap probe (check and apply); none on the two disjoint edits
- Target IDs: main-story paragraph IDs; two revision pairs 1/2 and 3/4 in M.P0002
- Markup types: deleted-run and inserted-run pairs with requested author and timestamp
- Counts: 20 commands, 1 failing check plus 1 failing apply on the overlap probe, 0 retries; about 297k input tokens with about 264k cached
- Artifacts: artifacts/agent-challenges/iterative-tracked-edit/20260929T135941Z/summary.json
- Notes: live codex-exec run (ephemeral, prompt-only, codex-cli 0.155.1) on the public fixture stand-in, hash-identical before and after. The agent led with five help calls, applied two guarded replacements under require with the first revision pair preserved, then probed the overlap: E6002 named the tracked-insertion boundary on both check and apply with no rejected output created. Structural validation clean. No agent-reported weaknesses; no text disclosed.

### 2026-09-29 — tracked-review-workflow (live codex run) — completed

- Commands: help, read, dump, changes, check, apply, changes, validate
- Diagnostics: none
- Target IDs: main-story paragraph IDs; edited and annotated M.P0002
- Markup types: deleted-run, inserted-run, comment with range and reference markup
- Counts: 16 commands, 0 failing checks, 0 retries; about 164k input tokens with about 131k cached
- Artifacts: artifacts/agent-challenges/tracked-review-workflow/20260929T140301Z/summary.json
- Notes: live codex-exec run (ephemeral, prompt-only, codex-cli 0.155.1) on the public fixture stand-in, hash-identical before and after. The agent wrote one patch combining a guarded replacement with an anchored comment under require with explicit author and timestamp; notably the comment guard used the post-edit paragraph text, showing sequential-guard understanding. Verified one deletion run, one insertion run with revisions 1 and 2, a new comment 0 anchored to the edited paragraph, and all three read views plus structural validation. No agent-reported weaknesses; no text disclosed.

### 2026-09-29 — unsupported-recovery (live codex run) — partial

- Commands: help, read, changes, check, capabilities, validate
- Diagnostics: E4316 on the delete-column probe; E1201 on a cell-capabilities probe; nothing else
- Target IDs: none usable (fixture holds 2 paragraphs and 0 tables)
- Markup types: not applicable
- Counts: 16 commands, 1 failing check, 0 applies, 0 retries; about 192k input tokens with about 174k cached
- Artifacts: artifacts/agent-challenges/unsupported-recovery/20260929T140556Z/summary.json
- Notes: live codex-exec run (ephemeral, prompt-only, codex-cli 0.155.1) on the tableless public fixture stand-in. The agent correctly declined to fabricate a table, probed the E4316 refusal path with its reason and help pointer, identified the guarded set-cell recovery while explaining why it could not perform it, and stopped without applying. Runner reports harness success; challenge outcome partial by design. Finding: E4316 fires even when the table target does not exist, masking target-existence feedback behind the unsupported-operation refusal; whether refusal precedence should yield to selector validation is a design call. No text disclosed.

### 2026-09-29 — whitespace-selector (live codex run) — completed

- Commands: help, changes, dump, read, check, apply, validate
- Diagnostics: E3201 on the normalized guard probe; none elsewhere
- Target IDs: main-story paragraph IDs; edited M.P0002
- Markup types: none (untracked edit)
- Counts: 23 commands, 1 failing check, 0 retries; about 348k input tokens with about 313k cached
- Artifacts: artifacts/agent-challenges/whitespace-selector/20260929T141009Z/summary.json
- Notes: live codex-exec run (ephemeral, prompt-only, codex-cli 0.155.1) on the public fixture stand-in, hash-identical before and after. The fixture had no repeated whitespace, so the agent prepared a separate doubled-space baseline with docxedit, then proved the D09 contract live: a single-spaced text selector resolved while the single-spaced guard failed E3201 and the exact-spaced guard passed, with only the find span changed and spacing preserved. It wrote its own machine-checked verification summary plus a verifier script. Finding: help patch never states the normalized-selector versus exact-guard whitespace contract generally (patch-format.md does); the agent succeeded anyway. No text disclosed.

### 2026-09-29 — alias-compose-probe (live codex run) — completed

- Commands: help, read, changes, dump, styles, check, apply, validate
- Diagnostics: none
- Target IDs: M.P0002 anchor; created paragraph addressed as @insertedNote for two follow-up edits
- Markup types: none (untracked edits)
- Counts: 15 commands, 0 failing checks, 0 retries; about 178k input tokens with about 145k cached
- Artifacts: artifacts/agent-challenges/alias-compose-probe/20260929T141853Z/summary.json
- Notes: live codex-exec run (ephemeral, prompt-only, codex-cli 0.155.1) on the public fixture stand-in, hash-identical before and after. The agent inserted with as plus a style, then addressed the alias from set-style guarded by expect-style and from replace-text guarded by expect-text; all three operations succeeded and readback confirmed the styled replacement text. Finding: per-operation help insert-after lists as but shows no alias composition example, while help patch explains the mechanism; a registry example now closes that gap. No text disclosed.
