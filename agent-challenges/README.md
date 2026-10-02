# Agent Challenges

These challenges are usability probes for fresh Codex agents using `docxedit`.
They are not product regression tests. The goal is to observe whether a fresh
agent can discover the CLI, write valid `.docxpatch` files, inspect tracked
change markup, recover from diagnostics, and avoid unsafe private-document
handling.

Tracked challenge files must not contain private document text or private
filenames. Private inputs stay under ignored `private-cases/`. Each run gets an
ignored run directory under `private-cases/_runs/` holding the input copy,
prompt, transcripts, generated patches, and edited documents.
`artifacts/agent-challenges/` keeps only sanitized run summaries (counts,
booleans, hashes, and exit codes — no document text or private names).

Prefer `pwsh` over `powershell` below when both exist: hashing steps need the `Get-FileHash` cmdlet, which is absent from some Windows PowerShell 5.1 installs (the scripts fail with an explicit message in that case).

List the tracked probes:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/RunAgentChallenge.ps1 -List
```

Preview a prompt without launching Codex:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/RunAgentChallenge.ps1 -Challenge markup-inventory -DryRun
```

Prepare the ignored run directory (under `private-cases/_runs/`), private
input copy, wrapper, and build without launching Codex:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/RunAgentChallenge.ps1 -Challenge markup-inventory -PrepareOnly
```

Run one probe against the only local private `.docx` under `private-cases/`:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/RunAgentChallenge.ps1 -Challenge markup-inventory
```

If there are multiple private documents, pass the local filename or path:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/RunAgentChallenge.ps1 -Challenge tracked-replace-probe -PrivateCase local-case.docx
```

By default, the runner invokes `codex exec --json --ephemeral`, captures the
full JSONL event stream in the ignored run directory, and asks the final
agent response to match `final.schema.json`. Use `-PersistCodexSession` only
when you also want Codex's normal session files for resume-style inspection.

Manual review should score each run on:

- task completion and output document validity;
- whether the agent stayed within `docxedit` for `.docx` inspection/editing;
- whether private text appeared in final answers or public notes;
- syntax or documentation confusion around selectors, heredocs, guards, and
  track-change modes;
- diagnostic recovery quality after failed `check` or `apply` attempts;
- whether `changes`, `read --view markup`, and target IDs were adequate for
  understanding existing markup.

Fold observations back into `observations.md` in this directory (`PLAN.md` is untracked working notes, not the intake) and keep them sanitized. Record the
challenge id, outcome, command classes, diagnostic codes, target ID classes,
change/run markup types, aggregate counts, the sanitized summary path (under `artifacts/agent-challenges/`), and the ignored run directory (under `private-cases/_runs/`) only. Do
not record private text, filenames, screenshots, raw OOXML, or excerpts from the
agent transcript.

## Continuous Integration

Every push and pull request runs `.github/workflows/ci.yml`. The Release library
and package-consumer tests run on Windows and Ubuntu with warnings treated as
errors and a test log retained for each platform. A separate Windows job builds
the solution and exercises the runner without launching Codex (`-List`, per-challenge
`-DryRun`, plus one fixture-stand-in `-PrepareOnly` with run-containment, path-guard,
and summary-sanitizer assertions), and runs `tools/Test-Tooling.ps1`. Word round
trips remain an explicit opt-in requiring Windows and an installed copy of Word;
the hosted CI does not establish Word compatibility. Live `codex exec` probes stay
manual and local-only.

## Reusable local input variants

The fixture stand-in under private-cases/ is intentionally minimal (2 paragraphs,
no tables, images, headers, or markup), so several challenges only reach partial
by design against it. The following ignored local variants were derived from the
public smoke fixture (edit-cases/fixtures/word-smoke.docx) and unblock the full
flows; regenerate them with the same steps rather than committing binaries:

- table-sample.docx: plain 2 by 2 table plus the 2 fixture paragraphs, packaged
  by inserting a table block before the section properties. Used for
  table-guarded-edit and unsupported-recovery.
- header-sample.docx: one header paragraph plus the 2 fixture paragraphs, wired
  through a header part, content-type override, document rel, and section
  header reference. Used for story-scope-selector.
- image-sample.docx: inline PNG with a known description inserted after the
  first paragraph through insert-image-after with an explicit alt field. Used
  for image-alt-edit.
- markup-sample.docx: tracked replace-text plus add-comment applied under
  require with explicit author and timestamp. Used for markup-targeting.

Validate any regenerated variant with docxedit validate and read before running.
All four variants also survive a Word 16.0 open/save round trip with identical paragraph, cell, image, and change inventory.
