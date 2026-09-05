# Agent Challenges

These challenges are usability probes for fresh Codex agents using `docxedit`.
They are not product regression tests. The goal is to observe whether a fresh
agent can discover the CLI, write valid `.docxpatch` files, inspect tracked
change markup, recover from diagnostics, and avoid unsafe private-document
handling.

Tracked challenge files must not contain private document text or private
filenames. Private inputs stay under ignored `private-cases/`; transcripts,
generated patches, and edited documents stay under ignored
`artifacts/agent-challenges/`.

List the tracked probes:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/RunAgentChallenge.ps1 -List
```

Preview a prompt without launching Codex:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/RunAgentChallenge.ps1 -Challenge markup-inventory -DryRun
```

Prepare the ignored run directory, private input copy, wrapper, and build without
launching Codex:

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
full JSONL event stream in the ignored artifact directory, and asks the final
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
change/run markup types, aggregate counts, and ignored artifact paths only. Do
not record private text, filenames, screenshots, raw OOXML, or excerpts from the
agent transcript.

## Continuous Integration

Every push and pull request runs `.github/workflows/agent-tooling.yml`: it builds
the solution and exercises the runner without launching Codex (`-List` plus one
`-DryRun`). Live `codex exec` probes stay manual and local-only.
