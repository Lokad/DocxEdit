# CLI

Run locally with:

```powershell
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- <command> [options]
```

## Read And Explore

- `read input.docx [--summary] [--headers-footers] [--all-stories] [--view final|original|markup] [--max-text N] [--json] [--diagnostics path] [--strict]`
- `outline input.docx [--headers-footers] [--json] [--diagnostics path] [--strict]`
- `find input.docx "text" [--headers-footers] [--view final|original|markup] [--max-text N] [--json] [--diagnostics path] [--strict]`
- `dump input.docx --id M.P0001 [--runs] [--view final|original|markup] [--max-text N] [--json] [--diagnostics path] [--strict]`
- `styles input.docx [--json] [--diagnostics path] [--strict]`
- `media input.docx [--extract dir] [--json] [--diagnostics path] [--strict]`
- `changes input.docx [--json] [--diagnostics path] [--strict]`

`changes` lists existing tracked-change and comment markup without printing revision or comment text. It reports counts, IDs, type, story, part, target, revision/comment metadata, text length, and child element count.
JSON output includes `Summary` counts by type, `GroupSummary` counts by
`story`, `part`, `author`, and `target`, `TargetSummary` compact per-target
rollups, `CommentSummary` compact per-comment rollups, and private-text-free
`Changes` records.
Plain text output includes the same group summaries as lines like
`summary group=story key="main" type=inserted-run count=1`.
It also includes `target-summary` and `comment-summary` lines before individual
records.

Some records legitimately have `target=unknown`: for example package-level range
markers or markup not inside or adjacent to a modeled paragraph, table, cell, or
section target. Each record has `target-status` (`targeted`, `comment-anchor`, or
`targetless`) and targetless records include a short `target-note`. Comment records
are linked by `comment-id`; when possible, comment
body records also include `comment-anchor-target`, `comment-reference-target`,
`comment-anchor-story`, and `comment-anchor-part` so an agent can navigate from the
comment-story record back to the main document anchor without printing comment text.
`TimestampUtc` and `CommentTimestampUtc` serialize as nullable UTC ISO-8601 values.
Raw JSON uses UTC values such as `+00:00`; PowerShell `ConvertFrom-Json` may display
date values in the local timezone after parsing.
Use `dump --runs` on a target to see run-level `markup=...`, `revision-id`, and
`comment-id` annotations for nearby tracked-change/comment markup. With `--json`,
`dump --runs` also exposes those annotations as structured `Runs` objects.
Change IDs from `changes` identify markup records. Run IDs from `dump --runs` identify
rendered run/marker lines and are a separate namespace.
Read text views are `final` (default), `original`, and `markup`. Markup view includes inserted and deleted text with lightweight `[+text+]` and `[-text-]` markers.
`read --summary` prints package/story counts without listing every target, which is
useful for large-document validation.

## Patch

- `check input.docx edits.docxpatch [--track-changes mode] [--author name] [--timestamp-utc instant] [--json] [--report path] [--diagnostics path] [--strict]`
- `apply input.docx edits.docxpatch --output output.docx [--track-changes mode] [--author name] [--timestamp-utc instant] [--json] [--report path] [--diagnostics path] [--strict]`

Track-change modes are `off`, `preserve`, `suggest`, and `require`. `replace-text` generates simple `w:del`/`w:ins` tracked-change markup with the selected author and timestamp under `suggest` or `require`. Tracked output is limited to text-only matches without tabs or line breaks, protected OOXML boundaries, existing revision markup, or mixed direct run formatting. `require` fails unsupported operations or unsupported `replace-text` shapes; `suggest` warns and applies unsupported edits directly.
Plain text `check` and `apply` output includes one `operation index=...` line per
patch operation with operation name, target, and success. Use `--report` for the full
JSON operation report.

Exit codes:

- `0`: success
- `1`: operation failed
- `2`: invalid CLI usage
- `3`: strict mode saw warnings or errors
- `4`: unexpected CLI exception

## Help

Command-specific help is available for common agent workflows:

```powershell
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- help dump
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- help changes
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- help check
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- help apply
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- help patch
```

Examples use `--output` for PowerShell compatibility; `-o` remains supported as a
short alias.
