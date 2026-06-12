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
- `context input.docx --id M.P0001 [--radius N] [--headers-footers] [--view final|original|markup] [--max-text N] [--json] [--diagnostics path] [--strict]`
- `styles input.docx [--json] [--diagnostics path] [--strict]`
- `media input.docx [--extract dir] [--json] [--diagnostics path] [--strict]`
- `changes input.docx [--include-comment-text] [--max-comment-text N] [--json] [--diagnostics path] [--strict]`

`changes` lists existing tracked-change and comment markup. By default it does not
print revision text or comment body text. It reports counts, IDs, type, story, part,
target, revision/comment metadata, text length, and child element count.
JSON output includes `Summary` counts by type, `GroupSummary` counts by
`story`, `part`, `author`, and `target`, `TargetSummary` compact per-target
rollups, `CommentSummary` compact per-comment rollups, and `Changes` records.
Comment body snippets appear only when `--include-comment-text` is passed; bound them
with `--max-comment-text N`.
Plain text output includes the same group summaries as lines like
`summary group=story key="main" type=inserted-run count=1`.
It also includes `target-summary` and `comment-summary` lines before individual
records. With `--include-comment-text`, comment summaries and comment body records
add `comment-text-length`, `comment-text`, and `comment-text-truncated` fields.

Some records legitimately have `target=unknown`: for example package-level range
markers or markup not inside or adjacent to a modeled paragraph, table, cell, or
section target. Each record has `target-status` (`targeted`, `comment-anchor`, or
`targetless`), `target-source` (`ancestor`, `adjacent-range`, `comment-anchor`, or
`none`), and optional `target-reason`, `nearest-target`, and `target-note` fields.
`nearest-target` is context only, not exact ownership. Body-level range boundaries
that are linked by adjacent-target heuristics use `target-source=adjacent-range`.
Range start/end records that share a revision or comment ID expose
`paired-change-id`, which helps recover sparse end-marker metadata from the matching
start record. Comment records are linked by `comment-id`; when possible, comment
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
To extract structured run metadata in PowerShell:

```powershell
docxedit dump report.docx --id M.P0004 --runs --json |
  ConvertFrom-Json |
  Select-Object -ExpandProperty Runs
```

`context` summarizes nearby modeled structure around one target without broad document
text. Its default `--max-text` is `0`, so paragraph and cell text fields are present
but empty. Use `--radius` to include same-kind neighbors and raise `--max-text` only
when short snippets are needed.
Read text views are `final` (default), `original`, and `markup`. Markup view includes inserted and deleted text with lightweight `[+text+]` and `[-text-]` markers.
`read --summary` prints package/story counts without listing every target, which is
useful for large-document validation.
Paragraph lines may include `styleId=...` and resolved list metadata. List metadata
starts with the concrete `numId` and zero-based level, then includes resolved
`abstractNumId`, numbering `format`, `level-text`, paragraph style link, and
`source=style` or `source=style-inherited` when the list comes from paragraph style
inheritance rather than direct paragraph numbering.
`styles` output includes inheritance links such as `based-on`, `next`, `linked`, and
style-level numbering defaults when present.
`read` and `outline` list bookmark and content-control metadata when present.
Bookmark records include name, OOXML ID, story, part, start/end targets, and whether
the range is complete. Content-control records include kind, tag, alias, lock, story,
part, containing target, and text length. `context` annotates nearby targets with
`bookmark-names`, `content-controls`, `content-control-tags`, and
`content-control-aliases` so agents can connect selector names to stable target IDs.
If `bookmark:"Name"` or `content-control:"TagOrAlias"` is ambiguous, selector
diagnostics list candidate paragraph IDs; retry with an explicit ID.

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
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- help context
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- help changes
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- help check
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- help apply
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- help patch
```

Examples use `--output` for PowerShell compatibility; `-o` remains supported as a
short alias.
