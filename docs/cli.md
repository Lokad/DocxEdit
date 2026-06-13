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
- `validate input.docx [--profile structural|package] [--json] [--diagnostics path] [--strict]`
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
Patch operations can create comments on modeled paragraph IDs with `add-comment`.
Existing comment operations can target `comment:<id>` or comment body IDs such as
`C001.C0001`: `set-comment-text` replaces a comment body under explicit patch
control, `resolve-comment` and `reopen-comment` create or update modern
`commentsExtended` resolution metadata for basic comments, and `delete-comment`
removes the comment body plus matching range/reference markers. `add-comment-reply`
and `delete-comment-reply` are recognized but fail with `E4314` until threaded
reply metadata is safely modeled.

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
Modern Word comment resolution metadata appears as `comment-para-id`,
`comment-parent-para-id`, and `comment-resolved` on records, and as `para-id`,
`parent-para-id`, and `resolved` on `comment-summary` lines.
`TimestampUtc` and `CommentTimestampUtc` serialize as nullable UTC ISO-8601 values.
Raw JSON uses UTC values such as `+00:00`; JSON consumers that parse dates may display
those values in a local timezone, so inspect the raw serialized value when the offset
matters.
Use `dump --runs` on a target to see run-level `markup=...`, `revision-id`, and
`comment-id` annotations for nearby tracked-change/comment markup. With `--json`,
`dump --runs` also exposes those annotations as structured `Runs` objects.
Use `dump --id C001.C0001` or `dump --id comment:3` to inspect comment body
metadata without printing the comment body text.
Change IDs from `changes` identify markup records. Run IDs from `dump --runs` identify
rendered run/marker lines and are a separate namespace.
Use `docxedit dump report.docx --id M.P0004 --runs --json` when structured run
metadata is easier for an integration to consume than text output.

`context` summarizes nearby modeled structure around one target without broad document
text. Its default `--max-text` is `0`, so paragraph and cell text fields are present
but empty. Comment anchors are exposed as `comments` and `comment-bodies` fields, and
comment body IDs such as `C001.C0001` or `comment:3` can be used as metadata-only
targets. Use `--radius` to include same-kind neighbors and raise `--max-text` only
when short snippets are needed.
Read text views are `final` (default), `original`, and `markup`. Markup view includes inserted and deleted text with lightweight `[+text+]` and `[-text-]` markers.
`read --summary` prints package/story counts without listing every target, which is
useful for large-document validation.
Paragraph lines may include `styleId=...` and resolved list metadata. List metadata
starts with the concrete `numId` and zero-based level, then includes resolved
`abstractNumId`, numbering `format`, `level-text`, paragraph style link, and
`source=style` or `source=style-inherited` when the list comes from paragraph style
inheritance rather than direct paragraph numbering. Deterministic labels are exposed
as `label="..."` with `label-components` entries in `level:value:format:text` form
and `label-status=resolved|partial|unsupported`; `start`, `suffix`, `legal=true`,
`restart-after-level`, and `label-warnings` appear when they are known or needed.
Supported label formats include decimal, zero-padded decimal, upper/lower letters,
upper/lower roman numerals, bullets, and nested `lvlText` tokens whose referenced
counters are known. The same compact list metadata appears on numbered heading
`outline` lines and paragraph `find` matches.
`styles` output includes inheritance links such as `based-on`, `next`, `linked`, and
style-level numbering defaults when present.
`read` and `outline` list bookmark and content-control metadata when present.
Bookmark records include name, OOXML ID, story, part, start/end targets, duplicate
name candidate IDs, and whether the range is complete. Content-control records
include kind, tag, alias, placeholder/data-binding metadata, parent/child control
IDs, safe-edit status, duplicate tag/alias candidate IDs, lock, story, part,
containing target, text length, checkbox state, dropdown/combo item count,
repeating-section metadata, and date settings when present. `context` annotates nearby targets with
`bookmark-names`, `content-controls`, `content-control-tags`, and
`content-control-aliases` so agents can connect selector names to stable target IDs.
If `bookmark:"Name"` or `content-control:"TagOrAlias"` is ambiguous, selector
diagnostics list candidate paragraph IDs; retry with an explicit ID.
Patch operations can also target explicit bookmark/content-control IDs for safe shapes:
`set-content-control-text` updates plain-text controls while preserving the wrapper,
`set-content-control-checkbox` toggles checkbox controls and updates their displayed
state symbol, `set-content-control-choice` selects dropdown/combo items by `value` or
`display-text`, `set-content-control-date` updates date control `fullDate` values and
displayed text, and `replace-bookmark-text` updates simple same-paragraph bookmark
ranges while preserving the bookmark markers. `rename-bookmark` renames bookmark
markers and same-story internal hyperlink anchors when the old name is unambiguous;
`delete-bookmark` removes only complete unreferenced bookmark markers.
`read` and `outline` also list simple and complex field metadata. Field records
include kind, field code, containing target, result text length, dirty/lock flags,
and whether a complex begin/separate/end sequence is complete. `context` annotates
nearby targets with `fields`, `field-codes`, and `field-kinds`. DocxEdit preserves
field XML, can set field dirty/lock flags through `set-field-dirty` and
`set-field-lock` for one field or `target all`, and can update simple `w:fldSimple`
field code/result caches through `set-field-code` and `set-field-result`. DocxEdit
can mark the document for field updates after edits through
`MarkFieldsDirtyWhenEditing`; Word remains responsible for recalculating field
results.
Hyperlink records are listed by `read` and `outline`. Relationship-backed hyperlinks
expose relationship ID, relationship part, target mode, URI or target part, scheme,
validation status, and validation reason for unsupported schemes or
relative/malformed targets. Valid patch URI targets are absolute `http`, `https`,
or `mailto`; relative, malformed, `file`, UNC/file-style, and unsafe-scheme targets
are rejected. Internal anchor links expose missing and duplicate bookmark-anchor
flags. Relationship-backed internal part links expose `target-part` metadata and
emit `W1023` because patch edits support external URI or anchor targets only.
Hyperlinks expose tooltip, target-frame, and history metadata when present. Broken
relationship IDs are flagged.
`read` also aggregates invalid URI, missing-anchor, and duplicate-anchor diagnostics.
`context` annotates nearby targets with `hyperlinks` and `hyperlink-targets`; `dump --runs` annotates hyperlink runs with
`markup=hyperlink`, `hyperlink-relationship-id`, and `hyperlink-anchor` when present.
Hyperlink patch operations can update URI/anchor targets, tooltip, target-frame,
history, display text, insert new hyperlink paragraphs after a target, and remove
hyperlink markup while preserving the displayed runs.
Image records from `read`, `outline`, and `media` include `layout=inline` or
`layout=anchor`, the media relationship ID, containing paragraph/cell target,
DrawingML extent in EMUs, `docPr` description/title/name when present, wrap mode,
`behind-doc` for floating drawings, wrap distances, anchor relative positioning,
relative height, overlap/aspect-lock flags, and `crop-left-percent`/
`crop-top-percent`/`crop-right-percent`/`crop-bottom-percent` when DrawingML crop
metadata is present. Replacement, alt-text, and docPr metadata operations preserve
the existing drawing layout where supported. `set-image-size` updates DrawingML
extents, `set-image-wrap` updates anchored image wrap mode/distances, and
`set-image-position` updates anchored relative positions, offsets, and alignments.
`set-image-crop` updates DrawingML crop percentages; unsupported drawing shapes remain
preserve-only. Linked images are never fetched and are omitted from editable image
records; VML, grouped drawings, charts, SmartArt, OLE objects, equations, and generic
shapes are reported as diagnostics and preserved.
Table output includes table style ID, declared grid column count, header-row,
merged-cell, and nested-table flags when present. Row lines include physical cell
count, omitted grid columns (`grid-before`/`grid-after`), header status, and
`cant-split`. Cell lines include logical target column plus `physical-column`, span,
visual column end, merge group ID, vertical-merge root cell, vertical merge, and
nested-table metadata. Row operations still reject unsafe
non-rectangular tables unless `force true` is explicitly supplied. `set-table-style`
updates `w:tblStyle`; `set-row-header` sets or clears the row repeating-header flag.
`validate` supports `--profile structural|package`. `structural` is the default and
runs bounded structural package checks plus WordprocessingML invariants: known part
roots, paired bookmark/comment ranges, commentsExtended paraId consistency,
duplicate semantic selectors, complex field begin/end balance, field result
containment, drawing relationship references, image target/content-type checks,
duplicate drawing property IDs,
drawing extent/crop geometry, basic table row/cell shape, and table visual-grid
consistency. `package` limits validation to package/XML root checks. This is not full
ISO/IEC 29500 schema validation; it is intended to catch common corruption and
relationship mistakes with stable diagnostics such as `E9103`, `E9104`, `E9105`,
`E9106`, `E9107`, `E9108`, `E9109`, `E9110`, `E9113`, `E9114`, and `W9109`.

## Patch

- `check input.docx edits.docxpatch [--track-changes mode] [--author name] [--timestamp-utc instant] [--json] [--report path] [--diagnostics path] [--strict]`
- `apply input.docx edits.docxpatch --output output.docx [--track-changes mode] [--author name] [--timestamp-utc instant] [--json] [--report path] [--diagnostics path] [--strict]`

Track-change modes are `off`, `preserve`, `suggest`, and `require`. Supported tracked output includes simple `replace-text`, whole-paragraph replacement, inserted/deleted paragraph text, paragraph style changes, and simple single-paragraph table-cell text replacement. Tracked text output is limited to shapes without tabs or line breaks, protected OOXML boundaries, existing revision markup, or mixed direct run formatting. `require` fails preserve-only operations with `E6001` and unsupported tracked shapes with `E6002`; `suggest` warns with `W4001` or `W4002` and applies the direct edit.
`docxedit help patch` includes a track-change support table generated from
`DocxHelp.Catalog`, including operation-specific support values such as
`tracked-simple`, `tracked-paragraph`, `tracked-style`, `tracked-cell-simple`,
`preserve-only`, and `unsupported`. `preserve-only` means existing revision markup is
preserved but the operation does not create new revision markup; `suggest` applies
directly with `W4001`, and `require` fails with `E6001`. The `W4001` and `E6001`
messages include the catalog support value.
Plain text `check` and `apply` output includes one `operation index=...` line per
patch operation with operation name, target, and success. Table operations also emit
`affected id=...` row/cell lines with action, parent, row/column, and row-count
metadata. Use `--report` for the full JSON operation report.

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
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- help validate
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- help check
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- help apply
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- help patch
```

Examples use `--output` for clarity; `-o` remains supported as a short alias.
