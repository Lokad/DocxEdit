# CLI

`docxedit` is the command-line surface for inspecting a `.docx`, locating stable
targets, validating a `.docxpatch`, and writing a new edited `.docx`.

The normal loop is:

1. Inspect the document structure.
2. Locate one or more stable target IDs.
3. Inspect the target and nearby markup.
4. Write a `.docxpatch`.
5. Run `check`.
6. Run `apply`.
7. Validate the output.

## Running The CLI

When installed as a tool or exposed by an integration, examples use `docxedit`:

```text
docxedit read report.docx --summary
```

When running this repository from source, replace `docxedit` with:

```text
dotnet run --project src/Lokad.DocxEdit.Cli/Lokad.DocxEdit.Cli.csproj --
```

For example:

```text
dotnet run --project src/Lokad.DocxEdit.Cli/Lokad.DocxEdit.Cli.csproj -- read report.docx --summary
```

## Recommended Workflow

Start with low-text inspection. This gives counts, IDs, diagnostics, and markup
metadata without dumping the whole document:

```text
docxedit read report.docx --summary
docxedit validate report.docx
docxedit changes report.docx
```

Find or inspect the target:

```text
docxedit find report.docx "old wording"
docxedit dump report.docx --id M.P0004 --runs
docxedit context report.docx --id M.P0004
```

Write an `edits.docxpatch`, then validate it before producing a new document:

```text
docxedit check report.docx edits.docxpatch
docxedit apply report.docx edits.docxpatch --output report.edited.docx
docxedit validate report.edited.docx
```

Use `--json` when another program will consume the result. Use plain text when an
agent or human needs compact context.

## Commands

| Command | Purpose | Common use |
| --- | --- | --- |
| `read` | Structural document view | Broad inventory of paragraphs, tables, images, fields, links, bookmarks, content controls, sections |
| `outline` | Compact navigational view | Headings, tables, images, sections, headers, footers |
| `find` | Text search | Locate stable paragraph or cell targets from visible text |
| `dump` | Detailed target view | Inspect one target, optionally with run-level markup |
| `context` | Nearby target context | Inspect neighbors and attached metadata without broad text |
| `styles` | Style inventory | Discover valid paragraph, character, and table styles |
| `media` | Image inventory and extraction | List or extract embedded image parts |
| `changes` | Existing markup inventory | Track changes, comments, anchors, and comment bodies without text by default |
| `validate` | Package and structural checks | Detect common corruption, bad references, malformed markup, and table shape issues |
| `check` | Dry-run patch validation | Validate selectors, guards, assets, and track-change constraints |
| `apply` | Patch application | Write a new edited `.docx`; the input is not modified |
| `help` | Built-in command guidance | Show command-specific or patch-operation help |
| `catalog` | Machine-readable surface | Commands, options, output fields, and patch operations as text or `--json` |
| `version` | Build version | Assembly version, framework, and patch-operation count as text or `--json` |

## Common Options

| Option | Applies to | Meaning |
| --- | --- | --- |
| `--json` | Most commands | Emit the structured result object as JSON |
| `--diagnostics path` | Most commands | Write diagnostics JSON to a separate file |
| `--strict` | Most commands | Return exit code `3` when warnings are present |
| `--view final/original/markup` | Text reads and outline | Select how tracked inserted/deleted content is rendered |
| `--max-text N` | Text reads | Limit rendered text per field |
| `--headers-footers` | Read commands | Include modeled header and footer stories |
| `--all-stories` | `read` | Include every modeled story |
| `--report path` | `check`, `apply` | Write full operation report JSON |
| `--track-changes mode` | `check`, `apply` | Control generated revision markup |
| `--author name` | `check`, `apply` | Author used for generated revisions |
| `--timestamp-utc instant` | `check`, `apply` | ISO-8601 UTC timestamp for generated revisions (e.g. `2026-01-01T00:00:00Z`) |

Text views:

- `final`: default; shows the accepted visible text.
- `original`: shows text before tracked insertions/deletions.
- `markup`: includes inserted and deleted text with lightweight markers such as
  `[+inserted+]` and `[-deleted-]`.

For numbered paragraphs, list counters are computed from the paragraphs visible
in the selected view. Block-level inserted numbered paragraphs participate in
`final` and `markup` counters; block-level deleted numbered paragraphs
participate in `original` and `markup` counters. If paragraph property revisions
carry previous `w:numPr` numbering state, DocxEdit emits `W1026` because
original-view reconstruction of the previous numbering properties is not modeled.

## Target IDs

Most edits should use explicit IDs from `read`, `outline`, `find`, `dump`, or
`context`.

Common IDs:

- `M.P0001`: main-document paragraph.
- `H001.P0001` / `F001.P0001`: header or footer paragraph.
- `M.T0001`: table.
- `M.T0001.R02`: table row.
- `M.T0001.R02.C03`: table cell using visual grid coordinates.
- `M.T0001.MG0001`: horizontal or vertical merge-group root cell.
- `M.I0001`: image.
- `M.S0001`: section.
- `M.B0001`: bookmark.
- `M.CC0001`: content control.
- `M.F0001`: field.
- `M.L0001`: hyperlink.
- `C001.C0001` or `comment:3`: comment body target.

Semantic selectors such as `heading:"Exact heading"`,
`bookmark:"BookmarkName"`, and `content-control:"TagOrAlias"` are documented in
[patch-format.md](patch-format.md). Prefer explicit IDs after discovery,
especially when selector diagnostics report duplicates.

## Privacy Defaults

DocxEdit is designed so an agent can inspect document shape before exposing text.

- `read --summary` prints package and story counts without listing every target.
- `context` defaults to `--max-text 0`; paragraph and cell text fields are present
  but empty.
- `changes` does not print revision text or comment body text by default.
- `dump --id C001.C0001` and `dump --id comment:3` print comment body metadata,
  not comment body text.
- Metadata-only `context` surfaces threaded-comment IDs, durable IDs,
  parent/root para IDs, reply IDs, and resolved IDs without printing comment body
  text.
- Comment snippets require `changes --include-comment-text`; bound them with
  `--max-comment-text N`.

For private inputs, start with `read --summary`, `validate`, `changes`, and
metadata-only `context`.

## Existing Markup

Use `changes` first when a document may contain tracked changes or comments:

```text
docxedit changes report.docx
docxedit changes report.docx --json
```

`changes` reports summary counts, group summaries, target summaries, comment
summaries, and individual markup records. Records include IDs, type, story, part,
normalized parent type, author/timestamp metadata, target IDs when known, text
length, and child element counts.

When a check/apply JSON report is available, pass it back to `changes` to link
generated revision IDs to operation metadata:

```text
docxedit changes report.edited.docx --operation-report apply-report.json
```

Matching records include `operation-index`, `operation-name`, and
`operation-target` in text output, and `OperationIndex`, `OperationName`, and
`OperationTarget` in JSON.

Target fields help interpret sparse OOXML markup:

- `target-status`: whether the record is targeted, attached by a comment anchor,
  or targetless.
- `target-source`: whether the target came from an ancestor, adjacent range,
  comment anchor, or no source.
- `paired-change-id`: links related range start/end records.
- `nearest-target`: nearby context only; not exact ownership.

Use `dump --runs` on the target to see run-level annotations:

```text
docxedit dump report.docx --id M.P0004 --runs --view markup
```

Run lines expose markup such as inserted/deleted runs, revision IDs, authors,
timestamps, comment IDs, comment ranges, and hyperlink annotations. In JSON output,
the same data is in `Runs`.
When the target owns tracked markup records, `dump` appends a privacy-safe
`changes:` block with change IDs, types, parent type, revision metadata, and
child element counts. This exposes property revisions on paragraph, table, row,
cell, and section targets without printing raw OOXML or broad document text.

## Feature Discovery

| Feature | Inspect with | Editing notes |
| --- | --- | --- |
| Lists and numbering | `read`, `outline`, `find` | Paragraphs may include resolved list labels, numbering format, level text, style-linked numbering, view-aware tracked paragraph counters, and diagnostics for unsupported custom formats or tracked numbering property revisions |
| Bookmarks | `read`, `context` | Use bookmark IDs or unambiguous `bookmark:"Name"` selectors; duplicate names are diagnosed; table-spanning replacements require one replacement line per visible text slot |
| Content controls | `read`, `context` | Metadata includes kind, tag, alias, lock state, safe-edit status/reason, checkbox/dropdown/date details, hierarchy IDs, and duplicate selector candidates |
| Comments | `changes`, `context`, `dump` | Comment operations target a paragraph, `comment:<id>`, or a comment body ID; context exposes thread topology as IDs without comment body text |
| Fields | `read`, `outline`, `context` | Field metadata includes parsed arguments/switches and refresh policy; DocxEdit can refresh REF-style bookmark fields and QUOTE literal fields, but Word remains responsible for general recalculation |
| Hyperlinks | `read`, `outline`, `context`, `dump --runs` | External patch targets must be absolute `http`, `https`, or `mailto`; internal targets use bookmark anchors |
| Images | `read`, `outline`, `media` | Editable image records are inline or anchored DrawingML images; linked images and complex drawing shapes are preserve-only diagnostics |
| Tables | `read`, `context` | Cell IDs use visual grid coordinates; `set-cell` and `set-cell-shading` can target merge-group IDs and horizontal spans, but reject vertical-merge continuations; direct row edits can clone safe visual-grid row shapes, while tracked row revisions require simple rectangular tables |
| Sections | `read`, `outline` | Section operations target main-document section IDs |

For a full inventory of supported and unsupported shapes, see
[status.md](status.md).

## Check And Apply

`check` validates a patch without writing output:

```text
docxedit check report.docx edits.docxpatch
```

`apply` writes a new `.docx`:

```text
docxedit apply report.docx edits.docxpatch --output report.edited.docx
```

Both commands report one operation line per patch operation. Table operations also
report affected row/cell IDs with visual-grid metadata such as column spans,
omitted-column offsets, merge groups, and nested-table paths when relevant. Use
`--report path` for the full JSON report.
Generated revision IDs from the report can be correlated with
`changes --operation-report`.

Track-change modes:

| Mode | Behavior |
| --- | --- |
| `off` | Apply direct edits |
| `preserve` | Apply direct edits while preserving existing tracked-change markup where possible |
| `suggest` | Generate new revision markup for supported operations; warn and apply directly for preserve-only operations or unsupported shapes |
| `require` | Require generated revision markup; fail preserve-only operations and unsupported tracked shapes |

Generated tracked output is intentionally narrow. It supports simple text
replacement, whole-paragraph replacement, inserted/deleted paragraph text,
paragraph style changes, simple plain-text content controls, simple hyperlink
display text, simple field result text, simple same-paragraph bookmark text,
guarded paragraph-only rich-text content controls, and simple table-cell
replacement including compatible multi-paragraph text-only cells. Simple
paragraph-only comment body replacement can also emit tracked revisions in
`comments.xml`, and simple inserted hyperlink display text can be wrapped in
tracked insertion markup. Table style and row-header updates emit property
revision markup. Simple rectangular row insertion and deletion emit row
structure revisions through `w:trPr/w:ins` and `w:trPr/w:del`; complex or
forced row edits fall back with `W4002` or fail with `E6002`. Section column and
orientation updates emit section property revision markup unless the section
already contains `w:sectPrChange`; in that case `suggest` falls back to a direct
edit with `W4002`, and `require` fails with `E6002`.
Unsupported tracked shapes produce `W4001`, `W4002`, `E6001`, or `E6002`
depending on the selected mode.

## Validation And Exit Codes

Use `validate` before and after editing:

```text
docxedit validate report.docx
docxedit validate report.edited.docx --profile structural
```

Profiles:

- `structural`: default; package/XML root checks plus common WordprocessingML
  invariants.
- `package`: package/XML root checks only.

Validation is not full ISO/IEC 29500 schema validation. It is a bounded set of
stable checks for package roots, relationships, comments, fields, content
controls, drawings, headers/footers, sections, numbering references, and table
shape.

Exit codes:

- `0`: success.
- `1`: operation failed.
- `2`: invalid CLI usage.
- `3`: `--strict` with a successful result that carries warnings (errors fail as `1`).
- `4`: unexpected CLI exception.

In text mode, error diagnostics print to standard error (warnings too under `--strict`); `--json` carries diagnostics in the result object. Diagnostics are documented in [diagnostics.md](diagnostics.md). Validation details
are documented in [validation.md](validation.md).

## Built-In Help

The CLI includes task-specific help:

```text
docxedit --help
docxedit help dump
docxedit help context
docxedit help changes
docxedit help validate
docxedit help check
docxedit help apply
docxedit help catalog
docxedit help patch
```

`docxedit help patch` is generated from the shared library command catalog. It is
the most compact source for exact operation fields and track-change support.

For `apply --track-changes suggest|require`, operation summaries and report JSON
include `GeneratedRevisionIds` when an operation actually creates revision markup.
