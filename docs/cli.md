# CLI

`docxedit` is the command-line surface for creating and inspecting a `.docx`, locating stable
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

Examples use `docxedit` as the command name. No tool package is published:
run the CLI from source (integrations embed the library instead):

```text
dotnet run --project src/Lokad.DocxEdit.Cli/Lokad.DocxEdit.Cli.csproj -- read report.docx --summary
```

## Standard Input And Output

The input `.docx`, the patch file, and `--output` accept `-` for stdin/stdout
(UTF-8 for the patch; raw bytes otherwise), so chained runs need no temp files:

```text
docxedit check report.docx edits.docxpatch --report check.json
docxedit apply report.docx edits.docxpatch --output - > report.edited.docx
docxedit validate - < report.edited.docx
```

A patch read from stdin has no directory: relative image `asset` paths resolve
against the invoking working directory only (see [patch-format.md](patch-format.md)).
`--report`, `--diagnostics`, and `--operation-report` stay file-only; `-` is not
accepted there because it would collide with `--json` or binary stdout.

Binary-stdout ownership: with `create --output -` or `apply --output -`, standard output carries
exactly the generated package and nothing else. Status lines and the operation
summary go to standard error (use `--report <path>` for the machine-readable
report); `--json` is rejected with `--output -`. The input document and the
patch file cannot both use `-`. A failed binary run writes no bytes to
standard output. `--output`/`--report`/`--diagnostics` must differ from the
input document, the patch file, and each other, failing with exit code `2`
before any writer opens.

## Creating A Document

```text
docxedit create --output report.docx
docxedit create --output report.docx --paper a4 --orientation landscape
docxedit create --output report.docx --paper letter --orientation portrait
```

`create` takes no input document or positional arguments. `--output` (alias `-o`) is
required. Paper values are `a4` and `letter`; orientation values are `portrait` and
`landscape`. Defaults are A4 portrait, independent of locale. Margins are one inch
(1440 twips), header/footer distances are half an inch (720 twips), and the gutter
is zero. Landscape swaps page width and height and sets the orientation attribute.

The generated document has one empty paragraph (`M.P0001`) and one section
(`M.S0001`). Use `replace-paragraph` to fill that paragraph and `insert-after` to
add more content through the usual patch workflow. Styles include Normal and
Heading1–Heading9. Body defaults use Arial 11pt, single line spacing, and 8pt after
paragraphs. Headings are bold with explicit outline levels, 12pt before and 6pt
after; sizes are 16pt, 14pt, 12pt, then 11pt for levels 4–9.

The complete package is staged and validated before publication. Successful creation
replaces an existing destination; failed creation preserves it. `--json`, `--compact`,
`--report`, `--diagnostics`, and `--strict` follow the existing command conventions.
Report and diagnostics paths must differ from the document and each other.
`--output -` writes binary stdout, with status on stderr; use a file report instead of
`--json` in that mode.

## Section Layout

`read` exposes section `PageWidthTwips`, `PageHeightTwips`, `MarginTopTwips`,
`MarginBottomTwips`, `MarginLeftTwips`, `MarginRightTwips`, `HeaderDistanceTwips`,
`FooterDistanceTwips`, and `GutterTwips` in JSON. One twip is 1/1440 inch. Text output
uses corresponding names such as `page-width-twips` and `margin-top-twips`.
Missing or non-integer values are null in JSON and omitted in text. Signed top and
bottom margins are preserved. These are stored section properties, with no inferred
layout defaults or paper preset for existing documents.

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
docxedit capabilities report.docx --id M.P0004
docxedit template report.docx --id M.P0004
```

Write an `edits.docxpatch`, lint it without a document, then validate it before producing a new document:

```text
docxedit lint edits.docxpatch
docxedit check report.docx edits.docxpatch
docxedit apply report.docx edits.docxpatch --output report.edited.docx
docxedit validate report.edited.docx
```

Use `--json` when another program will consume the result. Add `--compact` to shrink JSON output. Use plain text when an
agent or human needs compact context. Keep token cost down with `read --summary`, `context --max-text 0` (the default), and `--compact` for machine reads.

## Commands

| Command | Purpose | Common use |
| --- | --- | --- |
| `create` | Empty document creation | A4 or US Letter, portrait or landscape; one paragraph and one section |
| `read` | Structural document view | Broad inventory of paragraphs, tables, images, fields, links, bookmarks, content controls, sections |
| `outline` | Compact navigational view | Headings, tables, images, sections, headers, footers |
| `find` | Text search | Locate stable paragraph or cell targets from visible text |
| `dump` | Detailed target view | Inspect one target, optionally with run-level markup |
| `context` | Nearby target context | Inspect neighbors and attached metadata without broad text |
| `capabilities` | Target editing capabilities | Supported, conditional, and unsupported edits for one target under a track-change policy |
| `template` | Guarded patch template | Check-clean starter plus commented examples for one target |
| `lint` | Patch shape validation | Syntax, required fields, and field groups without a document |
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
| `--compact` | JSON-emitting commands | Emit JSON without indentation (stdout `--json` plus `--report`/`--diagnostics` files) |
| `--diagnostics path` | Most commands | Write diagnostics JSON to a separate file |
| `--strict` | Most commands | Return exit code `3` when warnings are present |
| `--view final/original/markup` | Text reads and outline | Select how tracked inserted/deleted content is rendered |
| `--max-text N` | Text reads | Limit body text per field in text and structured/JSON output (0 drops; IDs, counts, and structural metadata retained) |
| `--headers-footers` | Read commands and `media` | Include modeled header and footer stories |
| `--report path` | `check`, `apply` | Write full operation report JSON |
| `--track-changes mode` | `check`, `apply`, `capabilities`, `template` | Control generated revision markup (effective policy for capabilities and templates) |
| `--author name` | `check`, `apply` | Author used for generated revisions |
| `--timestamp-utc instant` | `check`, `apply` | ISO-8601 UTC timestamp for generated revisions (e.g. `2026-01-01T00:00:00Z`) |
| `--max-preview-chars N` | `check`, `apply` | Bounded before/after preview text per operation report side (0 disables previews and keeps reports metadata-only) |
| `--extract dir` | `media` | Extract embedded image parts to a directory |
| `--id <image-id>` | `media` | Select a single discovered image placement for listing or extraction |
| `--radius N` | `context` | Number of same-kind neighbors to include |
| `--max-diagnostics N` | `validate` | Maximum diagnostics to return; default 500 |

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

- `read --summary` prints package and story counts without listing every target; with `--json` it emits a counts-only object (no target lists, no text).
- `apply` writes the edited document to a temporary sibling and replaces `--output` only after successful editing; failed runs preserve any pre-existing destination file. `--output`/`--report`/`--diagnostics` must differ from the input document, the patch file, and each other, failing with exit code `2` otherwise.
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
- `--max-text 0` (and the `ReadSummary`/`ContextMetadataOnly` presets) drops
  paragraph/run/cell text, field cached results, and table/image
  captions/descriptions/titles in text and structured/JSON output. IDs, counts,
  style IDs/names, bookmark names, content-control tags/aliases, field
  codes/kinds/types, hyperlink URIs/anchors, authors, and revision IDs are
  retained as structural metadata and never imply redaction of those fields.

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
| Tables | `read`, `context` | Cell IDs use visual grid coordinates; a multi-paragraph cell reads and finds as flat concatenated text with no boundary marker; `set-cell` and `set-cell-shading` can target merge-group IDs and horizontal spans, but reject vertical-merge continuations; direct row edits can clone safe visual-grid row shapes, while tracked row revisions require simple rectangular tables |
| Sections | `read`, `outline` | Section operations target main-document section IDs |

Image identifiers name drawing placements, not media parts: repeated use of one media part yields one identifier per placement. Drawings hidden from the Final view (deleted blocks and rows, block-level structured-document-tag content) own no identifier, so every published ID addresses exactly the drawing discovery describes.

Image workflow: discover placements with `read --headers-footers` or `media --headers-footers`, extract original bytes with `media --extract <dir>` optionally narrowed by `--id` (exported filenames follow the actual media content type, exposed per file, so they feed back as asset hints), process the files externally, then bring bytes back with a guarded `replace-image` (optionally swapping PNG and JPEG, which updates the media content type) or `insert-image-after`, and confirm with a readback. Replacement affects only the selected placement: shared media is isolated to fresh parts and relationships while exclusive media is updated in place, preserving size, wrapping, position, crop, and metadata unless explicitly changed. Removing a placement stays an explicit `delete-image` operation.

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

Relative image `asset` paths in the patch resolve against the patch file directory first, then the invoking working directory (see [patch-format.md](patch-format.md)).

Both commands report one operation line per patch operation. Table operations also
report affected row/cell IDs with visual-grid metadata such as column spans,
omitted-column offsets, merge groups, and nested-table paths when relevant. Text
and structural paragraph operations report the resolved paragraph ID (or the table
anchor for table targets) with update, insert, or delete actions. Other explicit-ID mutations
report the resolved hyperlink, bookmark, content-control, field, image, section, row, cell,
or merge-group ID the same way. Creations additionally report created IDs per operation. Use
`--report path` for the full JSON report. Affected targets carry lifetime explicitly: Id keeps the historical input or operation-time identity named by Coordinate while FinalId alone carries the live final ID, staying absent for deleted objects, so a historical ordinal that coincides with a survivor never reads as live.
Generated revision IDs from the report can be correlated with
`changes --operation-report`.

Both commands echo the effective `--author`, `--timestamp-utc`, and producing-library version in the text summary line and the JSON report, so a run can be replayed and attributed from its report alone. Agent workflows should always pass explicit `--author` and `--timestamp-utc` for reproducible outputs; omitting them records the `docxedit` default author and the current UTC time. Reproducibility covers part payloads and rendered text: archive bytes still differ because ZIP entry timestamps record each save.

Track-change modes:

| Mode | Behavior |
| --- | --- |
| `off` | Apply direct edits |
| `preserve` | Apply direct edits while preserving existing tracked-change markup where possible |
| `suggest` | Generate new revision markup for supported operations; warn and apply directly for preserve-only operations or unsupported shapes |
| `require` | Require generated revision markup for content edits while permitting declared annotation (comment) operations; fail other preserve-only operations and unsupported tracked shapes |

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
- `3`: `--strict` with a successful result that carries warnings (errors fail as `1`). A published output document, if any, is still written; only the exit code signals the warnings.
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
docxedit help <operation> (for example: replace-text, set-cell, add-comment)
```

`docxedit help patch` is generated from the shared library command catalog. It is
the most compact source for exact operation fields and track-change support.

For `apply --track-changes suggest|require`, operation summaries and report JSON
include `GeneratedRevisionIds` when an operation actually creates revision markup.
Revision IDs appear only for committed output: check reports never carry them, and a failed patch reports none because nothing was published.
