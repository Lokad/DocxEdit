# Patch Format

`.docxpatch` is DocxEdit's line-based DSL for Word edits. A patch names one or
more operations, each operation targets a stable ID or selector discovered through
the CLI, and guarded fields verify that the document still has the expected shape
before anything is written.

Use [cli.md](cli.md) to discover targets and run `check` before `apply`.

## File Shape

Every patch starts with `docxpatch 1`:

```text
docxpatch 1

# Blank lines and comment lines are ignored.
op replace-text
target M.P0004
expect-text <<<
The current paragraph text.
>>>
find current
with revised
end
```

Rules:

- The first non-empty line must be `docxpatch 1`.
- Each operation starts with `op <name>` and ends with `end`.
- Field lines use `field value`.
- Multi-line values use heredocs:

  ```text
  text <<<
  First line.
  Second line.
  >>>
  ```

- Field names are operation-specific. Unknown fields fail validation.
- Boolean fields must be `true` or `false`.
- Integer fields must parse as integers.
- Repeated fields are preserved for operations that need them, such as repeated
  `cell` fields in row operations.
- `expect-hash` and `preserve-size` are not supported.

## Minimal Examples

Replace text inside one paragraph:

```text
docxpatch 1

op replace-text
target M.P0004
expect-text <<<
The paragraph as it currently appears.
>>>
find currently appears
with now reads
end
```

Replace one table cell with row and column guards:

```text
docxpatch 1

op set-cell
target M.T0001.R02.C03
expect-row-count 4
expect-column-count 3
expect-text <<<
old cell text
>>>
text <<<
updated cell text
>>>
end
```

Add a paragraph-level comment:

```text
docxpatch 1

op add-comment
target M.P0004
expect-text <<<
Reviewed paragraph text.
>>>
text <<<
Please verify this statement.
>>>
author Reviewer
end
```

Anchor a comment to one phrase inside a paragraph:

```text
docxpatch 1

op add-comment
target M.P0004
expect-text Reviewed paragraph text with repeated phrase.
anchor-text repeated phrase
occurrence 1
text Please verify this phrase.
author Reviewer
end
```

Replace an image while preserving its existing supported drawing layout:

```text
docxpatch 1

op replace-image
target M.I0001
asset chart.png
expect-content-type image/png
alt Updated chart
end
```

## Selectors

Explicit IDs are the safest selectors:

- `M.P0001`: main paragraph.
- `H001.P0001` / `F001.P0001`: header or footer paragraph.
- `M.T0001`: table.
- `M.T0001.R02`: row.
- `M.T0001.R02.C03`: cell.
- `H001.T0001.R02.C03` / `F001.T0001.R02.C03`: header or footer cell.
- `M.I0001`: image.
- `H001.I0001` / `F001.I0001`: header or footer image.
- `M.S0001`: section.
- `M.B0001`: bookmark.
- `M.CC0001`: content control.
- `M.F0001`: field.
- `M.L0001`: hyperlink.
- `C001.C0001` or `comment:3`: comment body.

Paragraph operations also support semantic selectors:

- `heading:"Exact heading"`.
- `heading:2:"Exact heading"`.
- `text:"contained paragraph text"`.
- `bookmark:"BookmarkName"`.
- `content-control:"TagOrAlias"`.

Ambiguous selectors fail with `E1202`; use an explicit ID from `read`,
`context`, or `find`. Selectors with no match fail with `E1201`.

## Guards

Guard fields make patches reviewable and safer to rerun. Prefer them whenever the
current content or structure is known.

| Guard | Typical operations | Meaning |
| --- | --- | --- |
| `expect-text` | Paragraphs, cells, comments, content controls | Visible text must match before editing |
| `expect-row-count` | Table and row edits | Target table must have the expected number of rows |
| `expect-column-count` | Table and row edits | Target table must have the expected logical column count |
| `expect-cell-count` | Row insert/delete | Target row must have the expected physical cell count |
| `expect-contains` | `delete-row` | Row text must contain the exact guard value |
| `expect-fill` | `set-cell-shading` | Current cell shading fill must match a hex color, `auto`, or `none` |
| `expect-style` | `set-table-style` | Current table style must match |
| `expect-caption` | `set-table-metadata` | Current table caption must match; missing and empty are equivalent |
| `expect-description` | `set-table-metadata` | Current table description must match; missing and empty are equivalent |
| `expect-header` | `set-row-header` | Current repeating-header flag must match |
| `expect-content-type` | Image operations | Current media content type must match |
| `expect-code` | Field code edits | Normalized field code must match |
| `expect-result` | Field result edits | Cached field result must match |
| `expect-columns` | Section edits | Current section column count must match |
| `expect-orientation` | Section edits | Current section orientation must match |

Guard failures are reported as `E32xx` diagnostics.

## Operation Reference

The built-in reference is always available:

```text
docxedit help patch
```

The tables below group the same operations by editing area.

### Paragraphs And Blocks

| Operation | Required fields | Optional fields | Notes |
| --- | --- | --- | --- |
| `replace-text` | `target`, `find`, `with` | `expect-text`, `preserve-runs`, `occurrence` | Replaces matching text inside one target |
| `replace-paragraph` | `target`, `text` | `expect-text`, `style` | Replaces the paragraph text, optionally setting style |
| `insert-before` | `target`, `text` | `style`, `copy-paragraph-properties` | Inserts a paragraph/block before the target |
| `insert-after` | `target`, `text` | `style`, `copy-paragraph-properties` | Inserts a paragraph/block after the target |
| `delete-block` | `target` | `expect-text` | Deletes the target block |
| `set-style` | `target`, `style` | | Sets paragraph style |

For list-like insertions, use `copy-paragraph-properties true` with a paragraph
target. The new paragraph copies the target paragraph properties, including list
numbering, while dropping copied `w:pPrChange` and `w:sectPr`. An explicit
`style` overrides only the copied paragraph style.

### Content Controls

| Operation | Required fields | Optional fields | Notes |
| --- | --- | --- | --- |
| `set-content-control-text` | `target`, `text` | `expect-text` | Plain-text controls are supported; guarded simple rich-text controls are supported when safe; picture/group controls fail with kind-specific guidance |
| `set-content-control-checkbox` | `target`, `checked` | | Updates checkbox state and displayed symbol |
| `set-content-control-choice` | `target` plus `value` or `display-text` | | Selects a dropdown/combo item |
| `set-content-control-date` | `target`, `value` | `display-text` | Updates date value and visible text |
| `add-repeating-section-item` | `target` | `source`, `index`, `text` | Recognized but fails with `E4315` |
| `delete-repeating-section-item` | `target` | `index` | Recognized but fails with `E4315` |

Content-control edits preserve the `w:sdt` wrapper and metadata when supported.
Controls with a lock value other than unlocked are rejected.

### Bookmarks

| Operation | Required fields | Optional fields | Notes |
| --- | --- | --- | --- |
| `add-bookmark` | `target`, `name` | `expect-text` | Creates a guarded paragraph bookmark |
| `replace-bookmark-text` | `target`, `text` | | Replaces a complete paragraph-bounded bookmark range; simple table-spanning ranges require one replacement line per visible text slot |
| `rename-bookmark` | `target`, `name` | | Renames markers and same-story internal hyperlink anchors when unambiguous |
| `delete-bookmark` | `target` | | Removes complete unreferenced bookmark markers, preserving content |

Bookmark names must be non-empty and contain no whitespace. Duplicate new names
are rejected.

### Comments

| Operation | Required fields | Optional fields | Notes |
| --- | --- | --- | --- |
| `add-comment` | `target`, `text` | `expect-text`, `anchor-text`, `occurrence`, `author`, `initials`, `date` | Anchors a new comment to a modeled paragraph, or to one selected text span inside it |
| `set-comment-text` | `target`, `text` | | Replaces one comment body |
| `resolve-comment` | `target` | | Creates or updates modern resolution metadata for basic comments |
| `reopen-comment` | `target` | | Clears modern resolution metadata for basic comments |
| `delete-comment` | `target` | | Removes body, range/reference markers, and matching extension records |
| `add-comment-reply` | `target`, `text` | `author`, `initials`, `date` | Adds a modern threaded reply under a comment |
| `delete-comment-reply` | `target` | | Removes a leaf threaded reply |

Existing comment operations target `comment:<id>` from `changes` or a comment body
ID such as `C001.C0001`. For `add-comment`, `anchor-text` selects one normalized
text span inside the target paragraph; specify `occurrence` when that text is
repeated. Selected ranges are intentionally limited to direct paragraph text runs
and fail on protected markup boundaries. `add-comment-reply` creates modern
`commentsExtended.xml` and `commentsIds.xml` records. `delete-comment-reply`
accepts `comment:<parent-id>.reply:<ordinal>` or an explicit reply comment target;
it fails with `E4314` when the reply has child replies.

### Fields

| Operation | Required fields | Optional fields | Notes |
| --- | --- | --- | --- |
| `set-field-dirty` | `target`, `dirty` | | `target` can be a field ID or `all` |
| `set-field-lock` | `target`, `locked` | | `target` can be a field ID or `all` |
| `set-field-code` | `target`, `code` | `expect-code` | Simple `w:fldSimple` fields only |
| `set-field-result` | `target`, `text` | `expect-result` | Simple `w:fldSimple` cached result or validated simple same-paragraph complex result |
| `refresh-field-result` | `target` | `expect-code`, `expect-result` | Limited refresh for simple REF/PAGEREF/NOTEREF bookmark fields and simple QUOTE literal fields |

Complex field code edits and unsafe complex-field result topologies fail with
`E4313`. General field recalculation is Word's responsibility. Apply emits
`W5103` when edits mark fields for Word-side refresh. Refresh diagnostics
classify unsupported field types that require layout, document properties,
formulas, mail merge data, date/time state, conditionals, or external state.

### Hyperlinks

| Operation | Required fields | Optional fields | Notes |
| --- | --- | --- | --- |
| `set-hyperlink-target` | `target` plus `uri` or `anchor` | `tooltip`, `target-frame`, `history` | Updates external URI or internal bookmark anchor |
| `set-hyperlink-text` | `target`, `text` | | Updates visible hyperlink text |
| `insert-hyperlink-after` | `target`, `text` plus `uri` or `anchor` | `tooltip`, `target-frame`, `history` | Inserts a new hyperlink paragraph after the target |
| `remove-hyperlink` | `target` | | Removes hyperlink markup and preserves display runs |

External `uri` values must be absolute `http`, `https`, or `mailto` URIs.
Relative targets, malformed URIs, `file`, UNC/file-style paths, and unsafe schemes
are rejected. Internal links use bookmark `anchor` values.

### Tables

| Operation | Required fields | Optional fields | Notes |
| --- | --- | --- | --- |
| `set-cell` | `target`, `text` | `expect-text`, `expect-row-count`, `expect-column-count`, `force` | Replaces one modeled cell by visual cell ID or merge-group ID |
| `set-cell-shading` | `target` plus `fill` or `clear true` | `expect-fill` | Sets or clears `w:tcPr/w:shd` fill |
| `set-table-style` | `target`, `style` | `expect-style` | Updates `w:tblStyle` |
| `set-table-metadata` | `target` plus `caption` or `description` | `expect-caption`, `expect-description` | Sets or clears table caption/description |
| `set-row-header` | `target`, `header` | `expect-header` | Sets or clears the repeating-header flag |
| `append-row` | `target`, repeated `cell` | `expect-row-count`, `expect-column-count` | Appends by cloning the last row shape when the visual grid is consistent |
| `insert-row-before` | `target`, repeated `cell` | `expect-row-count`, `expect-column-count`, `expect-cell-count`, `force` | Inserts before a row by cloning the target row shape when safe |
| `insert-row-after` | `target`, repeated `cell` | `expect-row-count`, `expect-column-count`, `expect-cell-count`, `force` | Inserts after a row by cloning the target row shape when safe |
| `delete-row` | `target` | `expect-row-count`, `expect-column-count`, `expect-cell-count`, `expect-contains`, `force` | Deletes a row; direct mode can promote the next vertical-merge continuation |
| `append-column` | `target`, repeated `cell` | `expect-row-count`, `expect-column-count`, `force` | Recognized but fails with `E4316` |
| `insert-column-before` | `target`, `column`, repeated `cell` | `expect-row-count`, `expect-column-count`, `expect-cell-count`, `force` | Recognized but fails with `E4316` |
| `insert-column-after` | `target`, `column`, repeated `cell` | `expect-row-count`, `expect-column-count`, `expect-cell-count`, `force` | Recognized but fails with `E4316` |
| `delete-column` | `target`, `column` | `expect-row-count`, `expect-column-count`, `expect-cell-count`, `expect-contains`, `force` | Recognized but fails with `E4316` |

Table and cell IDs use visual grid coordinates from `read` or `context`, not raw
OOXML cell ordinals. `set-cell` also accepts merge-group IDs such as
`M.T0001.MG0001`; visual columns inside a horizontal span resolve to the
spanning cell, while vertical-merge continuation cells remain rejected. Table
metadata exposes spans, omitted grid columns, merge-group IDs, vertical-merge
roots, and nested-table flags so an agent can decide whether a table is safe to
edit.

`set-cell-shading` uses the same cell target forms as `set-cell`. Use `fill`
with a 6-digit hexadecimal color or `auto`, or use `clear true` to remove the
cell shading element. Under tracked changes it records the previous cell
properties with `w:tcPrChange`.

Direct row operations support consistent visual-grid tables by cloning the
template row shape, including `gridBefore`, `gridAfter`, and `gridSpan`
metadata. Insertion is rejected when the insertion boundary would cross an
active vertical merge chain. Deleting a vertical-merge root promotes the next
continuation to `restart` when the spans match. Generated tracked row revisions
remain limited to simple rectangular tables; visual-grid row edits under
`suggest` fall back with `W4002`, and `require` fails with `E6002`.

Use an empty heredoc for `caption` or `description` to remove that metadata
element:

```text
caption <<<
>>>
```

### Images

| Operation | Required fields | Optional fields | Notes |
| --- | --- | --- | --- |
| `replace-image` | `target`, `asset` | `expect-content-type`, `alt` | Replaces media bytes and preserves supported drawing layout |
| `insert-image-after` | `target`, `asset` | `expect-content-type`, `width`, `height`, `alt` | Inserts an inline image paragraph after a paragraph target |
| `set-image-alt` | `target`, `alt` | `expect-content-type` | Updates DrawingML description |
| `set-image-metadata` | `target` plus `alt`, `title`, or `name` | `expect-content-type` | Updates DrawingML `docPr` metadata |
| `set-image-size` | `target` plus `width` or `height` | `expect-content-type` | Updates DrawingML extents |
| `set-image-wrap` | `target` plus `mode` or a distance field | `expect-content-type` | Anchored images only |
| `set-image-position` | `target` plus relative, offset, or align field | `expect-content-type` | Anchored images only |
| `set-image-crop` | `target` plus one crop percentage | `expect-content-type` | Updates DrawingML crop percentages |
| `delete-image` | `target` | `expect-content-type` | Deletes the modeled image |

Image dimensions and distances accept `emu`, `in`, `cm`, `pt`, and `px` suffixes.
If only `width` or `height` is supplied, DocxEdit preserves the current aspect
ratio when it can infer one. Crop fields are percentages:
`left-percent`, `top-percent`, `right-percent`, and `bottom-percent`; opposing
side sums must remain below 100.

Linked images are not fetched or listed as editable image records. VML, grouped
drawings, charts, SmartArt, OLE objects, equations, and generic shapes are
preserved but not edited.

### Sections

| Operation | Required fields | Optional fields | Notes |
| --- | --- | --- | --- |
| `set-section-columns` | `target`, `count` | `expect-columns`, `expect-orientation` | Column count must be 1 through 4 |
| `set-section-orientation` | `target`, `orientation` | `expect-columns`, `expect-orientation` | `orientation` is `portrait` or `landscape` |

Section operations target main-document section IDs such as `M.S0001`.

## Header And Footer Targets

Explicit header/footer paragraph IDs can be used for paragraph text/style edits,
block insertion/deletion, and image insertion after the paragraph.

Explicit header/footer table IDs can be used for simple table edits and as block
insertion/deletion anchors.

Explicit header/footer image IDs can be used for image replacement, alt text,
metadata, size, crop, wrap/position where supported, and deletion.

## Track Changes

Track-change behavior is selected by `check` or `apply`, not inside the patch:

```text
docxedit apply report.docx edits.docxpatch --output report.edited.docx --track-changes require --author Agent
```

Generated revisions use author `docxedit` when `--author` is omitted. The author
must be non-empty after trimming. `--timestamp-utc` is normalized to UTC; when it
is omitted, DocxEdit uses the current UTC timestamp from the apply/check
invocation.

Generated tracked output is intentionally narrow. The support matrix below is
generated from `DocxHelp.RenderPatchTrackChangesSupportTable()` and is the same
table printed by `docxedit help patch`.

<!-- BEGIN GENERATED TRACK-CHANGES SUPPORT TABLE -->
operation | support class | support value | behavior
--- | --- | --- | ---
replace-text | text-run | tracked-simple | Suggest/Require emit tracked w:del/w:ins for supported simple text-only matches; unsupported shapes warn with W4002 or fail with E6002.
replace-paragraph | paragraph-block | tracked-paragraph | Suggest/Require emit whole-paragraph w:del/w:ins for simple text replacements and add w:pPrChange when a compatible style change is included; complex shapes warn with W4002 or fail with E6002.
insert-before | paragraph-block | tracked-paragraph-insert | Suggest/Require emit inserted paragraph text as w:ins when the inserted text has no tabs or line breaks; unsupported shapes warn with W4002 or fail with E6002.
insert-after | paragraph-block | tracked-paragraph-insert | Suggest/Require emit inserted paragraph text as w:ins when the inserted text has no tabs or line breaks; unsupported shapes warn with W4002 or fail with E6002.
delete-block | paragraph-block | tracked-paragraph-delete | Suggest/Require emit deleted paragraph text as w:del for simple paragraph targets; table/block or complex shapes warn with W4002 or fail with E6002.
set-style | paragraph-property | tracked-style | Suggest/Require emit paragraph property revisions with w:pPrChange.
set-content-control-text | text-run | tracked-content-control-text | Suggest/Require emit w:del/w:ins inside simple plain-text and guarded paragraph-only rich-text content controls while preserving wrappers, bindings, locks, and paragraph containers; complex content controls warn with W4002 or fail with E6002.
set-content-control-checkbox | preserve-only | preserve-only | Checkbox content controls update state metadata, not a simple Word revision range. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
set-content-control-choice | preserve-only | preserve-only | Dropdown and combo-box content controls update list value metadata and display text together; generated revision markup is not modeled yet. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
set-content-control-date | preserve-only | preserve-only | Date content controls update date metadata and display text together; generated revision markup is not modeled yet. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
add-repeating-section-item | unsupported | unsupported | Repeating-section item insertion is not safely modeled yet; check/apply fails with E4315.
delete-repeating-section-item | unsupported | unsupported | Repeating-section item deletion is not safely modeled yet; check/apply fails with E4315.
add-bookmark | preserve-only | preserve-only | Bookmark creation adds anchor metadata; Word has no useful generated revision range for the bookmark markers. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
replace-bookmark-text | text-run | tracked-bookmark-text | Suggest/Require emit w:del/w:ins inside simple same-paragraph bookmark ranges while preserving bookmark markers; direct mode also supports guarded multi-paragraph and simple table-spanning text-slot replacements. Multi-paragraph/table-spanning tracked output or protected ranges warn with W4002 or fail with E6002.
rename-bookmark | preserve-only | preserve-only | Bookmark rename changes anchor metadata; Word has no useful generated revision range for the name update. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
delete-bookmark | preserve-only | preserve-only | Bookmark deletion removes anchor metadata; Word has no useful generated revision range for the marker removal. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
add-comment | preserve-only | preserve-only | Comments are already review markup, so adding a comment does not create an additional tracked edit. Optional anchor-text selects one normalized text span inside the target paragraph; use occurrence when the span is repeated. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
set-comment-text | text-run | tracked-comment-text | Suggest/Require emit w:del/w:ins inside simple paragraph-only comment bodies while preserving comment metadata; complex comment bodies warn with W4002 or fail with E6002.
resolve-comment | preserve-only | preserve-only | Comment resolution changes review metadata, not visible document text. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
reopen-comment | preserve-only | preserve-only | Comment reopening changes review metadata, not visible document text. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
delete-comment | preserve-only | preserve-only | Comment deletion removes review markup, not a separate generated tracked edit. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
add-comment-reply | preserve-only | preserve-only | Threaded comment replies are review metadata, so adding a reply does not create an additional tracked edit. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
delete-comment-reply | preserve-only | preserve-only | Threaded comment reply deletion removes review metadata, not a separate generated tracked edit. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
set-field-dirty | preserve-only | preserve-only | Field dirty flags are field metadata and have no useful generated visible revision representation. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
set-field-lock | preserve-only | preserve-only | Field lock flags are field metadata and have no useful generated visible revision representation. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
set-field-code | preserve-only | preserve-only | Field codes are instruction metadata; generated revisions for field instructions are not modeled yet. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
set-field-result | text-run | tracked-field-result | Suggest/Require emit w:del/w:ins inside simple w:fldSimple cached result text while preserving the field instruction; direct mode also supports simple same-paragraph complex field result runs. Complex-field tracked output or unsafe topologies warn with W4002 or fail with E6002/E4313.
refresh-field-result | preserve-only | preserve-only | Field refresh updates cached result text from modeled document state for REF/PAGEREF/NOTEREF bookmark fields and QUOTE literal fields; unsupported refresh types return categorized E4313 diagnostics. Generated revision markup for the refresh is not modeled yet. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
set-hyperlink-target | preserve-only | preserve-only | Hyperlink target updates modify relationship or anchor metadata, not visible text. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
set-hyperlink-text | text-run | tracked-hyperlink-text | Suggest/Require emit w:del/w:ins inside the hyperlink wrapper for simple display text while preserving the relationship or anchor; protected or complex hyperlink content warns with W4002 or fails with E6002.
insert-hyperlink-after | text-run | tracked-hyperlink-insert | Suggest/Require emit the inserted hyperlink display text as w:ins inside the hyperlink wrapper while preserving relationship or anchor metadata; text with tabs or line breaks warns with W4002 or fails with E6002.
remove-hyperlink | preserve-only | preserve-only | Hyperlink removal changes wrapper and relationship metadata while preserving display text. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
set-cell | text-run | tracked-cell-simple | Targets can be visual-grid cell IDs or merge-group IDs. Suggest/Require emit w:del/w:ins for simple text-only cells, including compatible multi-paragraph and horizontally merged cells; vertical-merge continuations, force, or complex cells warn with W4002 or fail with E6002.
set-cell-shading | cell-property | tracked-cell-shading | Sets or clears w:tcPr/w:shd fill on a visual-grid cell ID or merge-group ID. Suggest/Require emit cell property revisions with w:tcPrChange while preserving previous cell properties; vertical-merge continuations fail with E4301.
set-table-style | table-property | tracked-table-style | Suggest/Require emit table property revisions with w:tblPrChange while preserving previous table properties.
set-table-metadata | preserve-only | preserve-only | Table caption and description updates are table metadata, not visible document text. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
set-row-header | row-property | tracked-row-header | Suggest/Require emit row property revisions with w:trPrChange while preserving previous row properties.
append-row | row-structure | tracked-row-insert | Direct mode appends by cloning the last row shape when the table has a consistent visual grid and the last row does not contain vertical merge cells. Suggest/Require emit row insertion revisions with w:trPr/w:ins for simple rectangular tables; visual-grid or other complex shapes warn with W4002 or fail with E6002.
insert-row-before | row-structure | tracked-row-insert | Direct mode clones the target row shape for consistent visual-grid tables when the insertion boundary does not cross an active vertical merge chain. Suggest/Require emit row insertion revisions with w:trPr/w:ins for simple rectangular tables; force, visual-grid, or other complex shapes warn with W4002 or fail with E6002.
insert-row-after | row-structure | tracked-row-insert | Direct mode clones the target row shape for consistent visual-grid tables when the insertion boundary does not cross an active vertical merge chain. Suggest/Require emit row insertion revisions with w:trPr/w:ins for simple rectangular tables; force, visual-grid, or other complex shapes warn with W4002 or fail with E6002.
delete-row | row-structure | tracked-row-delete | Direct mode deletes rows in consistent visual-grid tables and promotes the next vertical-merge continuation when deleting a merge root. Suggest/Require emit row deletion revisions with w:trPr/w:del for simple rectangular tables; force, visual-grid, or other complex shapes warn with W4002 or fail with E6002.
append-column | unsupported | unsupported | Table-column transforms are not safely modeled yet; check/apply fails with E4316.
insert-column-before | unsupported | unsupported | Table-column transforms are not safely modeled yet; check/apply fails with E4316.
insert-column-after | unsupported | unsupported | Table-column transforms are not safely modeled yet; check/apply fails with E4316.
delete-column | unsupported | unsupported | Table-column transforms are not safely modeled yet; check/apply fails with E4316.
replace-image | preserve-only | preserve-only | Image replacement updates DrawingML and package media; generated drawing-level revision markup is not modeled yet. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
insert-image-after | preserve-only | preserve-only | Image insertion creates DrawingML and package media; generated drawing-level revision markup is not modeled yet. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
set-image-alt | preserve-only | preserve-only | Image alt-text updates DrawingML metadata, not visible document text. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
set-image-metadata | preserve-only | preserve-only | Image title/name/alt updates DrawingML metadata, not visible document text. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
set-image-size | preserve-only | preserve-only | Image size updates DrawingML layout metadata, not visible document text. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
set-image-wrap | preserve-only | preserve-only | Image wrapping updates DrawingML layout metadata, not visible document text. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
set-image-position | preserve-only | preserve-only | Image position updates DrawingML layout metadata, not visible document text. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
set-image-crop | preserve-only | preserve-only | Image crop updates DrawingML layout metadata, not visible document text. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
delete-image | preserve-only | preserve-only | Image deletion removes DrawingML and package media; generated drawing-level revision markup is not modeled yet. Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.
set-section-columns | section-property | tracked-section-columns | Suggest/Require emit section property revisions with w:sectPrChange while preserving previous section properties and references; existing section property revisions fall back or fail instead of being replaced.
set-section-orientation | section-property | tracked-section-orientation | Suggest/Require emit section property revisions with w:sectPrChange while preserving previous page size, section properties, and references; existing section property revisions fall back or fail instead of being replaced.
<!-- END GENERATED TRACK-CHANGES SUPPORT TABLE -->

Other known operations are classified as `preserve-only`: they preserve existing
tracked-change markup but do not create new revision markup. In `suggest` mode
they apply directly with `W4001`; in `require` mode they fail with `E6001`.

`docxedit help patch` prints both the support class and the operation-specific
support value. Support classes include `text-run`, `paragraph-block`,
`paragraph-property`, `table-property`, `row-property`, `cell-property`,
`row-structure`, `section-property`, `relationship-metadata`, `preserve-only`,
and `unsupported`.
The patch engine uses the support class, while the detailed support value explains
the narrower operation shape, such as `tracked-simple` or `tracked-cell-simple`.

Even supported tracked operations can fail for unsupported shapes, such as tabs,
line breaks, soft hyphens, symbols, other non-text run content, protected OOXML
boundaries, existing revision markup, or genuinely mixed direct run formatting.
Equivalent direct run formatting is compared through canonicalized `w:rPr`, so
harmless child or attribute ordering differences do not force fallback.
Unsupported shapes produce `W4002` in `suggest` mode or `E6002` in `require`
mode. The diagnostic includes the operation name, target ID, catalog support
value, and exact unsupported-shape reason; `W4002` also says the edit is being
applied directly.

Overlap policy:

- Unrelated existing tracked-change, move, custom XML revision, comment,
  bookmark, content-control, field, hyperlink, image, table, and section markup
  is preserved.
- Generated tracked edits may be adjacent to existing revision or comment
  markup, but they must not replace through it.
- Existing `w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`, move range markers, and
  custom XML revision range markers are protected boundaries for generated text
  edits. `suggest` falls back to a direct edit with `W4002`; `require` fails with
  `E6002`.
- Generic paragraph/cell text edits do not cross comment ranges, bookmarks,
  content controls, fields, or hyperlinks. Use the dedicated comment, bookmark,
  content-control, field-result, or hyperlink-text operation when the wrapper is
  the intended edit surface.
- Property revisions are not nested or replaced. If a target already owns the
  same tracked property revision shape, `suggest` falls back and `require` fails.

When `apply` creates revision markup, operation summaries and report JSON include
the generated revision IDs for that operation. Use those IDs with `changes` to
verify the resulting markup without reading raw OOXML.
Table operation reports include affected row/cell IDs plus visual-column ends,
grid-before/grid-after offsets, merge-group IDs, and nested-table paths when the
target shape has them.

## Assets

Image operations read `asset` from the local filesystem. Relative paths are
resolved against the patch file directory first, then the current working
directory of the CLI process; absolute paths work unchanged. Keep assets next
to the patch so the same patch resolves identically regardless of where the
CLI is invoked from. A patch read from stdin (`-`) has no directory, so relative
assets resolve against the invoking working directory only. Use `expect-content-type` when replacing or deleting
an image so accidental target mixups fail early.

## Diagnostics

Run `check` first:

```text
docxedit check input.docx edits.docxpatch
```

Common diagnostics:

- `E12xx`: selector parse, not found, or ambiguous.
- `E20xx`: patch syntax error.
- `E32xx`: guard failure.
- `E42xx`: missing or invalid operation field.
- `E43xx`: unsafe edit shape, unsupported protected boundary, or unsupported
  target kind. Content-control kind mismatches include actual kind and next-step
  guidance for picture, group, checkbox, choice, date, and repeating controls.
- `E52xx`: image asset or DrawingML issue.
- `E60xx`: track-change mode issue.
- `E62xx`: section edit issue.

See [diagnostics.md](diagnostics.md) for the full diagnostic map.
