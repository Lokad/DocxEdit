# Patch Format

Patch files start with `docxpatch 1` and then one or more operation blocks.

```text
docxpatch 1

op replace-text
target M.P0001
expect-text <<<
current text
>>>
find current
with updated
end
```

Fields are line based. Heredocs use `<<<` and close with `>>>`. Repeated fields are preserved for operations such as table row edits.

## Selectors

Explicit IDs remain the most stable selectors:

- `M.P0001`: main paragraph
- `H001.P0001` / `F001.P0001`: header/footer paragraph
- `M.T0001`, `M.T0001.R02`, `M.T0001.R02.C03`: table, row, cell
- `H001.T0001`, `H001.T0001.R02`, `H001.T0001.R02.C03`: header table, row, cell
- `F001.T0001`, `F001.T0001.R02`, `F001.T0001.R02.C03`: footer table, row, cell
- `M.I0001`: inline main-document image
- `H001.I0001` / `F001.I0001`: inline header/footer image
- `M.S0001`: main-document section

Paragraph operations also support:

- `heading:"Exact heading"`
- `heading:2:"Exact heading"`
- `text:"contained paragraph text"`
- `bookmark:"BookmarkName"`
- `content-control:"TagOrAlias"`

Ambiguous selectors fail with `E1202`; use a more specific selector or an explicit ID.
Parsed selectors that match no target fail with `E1201` and include nearby target IDs without
including nearby paragraph text.

## Supported Operations

- `replace-text`: `target`, `find`, `with`, optional `expect-text`, `preserve-runs`, `occurrence`. Under `TrackChangesMode.Suggest` or `Require`, simple text-only replacements are emitted as tracked `w:del`/`w:ins` markup with the configured author and timestamp.
- `replace-paragraph`: `target`, `text`, optional `expect-text`, `style`
- `insert-before`, `insert-after`: `target`, `text`, optional `style`, `copy-paragraph-properties`
- `delete-block`: `target`, optional `expect-text`
- `set-style`: `target`, `style`
- `set-cell`: `target`, `text`, optional `expect-text`, `expect-row-count`, `expect-column-count`, `force`
- `append-row`: `target`, repeated `cell`, optional `expect-row-count`, `expect-column-count`
- `insert-row-before`, `insert-row-after`: `target`, repeated `cell`, optional `expect-row-count`, `expect-column-count`, `expect-cell-count`, `force`
- `delete-row`: `target`, optional `expect-row-count`, `expect-column-count`, `expect-cell-count`, `expect-contains`, `force`
- `replace-image`: `target`, `asset`, optional `expect-content-type`, `alt`
- `insert-image-after`: `target`, `asset`, optional `expect-content-type`, `width`, `height`, `alt`
- `set-image-alt`: `target`, `alt`, optional `expect-content-type`
- `delete-image`: `target`, optional `expect-content-type`
- `set-section-columns`: `target`, `count`, optional `expect-columns`, `expect-orientation`
- `set-section-orientation`: `target`, `orientation`, optional `expect-columns`, `expect-orientation`

Explicit header/footer paragraph IDs can be used for paragraph text/style edits, block insertion/deletion, and image insertion after the paragraph. Explicit header/footer table IDs can be used for simple table edits and as block insertion/deletion anchors. Explicit header/footer image IDs can be used for image replacement, alt text, and deletion.

For list-like insertions, set `copy-paragraph-properties true` on `insert-before` or
`insert-after` with a paragraph target. The inserted paragraph copies the target
paragraph's `w:pPr`, including style and numbering properties; an explicit `style`
field overrides the copied paragraph style while preserving the other copied
properties.

Use table guards whenever possible:

- `expect-text` verifies the selected cell's current visible text for `set-cell`.
- `expect-row-count` verifies the target table's row count.
- `expect-column-count` verifies the target table's logical column count.
- `expect-cell-count` verifies a targeted row's physical cell count for row insert/delete.
- `expect-contains` verifies a row's visible text before `delete-row`.

CLI examples prefer `--output output.docx`; `-o output.docx` is also accepted.

Unsupported fields are rejected. `expect-hash`, `preserve-size`, and `caption` are not supported.

`delete-row` `expect-contains` is a row-text guard: the operation fails unless the
resolved row's final visible text contains the supplied value exactly.

`replace-image` `alt` updates the target inline or anchored DrawingML object's
description while replacing the media bytes. Use `set-image-alt` when only the
description should change.

Tracked `replace-text` output is intentionally narrow. It supports replacements whose
matched and replacement text contain no tabs or line breaks, whose paragraph does not
cross protected OOXML boundaries such as hyperlinks, fields, comments, bookmarks,
content controls, drawings, or existing revision markup, and whose visible direct text
runs share one run-property shape. `TrackChangesMode.Require` fails unsupported shapes
with `E6002`; `TrackChangesMode.Suggest` warns with `W4002` and applies the direct edit
instead.
