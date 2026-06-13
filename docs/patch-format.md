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
- `set-content-control-text`: `target`, `text`
- `set-content-control-checkbox`: `target`, `checked`
- `set-content-control-choice`: `target`, exactly one of `value` or `display-text`
- `set-content-control-date`: `target`, `value`, optional `display-text`
- `replace-bookmark-text`: `target`, `text`
- `rename-bookmark`: `target`, `name`
- `delete-bookmark`: `target`
- `add-comment`: `target`, `text`; optional `expect-text`, `author`, `initials`, `date`
- `set-comment-text`: `target`, `text`
- `resolve-comment`: `target`
- `reopen-comment`: `target`
- `delete-comment`: `target`
- `add-comment-reply`: `target`, `text`, optional `author`, `initials`, `date`; recognized but fails with `E4314`
- `delete-comment-reply`: `target`; recognized but fails with `E4314`
- `set-field-dirty`: `target`, `dirty`
- `set-field-lock`: `target`, `locked`
- `set-field-code`: `target`, `code`; optional `expect-code`
- `set-field-result`: `target`, `text`; optional `expect-result`
- `set-hyperlink-target`: `target`, exactly one of `uri` or `anchor`, optional `tooltip`, `target-frame`, `history`
- `set-hyperlink-text`: `target`, `text`
- `insert-hyperlink-after`: `target`, `text`, exactly one of `uri` or `anchor`, optional `tooltip`, `target-frame`, `history`
- `remove-hyperlink`: `target`
- `set-cell`: `target`, `text`, optional `expect-text`, `expect-row-count`, `expect-column-count`, `force`
- `set-table-style`: `target`, `style`, optional `expect-style`
- `set-table-metadata`: `target`, `caption` and/or `description`, optional `expect-caption`, `expect-description`
- `set-row-header`: `target`, `header`, optional `expect-header`
- `append-row`: `target`, repeated `cell`, optional `expect-row-count`, `expect-column-count`
- `insert-row-before`, `insert-row-after`: `target`, repeated `cell`, optional `expect-row-count`, `expect-column-count`, `expect-cell-count`, `force`
- `delete-row`: `target`, optional `expect-row-count`, `expect-column-count`, `expect-cell-count`, `expect-contains`, `force`
- `replace-image`: `target`, `asset`, optional `expect-content-type`, `alt`
- `insert-image-after`: `target`, `asset`, optional `expect-content-type`, `width`, `height`, `alt`
- `set-image-alt`: `target`, `alt`, optional `expect-content-type`
- `set-image-metadata`: `target`, at least one of `alt`, `title`, or `name`, optional `expect-content-type`
- `set-image-size`: `target`, `width` and/or `height`, optional `expect-content-type`
- `set-image-wrap`: `target`, optional `mode`, `dist-top`, `dist-bottom`, `dist-left`, `dist-right`, optional `expect-content-type`
- `set-image-position`: `target`, optional `horizontal-relative`, `horizontal-offset`, `horizontal-align`, `vertical-relative`, `vertical-offset`, `vertical-align`, optional `expect-content-type`
- `set-image-crop`: `target`, at least one of `left-percent`, `top-percent`, `right-percent`, or `bottom-percent`, optional `expect-content-type`
- `delete-image`: `target`, optional `expect-content-type`
- `set-section-columns`: `target`, `count`, optional `expect-columns`, `expect-orientation`
- `set-section-orientation`: `target`, `orientation`, optional `expect-columns`, `expect-orientation`

Explicit header/footer paragraph IDs can be used for paragraph text/style edits, block insertion/deletion, and image insertion after the paragraph. Explicit header/footer table IDs can be used for simple table edits and as block insertion/deletion anchors. Explicit header/footer image IDs can be used for image replacement, alt text, and deletion.

For list-like insertions, set `copy-paragraph-properties true` on `insert-before` or
`insert-after` with a paragraph target. The inserted paragraph copies the target
paragraph's `w:pPr`, including style and numbering properties; an explicit `style`
field overrides the copied paragraph style while preserving the other copied
properties.

Use `set-content-control-text` with plain-text content-control IDs such as `M.CC0001`.
Use `set-content-control-checkbox` with checkbox controls and `checked true|false`;
it updates `w:checked` and the displayed state symbol. The wrapper and `w:sdtPr`
metadata are preserved. Use `set-content-control-choice` with dropdown or combo box
controls and exactly one of `value` or `display-text`; it verifies the list item and
updates the displayed content. Use `set-content-control-date` with date controls; it
updates `w:fullDate` to `value` and uses `display-text` for the visible content when
provided. Content-control edit operations reject controls whose `w:lock` value is not
`unlocked`. Use `replace-bookmark-text` with simple same-paragraph bookmark IDs such
as `M.B0001`; the bookmark start/end markers are preserved and unsupported ranges
fail instead of flattening surrounding XML. Use `rename-bookmark` to change a bookmark
name; DocxEdit rejects duplicate new names and updates same-story internal hyperlink
anchors when the old name is unambiguous. Use `delete-bookmark` to remove complete
unreferenced bookmark markers while preserving the bookmarked content.

`add-comment` targets a modeled paragraph such as `M.P0004`, creates the comments
part/relationship/content type when needed, appends a new comment body, and anchors
the whole paragraph with matching range/reference markers. It accepts optional
`expect-text`, `author`, `initials`, and ISO-8601 `date` fields. Existing comment
operations target either `comment:<id>` from `changes` output or a comment body target
such as `C001.C0001`. `set-comment-text` replaces the comment body with a single
paragraph while preserving comment metadata. `resolve-comment` and
`reopen-comment` create or update the matching `commentsExtended.xml` `w15:done`
flag for basic comments; unsupported body shapes fail with `E4312`. `delete-comment`
removes the comment body, matching range/reference markers from document stories, and
matching `commentsExtended.xml` records when present. Threaded reply operations are
recognized as `add-comment-reply` and `delete-comment-reply` so agents receive stable
`E4314` diagnostics instead of generic unknown-operation errors; DocxEdit does not
edit threaded reply metadata yet.

Field operations target field IDs from `read` or `outline`, such as `M.F0001`,
`H001.F0001`, or `F001.F0001`. `set-field-dirty` updates `w:dirty` and
`set-field-lock` updates `w:fldLock` on `w:fldSimple` or the complex field begin
`w:fldChar`; use `target all` to update every modeled field in main/header/footer
stories. `set-field-code` updates `w:fldSimple/@w:instr`, supports optional
normalized `expect-code`, preserves the cached result, and marks that field dirty.
`set-field-result` replaces the cached result runs inside `w:fldSimple`, supports
optional exact `expect-result`, preserves the field code/boundary, and avoids
document-level field-update marking when it is the only patch operation. Complex
field code/result edits fail with `E4313`. DocxEdit does not recalculate field
results; apply emits `W5103` when a document containing fields is marked for
Word-side refresh.

Hyperlink operations target hyperlink IDs from `read` or `outline`, such as
`M.L0001`, `H001.L0001`, or `F001.L0001`. Use `uri` for external absolute
`http`, `https`, or `mailto` links and `anchor` for internal bookmark anchors.
Relative targets, malformed URIs, `file`, UNC/file-style targets, and unsafe schemes
are rejected. Relationship-backed internal part links are preserved and surfaced as
`target-part` metadata, but patch edits support external URI or bookmark-anchor
targets only. `target-frame` writes `w:tgtFrame`, and `history true|false` writes
`w:history`. `remove-hyperlink` unwraps the hyperlink and keeps its child runs as
ordinary document content.

Use table guards whenever possible:

- `expect-text` verifies the selected cell's current visible text for `set-cell`.
- `expect-style` verifies the selected table's current `w:tblStyle`.
- `expect-caption` verifies the selected table's current `w:tblCaption`; missing
  and empty values are equivalent.
- `expect-description` verifies the selected table's current `w:tblDescription`;
  missing and empty values are equivalent.
- `expect-header` verifies the selected row's current repeating-header flag.
- `expect-row-count` verifies the target table's row count.
- `expect-column-count` verifies the target table's logical column count.
- `expect-cell-count` verifies a targeted row's physical cell count for row insert/delete.
- `expect-contains` verifies a row's visible text before `delete-row`.

CLI examples prefer `--output output.docx`; `-o output.docx` is also accepted.

Unsupported fields are rejected. `expect-hash` and `preserve-size` are not supported.

`delete-row` `expect-contains` is a row-text guard: the operation fails unless the
resolved row's final visible text contains the supplied value exactly.
`set-table-style` updates `w:tblPr/w:tblStyle` and creates `w:tblPr` when missing.
`set-table-metadata` updates table `w:tblPr/w:tblCaption` and
`w:tblPr/w:tblDescription`, creating `w:tblPr` when missing. Use an empty heredoc
value for `caption` or `description` to remove that metadata element.
`set-row-header` sets or clears the row's `w:tblHeader` flag while preserving other
row properties.

`replace-image` `alt` updates the target inline or anchored DrawingML object's
description while replacing the media bytes. Use `set-image-alt` when only the
description should change. Use `set-image-metadata` to update DrawingML `docPr`
description (`alt`), `title`, and `name` without replacing media bytes. Use
`set-image-size` to update DrawingML `wp:extent` and picture transform extents; if
only `width` or `height` is provided, DocxEdit preserves the current aspect ratio
when it can infer one. Use `set-image-wrap` on anchored images to update `wp:wrap*`
mode and anchor wrap distances; inline images reject wrap edits. Use
`set-image-position` on anchored images to update `wp:positionH`/`wp:positionV`
relative bases, signed offsets, or alignments. Use `set-image-crop` to update
DrawingML `a:srcRect` crop percentages without replacing media bytes; omitted crop
sides keep their current value, zero-valued sides are removed, and opposing side sums
must remain below 100 percent.

Tracked output is intentionally narrow. It supports simple `replace-text`,
whole-paragraph replacement, inserted/deleted paragraph text, paragraph style changes,
and simple single-paragraph table-cell text replacement. Tracked text shapes must
contain no tabs or line breaks, must not cross protected OOXML boundaries such as
hyperlinks, fields, comments, bookmarks, content controls, drawings, or existing
revision markup, and must have compatible direct run-property shape.
Known operations without generated revision output are classified as `preserve-only`
in the shared help catalog: they preserve existing revision markup, apply directly
under `TrackChangesMode.Suggest` with `W4001`, and fail under
`TrackChangesMode.Require` with `E6001`. The `W4001` and `E6001` messages include
the catalog support value so integrations can distinguish intentionally preserve-only
operations from unclassified operations. Supported tracked operations still fail
unsupported target shapes with `E6002`; `TrackChangesMode.Suggest` warns with `W4002`
and applies the direct edit instead.
Existing tracked-change and comment markup is preserved unless the targeted operation
would directly replace that protected boundary. Generated revision IDs are allocated
after existing `w:id` values to avoid collisions.
