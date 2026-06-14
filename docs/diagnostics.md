# Diagnostics

Diagnostics have a severity, code, message, and optional target, part, story, feature, fallback, operation index, line, and column.

Common code ranges:

- `E0001`: expected document/package/XML failure normalized by the public API.
- `W10xx`: read-model unsupported-feature warnings.
- `E12xx`: selector parse, not found, or ambiguity.
- `E20xx`: patch syntax errors.
- `E32xx`: guard failures.
- `E42xx`: missing or invalid operation fields.
- `E43xx`: unsafe table/text structure or protected edit boundary.
- `E52xx`: image asset, type, or DrawingML issue.
- `E60xx`: unsupported track-change edit mode.
- `E62xx`: section edit issue.
- `E71xx`: style resolution issue.
- `E90xx`: post-edit validation failure.
- `E91xx`: structural package validation failure from `validate`.
- `W91xx`: validation warning from `validate`.
- `W4001`: `TrackChangesMode.Suggest` warning for an operation without tracked-change
  output; the message includes the shared catalog support value, and the edit is
  applied directly. `Feature` is `track-changes-no-revision-representation` and
  `Fallback` is `direct-edit-preserve-existing-revisions`.
- `W4002`: `TrackChangesMode.Suggest` warning for a tracked-capable operation
  whose specific target shape cannot be represented as generated revision markup.
  The message includes the operation name, target ID, catalog support value, and
  unsupported-shape reason; the edit is applied directly. `Feature` is
  `track-changes-unsupported-target-shape` and `Fallback` is
  `direct-edit-preserve-existing-revisions`.
- `E6001`: `TrackChangesMode.Require` error for an operation without generated
  tracked-change output; the message includes the shared catalog support value.
  `Feature` is `track-changes-no-revision-representation` and `Fallback` is
  `require-failed`.
- `E6002`: `TrackChangesMode.Require` error for a tracked-capable operation whose
  specific target shape cannot be represented as generated revision markup. The
  message includes the operation name, target ID, catalog support value, and
  unsupported-shape reason. `Feature` is
  `track-changes-unsupported-target-shape` and `Fallback` is `require-failed`.
- `E6003`: `TrackChangesMode.Suggest` or `TrackChangesMode.Require` error for
  invalid generated revision metadata before operation execution. Empty revision
  authors are rejected with `Feature = track-changes-revision-metadata` and
  `Fallback = no-output-written`.
- `E4310`: content-control edit requested a control kind, content container, or
  lock state that is not safely editable.
- `E4312`: comment resolution state cannot be edited because the comment body shape
  cannot safely host modern resolution metadata.
- `E4313`: field code/result replacement was requested for a field shape that is not
  safely editable by the requested operation.
- `E4314`: threaded comment reply operations are recognized but not supported because
  DocxEdit does not safely model commentsIds/threaded-comments metadata yet.
- `E4315`: repeating-section item operations are recognized but not supported because
  DocxEdit does not safely model repeating-section subtree insertion/deletion yet.
- `E4316`: table column operations are recognized but not supported because DocxEdit
  does not safely model column transforms across grids, spans, omitted cells, and
  vertical merges yet.
- `W5103`: apply marked a document containing fields for Word-side refresh because
  DocxEdit does not recalculate field results.
- `E9199`/`W9199`: validation diagnostics were capped; `E9199` means omitted
  diagnostics include errors, while `W9199` means only warnings were omitted.

Common read-model warning features include `tracked-changes`, `hyperlink`, `field`, `comment`,
`bookmark`, `content-control`, `floating-image`, `external-image`, `chart`, `smart-art`,
`equation`, `shape`, `linked-image`, `vml`, `grouped-drawing`, `ole-object`, `alt-chunk`,
`section-flow`, and `numbering`.

`W10xx` read-model warnings include `PartName`, `Story`, `Feature`, and `Fallback` metadata.
Fallback values describe the behavior used by the reader, such as `selected-text-view`, `plain-text`,
`changes-metadata`, `preserve-only`, `omit-from-editable-images`, `invalid-uri`,
`missing-anchor`, `duplicate-anchor`, `unsupported-internal-part-link`, or
`basic-section-model`, `unsupported-picture-bullet`, or `unsupported-numbering-format`.

Unsupported drawing warnings distinguish editable image records from preserve-only
OOXML shapes. Linked images (`W1019`) are not fetched or listed as editable images.
VML (`W1020`), grouped drawings (`W1021`), OLE objects (`W1022`), charts, SmartArt,
equations, and generic shapes are preserved but not modeled.

Hyperlink warnings distinguish broken relationships (`W1015`), invalid external URI
targets (`W1016`), missing or duplicate anchors (`W1017`/`W1018`), and preserved but
unsupported internal part links (`W1023`).

Numbering warnings distinguish preserved picture bullets (`W1024`) and `numFmt`
values that DocxEdit cannot expand into deterministic labels (`W1025`).

Use `--diagnostics path` to write diagnostics JSON. Use `--strict` to turn warnings into exit code `3`.

`validate` currently emits:

- `E9101`: XML part has no root element.
- `E9102`: known WordprocessingML part has an unexpected root.
- `E9103`: bookmark or comment range start/end IDs are unbalanced.
- `E9104`: complex field begin/end markers are unbalanced.
- `E9105`: drawing references a missing relationship ID.
- `E9106`: table or row is missing required row/cell structure.
- `E9107`: duplicate drawing `wp:docPr` ID within one Word part.
- `E9108`: `commentsExtended` metadata has missing, duplicate, or orphan paraId records.
- `E9109`: DrawingML extents are missing or non-positive.
- `E9110`: DrawingML crop percentages are malformed or leave no visible image area.
- `E9111`: comment body, anchor, or reference IDs are missing, duplicated, or inconsistent.
- `E9112`: field instruction text, result containment, or dirty/lock flag metadata
  is malformed.
- `E9113`: DrawingML image references use a non-image relationship, missing target
  part, or non-image media content type.
- `E9114`: table visual-grid metadata is inconsistent, such as invalid `gridSpan`,
  vertical merge continuation mismatch, or row width exceeding `tblGrid`.
- `E9115`: content-control metadata is malformed, such as invalid `w:id`,
  duplicate `w:id`, invalid `w:lock`, or invalid checkbox `w:checked` values.
- `E9118`: settings metadata is malformed, such as an invalid `w:updateFields`
  boolean value.
- `E9119`: header or footer section references are missing a relationship, point
  to a missing relationship ID, or use the wrong relationship type.
- `E9120`: section properties are malformed, such as invalid column counts or page
  orientation values.
- `E9199`: validation diagnostics were capped and omitted diagnostics include errors.
- `W9109`: duplicate bookmark names or duplicate content-control tag/alias values make
  semantic selectors ambiguous; the message lists candidate IDs.
- `W9116`: a paragraph references a style ID that is missing from `styles.xml`.
- `W9117`: a paragraph numbering reference, numbering definition, or abstract
  numbering definition is missing.
- `W9199`: validation diagnostics were capped and omitted diagnostics are warnings only.
