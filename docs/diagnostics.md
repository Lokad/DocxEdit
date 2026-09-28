# Diagnostics

Diagnostics have a severity, code, message, and optional target, part, story, feature, fallback, operation index, line, and column.

Patch-operation findings reuse source positions: selector findings point at the target field, guard findings at the guard field, and find or anchor findings at the find or anchor-text field.

Common code ranges:

- `E0001`: expected document/package/XML failure normalized by the public API.
- `I00xx`: informational notes that never fail; success is unaffected.
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
- `E1201`: selector matched no targets, a target ID was not found, or the target shape is wrong for the operation.
- `E1202`: selector matched multiple targets ambiguously (the message lists candidate IDs), or find/anchor text matched multiple ranges; the message reports the match count and how to select one match or every match.
- `E1203`: selector text is malformed.
- `E2001`: patch text is empty.
- `E2002`: patch is missing the required `docxpatch 1` preamble.
- `E2003`: unsupported docxpatch major version.
- `E2004`: an `expect-hash` patch field was supplied; the feature is not supported (the same words inside a heredoc value stay literal).
- `E2005`: unexpected patch line.
- `E2006`: operation name is required.
- `E2007`: invalid field line.
- `E2008`: unterminated heredoc for a field.
- `E2009`: operation is missing `end`.
- `E2010`: unknown operation.
- `E2011`: unknown field for the operation.
- `E2012`: boolean field value must be true or false.
- `E2013`: integer field value must be an integer.
- `E2014`: patch text exceeds the configured maximum patch size.
- `E2015`: a non-repeatable patch field is specified more than once in one operation.
- `I0001`: an operation completed as a semantic no-op: guards passed but the requested end state already held, so nothing was written, no revisions were generated, and fields were not marked dirty.
- `E3201`: guard failed: an `expect-*` value differs from the current value.
- `E4201`: operation is recognized but not supported in this mode.
- `E4202`: required field or field combination is missing.
- `E4203`: find or anchor text was not found.
- `E4204`: boolean field value must be true or false.
- `E4205`: field value fails its shape rule.
- `E4301`: table, row, or cell target is unsafe (visual grid or vertical merges).
- `E4302`: cell contains unsupported content; `force true` is required to replace it all.
- `E4303`: cell field count does not match the expected count.
- `E4304`: the last table row cannot be deleted.
- `E4305`: the edit would remove protected OOXML boundary markup.
- `E4306`: run-preserving replacement is not supported for the target shape.
- `E4307`: `copy-paragraph-properties` requires a paragraph target.
- `E4311`: bookmark operation target is unsafe (duplicates, incomplete range, hyperlink anchors, or protected boundary).
- `E4310`: content-control edit requested a control kind, content container, or
  lock state that is not safely editable.
- `E4312`: comment resolution state cannot be edited because the comment body shape
  cannot safely host modern resolution metadata.
- `E4313`: field code/result replacement or refresh was requested for a field shape
  or field type that is not safely editable by the requested operation. Refresh
  diagnostics distinguish layout/pagination, document-property, formula,
  mail-merge, date/time, conditional, and external-state requirements.
- `E4314`: a threaded comment reply operation would require unsafe thread metadata
  changes, such as missing parent metadata, a non-reply target, or deleting a reply
  that still has child replies.
- `E4315`: repeating-section item operations are recognized but not supported because
  DocxEdit does not safely model repeating-section subtree insertion/deletion yet.
- `E4316`: table column operations are recognized but not supported because DocxEdit
  does not safely model column transforms across grids, spans, omitted cells, and
  vertical merges yet.
- `E4317`: selected-range comment creation was requested for a paragraph shape that
  cannot safely host direct comment range markers, such as non-direct text runs or
  tabs/line breaks inside the selected span.
- `E5201`: no asset provider is configured for image operation assets.
- `E5202`: image asset could not be resolved.
- `E5203`: image asset is not a supported PNG or JPEG image (missing signature, hint/content disagreement, or structural defect such as truncation).
- `E5204`: replacing the image content type is not supported for existing media.
- `E5205`: image lacks editable DrawingML (properties, crop, or anchored-only metadata).
- `E5206`: invalid image width or height.
- `E5207`: image asset exceeds the configured maximum single-part size.
- `E5208`: image crop percentages are invalid.
- `E5209`: image wrap mode or wrap distance is invalid.
- `E5210`: image position axis is invalid.
- `E6201`: section column count must be between 1 and 4.
- `E6202`: section orientation must be portrait or landscape.
- `E7101`: style was not found.
- `E7102`: style name is ambiguous.
- `E7103`: style reference is the wrong kind (for example a table style where a paragraph style is required).
- `E9001`: post-edit validation failed.
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
`basic-section-model`, `unsupported-picture-bullet`, `unsupported-numbering-format`,
or `tracked-numbering-property-revision`.

Unsupported drawing warnings distinguish editable image records from preserve-only
OOXML shapes. Linked images (`W1019`) are not fetched or listed as editable images.
VML (`W1020`), grouped drawings (`W1021`), OLE objects (`W1022`), charts, SmartArt,
equations, and generic shapes are preserved but not modeled.

Content-control edit diagnostics use `E4310` when the target control kind does
not match the requested operation. The message reports the actual kind and gives
generic next-step guidance, such as inspecting media/image targets for picture
controls or targeting an editable child content control for group controls.

Hyperlink warnings distinguish broken relationships (`W1015`), invalid external URI
targets (`W1016`), missing or duplicate anchors (`W1017`/`W1018`), and preserved but
unsupported internal part links (`W1023`).

Numbering warnings distinguish preserved picture bullets (`W1024`), `numFmt`
values that DocxEdit cannot expand into deterministic labels (`W1025`), and
tracked paragraph property revisions that carry previous `w:numPr` state
(`W1026`). For `W1026`, final-view labels use current numbering properties, while
original-view reconstruction of the previous numbering properties is not modeled.


Read-model unsupported-feature warnings by code:

- `W1001`: tracked-change markup is present.
- `W1002`: hyperlinks are surfaced as metadata.
- `W1003`: fields are surfaced as metadata; results are not evaluated.
- `W1004`: comment anchors are surfaced; comment text is hidden from the read model.
- `W1005`: bookmarks are surfaced as metadata.
- `W1006`: content controls are surfaced as metadata.
- `W1007`: floating images are surfaced with limited layout editing.
- `W1008`: external images are not fetched or listed.
- `W1009`: charts are preserved but not modeled.
- `W1010`: SmartArt is preserved but not modeled.
- `W1011`: equations are preserved but not modeled.
- `W1012`: shapes are preserved but not modeled.
- `W1013`: `altChunk` content is preserved but not imported.
- `W1014`: complex section flow with a basic section model.
- `W1017`: internal hyperlink anchors match no bookmark.
- `W1018`: internal hyperlink anchors match multiple bookmarks.
- `W1021`: grouped drawings are preserved but not modeled.
- `W1022`: OLE objects are preserved but not executed.
Use `--diagnostics path` to write diagnostics JSON. Use `--strict` to turn warnings into exit code `3`.

`validate` currently emits:

- `E9101`: XML part has no root element.
- `E9102`: known WordprocessingML part has an unexpected root.
- `E9103`: bookmark or comment range start/end IDs are unbalanced.
- `E9104`: complex field begin/end markers are unbalanced.
- `E9105`: drawing references a missing relationship ID.
- `E9106`: table or row is missing required row/cell structure.
- `E9107`: duplicate drawing `wp:docPr` ID within one Word part.
- `E9108`: `commentsExtended` metadata has missing, duplicate, orphan, or
  self-parented paraId records.
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
- `E9121`: tracked revision markup is malformed, such as missing revision metadata,
  orphan `w:delText`, or malformed paragraph, table, row, cell, or section
  property revisions.
- `E9122`: `commentsIds.xml` metadata is malformed, such as missing or duplicate
  `w16cid:paraId` / `w16cid:durableId`, or a paraId without a matching comment
  paragraph.
- `E9199`: validation diagnostics were capped and omitted diagnostics include errors.
- `W9109`: duplicate bookmark names or duplicate content-control tag/alias values make
  semantic selectors ambiguous; the message lists candidate IDs.
- `W9116`: a paragraph references a style ID that is missing from `styles.xml`.
- `W9117`: a paragraph numbering reference, numbering definition, or abstract
  numbering definition is missing.
- `W9199`: validation diagnostics were capped and omitted diagnostics are warnings only.
