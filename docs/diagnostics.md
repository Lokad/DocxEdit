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
- `W4001`: `TrackChangesMode.Suggest` warning for an operation without tracked-change output; the edit is applied directly.
- `W4002`: `TrackChangesMode.Suggest` warning for a `replace-text` shape that cannot be represented as simple tracked-change output; the edit is applied directly.
- `E4312`: comment resolution state cannot be edited because the comment lacks
  modern `w15:paraId`/`commentsExtended` metadata or no matching extension record
  exists.

Common read-model warning features include `tracked-changes`, `hyperlink`, `field`, `comment`,
`bookmark`, `content-control`, `floating-image`, `external-image`, `chart`, `smart-art`,
`equation`, `shape`, `alt-chunk`, and `section-flow`.

`W10xx` read-model warnings include `PartName`, `Story`, `Feature`, and `Fallback` metadata.
Fallback values describe the behavior used by the reader, such as `selected-text-view`, `plain-text`,
`changes-metadata`, `preserve-only`, `omit-from-editable-images`, `invalid-uri`,
`missing-anchor`, `duplicate-anchor`, or `basic-section-model`.

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
- `W9109`: duplicate bookmark names or duplicate content-control tag/alias values make
  semantic selectors ambiguous; the message lists candidate IDs.
