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
- `W4001`: `TrackChangesMode.Suggest` warning; the edit is applied directly.

Use `--diagnostics path` to write diagnostics JSON. Use `--strict` to turn warnings into exit code `3`.
