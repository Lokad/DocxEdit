# Changelog

## 0.1.0

- Added the stream-first `DocxEditor` public API and CLI commands for read, outline, find, dump, styles, media, tracked-change markup summaries, check, and apply.
- Added `.docxpatch` parsing and edit support for paragraph text, blocks, styles, simple tables, inline images, and basic section settings.
- Added package safety checks for ZIP paths, size limits, relationships, macro policy, unknown part preservation, and touched-part post-edit validation.
- Added no-content private validation tooling for ignored local `.docx` cases.
- Added initial documentation for CLI usage, patch syntax, diagnostics, validation, and architecture.
