# Changelog

## 0.1.0 - First NuGet release

Shipped as a net10.0 library with no runtime package dependencies, plus source CLI for development. The CLI is not installed by the package.

Library and patch behavior shipped:

- Stream-first DocxEditor for read, outline, find, dump, styles, media extraction, tracked-change summaries, check, and apply, with caller-owned streams.
- Docxpatch DSL for paragraph text, blocks, styles, main header footer tables with visual grid coordinates, inline images with PNG JPEG interchange, bookmarks, hyperlinks, comments, fields, content controls, sections, and rows with guarded templates.
- Positional target IDs from fresh read output, with explicit input binding, live semantic resolution, and patch-local aliases. Reports carry Id with Coordinate input or operation-time and FinalId live, absent when deleted. Deleted objects never acquire survivor identity.
- Asset boundary where image edits use caller-supplied bytes via AssetProvider. Extraction returns bytes for external tools, replacement inserts caller bytes. No pixel editing in the library.
- Tracked output with off suggest require modes and generated revision IDs. Annotation operations stay permitted under Require. Complex shapes warn or fail with explicit diagnostics.
- Package validation for ZIP paths, size limits, relationships, macro policy, unknown part preservation, and touched-part checks. Failed CLI apply preserves any pre-existing destination file byte for byte. Library callers stage output and publish only on success.
- No layout, rendering, pagination, or Word fidelity guarantees. Word open and save checks are bounded compatibility only.

Notes:

- This is the first published version. There is no migration from a prior NuGet release.
- Source documentation lives in docs in the repository. Package consumers use the library API and the guidance in this README.
