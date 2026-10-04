# Changelog

## Unreleased

- Added native equation discovery and guarded whole-equation insertion, replacement,
  and deletion with a bounded LaTeX-like syntax. Text rewrites now protect equations.
  Includes inline/display placement, snapshot IDs, aliases, OMML extraction,
  capability/template guidance, and a local Word-to-PDF smoke-test script.
- Added `DocxEditor.Create` and the shared `create` command for empty A4 or US Letter
  documents in portrait or landscape, with fixed margins and Normal/Heading1–Heading9
  styles. Includes quotas, stream ownership, cancellation, staged command publication,
  binary stdout, and JSON reports.
- Section reads now expose stored page dimensions, margins, header/footer distances,
  and gutter in twips, with nullable values when unspecified or non-integer.
  XML integer measurements use invariant number syntax regardless of host culture.

## 0.1.0 - 2026-10-02

First release as a net10.0 library with no runtime package dependencies, plus source CLI for development. The CLI is not installed by the package.

Library and patch behavior:

- `DocxCommand.RunAsync` and a host I/O contract let embedded callers and the development CLI share command parsing, execution, rendering, and exit codes. `DocxJson.CreateOptions` exposes the canonical JSON wire format. Document and patch input, file publication, and command output support asynchronous hosts.
- Hosted image assets use `IDocxAsyncAssetProvider` with cancellation and owned streams. Each decoded reference is read once per command; the existing image validation and diagnostic behavior is shared with direct library edits.
- Stream-first DocxEditor for read, outline, find, dump, styles, media extraction, tracked-change summaries, check, and apply, with caller-owned streams.
- Docxpatch DSL for paragraph text, blocks, styles, main header footer tables with visual grid coordinates, inline images with PNG JPEG interchange, bookmarks, hyperlinks, comments, fields, content controls, sections, and rows with guarded templates.
- Positional target IDs from fresh read output, with explicit input binding, live semantic resolution, and patch-local aliases. Reports carry Id with Coordinate input or operation-time and FinalId live, absent when deleted. Deleted objects never acquire survivor identity.
- Asset boundary where image edits use caller-supplied bytes via AssetProvider. Extraction returns bytes for external tools, replacement inserts caller bytes. No pixel editing in the library.
- Tracked output with off suggest require modes and generated revision IDs. Annotation operations stay permitted under Require. Complex shapes warn or fail with explicit diagnostics.
- Package validation for ZIP paths, size limits, relationships, macro policy, unknown part preservation, and touched-part checks. Failed CLI apply preserves any pre-existing destination file byte for byte. Library callers stage output and publish only on success.
- No layout, rendering, pagination, or Word fidelity guarantees. Word open and save checks are bounded compatibility only.

Notes:

- This is the first package version. There is no migration from a prior NuGet release.
- Source documentation lives in docs in the repository. Package consumers use the library API and the guidance in this README.
