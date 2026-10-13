# Changelog

## 0.2.0 - Unreleased

- Added `DocxEditor.Create` and the shared `create` command for empty A4 or US Letter
  documents in portrait or landscape, with one-inch margins and Normal/Heading1–Heading9
  styles. Supports quotas, stream ownership, cancellation, staged command publication,
  binary stdout, and JSON reports.
- Added stored section page dimensions, margins, header/footer distances, and gutter
  to read results and text output, in twips with nulls for unspecified or non-integer values.
- Added comment anchors for table cells and selected text within paragraphs or cells,
  with occurrence disambiguation and preservation of surrounding protected markup.
- Added native Word equation discovery, OMML extraction, and guarded whole-equation
  insertion, replacement, and deletion through a bounded LaTeX-like syntax. Supports
  inline/display placement, snapshot IDs, aliases, and capability/template guidance.
  Text rewrites protect equations from accidental removal.
- Added original-byte SVG insertion, raster/vector replacement, sizing, metadata,
  cropping, and extraction through the library and commands. Existing Word SVG previews
  share one image identity; replacement/deletion preserve shared assets. Insertion
  generates no raster preview and requires a reader with SVG support.
- Added native still WebP insertion and format interchange with original-byte
  extraction, sizing, metadata, and cropping. Bounded RIFF/VP8/VP8L/VP8X checks
  reject malformed containers and animated assets without decoding or conversion;
  display requires a reader with WebP support.
- Fixed comment threading to use the last comment paragraph's identity. Editing a
  comment preserves resolution state and reply linkage; deleting a parent with replies
  or a reply with children is refused until the leaf replies are removed.
- Fixed XML integer measurement parsing to use invariant syntax regardless of host culture.
- Fixed duplicate unsupported-feature warnings for shared header/footer parts;
  diagnostics now use the same story labels as the read inventory.
- Reduced repeated XML parsing and traversal by reusing package styles, relationships,
  part roles, story trees, and validation trees. Text editing and content-control
  hierarchy scanning avoid repeated rebuilding and descendant scans.
- Reduced buffer copies during package loading, hosted asset handling, image inspection,
  and command output. Read text is bounded during collection, summaries skip run/text
  rendering, JSON reports serialize directly to staging, and buffered text flushes in chunks.
- Added optional local Word/PDF smoke tests for equations and comments. Creation
  compatibility tests now require an actual Word save rather than accepting a no-op.

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
