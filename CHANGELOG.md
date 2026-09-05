# Changelog

## 0.1.0

- Added the stream-first `DocxEditor` public API and CLI commands for read, outline, find, dump, styles, media, tracked-change markup summaries, check, and apply.
- Added `.docxpatch` parsing and edit support for paragraph text, blocks, styles, simple main/header/footer tables, inline images, and basic section settings.
- Added simple tracked-change output for `replace-text` with author, timestamp, and revision IDs.
- Added comment/revision markup metadata scanning without exposing comment or revision text.
- Added package safety checks for ZIP paths, size limits, relationships, macro policy, unknown part preservation, and touched-part post-edit validation.
- Made NuGet packaging explicit and Release-only, producing both `.nupkg` and `.snupkg` artifacts under `artifacts/nuget/`.
- Added no-content private validation tooling for ignored local `.docx` cases.
- Added initial documentation for CLI usage, patch syntax, diagnostics, validation, and architecture.
- Changed `DocxSectionInfo.Orientation` from `string` to the `DocxOrientation` enum (`Portrait`/`Landscape`); text, JSON, and patch wire values stay lowercase strings.
- Changed change-target annotation (`TargetStatus`/`TargetSource`/`TargetReason`) and field `RefreshPolicy` from `string` to enums (`DocxTargetStatus`, `DocxTargetSource`, `DocxTargetReason`, `DocxRefreshPolicy`); text and JSON wire values stay lowercase strings.
- Documented the full public API surface (missing-docs warning enforced by the build) and cataloged every diagnostic code in `docs/diagnostics.md`.
- Added `docxedit version`, catalog-sourced usage errors with help pointers, and unknown-command guidance.
- Changed numbering label `LabelStatus`/`Source` from `string` to enums (`DocxLabelStatus`, `DocxLabelSource`); text and JSON wire values stay lowercase strings.
- Changed table `VerticalMerge` from `string` to the `DocxVerticalMerge` enum (`Restart`/`Continue`, null when unmerged); text and JSON wire values stay lowercase strings.
- Added the public `DocxChangeId` struct for change-record IDs (`M.CH0001`, `H001.CH0002`, comments `C001.CH0001`); change-record text and JSON shapes are unchanged.
- Changed bookmark, content-control, field, and hyperlink `Id` plus their target-ID fields from `string` to the `DocxTargetId` struct (`required` IDs, nullable targets); text and JSON wire values are unchanged.
