# DocxEdit

DocxEdit is a stream-first .NET library and local CLI for inspecting and editing `.docx` files with deterministic, reviewable patch files.

The package identity is `Lokad.DocxEdit`. The production library has no NuGet dependencies beyond the .NET platform libraries.

## DocxPatch DSL

The core editing interface is `.docxpatch`: a small text DSL for describing Word document edits without touching raw WordprocessingML. The CLI helps an agent discover stable targets in the document, then the patch file describes what should change.

First, inspect the document and locate a target:

```text
docxedit --help
docxedit read report.docx --summary
docxedit validate report.docx
docxedit find report.docx "old wording"
docxedit dump report.docx --id M.P0004 --runs
```

Then write a patch against the stable target ID:

```text
docxpatch 1

op replace-text
target M.P0004
expect-text <<<
old wording in the paragraph
>>>
find old wording
with new wording
end
```

Validate before writing a new `.docx`:

```text
docxedit check report.docx edits.docxpatch
docxedit apply report.docx edits.docxpatch --output report.edited.docx
```

Patch operations are explicit and guarded. A table-cell edit can assert the expected table shape:

```text
op set-cell
target M.T0001.R02.C03
expect-row-count 4
expect-column-count 3
text <<<
updated cell text
>>>
end
```

Image edits use document image IDs and external assets:

```text
op replace-image
target M.I0001
asset chart.png
expect-content-type image/png
alt Updated chart
end
```

Track-change behavior is controlled at check/apply time:

```text
docxedit apply report.docx edits.docxpatch --output report.edited.docx --track-changes require --author Agent
```

Fresh agents are expected to rely on `docxedit help patch`, `docxedit help changes`, and `docxedit help dump` for the exact syntax. For product integrations, the same guidance is available from the NuGet library through `DocxHelp.Catalog`, and the CLI text output is reusable through `DocxTextRenderer`.

## Library Quick Start

```csharp
using DocxEdit;

await using Stream input = File.OpenRead("report.docx");
DocxReadResult read = new DocxEditor().Read(input);

await using Stream applyInput = File.OpenRead("report.docx");
using var patch = File.OpenText("edits.docxpatch");
await using Stream edited = File.Create("report.edited.docx");
DocxApplyResult result = new DocxEditor().Apply(applyInput, patch, edited);
```

Input and output streams are left open by default. Set `LeaveInputOpen` or `LeaveOutputOpen` to `false` on the relevant options when the editor should dispose them.

Agent-facing help, output formatting, and privacy-safe workflow presets are available from the library:

```csharp
string help = DocxHelp.RenderTopic("changes");

await using Stream contextInput = File.OpenRead("report.docx");
DocxContextResult context = new DocxEditor().Context(
    contextInput,
    "M.P0004",
    DocxPrivacyPresets.ContextMetadataOnly);
string contextText = DocxTextRenderer.RenderContext(context);
```

## CLI Quick Start

```powershell
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- read report.docx
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- changes report.docx
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- check report.docx edits.docxpatch
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- apply report.docx edits.docxpatch -o report.edited.docx
```

## Documentation

See [docs/cli.md](docs/cli.md), [docs/patch-format.md](docs/patch-format.md), [docs/diagnostics.md](docs/diagnostics.md), and [docs/validation.md](docs/validation.md).

## Build And Test

```powershell
dotnet build DocxEdit.slnx
dotnet test DocxEdit.slnx
dotnet pack src/DocxEdit/DocxEdit.csproj -c Release
```

Generated packages are written under ignored `artifacts/nuget/`.

## Status

This is a pre-release editor. It supports useful structural reads and a focused set of edits, but it is not a complete WordprocessingML implementation.

Implemented areas include:

- safe ZIP/package loading without filesystem extraction;
- stable IDs for paragraphs, tables, cells, images, sections, headers, and footers;
- resolved paragraph numbering/list metadata, including abstract numbering IDs, formats,
  level text, style-linked list sources, start/suffix metadata, and visible labels for
  deterministic decimal, letter, roman, bullet, and nested `lvlText` patterns;
- merged/nested/styled table and row/grid read metadata, including merge groups and
  vertical-merge root cells;
- styles, media, outline, find, dump, and tracked-change markup summaries;
- bookmark and content-control selectors for safe paragraph targeting;
- patch operations for paragraph text, blocks, styles, simple main/header/footer tables,
  inline images and image metadata, and basic sections;
- no-content tracked-change and comment markup viewing through `changes`;
- opt-in bounded comment body snippets through `changes`;
- comment body editing and deletion through explicit patch operations keyed by
  `comment:<id>` or comment body IDs;
- comment anchor/context metadata and metadata-only comment body `dump`/`context`
  targets;
- comment resolution metadata from `commentsExtended.xml`, plus `resolve-comment` and
  `reopen-comment` for comments that already have modern extension records;
- bookmark and content-control read/context metadata with selector guidance, including
  duplicate selector candidate IDs, checkbox state, dropdown/combo item counts, and
  date settings;
- safe bookmark/content-control patch operations for plain-text content controls,
  checkbox state, dropdown/combo selections, date values, bookmark rename/delete,
  and simple same-paragraph bookmark ranges;
- simple and complex field read/context metadata, explicit dirty/lock flag patch
  operations, and field-update marking after edits;
- hyperlink read/context/dump metadata for external links, internal anchors, broken relationship IDs,
  URI scheme validation, missing/duplicate internal anchors, target frames, and history flags;
- hyperlink patch operations for target URI/anchor updates, tooltip/frame/history updates,
  display text updates, insertion, and unlinking while preserving display runs;
- check/apply operation reports with affected row/cell summaries for table edits;
- inline and anchored image metadata with layout kind, size, wrap mode, wrap distances,
  anchor positioning, aspect-lock, crop percentages, alt text, and containing target;
- safe image patch operations for media replacement, alt/title/name metadata, extents,
  anchored wrap mode/distances, anchored positioning, and crop percentages;
- tracked-change output for simple text replacement, whole-paragraph replacement,
  paragraph insertion/deletion, paragraph style changes, and simple table-cell text
  replacement with author, timestamp, and revision IDs;
- operation-level track-change capability metadata in the shared help catalog,
  including tracked and preserve-only operation classifications;
- structural package validation through `validate`, including known part roots, paired
  ranges, complex field balance, comment extension consistency, drawing relationships,
  drawing property ID uniqueness, duplicate semantic selector warnings, and basic
  table shape;
- post-edit validation for touched XML parts before writing output.

Known limits include exotic/custom visual numbering expansion beyond deterministic
decimal, letter, roman, bullet, and nested `lvlText` labels, comment creation,
resolution creation when modern extension metadata is absent, and full threaded comment models, advanced
bookmark/content-control editing beyond plain-text controls, checkbox toggles,
dropdown/combo selections, date values, bookmark rename/delete, and simple same-paragraph bookmark ranges, field result recalculation and field-specific
edit workflows beyond dirty/lock flags,
advanced hyperlink edit workflows beyond target URI/anchor/text/tooltip/frame/history,
relative/UNC/file URI policies, advanced floating-image edit operations beyond
size/wrap/position/crop metadata, complex
merged/nested table edit transformations, full tracked-change edit coverage, and full
ISO/IEC 29500 schema validation beyond DocxEdit's layered structural invariants.
