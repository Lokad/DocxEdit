# DocxEdit

DocxEdit is a stream-first .NET library and local CLI for inspecting and editing `.docx` files with deterministic, reviewable patch files.

The package identity is `Lokad.DocxEdit`. The production library has no NuGet dependencies beyond the .NET platform libraries.

## Status

This is a pre-release editor. It supports useful structural reads and a focused set of edits, but it is not a complete WordprocessingML implementation.

Implemented areas include:

- safe ZIP/package loading without filesystem extraction;
- stable IDs for paragraphs, tables, cells, images, sections, headers, and footers;
- styles, media, outline, find, dump, and tracked-change markup summaries;
- patch operations for paragraph text, blocks, styles, simple tables, inline images, and basic sections;
- no-content tracked-change viewing through `changes`;
- post-edit validation for touched XML parts before writing output.

Known limits include numbering, comments, bookmarks, content controls, fields, hyperlinks, floating images, complex tables, real tracked-change generation, and full OOXML schema validation.

## Build And Test

```powershell
dotnet build DocxEdit.slnx
dotnet test DocxEdit.slnx
dotnet pack src/DocxEdit/DocxEdit.csproj -c Release
```

Generated packages are written under ignored `artifacts/nuget/`.

## CLI Quick Start

```powershell
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- read report.docx
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- changes report.docx
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- check report.docx edits.docxpatch
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- apply report.docx edits.docxpatch -o report.edited.docx
```

See [docs/cli.md](docs/cli.md), [docs/patch-format.md](docs/patch-format.md), [docs/diagnostics.md](docs/diagnostics.md), and [docs/validation.md](docs/validation.md).

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

Private documents belong under ignored `private-cases/` and must never be committed.
