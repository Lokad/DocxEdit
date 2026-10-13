# DocxEdit

DocxEdit is a .NET 10 library for creating, inspecting, and editing Word (`.docx`)
files through deterministic `.docxpatch` scripts. It works with streams, requires
no Word installation, and has no runtime package dependencies beyond .NET.

```sh
dotnet add package Lokad.DocxEdit --version 0.2.0
```

The package contains the library. Run the development CLI from source; see
[CLI guidance](docs/cli.md).

## Editing a document

Inspect the document with `DocxEditor.Read`, `Outline`, or `Find` to discover
targets, then write a patch using their IDs and guards. For example, if the first
paragraph contains exactly `Draft report`, save this as `edits.docxpatch`:

```text
docxpatch 1

op replace-text
target M.P0001
expect-text Draft report
find Draft
with Final
end
```

`DocxEditor.Check` validates a patch against the input without writing output.
Apply it to a staging stream and publish only when `Success` is true:

```csharp
using Lokad.DocxEdit;

await using Stream input = File.OpenRead("report.docx");
using var patch = File.OpenText("edits.docxpatch");
await using var staged = new MemoryStream();
DocxApplyResult result = new DocxEditor().Apply(input, patch, staged);
if (!result.Success)
{
    foreach (DocxDiagnostic diagnostic in result.Diagnostics)
    {
        Console.WriteLine(diagnostic.Code + ": " + diagnostic.Message);
    }

    return;
}

staged.Position = 0;
await using Stream output = File.Create("report.edited.docx");
await staged.CopyToAsync(output);
```

Caller-owned input and output streams stay open by default; change the relevant
`LeaveInputOpen` or `LeaveOutputOpen` option to transfer ownership. Handle file I/O
errors in your application. The copy above can leave a partial file if publication
fails. To preserve an existing destination, write to a temporary file in the same
directory and move it over the destination after a successful copy. The source CLI
already stages file publication this way.

## Supported workflows

- **Create:** `new DocxEditor().Create(output)` produces an empty A4 portrait
  document with one-inch margins and Normal/Heading1–Heading9 styles. Use
  `DocxCreateOptions` for US Letter or landscape. Check the returned `Success` and
  `Diagnostics` and stage output before publication, as above.
- **Inspect:** read document structure, styles, tables, sections, images, and
  review markup. Section dimensions and margins are stored values in twips
  (1/1440 inch), with nulls for unspecified or non-integer values.
- **Edit:** guarded operations cover text, paragraphs, tables, rows, styles,
  images, bookmarks, hyperlinks, comments, fields, content controls, and sections,
  including supported header/footer content. Native equations support discovery,
  OMML extraction, and whole-object edits through a bounded LaTeX-like syntax;
  see [equation guidance](docs/equations.md).
- **Review:** tracked-change modes are `off`, `suggest`, and `require`.
  Supported edits generate revision markup; unsupported shapes produce explicit
  diagnostics. Comment annotations remain permitted under `require`. See the
  [operation support matrix](docs/patch-format.md#track-changes).
- **Images:** insert, replace, size, crop, and extract PNG, JPEG, SVG, and still
  WebP while preserving encoded bytes. Process bytes externally and supply
  replacements through `AssetProvider`; the library does not read asset files.
  SVG insertion generates no raster preview; animated WebP assets are rejected.
  Reader support varies; see [image guidance](docs/patch-format.md#images).
  Format interchange updates content types; package part paths are opaque and
  may retain the original extension.

Target IDs describe physical document order at inspection time. They are
repeatable for the same bytes and scanner version, but can change after edits.
Within a patch, explicit IDs bind to the input snapshot, semantic selectors
resolve against current content, and aliases address newly created objects.
Re-read the output for its current IDs; operation reports also provide `FinalId`
for surviving targets.

DocxEdit does not calculate layout, pagination, or rendered output. Structural
validation and optional Word open/save tests cover bounded cases and do not
guarantee fidelity for arbitrary documents.

## Documentation and integration

- [CLI workflow and commands](docs/cli.md)
- [Patch syntax, selectors, guards, and operations](docs/patch-format.md)
- [Diagnostics](docs/diagnostics.md) and [validation](docs/validation.md)
- [Current status](docs/status.md) and [changelog](CHANGELOG.md)
- [Hosting commands over your own storage](docs/hosting.md)

`DocxCommand.RunAsync` with an `IDocxCommandHost` reuses the CLI commands over your
own storage and standard streams. `DocxHelp.Catalog` exposes the command guidance,
`DocxTextRenderer` provides text output, and `DocxJson.CreateOptions` provides the
shared JSON wire format.

## Build from source

```sh
dotnet build Lokad.DocxEdit.slnx
dotnet test tests/Lokad.DocxEdit.Tests/Lokad.DocxEdit.Tests.csproj -c Release
dotnet pack src/Lokad.DocxEdit/Lokad.DocxEdit.csproj -c Release
```

Release packs produce `.nupkg` and `.snupkg` files under ignored `artifacts/nuget`.
The package includes XML documentation, README, changelog, license, and icon.
No CLI tool package is published. See the [release procedure](docs/release.md).

CI runs Release library and package-consumer tests on Windows and Ubuntu with
warnings treated as errors. Word compatibility tests require an explicit opt-in
on Windows with Word installed; see [test layers](docs/test-layers.md).
