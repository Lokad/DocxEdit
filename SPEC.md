# `docxedit` implementation specification

## 1. Project purpose

Implement **`docxedit`** as a C#/.NET 10 library for reading, exploring, validating, and editing `.docx` files through a deterministic patch DSL.

The project must produce:

1. A NuGet-packaged library published as **`Lokad.DocxEdit`**.
2. A non-packaged CLI project: **`docxedit`**, used only as a test/debug harness for coding agents.
3. xUnit tests.
4. Optional Windows Office integration tests for fixture generation and compatibility validation.

The library must be **command-line agnostic** and must support **stream-based operation**. It must be possible to read and edit a `.docx` without direct access to the file system.

The core library must have **no production dependencies beyond the .NET platform/system libraries**. Do **not** use the Open XML SDK, `python-docx`, Pandoc, Mammoth, third-party CLI parsers, YAML parsers, JSON libraries outside `System.Text.Json`, or any other NuGet package in the production library.

Use `.NET 10` and target `net10.0`. Microsoft documents .NET 10 as an LTS release and the successor to .NET 9. ([Microsoft Learn][1]) The solution should use `.slnx`; starting with .NET 10, `dotnet new sln` creates SLNX-format solution files by default. ([Microsoft Learn][2])

---

## 2. Key design principle

The library is **not** a Word clone and must not attempt pixel-perfect document layout editing.

It should provide a safe, deterministic editing layer over common corporate document structures:

* paragraphs
* headings
* runs
* simple text replacements
* tables
* simple table row/cell edits
* inline images
* headers and footers
* section metadata, including columns
* styles
* basic tracked-change output for simple text edits

A `.docx` file is an Open XML package made of parts, content types, and relationships; Microsoft documents that an Open XML package contains `[Content_Types].xml`, path-named parts such as `/word/theme/theme1.xml`, and relationship parts ending in `.rels`. ([Microsoft Learn][3]) WordprocessingML content is XML; the document body contains block-level elements such as paragraphs, paragraphs contain runs, and runs contain text. ([Microsoft Learn][4]) Tables are block-level content represented by `w:tbl`, containing rows `w:tr` and cells `w:tc`. ([Microsoft Learn][5])

Because the library cannot depend on the Open XML SDK, implement package and XML handling directly using:

* `System.IO`
* `System.IO.Compression`
* `System.Xml`
* `System.Xml.Linq`
* `System.Text`
* `System.Text.Json`
* other BCL/system APIs

`System.IO.Compression.ZipArchive` is the intended ZIP container API. Microsoft’s ZipArchive examples show creating and writing ZIP entries through streams, which fits the stream-first requirement. ([Microsoft Learn][6])

---

## 3. Explicit non-goals

The library must document and enforce these non-goals:

* No pixel/page-coordinate addressing such as “page 3, right column, second paragraph.”
* No full Word layout engine.
* No full support for floating shapes.
* No SmartArt editing.
* No chart data editing.
* No equation editing beyond preserving existing XML.
* No VBA or macro editing.
* No guaranteed TOC, field, cross-reference, or page-number rendering updates.
* No arbitrary XML mutation as the default workflow.
* No full Open XML schema validator.
* No `expect-hash` feature. Do not implement it.

The library may preserve unsupported structures, report them, and avoid modifying them.

---

## 4. Solution layout

Create this structure:

```text
.gitignore
DocxEdit.slnx
Directory.Build.props
Directory.Build.rsp
Directory.Packages.props                # optional; no production package refs
README.md
CHANGELOG.md
LICENSE.txt
src/
  DocxEdit/
    DocxEdit.csproj
    Public/
    Ooxml/
    Model/
    Patch/
    Rendering/
    Validation/
  DocxEdit.Cli/
    DocxEdit.Cli.csproj
    Program.cs
tests/
  DocxEdit.Tests/
    DocxEdit.Tests.csproj
    Fixtures/
    Unit/
    Golden/
  DocxEdit.OfficeTests/
    DocxEdit.OfficeTests.csproj
    OfficeInterop/
tools/
  CheckDocxCase.ps1
  CheckPrivateCase.ps1
  ValidateDocxCases.ps1
edit-cases/
  cases/
  families/
private-cases/                          # ignored local-only cases
artifacts/                              # ignored generated outputs
docs/
  patch-format.md
  cli.md
  architecture.md
  diagnostics.md
  validation.md
```

Production project:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
    <Deterministic>true</Deterministic>
    <GenerateDocumentationFile>true</GenerateDocumentationFile>
    <PackageId>Lokad.DocxEdit</PackageId>
    <Version>0.1.0</Version>
    <Authors>Lokad</Authors>
    <Company>Lokad</Company>
    <Description>Stream-first .docx reader and patch editor for coding agents.</Description>
    <PackageReadmeFile>README.md</PackageReadmeFile>
    <PackageLicenseFile>LICENSE.txt</PackageLicenseFile>
    <PackageOutputPath>..\..\artifacts\nuget\</PackageOutputPath>
    <IncludeSymbols>true</IncludeSymbols>
    <SymbolPackageFormat>snupkg</SymbolPackageFormat>
    <GeneratePackageOnBuild>false</GeneratePackageOnBuild>
  </PropertyGroup>
  <ItemGroup>
    <None Include="..\..\README.md" Pack="true" PackagePath="\" Visible="false" />
    <None Include="..\..\CHANGELOG.md" Pack="true" PackagePath="\" Visible="false" />
    <None Include="..\..\LICENSE.txt" Pack="true" PackagePath="\" Visible="false" />
  </ItemGroup>
</Project>
```

The `DocxEdit` project must contain **no `PackageReference` entries**.

`Directory.Build.props` should centralize:

```xml
<Project>
  <PropertyGroup>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
    <Deterministic>true</Deterministic>
    <ContinuousIntegrationBuild Condition="'$(CI)' == 'true'">true</ContinuousIntegrationBuild>
    <LangVersion>latestMajor</LangVersion>
  </PropertyGroup>
</Project>
```

`Directory.Build.rsp` should contain:

```text
-tl:off
```

This keeps command-line builds on stable console logging for automation and coding agents.

`.gitignore` must ignore:

```text
artifacts/
private-cases/
bin/
obj/
```

Test projects may reference xUnit and the minimal runner/test SDK needed to execute xUnit tests. Those dependencies must remain test-only and must not become transitive dependencies of the `Lokad.DocxEdit` NuGet package.

The CLI project is not packaged with the `Lokad.DocxEdit` NuGet package. It may reference only `DocxEdit` and system libraries.

The CLI can be published as a single executable for local use. .NET supports single-file deployment for framework-dependent and self-contained apps; publishing requires a runtime identifier and `PublishSingleFile`. ([Microsoft Learn][7])

Example:

```bash
dotnet publish src/DocxEdit.Cli/DocxEdit.Cli.csproj \
  -c Release \
  -r win-x64 \
  --self-contained false \
  -p:PublishSingleFile=true
```

---

## 5. Public API requirements

The public API must be stream-first and command-line agnostic.

### 5.1 Main entry point

Expose a simple façade:

```csharp
namespace DocxEdit;

public sealed class DocxEditor
{
    public DocxReadResult Read(
        Stream input,
        DocxReadOptions? options = null,
        CancellationToken cancellationToken = default);

    public DocxOutlineResult Outline(
        Stream input,
        DocxOutlineOptions? options = null,
        CancellationToken cancellationToken = default);

    public DocxFindResult Find(
        Stream input,
        string query,
        DocxFindOptions? options = null,
        CancellationToken cancellationToken = default);

    public DocxDumpResult Dump(
        Stream input,
        string targetId,
        DocxDumpOptions? options = null,
        CancellationToken cancellationToken = default);

    public DocxStylesResult Styles(
        Stream input,
        DocxStylesOptions? options = null,
        CancellationToken cancellationToken = default);

    public DocxMediaResult Media(
        Stream input,
        DocxMediaOptions? options = null,
        CancellationToken cancellationToken = default);

    public DocxPatch ParsePatch(
        TextReader patchReader,
        DocxPatchParseOptions? options = null,
        CancellationToken cancellationToken = default);

    public DocxCheckResult Check(
        Stream input,
        TextReader patchReader,
        DocxEditOptions? options = null,
        CancellationToken cancellationToken = default);

    public DocxApplyResult Apply(
        Stream input,
        TextReader patchReader,
        Stream output,
        DocxEditOptions? options = null,
        CancellationToken cancellationToken = default);
}
```

The library must not require file paths. The CLI may provide path-based wrappers. Every long-running public operation must observe the cancellation token while loading the package, parsing XML, scanning document stories, resolving assets, validating, and writing output.

### 5.2 Asset provider

Patches need to reference image assets. Since the library must work without file-system access, asset references are logical names resolved through an asset provider.

```csharp
namespace DocxEdit;

public interface IDocxAssetProvider
{
    bool TryOpen(
        string reference,
        out Stream stream,
        out string? contentTypeHint,
        out string? fileNameHint);
}
```

`DocxEditOptions` includes:

```csharp
public sealed class DocxEditOptions
{
    public TrackChangesMode TrackChanges { get; init; } = TrackChangesMode.Off;
    public string Author { get; init; } = "docxedit";
    public DateTimeOffset TimestampUtc { get; init; } = DateTimeOffset.UtcNow;

    public IDocxAssetProvider? AssetProvider { get; init; }

    public bool LeaveInputOpen { get; init; } = true;
    public bool LeaveOutputOpen { get; init; } = true;

    public int MaxZipEntries { get; init; } = 10_000;
    public long MaxUncompressedBytes { get; init; } = 512L * 1024 * 1024;
    public long MaxSinglePartBytes { get; init; } = 128L * 1024 * 1024;

    public bool AllowMacroEnabledDocuments { get; init; } = false;
    public bool MarkFieldsDirtyWhenEditing { get; init; } = true;
}
```

### 5.3 Result model

Every public method must return diagnostics instead of throwing for expected document/patch problems.

Throw only for programmer errors such as `null` arguments, non-readable streams, or non-writable output streams.

```csharp
public enum DocxSeverity
{
    Info,
    Warning,
    Error
}

public sealed record DocxDiagnostic(
    DocxSeverity Severity,
    string Code,
    string Message,
    string? TargetId = null,
    string? PartName = null,
    string? Story = null,
    string? Feature = null,
    string? Fallback = null,
    int? OperationIndex = null,
    int? Line = null,
    int? Column = null);

public abstract record DocxOperationResult
{
    public required bool Success { get; init; }
    public required IReadOnlyList<DocxDiagnostic> Diagnostics { get; init; }
}
```

`Check` and `Apply` must include operation-level results:

```csharp
public sealed record DocxPatchOperationReport(
    int Index,
    string OperationName,
    string? Target,
    bool Success,
    IReadOnlyList<DocxDiagnostic> Diagnostics);
```

---

## 6. Internal package model

Implement an internal package layer, not exposed as a general-purpose Open XML SDK.

```csharp
internal sealed class OoxmlPackage
{
    public IReadOnlyDictionary<string, OoxmlPart> Parts { get; }
    public OoxmlPart ContentTypesPart { get; }
    public OoxmlContentTypes ContentTypes { get; }

    public static OoxmlPackage Load(
        Stream input,
        DocxEditOptions options,
        CancellationToken cancellationToken = default);

    public OoxmlPart? GetPart(string partName);
    public IReadOnlyList<OoxmlRelationship> GetRelationships(
        string sourcePartName,
        CancellationToken cancellationToken = default);

    public void Save(Stream output, CancellationToken cancellationToken = default);
}
```

### 6.1 ZIP loading rules

* Accept readable streams.
* If the input stream is not seekable, copy it into a `MemoryStream`.
* Read each ZIP entry into memory, enforcing max entry count, max single-part uncompressed size, and max total uncompressed size before higher-level parsers touch XML or binary media.
* Preserve original entry names and unrelated entry bytes.
* Normalize package part names to a leading-slash form such as `/word/document.xml`.
* Reject entries with:

  * empty names
  * absolute paths
  * `..` segments
  * `.` segments
  * drive letters or `:` characters
  * backslash path traversal
  * duplicate names after normalization
* Never extract ZIP entries to disk in the production library.
* Ignore directory entries.
* Preserve ZIP entries with unknown content types as raw bytes when they are safe package parts; do not interpret them.
* Surface package-limit failures as stable `E0xxx` diagnostics instead of raw `ZipArchive` exceptions when called through public APIs.

### 6.2 XML loading rules

Load XML parts with:

* `XmlReaderSettings.DtdProcessing = DtdProcessing.Prohibit`
* `XmlResolver = null`
* `LoadOptions.PreserveWhitespace`
* `SaveOptions.DisableFormatting`

Preserve unknown elements and attributes.

Use `XDocument` only for parts that must be inspected or edited. Keep other parts as raw bytes.

All XML loading must go through one internal `SafeXml` helper so DTD/external-entity hardening cannot be bypassed accidentally.

### 6.3 Required parts

The loader must parse:

```text
/[Content_Types].xml
/_rels/.rels
/word/document.xml
/word/_rels/document.xml.rels
```

It must discover the main document part through the office document relationship in `/_rels/.rels`, not by assuming `/word/document.xml` always exists.

The scanner should also discover, when present:

```text
/word/styles.xml
/word/numbering.xml
/word/settings.xml
/word/comments.xml
/word/header*.xml
/word/footer*.xml
/word/media/*
/docProps/core.xml
/docProps/app.xml
```

### 6.4 Content types and relationships

Implement explicit parsers for `[Content_Types].xml` and `.rels` files.

Rules:

* Content type lookup must support both `<Default>` extension mappings and `<Override>` part mappings.
* Relationship part names must be derived from the source part, e.g. `/word/document.xml` -> `/word/_rels/document.xml.rels`; package-level relationships use `/_rels/.rels`.
* Internal relationship targets must be resolved relative to the source part and normalized to package part names.
* Relationship target resolution must reject targets that escape the package root.
* External relationships must be represented with `TargetMode = "External"` and no resolved package part.
* Public operations must never fetch external relationships.
* Duplicate relationship IDs inside one `.rels` part are validation errors.

---

## 7. OOXML constants

Create a single internal constants class:

```csharp
internal static class OoxmlNs
{
    public static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public static readonly XNamespace R =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static readonly XNamespace Rel =
        "http://schemas.openxmlformats.org/package/2006/relationships";

    public static readonly XNamespace A =
        "http://schemas.openxmlformats.org/drawingml/2006/main";

    public static readonly XNamespace Wp =
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

    public static readonly XNamespace Pic =
        "http://schemas.openxmlformats.org/drawingml/2006/picture";

    public static readonly XNamespace Ct =
        "http://schemas.openxmlformats.org/package/2006/content-types";

    public static readonly XNamespace Xml =
        "http://www.w3.org/XML/1998/namespace";
}
```

Relationship type constants:

```csharp
internal static class OoxmlRelTypes
{
    public const string OfficeDocument =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

    public const string Header =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";

    public const string Footer =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";

    public const string Image =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

    public const string Styles =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

    public const string Numbering =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";

    public const string Settings =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";

    public const string Comments =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
}
```

---

## 8. Logical document model

The scanner must convert relevant OOXML into a logical model used by read, dump, find, selectors, and patch operations.

### 8.1 Stories

A WordprocessingML document is composed of stories such as main document, headers, footers, comments, text boxes, footnotes, and endnotes. Microsoft documents main document, headers and footers, comments, text boxes, footnotes, and endnotes as WordprocessingML stories. ([Microsoft Learn][3])

Support these in v0.1:

```text
main document      editable
headers            editable for paragraph/text operations
footers            editable for paragraph/text operations
comments           read-only initially
footnotes          read-only initially
endnotes           read-only initially
text boxes         read-only initially
```

### 8.2 Target IDs

IDs are generated deterministically during scanning. They are not written into the `.docx`.

IDs are stable for the same document bytes and same scanner version.

Use these forms:

```text
M.P0001                         main-story top-level paragraph
M.T0001                         main-story top-level table
M.T0001.R02.C03                 table cell
M.T0001.R02.C03.P0001           paragraph inside a table cell
M.I0001                         image in main story
M.S0001                         section

H001.P0001                      first header part, paragraph
H001.T0001                      first header part, table
H001.I0001                      first header part, image

F001.P0001                      first footer part, paragraph
F001.T0001                      first footer part, table
F001.I0001                      first footer part, image
```

Counters are 1-based and zero-padded.

### 8.3 Paragraph model

```csharp
public sealed record DocxParagraphInfo
{
    public required string Id { get; init; }
    public required string StoryId { get; init; }
    public required string Text { get; init; }
    public string? StyleId { get; init; }
    public string? StyleName { get; init; }
    public int? HeadingLevel { get; init; }
    public bool ContainsField { get; init; }
    public bool ContainsHyperlink { get; init; }
    public bool ContainsCommentReference { get; init; }
    public bool ContainsTrackedChange { get; init; }
    public IReadOnlyList<DocxRunInfo> Runs { get; init; } = [];
    public IReadOnlyList<string> ImageIds { get; init; } = [];
}
```

### 8.4 Run model

```csharp
public sealed record DocxRunInfo
{
    public required string Id { get; init; }
    public required string Text { get; init; }
    public string? StyleId { get; init; }
    public bool Bold { get; init; }
    public bool Italic { get; init; }
    public bool Underline { get; init; }
    public bool IsHyperlink { get; init; }
    public bool IsFieldCode { get; init; }
    public bool IsInsertedRevision { get; init; }
    public bool IsDeletedRevision { get; init; }
}
```

Text extraction rules:

* `w:t` contributes its text.
* `w:tab` contributes `\t`.
* `w:br` and `w:cr` contribute `\n`.
* Field instruction text `w:instrText` is not visible text by default.
* Deleted revision text is omitted in final-view mode.
* Inserted revision text is included in final-view mode.
* The default read mode is final-view mode.

### 8.5 Table model

```csharp
public sealed record DocxTableInfo
{
    public required string Id { get; init; }
    public required string StoryId { get; init; }
    public string? StyleId { get; init; }
    public string? StyleName { get; init; }
    public required int RowCount { get; init; }
    public required int ColumnCount { get; init; }
    public required IReadOnlyList<DocxTableCellInfo> Cells { get; init; }
    public bool HasMergedCells { get; init; }
}
```

Cell model:

```csharp
public sealed record DocxTableCellInfo
{
    public required string Id { get; init; }
    public required int RowIndex { get; init; }
    public required int ColumnIndex { get; init; }
    public required string Text { get; init; }
    public int GridSpan { get; init; } = 1;
    public bool IsVerticalMergeStart { get; init; }
    public bool IsVerticalMergeContinuation { get; init; }
}
```

### 8.6 Image model

```csharp
public sealed record DocxImageInfo
{
    public required string Id { get; init; }
    public required string StoryId { get; init; }
    public required string RelationshipId { get; init; }
    public required string RelationshipPartName { get; init; }
    public required string TargetPartName { get; init; }
    public required string ContentType { get; init; }
    public string? FileName { get; init; }
    public string? AltText { get; init; }
    public long? WidthEmu { get; init; }
    public long? HeightEmu { get; init; }
    public bool IsInline { get; init; }
    public bool IsFloating { get; init; }
    public string? ParentParagraphId { get; init; }
}
```

v0.1 must support replacing image bytes for both inline and floating images when a relationship-backed image is found. v0.1 must support inserting only inline images.

### 8.7 Section model

```csharp
public sealed record DocxSectionInfo
{
    public required string Id { get; init; }
    public string? StartsAfterParagraphId { get; init; }
    public string? PageOrientation { get; init; }
    public int? ColumnCount { get; init; }
    public string? ColumnSpaceTwips { get; init; }
    public int? PageWidthTwips { get; init; }
    public int? PageHeightTwips { get; init; }
}
```

---

## 9. Read/explore/probe output format

The library must expose structured results and a deterministic agent-readable text renderer.

The CLI will use these renderers.

### 9.1 `read`

Purpose: produce a compact structural view.

Example CLI output:

```text
# docxedit read input.docx
# Use: docxedit dump input.docx --id M.P0004 --runs
# Use: docxedit check input.docx edits.docxpatch
# Use: docxedit apply input.docx edits.docxpatch -o output.docx

document:
  stories: main, header[1], footer[1]
  sections: 2
  top-level-paragraphs: 18
  top-level-tables: 3
  images: 4

[M.S0001] section
  orientation: portrait
  columns: 1

[M.P0001] paragraph style="Title" styleId="Title"
  text: "Q4 Board Report"

[M.P0002] paragraph style="Subtitle" styleId="Subtitle"
  text: "Prepared for the Executive Committee"

[M.P0003] heading level=1 style="Heading 1" styleId="Heading1"
  text: "Executive Summary"

[M.P0004] paragraph style="Normal" styleId="Normal"
  text: "Revenue increased by 8.4% compared with the prior quarter."

[M.T0001] table style="Light Shading" styleId="LightShading" rows=4 cols=3
  [M.T0001.R01.C01] text: "Metric"
  [M.T0001.R01.C02] text: "Q3"
  [M.T0001.R01.C03] text: "Q4"
  [M.T0001.R02.C01] text: "Revenue"
  [M.T0001.R02.C02] text: "$12.4m"
  [M.T0001.R02.C03] text: "$13.5m"

[M.P0010] paragraph style="Normal" styleId="Normal"
  text: ""
  image: M.I0001 rel=rId9 type=image/png inline width=2926080emu height=1645920emu alt="Revenue chart"
```

### 9.2 `outline`

Purpose: show only sections, headings, tables, images, headers, and footers.

```text
[M.S0001] section columns=1 orientation=portrait
[M.P0001] title "Q4 Board Report"
[M.P0003] H1 "Executive Summary"
[M.T0001] table rows=4 cols=3 after=M.P0003
[M.P0011] H1 "Financial Summary"
[M.S0002] section columns=2 orientation=portrait
[H001.P0001] header paragraph "Confidential"
[F001.P0001] footer paragraph "Page "
```

### 9.3 `find`

Purpose: find targetable text.

```text
# docxedit find input.docx "Revenue"
matches: 2

[M.P0004] paragraph
  text: "Revenue increased by 8.4% compared with the prior quarter."

[M.T0001.R02.C01] table-cell
  text: "Revenue"
```

### 9.4 `dump`

Purpose: inspect a single target in detail.

```text
[M.P0004] paragraph style="Normal" styleId="Normal"
text: "Revenue increased by 8.4% compared with the prior quarter."

runs:
  [M.P0004.R0001] text="Revenue increased by " bold=false italic=false
  [M.P0004.R0002] text="8.4%" bold=true italic=false
  [M.P0004.R0003] text=" compared with the prior quarter." bold=false italic=false

flags:
  contains-field: false
  contains-hyperlink: false
  contains-tracked-change: false
```

### 9.5 `styles`

Purpose: list styles by style ID and display name.

```text
paragraph styles:
  styleId="Normal" name="Normal"
  styleId="Heading1" name="Heading 1"
  styleId="Title" name="Title"

character styles:
  styleId="Hyperlink" name="Hyperlink"

table styles:
  styleId="LightShading" name="Light Shading"
```

### 9.6 `media`

Purpose: list images and allow stream-based extraction.

CLI output:

```text
[M.I0001] image/png inline target="/word/media/image1.png" parent=M.P0010 width=2926080emu height=1645920emu alt="Revenue chart"
[M.I0002] image/jpeg floating target="/word/media/image2.jpeg" parent=M.P0014 width=1828800emu height=1828800emu alt=""
```

Library extraction API should allow copying a media item to a caller-provided stream:

```csharp
public void CopyMediaTo(Stream input, string imageId, Stream output);
```

---

## 10. Patch DSL: `.docxpatch`

### 10.1 Format goals

The patch DSL must be:

* line-oriented
* easy for coding agents to edit with textual patches
* deterministic
* strict enough to catch mistakes early
* free of indentation-sensitive semantics
* usable without JSON/YAML escaping pain
* validatable before writing output

The patch file extension is:

```text
.docxpatch
```

### 10.2 Required preamble

Every patch starts with:

```text
docxpatch 1
```

### 10.3 Comments

A line whose first non-whitespace character is `#` is a comment.

Blank lines are ignored between fields and operations.

### 10.4 Operation block

General form:

```text
op <operation-name>
<field-name> <value>
<field-name> <<<
multiline value
>>>
end
```

### 10.5 Strings

Support two string forms.

Quoted string:

```text
target heading "Executive Summary"
style "Normal"
```

Escapes:

```text
\"  double quote
\\  backslash
\n  newline
\t  tab
```

Heredoc string:

```text
text <<<
Revenue increased by 9.1% compared with the prior quarter.
>>>
```

Heredoc content:

* begins after the newline following `<<<`
* ends at a line containing only `>>>`
* normalizes internal line endings to `\n`
* does not include the terminating `>>>` line

### 10.6 Selectors

Supported selectors:

```text
target M.P0004
target M.T0001.R02.C03
target M.I0001
target H001.P0001
target F001.P0001

target heading "Executive Summary"
target paragraph containing "Revenue increased"
target table after heading "Financial Summary"
target bookmark "ClientName"
target content-control tag "client_name"
```

Selector rules:

* A selector must resolve to exactly one target unless the operation explicitly supports multiple targets.
* Case-sensitive exact matching by default.
* Whitespace is normalized for semantic text selectors.
* On zero matches, return an error with nearest candidate targets.
* On multiple matches, return an error listing the matched targets.

### 10.7 Guards

The DSL must not support `expect-hash`.

Supported guards:

```text
expect-text <<<
exact visible text
>>>

expect-contains "visible substring"
```

For destructive operations, `expect-text` is mandatory unless the operation has an equally specific structural guard. Destructive operations include:

```text
replace-text
replace-paragraph
delete-block
set-cell
delete-row
replace-image
delete-image
```

---

## 11. Patch operations v0.1

### 11.1 `replace-text`

Replace text inside one paragraph.

```text
op replace-text
target M.P0004
expect-text <<<
Revenue increased by 8.4% compared with the prior quarter.
>>>
find <<<
8.4%
>>>
with <<<
9.1%
>>>
preserve-runs true
end
```

Rules:

* Target must be a paragraph.
* `expect-text` is mandatory.
* `find` must occur exactly once unless `occurrence` is specified.
* If replacement is within a single run, split the run and preserve run properties.
* If replacement spans simple adjacent runs, preserve the first matched run’s properties for the replacement.
* Fail if replacement crosses:

  * field boundaries
  * hyperlink boundaries
  * comment range boundaries
  * structured document tag boundaries
  * unsupported revision boundaries
* Preserve `xml:space="preserve"` when replacement text has leading/trailing spaces or repeated spaces.

Optional field:

```text
occurrence 2
```

### 11.2 `replace-paragraph`

Replace all visible text in a paragraph.

```text
op replace-paragraph
target M.P0004
expect-text <<<
Revenue increased by 8.4% compared with the prior quarter.
>>>
style "Normal"
text <<<
Revenue increased by 9.1% compared with the prior quarter.
>>>
end
```

Rules:

* Target must be a paragraph.
* Preserve paragraph properties unless `style` is supplied.
* Replace paragraph content with one or more runs.
* Preserve bookmarks and comment anchors only when safe; otherwise fail with diagnostic.

### 11.3 `insert-before`

```text
op insert-before
target M.P0004
style "Normal"
text <<<
This paragraph is inserted before the revenue paragraph.
>>>
end
```

Rules:

* Target must be a paragraph or table.
* Insert a new paragraph before the target.
* If `style` is omitted, use target paragraph style when target is a paragraph; otherwise use `Normal`.

### 11.4 `insert-after`

```text
op insert-after
target heading "Executive Summary"
style "Normal"
text <<<
The quarter closed ahead of plan, with growth concentrated in enterprise accounts.
>>>
end
```

Rules:

* Target must be a paragraph or table.
* Insert a new paragraph after the target.
* If target is a paragraph with section properties, do not insert after the section break incorrectly; insert before the section break if necessary and warn.

### 11.5 `delete-block`

```text
op delete-block
target M.P0008
expect-text <<<
This paragraph should be removed.
>>>
end
```

Rules:

* Target must be a paragraph or top-level table.
* For paragraphs, delete the paragraph XML element.
* For tables, delete the table XML element.
* Do not delete the only paragraph in a table cell; replace it with an empty paragraph instead.

### 11.6 `set-style`

```text
op set-style
target M.P0004
style "Heading 2"
end
```

Rules:

* Target must be a paragraph.
* Style may be specified by style name or style ID.
* If ambiguous, fail and list candidates.
* Preserve paragraph content.

### 11.7 `set-cell`

```text
op set-cell
target M.T0001.R02.C03
expect-text <<<
$13.5m
>>>
text <<<
$13.8m
>>>
end
```

Rules:

* Target must be a table cell.
* Preserve `w:tcPr`.
* Replace cell content with a single paragraph.
* If the cell contains multiple paragraphs, nested tables, images, or fields, fail unless `force true` is supplied.
* `force true` still preserves `w:tcPr`.

### 11.8 `append-row`

```text
op append-row
target M.T0001
cell <<<
Gross margin
>>>
cell <<<
42.1%
>>>
cell <<<
43.0%
>>>
end
```

Rules:

* Target must be a table.
* Only support rectangular tables without vertical merges in v0.1.
* Clone the last row’s row properties and cell properties.
* Number of `cell` fields must equal logical column count.
* Preserve table style and grid.

### 11.9 `insert-row-before`

```text
op insert-row-before
target M.T0001.R03
cell <<<New metric>>>
cell <<<Q3 value>>>
cell <<<Q4 value>>>
end
```

Row target form:

```text
M.T0001.R03
```

Rules are the same as `append-row`.

### 11.10 `insert-row-after`

Same as `insert-row-before`, but inserts after the row target.

### 11.11 `delete-row`

```text
op delete-row
target M.T0001.R03
expect-contains "Obsolete metric"
end
```

Rules:

* Target must be a table row.
* Do not allow deleting the only row of a table.
* Fail on merged-cell tables in v0.1 unless `force true`.

### 11.12 `replace-image`

```text
op replace-image
target M.I0001
asset "assets/revenue-chart.png"
preserve-size true
alt "Updated revenue chart"
end
```

Rules:

* Target must be an image.
* Resolve `asset` through `IDocxAssetProvider`.
* CLI resolves asset paths relative to the patch file directory.
* Support PNG and JPEG initially.
* Detect content type from magic bytes, not only extension.
* If `preserve-size true`, keep existing drawing extents.
* If `preserve-size false`, compute image dimensions and convert to EMUs.
* Replacing a floating image’s bytes is allowed; changing floating layout is not.
* Update content type declarations as needed.

### 11.13 `insert-image-after`

```text
op insert-image-after
target M.P0012
asset "assets/architecture.png"
width 5.5in
alt "Architecture diagram"
caption "Figure 3. Target architecture"
end
```

Rules:

* Target must be a paragraph.
* Insert an inline image in a new paragraph after target.
* Support PNG and JPEG initially.
* Width and height may be specified in:

  * `in`
  * `cm`
  * `pt`
  * `px`
  * `emu`
* If only width is given, preserve aspect ratio.
* If neither width nor height is given, use image pixel dimensions at 96 DPI.
* If `caption` is supplied, insert a following paragraph styled `Caption` if that style exists; otherwise `Normal`.

### 11.14 `set-image-alt`

```text
op set-image-alt
target M.I0001
alt "Updated revenue chart"
end
```

Rules:

* Update alt text on drawing properties where present.
* Preserve image bytes and size.

### 11.15 `delete-image`

```text
op delete-image
target M.I0001
end
```

Rules:

* Remove the drawing object.
* Remove orphaned image relationship and media part only if no other relationship references it.
* Preserve containing paragraph; if paragraph becomes empty, leave an empty paragraph.

### 11.16 `set-section-columns`

```text
op set-section-columns
target M.S0002
count 2
space 0.3in
end
```

Rules:

* Target must be a section.
* Support equal-width columns only.
* `count` must be between 1 and 4 in v0.1.
* Update or create `w:sectPr/w:cols`.
* Do not attempt to move content between columns.

### 11.17 `set-section-orientation`

```text
op set-section-orientation
target M.S0002
orientation landscape
end
```

Rules:

* Target must be a section.
* Orientation must be `portrait` or `landscape`.
* Update `w:pgSz/@w:orient`.
* If changing orientation and page width/height are known, swap width and height when needed.

---

## 12. Track changes

Track changes are an ambient apply option, not part of individual operations.

```csharp
public enum TrackChangesMode
{
    Off,
    Preserve,
    Suggest,
    Require
}
```

CLI:

```bash
docxedit apply input.docx edits.docxpatch \
  -o output.docx \
  --track-changes suggest \
  --author "Coding Agent"
```

Modes:

```text
Off
  Apply edits directly.

Preserve
  Preserve existing tracked-change markup but do not create new tracked changes.

Suggest
  Represent supported edits as tracked changes.
  Unsupported operations are applied directly with warnings.

Require
  Represent all edits as tracked changes.
  Fail if any operation cannot be represented as tracked changes.
```

Microsoft documents `w:trackRevisions` as the setting that specifies whether applications track document revisions; if omitted, revisions are not generated by changes to document contents. ([Microsoft Learn][8]) Microsoft also documents inserted run content as `w:ins`, which marks inline content as inserted revision content. ([Microsoft Learn][9])

v0.1 tracked-change support:

```text
replace-text          supported when replacement is within one paragraph and not across protected boundaries
replace-paragraph     supported as delete old runs + insert new runs
insert-before         supported
insert-after          supported
delete-block          supported for paragraph deletion only
set-cell              supported only when cell has a single simple paragraph
```

Unsupported tracked-change operations in v0.1:

```text
append-row
insert-row-before
insert-row-after
delete-row
replace-image
insert-image-after
set-image-alt
delete-image
set-section-columns
set-section-orientation
```

In `Suggest`, these unsupported operations apply directly and return warnings.

In `Require`, they fail.

Revision metadata:

* Use `DocxEditOptions.Author`.
* Use `DocxEditOptions.TimestampUtc`.
* Generate monotonically increasing `w:id` values by scanning existing revision IDs and incrementing from the max.

---

## 13. Validation behavior

### 13.1 `check`

`Check` must parse and validate a patch against an input document without writing output.

It must:

1. Load the package.
2. Scan the logical document.
3. Parse the patch.
4. Resolve selectors.
5. Validate guards.
6. Simulate operations against an in-memory copy.
7. Run internal package validation.
8. Return a report.

Example output rendered by CLI:

```text
docxedit check: OK

Input:
  input.docx
Patch:
  edits.docxpatch

Operations:
  1 replace-text        M.P0004          OK
  2 insert-after        heading(...)     OK
  3 set-cell            M.T0001.R02.C03  OK
  4 replace-image       M.I0001          OK

Warnings:
  W2101 M.P0004 contains multiple runs; replacement preserves the first matched run style.
  W5103 document contains fields; output marks fields dirty but does not render-update fields.

No output written.
```

### 13.2 `apply`

`Apply` must do everything `Check` does and then write a new `.docx` to the output stream.

It must never modify the input stream.

### 13.3 Internal validation

Because the production library cannot use the Open XML SDK, internal validation is not full schema validation. It must check:

* ZIP package is readable.
* Required package parts exist.
* Required relationship targets exist.
* XML parts touched by the library are well-formed.
* Relationship IDs referenced from edited XML exist.
* Image parts referenced by drawings exist.
* Added image content types exist.
* No duplicate ZIP entries.
* No relationship target path traversal.
* Main document root has expected WordprocessingML namespace.
* Edited tables still contain valid basic `w:tbl/w:tr/w:tc` nesting.
* Edited paragraphs still contain valid basic `w:p/w:r/w:t` nesting.

Return warning:

```text
W9001 internal validation is structural, not full ISO/IEC 29500 schema validation.
```

Only show this warning in verbose CLI output or machine-readable reports, not on every successful run unless requested.

---

## 14. Diagnostics

Diagnostics must be stable and actionable.

Every diagnostic code is a public contract once released. Prefer one durable code per observable behavior, and change messages without changing codes when wording improves.

Use code ranges:

```text
E0xxx package/load errors
E1xxx selector errors
E2xxx patch parse errors
E3xxx guard errors
E4xxx operation errors
E5xxx media errors
E6xxx track-change errors
E9xxx internal validation errors

W1xxx selector/read warnings
W2xxx text/run warnings
W3xxx table warnings
W4xxx operation fallback warnings
W5xxx unsupported, approximated, or read-only feature warnings
W9xxx validation limitations
```

Diagnostic records should include `PartName` when the issue is tied to an OOXML part, `Story` when it is tied to a story such as main document/header/footer/footnote, `Feature` for unsupported or approximated OOXML features, `Fallback` for the behavior the library used, and `OperationIndex` for patch operation diagnostics.

Unsupported or approximated feature diagnostics must be emitted as warnings during read/check/apply when they can affect the requested workflow. Repeated occurrences of the same unsupported feature in the same part should be aggregated into one warning unless the exact target list is useful for fixing the patch.

Required unsupported-feature warning coverage:

```text
W5001 comments detected
W5002 complex field detected
W5003 unsupported tracked-change markup detected
W5004 footnote or endnote reference detected
W5005 equation detected
W5006 OLE object detected
W5007 floating DrawingML detected
W5008 macro or VBA project detected
W5009 multi-column or unsupported section flow detected
W5010 external relationship detected and left untouched
```

These warnings do not mean the library may corrupt the document. They mean the feature is preserved, ignored, approximated, or made read-only according to the operation semantics.

Examples:

```text
E1201 selector matched 0 targets:
  target heading "Executive summary"

Closest headings:
  M.P0003 heading level=1 text="Executive Summary"
  M.P0041 heading level=1 text="Financial Summary"

Try:
  docxedit outline input.docx
  docxedit dump input.docx --id M.P0003
```

```text
E1202 selector matched 3 targets:
  target paragraph containing "Revenue"

Matches:
  M.P0004 text="Revenue increased by 8.4%..."
  M.T0001.R02.C01 text="Revenue"
  M.P0044 text="Revenue by region..."

Use a more specific target ID from docxedit read or docxedit find.
```

```text
E3201 guard failed for M.P0004.

Expected:
  "Revenue increased by 8.4% compared with the prior quarter."

Current:
  "Revenue increased by 9.1% compared with the prior quarter."

Refresh the target with:
  docxedit dump input.docx --id M.P0004 --runs
```

```text
E4207 replacement crosses a hyperlink boundary in M.P0018.

Matched text:
  "see Appendix B"

Reason:
  "Appendix B" is inside a hyperlink run.

Options:
  - target the full paragraph with replace-paragraph
  - inspect runs with:
    docxedit dump input.docx --id M.P0018 --runs
```

---

## 15. CLI test harness

The CLI is a wrapper around the library. It is not part of the `Lokad.DocxEdit` NuGet package.

Command name:

```text
docxedit
```

Required commands:

```text
docxedit -h
docxedit --help
docxedit help patch

docxedit read input.docx
docxedit outline input.docx
docxedit find input.docx "some text"
docxedit dump input.docx --id M.P0004 --runs
docxedit styles input.docx
docxedit media input.docx
docxedit media input.docx --extract M.I0001 --out image.png

docxedit check input.docx edits.docxpatch
docxedit apply input.docx edits.docxpatch -o output.docx
```

Options for `apply`:

```text
-o, --output <path>
--track-changes off|preserve|suggest|require
--author <name>
--timestamp-utc <iso-8601>
--report <path>
--diagnostics <path>
--strict
--json
--verbose
```

Options for read/probe commands:

```text
--runs
--headers-footers
--all-stories
--max-text <chars>
--diagnostics <path>
--strict
--json
```

`--diagnostics <path>` writes the full diagnostics array as JSON for any command. `--report <path>` remains the operation-level patch report for `check` and `apply`.

`--strict` must make an otherwise successful command return the strict-warning exit code when any warning or error diagnostic was emitted.

The CLI parser must be hand-written or implemented with system libraries only. Do not use a third-party command-line parser.

CLI exit codes:

```text
0 success
1 expected failure: invalid patch, selector failure, guard failure, validation failure
2 invalid CLI usage
3 success, but --strict saw one or more warning or error diagnostics
4 unexpected internal error
```

---

## 16. Patch examples

### 16.1 Update a paragraph value

```text
docxpatch 1

op replace-text
target M.P0004
expect-text <<<
Revenue increased by 8.4% compared with the prior quarter.
>>>
find <<<
8.4%
>>>
with <<<
9.1%
>>>
preserve-runs true
end
```

### 16.2 Insert paragraph after a heading

```text
docxpatch 1

op insert-after
target heading "Executive Summary"
style "Normal"
text <<<
The quarter closed ahead of plan, with growth concentrated in enterprise accounts.
>>>
end
```

### 16.3 Update a table cell

```text
docxpatch 1

op set-cell
target M.T0001.R02.C03
expect-text <<<
$13.5m
>>>
text <<<
$13.8m
>>>
end
```

### 16.4 Append a table row

```text
docxpatch 1

op append-row
target M.T0001
cell <<<
Gross margin
>>>
cell <<<
42.1%
>>>
cell <<<
43.0%
>>>
end
```

### 16.5 Replace an image

```text
docxpatch 1

op replace-image
target M.I0001
asset "assets/revenue-chart.png"
preserve-size true
alt "Updated revenue chart"
end
```

### 16.6 Set a section to two columns

```text
docxpatch 1

op set-section-columns
target M.S0002
count 2
space 0.3in
end
```

---

## 17. Text editing implementation details

### 17.1 Visible text map

For each paragraph, build a visible text map:

```csharp
internal sealed record TextSegment(
    XElement Element,
    int ElementTextStart,
    int ElementTextLength,
    int VisibleStart,
    int VisibleLength,
    TextSegmentKind Kind,
    XElement? OwningRun,
    XElement? ProtectedBoundary);
```

Kinds:

```text
Text
Tab
Break
InsertedText
DeletedText
FieldInstruction
```

Default view:

```text
Final
  include inserted text
  exclude deleted text
  exclude field instruction text
```

The replacement engine must use the text map to translate visible character offsets back to XML nodes.

### 17.2 Run splitting

When replacing text in a `w:t`:

1. Determine prefix, matched text, suffix.
2. Preserve original `w:rPr`.
3. Replace with up to three runs:

   * prefix run
   * replacement run
   * suffix run
4. Omit empty runs.
5. Set `xml:space="preserve"` on `w:t` when needed.

### 17.3 Protected boundaries

Replacement must fail if the match crosses these boundaries:

```text
w:hyperlink
w:fldSimple
field begin/separate/end sequence
w:sdt
w:commentRangeStart / w:commentRangeEnd
w:bookmarkStart / w:bookmarkEnd
w:ins / w:del when track-change mode would become ambiguous
```

v0.1 may support replacing entirely inside a hyperlink, but not across hyperlink boundaries.

### 17.4 Fields

When a document contains fields and edits are applied:

* Preserve field XML.
* Do not attempt to evaluate fields.
* If `MarkFieldsDirtyWhenEditing` is true, mark simple field structures dirty where straightforward.
* Return warning:

```text
W5103 document contains fields; output may require Word to update fields.
```

---

## 18. Table implementation details

### 18.1 Logical columns

Compute logical column count using:

1. `w:tblGrid/w:gridCol` count when available.
2. Otherwise, max sum of `w:gridSpan` across rows.
3. Treat missing `w:gridSpan` as 1.

### 18.2 Merged cells

Detect:

```text
w:gridSpan
w:vMerge
```

v0.1 behavior:

* `set-cell` works for a physical cell unless it is a vertical merge continuation.
* Row insert/delete operations fail on tables with vertical merges.
* Row insert/delete operations fail on non-rectangular tables unless `force true`.

### 18.3 Cell replacement

`set-cell` must preserve:

```text
w:tcPr
```

It may replace all block-level cell content with:

```xml
<w:p>
  <w:r>
    <w:t>new text</w:t>
  </w:r>
</w:p>
```

When the replacement text contains `\n`, use `w:br`.

---

## 19. Image implementation details

### 19.1 Supported image types

v0.1 must support:

```text
PNG   image/png
JPEG  image/jpeg
```

Optional later:

```text
GIF
BMP
TIFF
EMF
WMF
SVG
```

### 19.2 Content type detection

Detect by magic bytes:

```text
PNG:  89 50 4E 47 0D 0A 1A 0A
JPEG: FF D8 FF
```

Do not trust file extension alone.

### 19.3 Dimensions

Implement small PNG and JPEG dimension readers with system APIs only.

Conversion constants:

```text
1 inch = 914400 EMU
1 point = 12700 EMU
1 cm = 360000 EMU
default pixel DPI = 96
```

### 19.4 Insert inline image

Generate minimal DrawingML inline image XML. The output must be accepted by Microsoft Word.

Implementation must:

1. Add a media part under `/word/media/imageN.ext`.
2. Add an image relationship to the relevant story part’s `.rels`.
3. Add content type default or override if needed.
4. Insert a `w:drawing/wp:inline` structure in a paragraph.
5. Set extent width/height in EMUs.
6. Set `docPr` name and description from alt text.

---

## 20. Headers and footers

Headers and footers are discovered from section properties and relationships.

The scanner must assign physical header/footer part IDs:

```text
H001
H002
F001
F002
```

Paragraph and table operations work inside headers and footers using the same operation names:

```text
op replace-text
target H001.P0001
expect-text <<<Confidential>>>
find <<<Confidential>>>
with <<<Internal>>>
end
```

v0.1 does not need to understand first-page/even-page/default header semantics beyond reporting them in `outline`.

---

## 21. Styles

Read styles from `/word/styles.xml`.

For each style:

```csharp
public sealed record DocxStyleInfo
{
    public required string StyleId { get; init; }
    public required string Type { get; init; } // paragraph, character, table, numbering
    public string? Name { get; init; }
    public bool IsDefault { get; init; }
}
```

Style resolution rules:

* First match by exact `styleId`.
* Then match by exact display name.
* If multiple styles match, fail.
* If no style matches, fail.
* Do not create new styles in v0.1.

---

## 22. Security requirements

Production library must be safe for untrusted `.docx` input.

Required safeguards:

* Never extract archive entries to disk.
* Reject suspicious ZIP paths.
* Enforce max entry count.
* Enforce max total uncompressed bytes.
* Enforce max single-part bytes.
* Disable XML DTDs and external resolvers.
* Do not fetch external relationships.
* Do not execute macros.
* Reject macro-enabled Word documents unless `AllowMacroEnabledDocuments` is true.
* Treat external image relationships as read-only metadata; never download them.
* Preserve unknown parts as bytes but do not execute or interpret them.

---

## 23. Testing requirements

### 23.1 Unit tests

Required unit test groups:

```text
Patch parser
  parses valid patches
  rejects malformed heredocs
  rejects missing preamble
  reports line and column
  rejects expect-hash

Selector resolver
  resolves IDs
  resolves headings
  reports zero matches
  reports multiple matches

Text extraction
  reads w:t
  handles tabs
  handles line breaks
  ignores field instructions by default
  handles inserted/deleted revisions in final view

Text replacement
  single run replacement
  run splitting
  replacement across simple adjacent runs
  failure across hyperlink boundary
  xml:space preservation

Tables
  reads simple tables
  detects merged cells
  set-cell preserves tcPr
  append-row clones row shape
  delete-row rejects last row

Images
  lists inline images
  lists floating images
  replaces PNG
  replaces JPEG
  rejects unsupported asset type
  inserts inline PNG
  computes EMUs

Sections
  reads columns
  sets column count
  sets orientation

Package
  preserves unknown parts
  preserves unrelated media
  rejects zip traversal paths
  rejects missing required parts

Streams
  works with MemoryStream input/output
  works with non-seekable input stream
  asset provider works without file paths
```

### 23.2 Golden tests

For representative fixtures, compare rendered text output to checked-in golden files:

```text
read.simple.golden.txt
outline.simple.golden.txt
dump.paragraph-runs.golden.txt
styles.corporate.golden.txt
media.images.golden.txt
```

Golden output must be deterministic.

### 23.3 Fixture generation

Provide an internal fixture builder that creates minimal `.docx` packages using only system libraries.

Also allow optional Microsoft Word-generated fixtures for real-world compatibility.

### 23.4 Office integration tests

Create `DocxEdit.OfficeTests`.

These tests are optional and run only when:

```text
DOCXEDIT_ENABLE_OFFICE_TESTS=1
```

and only on Windows with Microsoft Word installed.

Use late-bound COM via `Type.GetTypeFromProgID("Word.Application")` and `dynamic`. Do not add a production dependency. Avoid adding a Microsoft Office interop package.

Office tests should:

1. Open generated `.docx` files in Word.
2. Save them to a temp output path.
3. Close Word cleanly.
4. Fail if Word cannot open/save the document.
5. Optionally export to PDF for visual inspection in local developer workflows.

Run Office automation tests in STA threads. Always close documents and quit Word in `finally`.

### 23.5 Edit case validation harness

In addition to unit and golden tests, provide a lightweight case catalog for end-to-end edit workflows.

Public case layout:

```text
edit-cases/
  cases/
    basic-replace/
      case.json
      input.docx
      edits.docxpatch
      expected.read.txt
  families/
    text.json
    tables.json
    media.json
```

Example case manifest:

```json
{
  "id": "basic-replace",
  "kind": "docx",
  "input": "./input.docx",
  "patch": "./edits.docxpatch",
  "tags": ["text", "smoke"],
  "expected": {
    "checkMustSucceed": true,
    "applyMustSucceed": true,
    "wordRoundtripMustSucceed": false,
    "diagnosticCodes": []
  }
}
```

`tools/ValidateDocxCases.ps1` must validate:

* normalized kebab-case case and family IDs;
* each case directory has a `case.json`;
* manifest `id` matches directory name;
* input and patch paths exist and stay inside the case directory unless explicitly allowed;
* every public case belongs to exactly one family unless the tool is run with an uncovered-case override;
* tags are normalized kebab-case;
* expected diagnostic codes use known prefixes.

`tools/CheckDocxCase.ps1` must:

1. Build the CLI.
2. Run `docxedit check`.
3. Run `docxedit apply`.
4. Run `docxedit read` on the output and compare expected golden output when present.
5. Write timestamped artifacts under `artifacts/edit-cases/<case-id>/<run-id>/`.

Generated artifacts are ignored by git.

### 23.6 Private validation

Support local-only private cases for confidential documents, modeled on the public edit case manifest.

Private inputs and manifests must live under ignored `private-cases/`. `tools/CheckPrivateCase.ps1` must reject:

* manifests outside `private-cases/`;
* inputs outside `private-cases/`;
* git-tracked private manifests or inputs;
* case IDs that are not single filename-safe path segments.

Private artifacts must be written under ignored `artifacts/private-edit/<case-id>/<run-id>/`.

Public notes derived from private cases must be anonymized: record feature gaps, operation class, diagnostics, and page/section/table shape only; do not copy private text, screenshots, filenames, or document contents.

---

## 24. CLI help requirements

The CLI must be self-explanatory for a fresh coding agent.

`docxedit -h`:

```text
docxedit - read and patch .docx documents

Usage:
  docxedit <command> [options]

Read / explore:
  read       Produce an agent-friendly structural view of a .docx
  outline    Show headings, tables, images, sections, headers, footers
  find       Find text and print stable edit targets
  dump       Dump one target in detail
  styles     List paragraph, character, and table styles
  media      List or extract embedded images

Patch:
  check      Validate a .docxpatch file without writing output
  apply      Apply a .docxpatch file and write a new .docx

Help:
  help patch Show the .docxpatch syntax with examples

Examples:
  docxedit read report.docx
  docxedit find report.docx "Revenue"
  docxedit dump report.docx --id M.P0004 --runs
  docxedit check report.docx edits.docxpatch
  docxedit apply report.docx edits.docxpatch -o report.edited.docx
```

`docxedit help patch` must include:

* DSL preamble
* operation block syntax
* heredoc syntax
* selector examples
* operation examples
* track-changes explanation
* warning that `expect-hash` is unsupported

---

## 25. Machine-readable output

The CLI should support `--json` for automation.

Use `System.Text.Json`.

Example shape:

```json
{
  "success": true,
  "diagnostics": [],
  "operations": [
    {
      "index": 1,
      "operationName": "replace-text",
      "target": "M.P0004",
      "success": true,
      "diagnostics": []
    }
  ]
}
```

The public library should expose data models directly; JSON is a CLI rendering concern.

---

## 26. Versioning

Use semantic versioning for the `Lokad.DocxEdit` NuGet package.

Patch DSL version is independent:

```text
docxpatch 1
```

Rules:

* The parser must reject unknown major DSL versions.
* The parser may accept minor additions only if backward-compatible.
* Unknown operation names are errors.
* Unknown fields inside known operations are errors unless explicitly documented as ignored.

---

## 27. Required acceptance criteria

The implementation is acceptable when all of the following are true:

1. `DocxEdit` production project has no NuGet dependencies.
2. All public APIs operate on streams.
3. Public operations observe cancellation tokens.
4. CLI can read, find, dump, check, and apply patches.
5. CLI can write diagnostics JSON and `--strict` returns exit code `3` on warnings.
6. `.docxpatch` parser supports the specified block/heredoc syntax.
7. `expect-hash` is not implemented and is rejected if present.
8. Package loader rejects unsafe paths, duplicate normalized names, excessive entry count, excessive part size, excessive total uncompressed size, unsafe relationship targets, and XML DTD/external entity input.
9. Paragraph text replacement works across simple run splits.
10. Table cell edits and simple row appends work.
11. Inline image replacement and insertion work for PNG and JPEG.
12. Header/footer paragraphs can be read and edited.
13. Section column count can be read and changed.
14. `check` validates selectors and guards without writing output.
15. `apply` writes a new `.docx` and never mutates input.
16. Unknown OOXML parts are preserved.
17. Existing unsupported structures are preserved unless directly targeted.
18. Track changes modes `Off`, `Preserve`, `Suggest`, and `Require` exist with the behavior specified above.
19. Golden tests prove deterministic read/probe output.
20. Stream-only tests prove no file-system dependency in the library.
21. Public edit-case validation tools can run at least one smoke case.
22. Private-case tooling rejects tracked or out-of-directory confidential inputs.
23. Optional Office tests can validate generated files on Windows when enabled.

---

## 28. Recommended implementation order

1. Create solution and projects.
2. Implement ZIP package loader/saver.
3. Implement content types and relationships parser.
4. Discover main document, styles, headers, footers, media.
5. Implement logical scanner for paragraphs, tables, images, sections.
6. Implement deterministic ID assignment.
7. Implement read/outline/find/dump/styles/media renderers.
8. Implement patch parser.
9. Implement selector resolver.
10. Implement guards.
11. Implement paragraph text replacement.
12. Implement insert/delete paragraph operations.
13. Implement table `set-cell` and `append-row`.
14. Implement image list and replace.
15. Implement inline image insertion.
16. Implement section column/orientation edits.
17. Implement basic track-changes mode.
18. Implement `check` simulation.
19. Implement `apply`.
20. Implement CLI.
21. Implement CLI diagnostics JSON and strict mode.
22. Add golden tests.
23. Add public edit-case validation tools.
24. Add private-case validation guardrails.
25. Add optional Office integration tests.
26. Publish local `Lokad.DocxEdit` NuGet package and verify the CLI consumes the library through project reference only.

The most important invariant throughout the implementation is:

```text
Every edit must specify:
  where to edit,
  what current visible content is expected,
  and what change to make.
```

That invariant is what makes `docxedit` safe for iterative coding-agent use.

[1]: https://learn.microsoft.com/en-us/dotnet/core/whats-new/dotnet-10/overview "What's new in .NET 10 | Microsoft Learn"
[2]: https://learn.microsoft.com/en-us/dotnet/core/compatibility/sdk/10.0/dotnet-new-sln-slnx-default "Breaking change - `dotnet new sln` defaults to SLNX file format - .NET | Microsoft Learn"
[3]: https://learn.microsoft.com/en-us/office/open-xml/about-the-open-xml-sdk "About the Open XML SDK for Office | Microsoft Learn"
[4]: https://learn.microsoft.com/en-us/office/open-xml/word/how-to-accept-all-revisions-in-a-word-processing-document "How to: Accept all revisions in a word processing document | Microsoft Learn"
[5]: https://learn.microsoft.com/en-us/office/open-xml/word/working-with-wordprocessingml-tables "Working with WordprocessingML tables | Microsoft Learn"
[6]: https://learn.microsoft.com/en-us/dotnet/api/system.io.compression.ziparchive?view=net-10.0 "ZipArchive Class (System.IO.Compression) | Microsoft Learn"
[7]: https://learn.microsoft.com/en-us/dotnet/core/deploying/single-file/overview "Create a single file for application deployment - .NET | Microsoft Learn"
[8]: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.trackrevisions?view=openxml-3.0.1 "TrackRevisions Class (DocumentFormat.OpenXml.Wordprocessing) | Microsoft Learn"
[9]: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.insertedrun?view=openxml-3.0.1 "InsertedRun Class (DocumentFormat.OpenXml.Wordprocessing) | Microsoft Learn"
