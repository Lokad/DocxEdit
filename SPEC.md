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
* No full editing support for floating shapes beyond modeled image layout metadata.
* No linked image fetching.
* No VML, grouped drawing, or OLE object editing.
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

    public DocxContextResult Context(
        Stream input,
        string targetId,
        DocxContextOptions? options = null,
        CancellationToken cancellationToken = default);

    public DocxStylesResult Styles(
        Stream input,
        DocxStylesOptions? options = null,
        CancellationToken cancellationToken = default);

    public DocxMediaResult Media(
        Stream input,
        DocxMediaOptions? options = null,
        CancellationToken cancellationToken = default);

    public DocxValidateResult Validate(
        Stream input,
        DocxValidateOptions? options = null,
        CancellationToken cancellationToken = default);

    public DocxChangesResult Changes(
        Stream input,
        DocxChangesOptions? options = null,
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

Read-only option classes share the same package-limit and `LeaveInputOpen` defaults,
including `DocxValidateOptions`.
`DocxReadOptions`, `DocxFindOptions`, `DocxDumpOptions`, and `DocxContextOptions`
also accept `DocxTextView` (`Final`, `Original`, or `Markup`) where visible text is
rendered. `DocxContextOptions.MaxText` defaults to `0`; callers opt in when context
items should include text snippets. `DocxChangesOptions` defaults to private-text-free
change/comment metadata and requires explicit opt-in for comment body snippets:

```csharp
public sealed class DocxChangesOptions
{
    public bool IncludeCommentText { get; init; } = false;
    public int MaxCommentText { get; init; } = 240;
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
    IReadOnlyList<DocxDiagnostic> Diagnostics)
{
    public IReadOnlyList<DocxPatchAffectedTarget> AffectedTargets { get; init; } = [];
}

public sealed record DocxPatchAffectedTarget(string Id, string Kind, string Action)
{
    public string? ParentId { get; init; }
    public int? RowIndex { get; init; }
    public int? ColumnIndex { get; init; }
    public int? RowCountBefore { get; init; }
    public int? RowCountAfter { get; init; }
    public int? ColumnCount { get; init; }
    public int? CellCount { get; init; }
}
```

Read/explore result records must expose structured data in addition to the CLI text
renderings:

```csharp
public sealed record DocxReadResult : DocxOperationResult
{
    public IReadOnlyList<string> PartNames { get; init; } = [];
    public string? MainDocumentPartName { get; init; }
    public string Text { get; init; } = string.Empty;
    public IReadOnlyList<DocxParagraphInfo> Paragraphs { get; init; } = [];
    public IReadOnlyList<DocxTableInfo> Tables { get; init; } = [];
    public IReadOnlyList<DocxImageInfo> Images { get; init; } = [];
    public IReadOnlyList<DocxSectionInfo> Sections { get; init; } = [];
    public IReadOnlyList<DocxBookmarkInfo> Bookmarks { get; init; } = [];
    public IReadOnlyList<DocxContentControlInfo> ContentControls { get; init; } = [];
    public IReadOnlyList<DocxFieldInfo> Fields { get; init; } = [];
    public IReadOnlyList<DocxHyperlinkInfo> Hyperlinks { get; init; } = [];
}

public sealed record DocxDumpResult : DocxOperationResult
{
    public string TargetId { get; init; } = string.Empty;
    public string? Text { get; init; }
    public IReadOnlyList<DocxDumpRunInfo> Runs { get; init; } = [];
}

public sealed record DocxContextResult : DocxOperationResult
{
    public string TargetId { get; init; } = string.Empty;
    public string Text { get; init; } = string.Empty;
    public IReadOnlyList<DocxContextItem> Items { get; init; } = [];
}

public sealed record DocxChangesResult : DocxOperationResult
{
    public IReadOnlyList<string> PartNames { get; init; } = [];
    public string? MainDocumentPartName { get; init; }
    public IReadOnlyList<DocxChangeInfo> Changes { get; init; } = [];
    public IReadOnlyList<DocxChangeSummary> Summary { get; init; } = [];
    public IReadOnlyList<DocxChangeGroupSummary> GroupSummary { get; init; } = [];
    public IReadOnlyList<DocxChangeTargetSummary> TargetSummary { get; init; } = [];
    public IReadOnlyList<DocxCommentThreadSummary> CommentSummary { get; init; } = [];
}

public sealed record DocxValidateResult : DocxOperationResult
{
    public DocxValidationProfile Profile { get; init; } = DocxValidationProfile.Structural;
    public IReadOnlyList<string> PartNames { get; init; } = [];
    public string? MainDocumentPartName { get; init; }
}

public sealed class DocxValidateOptions
{
    public DocxValidationProfile Profile { get; init; } = DocxValidationProfile.Structural;
    public int MaxDiagnostics { get; init; } = 500;
    public int MaxZipEntries { get; init; } = 10_000;
    public long MaxUncompressedBytes { get; init; } = 512L * 1024L * 1024L;
    public long MaxSinglePartBytes { get; init; } = 128L * 1024L * 1024L;
    public bool LeaveInputOpen { get; init; } = true;
}

public enum DocxValidationProfile
{
    Structural,
    Package
}
```

`Changes` must be private-text-free by default. It may report revision/comment
metadata, text lengths, child counts, IDs, targets, stories, and parts. It must not
copy revision text into the result. Comment body snippets may appear only when
`DocxChangesOptions.IncludeCommentText` is true, and then must be bounded by
`MaxCommentText`.

### 5.1 Public integration surfaces

The NuGet library must expose the agent-facing command guidance that the CLI uses.
Integrators must not have to duplicate CLI-local help text.

```csharp
public static class DocxHelp
{
    public static DocxCommandCatalog Catalog { get; }
    public static bool TryGetCommand(string name, out DocxCommandInfo command);
    public static bool TryGetPatchOperation(string name, out DocxPatchOperationInfo operation);
    public static string RenderOverview();
    public static bool TryRenderTopic(string topic, out string text);
    public static string RenderTopic(string topic);
}
```

`DocxCommandCatalog` must include structured commands, options, examples, output
fields, privacy notes, and patch operation specs. The CLI must render its overview
and command-specific help from this catalog.

```csharp
public sealed record DocxPatchOperationInfo
{
    public string Name { get; init; } = string.Empty;
    public IReadOnlyList<string> RequiredFields { get; init; } = [];
    public IReadOnlyList<string> OptionalFields { get; init; } = [];
    public string Description { get; init; } = string.Empty;
    public string TrackChangesSupport { get; init; } = "unsupported";
    public string TrackChangesNote { get; init; } = string.Empty;
}
```

`TrackChangesSupport` values include `tracked-simple`, `tracked-paragraph`,
`tracked-paragraph-insert`, `tracked-paragraph-delete`, `tracked-style`,
`tracked-cell-simple`, `preserve-only`, and `unsupported`. `preserve-only`
operations preserve existing tracked-change markup but do not create new revision
markup; they apply directly under `Suggest` with `W4001` and fail under `Require`
with `E6001`. `unsupported` is reserved for unknown or unclassified tracked-output
behavior. Supported tracked operations may still reject an unsafe shape with `W4002`
or `E6002`.

The library must also expose stable plain-text renderers for public result objects:

```csharp
public static class DocxTextRenderer
{
    public static string RenderRead(DocxReadResult result);
    public static string RenderReadSummary(DocxReadResult result);
    public static string RenderOutline(DocxOutlineResult result);
    public static string RenderFind(DocxFindResult result);
    public static string RenderDump(DocxDumpResult result);
    public static string RenderContext(DocxContextResult result);
    public static string RenderStyles(DocxStylesResult result);
    public static string RenderMedia(DocxMediaResult result);
    public static string RenderValidate(DocxValidateResult result);
    public static string RenderChanges(DocxChangesResult result);
    public static string RenderOperationSummary(IReadOnlyList<DocxPatchOperationReport> operations);
}
```

Privacy-safe integration presets must be public:

```csharp
public static class DocxPrivacyPresets
{
    public static DocxReadOptions ReadSummary { get; }
    public static DocxContextOptions ContextMetadataOnly { get; }
    public static DocxChangesOptions ChangesMarkupOnly { get; }
}
```

`ReadSummary` and `ContextMetadataOnly` use `MaxText = 0`. Consumers must prefer
`DocxTextRenderer.RenderReadSummary`, `DocxTextRenderer.RenderContext`, and
`DocxTextRenderer.RenderChanges` when they need privacy-safe agent output instead of
serializing full result objects.

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
public sealed record DocxParagraphInfo(
    string Id,
    string Story,
    string Text,
    int? HeadingLevel,
    DocxListInfo? List,
    IReadOnlyList<DocxRunInfo> Runs)
{
    public string? StyleId { get; init; }
    public string? StyleName { get; init; }
}

public sealed record DocxListInfo(string NumberingId, int Level)
{
    public string? AbstractNumberingId { get; init; }
    public string? Format { get; init; }
    public string? LevelText { get; init; }
    public string? LabelText { get; init; }
    public IReadOnlyList<DocxListLabelComponent> LabelComponents { get; init; } = [];
    public string LabelStatus { get; init; } = "not-resolved";
    public IReadOnlyList<string> LabelWarnings { get; init; } = [];
    public int? StartValue { get; init; }
    public string? Suffix { get; init; }
    public bool IsLegal { get; init; }
    public int? RestartAfterLevel { get; init; }
    public string? ParagraphStyleId { get; init; }
    public string Source { get; init; } = "direct";
}

public sealed record DocxListLabelComponent(int Level, int Value, string Text, string Format);
```

`DocxListInfo` resolves direct paragraph numbering and paragraph-style numbering when
available. `Source` is `direct`, `style`, or `style-inherited`. `LabelText` is the
visible label when DocxEdit can deterministically expand the level text from known
counters and supported formats. `LabelComponents` contains one structured component
per expanded `%n` token with source level, raw counter value, formatted text, and
format. `LabelStatus` is `resolved`, `partial`, `unsupported`, or `not-resolved`;
`LabelWarnings` names missing counters, missing definitions, and unsupported formats
without exposing document text. Supported label formats include decimal, zero-padded
decimal, upper/lower letters, upper/lower roman numerals, bullets, and nested
`lvlText` tokens whose referenced counters are known.

### 8.4 Run model

```csharp
public sealed record DocxRunInfo(string Text)
{
    public string? MarkupType { get; init; }
    public string? RevisionId { get; init; }
    public string? Author { get; init; }
    public DateTimeOffset? TimestampUtc { get; init; }
    public string? CommentId { get; init; }
    public string? HyperlinkRelationshipId { get; init; }
    public string? HyperlinkAnchor { get; init; }
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
public sealed record DocxTableInfo(
    string Id,
    string Story,
    int RowCount,
    int ColumnCount,
    IReadOnlyList<DocxTableCellInfo> Cells)
{
    public string? StyleId { get; init; }
    public string? Caption { get; init; }
    public string? Description { get; init; }
    public int? GridColumnCount { get; init; }
    public bool HasHeaderRow { get; init; }
    public bool HasMergedCells { get; init; }
    public bool HasNestedTables { get; init; }
    public IReadOnlyList<DocxTableRowInfo> Rows { get; init; } = [];
}
```

Row model:

```csharp
public sealed record DocxTableRowInfo
{
    public string Id { get; init; } = string.Empty;
    public int RowIndex { get; init; }
    public int CellCount { get; init; }
    public int GridBefore { get; init; }
    public int GridAfter { get; init; }
    public bool IsHeader { get; init; }
    public bool CantSplit { get; init; }
}
```

Cell model:

```csharp
public sealed record DocxTableCellInfo(
    string Id,
    int RowIndex,
    int ColumnIndex,
    string Text,
    int ColumnSpan,
    string? VerticalMerge,
    bool HasNestedTable)
{
    public int PhysicalColumnIndex { get; init; }
    public int VisualColumnEndIndex { get; init; }
    public string? MergeGroupId { get; init; }
    public string? VerticalMergeRootCellId { get; init; }
}
```

`ColumnIndex` is the one-based visual grid start column after `gridBefore` and
`gridSpan` expansion. `VisualColumnEndIndex` is populated for cells that span
multiple visual columns. `MergeGroupId` groups horizontal spans and vertical merge
chains within a table, and `VerticalMergeRootCellId` links `w:vMerge` continuations
back to the restart cell when the root is visible in the scanned story.

### 8.6 Image model

```csharp
public sealed record DocxImageInfo(
    string Id,
    string PartName,
    string? ContentType,
    long ByteLength)
{
    public string LayoutKind { get; init; } = "unknown";
    public string? RelationshipId { get; init; }
    public string? ContainingTargetId { get; init; }
    public long? WidthEmu { get; init; }
    public long? HeightEmu { get; init; }
    public string? Name { get; init; }
    public string? Description { get; init; }
    public string? Title { get; init; }
    public string? WrapMode { get; init; }
    public bool BehindDoc { get; init; }
    public long? WrapDistanceTopEmu { get; init; }
    public long? WrapDistanceBottomEmu { get; init; }
    public long? WrapDistanceLeftEmu { get; init; }
    public long? WrapDistanceRightEmu { get; init; }
    public long? RelativeHeight { get; init; }
    public bool? AllowOverlap { get; init; }
    public bool? LockAspectRatio { get; init; }
    public string? HorizontalPositionRelativeFrom { get; init; }
    public long? HorizontalPositionOffsetEmu { get; init; }
    public string? HorizontalPositionAlign { get; init; }
    public string? VerticalPositionRelativeFrom { get; init; }
    public long? VerticalPositionOffsetEmu { get; init; }
    public string? VerticalPositionAlign { get; init; }
    public decimal? CropLeftPercent { get; init; }
    public decimal? CropTopPercent { get; init; }
    public decimal? CropRightPercent { get; init; }
    public decimal? CropBottomPercent { get; init; }
}
```

The public image model exposes compact DrawingML layout metadata for inline and
anchored drawings. Anchored drawings expose wrap distances, relative positioning,
relative height, overlap flags, and aspect-ratio locks when present. Crop percentages
are read from DrawingML `a:srcRect` attributes when present and exposed as
human-readable percentages, for example `10` for a 10% left crop. Detailed
positioning edits are preserved in OOXML where possible but not exposed as public
patch operations; crop percentages can be edited through `set-image-crop`.

### 8.7 Section model

```csharp
public sealed record DocxSectionInfo(
    string Id,
    string Story,
    int Columns,
    string Orientation);
```

### 8.8 Bookmark and content-control model

`DocxEditor.Read` exposes bookmark and content-control metadata so agents can inspect
selector candidates without raw OOXML.

```csharp
public sealed record DocxBookmarkInfo
{
    public string Id { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;
    public string? OoxmlId { get; init; }
    public string Story { get; init; } = string.Empty;
    public string PartName { get; init; } = string.Empty;
    public string? StartTargetId { get; init; }
    public string? EndTargetId { get; init; }
    public bool IsComplete { get; init; }
    public bool IsNameDuplicate { get; init; }
    public IReadOnlyList<string> DuplicateNameBookmarkIds { get; init; } = [];
}

public sealed record DocxContentControlInfo
{
    public string Id { get; init; } = string.Empty;
    public string Story { get; init; } = string.Empty;
    public string PartName { get; init; } = string.Empty;
    public string? TargetId { get; init; }
    public string Kind { get; init; } = "unknown";
    public string? OoxmlId { get; init; }
    public string? Tag { get; init; }
    public string? Alias { get; init; }
    public string? PlaceholderDocPart { get; init; }
    public bool IsShowingPlaceholderText { get; init; }
    public string? DataBindingXPath { get; init; }
    public string? DataBindingStoreItemId { get; init; }
    public string? DataBindingPrefixMappings { get; init; }
    public string? RepeatingSectionTitle { get; init; }
    public int? RepeatingSectionItemCount { get; init; }
    public string? ParentContentControlId { get; init; }
    public IReadOnlyList<string> ChildContentControlIds { get; init; } = [];
    public string SafeEditStatus { get; init; } = "unknown";
    public bool IsTagDuplicate { get; init; }
    public IReadOnlyList<string> DuplicateTagControlIds { get; init; } = [];
    public bool IsAliasDuplicate { get; init; }
    public IReadOnlyList<string> DuplicateAliasControlIds { get; init; } = [];
    public string? Lock { get; init; }
    public bool? Checked { get; init; }
    public string? CheckedSymbol { get; init; }
    public string? UncheckedSymbol { get; init; }
    public IReadOnlyList<DocxContentControlListItemInfo> ListItems { get; init; } = [];
    public string? DateFormat { get; init; }
    public string? DateLanguage { get; init; }
    public string? DateCalendar { get; init; }
    public string? DateValue { get; init; }
    public int TextLength { get; init; }
}

public sealed record DocxContentControlListItemInfo(string? DisplayText, string? Value);
```

Bookmark IDs use the `B` namespace and content-control IDs use the `CC` namespace,
for example `M.B0001` and `M.CC0001`. They are metadata IDs, not patch edit targets.
Use `StartTargetId`, `EndTargetId`, or `TargetId` for edits unless a later patch
operation explicitly accepts the metadata ID.
Content controls expose placeholder doc-part IDs, placeholder-display state, custom
XML data-binding attributes, repeating-section titles/item counts, direct
parent/child content-control IDs, and a conservative `SafeEditStatus` such as
`plain-text`, `rich-text`, `choice`, `date`, `locked`, `unsupported-picture`,
`unsupported-group`, or `unsupported-repeating-section`.

Duplicate bookmark names, duplicate content-control tags, and duplicate
content-control aliases must be surfaced with boolean duplicate flags and candidate
metadata IDs. This lets agents avoid ambiguous `bookmark:"Name"` and
`content-control:"TagOrAlias"` selectors before attempting a patch.

### 8.9 Field model

`DocxEditor.Read` exposes simple fields and complex `fldChar` begin/separate/end
sequences as metadata. DocxEdit preserves field XML, can set dirty/lock flags on
existing fields, and can mark documents for field updates after edits, but it does
not evaluate or recalculate field results.

```csharp
public sealed record DocxFieldInfo
{
    public string Id { get; init; } = string.Empty;
    public string Story { get; init; } = string.Empty;
    public string PartName { get; init; } = string.Empty;
    public string? TargetId { get; init; }
    public string Kind { get; init; } = "unknown";
    public string? FieldType { get; init; }
    public string Code { get; init; } = string.Empty;
    public string CachedResultText { get; init; } = string.Empty;
    public int ResultTextLength { get; init; }
    public int NestingDepth { get; init; }
    public IReadOnlyList<string> BookmarkDependencies { get; init; } = [];
    public IReadOnlyList<string> HyperlinkDependencies { get; init; } = [];
    public string SafeEditStatus { get; init; } = "unknown";
    public bool? IsDirty { get; init; }
    public bool? IsLocked { get; init; }
    public bool IsComplete { get; init; }
}
```

Field IDs use the `F` namespace, for example `M.F0001`. They are metadata IDs, not
general text edit targets. `CachedResultText` is the stored visible result currently
in the document, not a recalculated value. `set-field-dirty` and `set-field-lock`
explicitly accept field metadata IDs.

### 8.10 Hyperlink model

`DocxEditor.Read` exposes hyperlink metadata for external relationship links,
internal anchors, relationship targets, and broken relationship IDs. External
targets include URI scheme validation metadata. Internal anchors include
missing-bookmark and duplicate-bookmark flags.

```csharp
public sealed record DocxHyperlinkInfo
{
    public string Id { get; init; } = string.Empty;
    public string Story { get; init; } = string.Empty;
    public string PartName { get; init; } = string.Empty;
    public string? TargetId { get; init; }
    public string? RelationshipId { get; init; }
    public string? RelationshipPartName { get; init; }
    public string? RelationshipTargetMode { get; init; }
    public string? Uri { get; init; }
    public string? UriScheme { get; init; }
    public bool? IsUriValid { get; init; }
    public string? UriValidationReason { get; init; }
    public string? Anchor { get; init; }
    public bool? IsAnchorMissing { get; init; }
    public bool? IsAnchorDuplicate { get; init; }
    public string? Tooltip { get; init; }
    public string? TargetFrame { get; init; }
    public bool? History { get; init; }
    public string? TargetPartName { get; init; }
    public bool IsExternal { get; init; }
    public bool IsBroken { get; init; }
    public int DisplayTextLength { get; init; }
}
```

Hyperlink IDs use the `L` namespace, for example `M.L0001`. They are metadata IDs,
not patch edit targets. `RelationshipPartName` identifies the `.rels` part that
contains `RelationshipId`, and `RelationshipTargetMode` exposes the raw relationship
`TargetMode` value when the relationship exists. `Uri` is populated for external
hyperlink relationships, `UriScheme`, `IsUriValid`, and `UriValidationReason` report
whether the external target is an allowed absolute `http`, `https`, or `mailto` URI. Relative external
targets use `UriValidationReason = "relative-uri"`; malformed targets use
`UriValidationReason = "malformed-uri"`; unsupported absolute schemes, including
`file` and UNC/file-style targets parsed as `file`, use
`UriValidationReason = "unsupported-uri-scheme"`.
`Anchor` is populated for internal anchors, `IsAnchorMissing` flags anchors without a
matching bookmark, `IsAnchorDuplicate` flags anchors that match multiple bookmarks,
`TargetPartName` is populated for relationship-backed internal part links,
`TargetFrame` exposes `w:tgtFrame`, `History` exposes `w:history`, and `IsBroken`
flags missing relationship IDs. `read` diagnostics aggregate invalid URI, missing
anchor, duplicate-anchor, and unsupported internal part-link cases with
`Feature = "hyperlink"` and stable `Fallback` values.

### 8.11 Tracked-change and comment markup model

`DocxEditor.Changes` scans existing tracked-change, move, custom XML, property-change,
and comment markup. The default result does not expose private revision or comment
body text. Bounded comment body snippets are exposed only when
`DocxChangesOptions.IncludeCommentText` is true.

```csharp
public sealed record DocxChangeInfo
{
    public string Id { get; init; } = string.Empty;
    public string Type { get; init; } = string.Empty;
    public string Story { get; init; } = string.Empty;
    public string PartName { get; init; } = string.Empty;
    public string? TargetId { get; init; }
    public string? Author { get; init; }
    public DateTimeOffset? TimestampUtc { get; init; }
    public string? RevisionId { get; init; }
    public int TextLength { get; init; }
    public int ChildElementCount { get; init; }
    public string? CommentId { get; init; }
    public string? CommentAuthor { get; init; }
    public DateTimeOffset? CommentTimestampUtc { get; init; }
    public string? CommentInitials { get; init; }
    public string? CommentParaId { get; init; }
    public string? CommentParentParaId { get; init; }
    public string? CommentRootParaId { get; init; }
    public bool? CommentIsReply { get; init; }
    public bool? CommentResolved { get; init; }
    public string? CommentAnchorTargetId { get; init; }
    public string? CommentReferenceTargetId { get; init; }
    public string? CommentAnchorStory { get; init; }
    public string? CommentAnchorPartName { get; init; }
    public int? CommentTextLength { get; init; }
    public string? CommentTextSnippet { get; init; }
    public bool CommentTextTruncated { get; init; }
    public string TargetStatus { get; init; } = "targetless";
    public string TargetSource { get; init; } = "none";
    public string? TargetReason { get; init; }
    public string? NearestTargetId { get; init; }
    public string? TargetNote { get; init; }
    public string? PairedChangeId { get; init; }
}
```

`CommentParaId`, `CommentParentParaId`, `CommentRootParaId`, `CommentIsReply`,
and `CommentResolved` are populated from Word's modern `commentsExtended.xml`
metadata when it is present. They are safe metadata fields and do not expose
comment body text. `DocxCommentThreadSummary` exposes the same root/reply and
resolution fields, plus the same opt-in text snippet fields, at the thread level:

```csharp
public sealed record DocxCommentThreadSummary
{
    public string CommentId { get; init; } = string.Empty;
    public string? AnchorTargetId { get; init; }
    public string? ReferenceTargetId { get; init; }
    public string? AnchorStory { get; init; }
    public string? AnchorPartName { get; init; }
    public string? Author { get; init; }
    public DateTimeOffset? TimestampUtc { get; init; }
    public string? Initials { get; init; }
    public string? ParaId { get; init; }
    public string? ParentParaId { get; init; }
    public string? RootParaId { get; init; }
    public bool? IsReply { get; init; }
    public bool? Resolved { get; init; }
    public int? TextLength { get; init; }
    public string? TextSnippet { get; init; }
    public bool TextTruncated { get; init; }
    public int Count { get; init; }
    public IReadOnlyList<DocxChangeSummary> Summary { get; init; } = [];
}
```

`TargetStatus` is `targeted`, `comment-anchor`, or `targetless`. `TargetSource`
distinguishes exact ancestor matches from adjacent range-boundary heuristics. For
targetless records, `TargetReason`, `NearestTargetId`, and `TargetNote` provide
context without claiming exact ownership. Range starts/ends that share a revision or
comment ID expose `PairedChangeId`.

### 8.12 Context model

`DocxEditor.Context` summarizes nearby modeled structure around a target. Its default
`MaxText` is `0`, making it safe for private-document navigation unless the caller
explicitly requests text snippets.

```csharp
public sealed record DocxContextItem
{
    public string Id { get; init; } = string.Empty;
    public string Kind { get; init; } = string.Empty;
    public string Relation { get; init; } = string.Empty;
    public string Story { get; init; } = string.Empty;
    public string? ParentId { get; init; }
    public string Text { get; init; } = string.Empty;
    public int? HeadingLevel { get; init; }
    public string? StyleId { get; init; }
    public string? StyleName { get; init; }
    public DocxListInfo? List { get; init; }
    public IReadOnlyList<string> BookmarkNames { get; init; } = [];
    public IReadOnlyList<string> ContentControlIds { get; init; } = [];
    public IReadOnlyList<string> ContentControlTags { get; init; } = [];
    public IReadOnlyList<string> ContentControlAliases { get; init; } = [];
    public IReadOnlyList<string> FieldIds { get; init; } = [];
    public IReadOnlyList<string> FieldCodes { get; init; } = [];
    public IReadOnlyList<string> FieldKinds { get; init; } = [];
    public IReadOnlyList<string> FieldTypes { get; init; } = [];
    public IReadOnlyList<string> HyperlinkIds { get; init; } = [];
    public IReadOnlyList<string> HyperlinkTargets { get; init; } = [];
    public IReadOnlyList<string> CommentIds { get; init; } = [];
    public IReadOnlyList<string> CommentBodyIds { get; init; } = [];
    public string? Caption { get; init; }
    public string? Description { get; init; }
    public int? RowCount { get; init; }
    public int? ColumnCount { get; init; }
    public int? RowIndex { get; init; }
    public int? ColumnIndex { get; init; }
    public int? ColumnSpan { get; init; }
    public int? VisualColumnEndIndex { get; init; }
    public string? MergeGroupId { get; init; }
    public string? VerticalMerge { get; init; }
    public string? VerticalMergeRootCellId { get; init; }
    public bool HasNestedTable { get; init; }
}
```

Paragraph and cell context items include `CommentIds` and `CommentBodyIds` when
comment anchors or references are attached to that target. Comment body targets such
as `C001.C0001` and `comment:<id>` return a `Kind = "comment"` context item whose
metadata links back to the anchor target and does not include comment body text.
Table context items expose `Caption` and `Description` when `w:tblCaption` or
`w:tblDescription` are present. Cell context items also expose visual-grid and merge metadata, including
`VisualColumnEndIndex`, `MergeGroupId`, `VerticalMerge`, and
`VerticalMergeRootCellId`, when those values are present on the target cell.

---

## 9. Read/explore/probe output format

The library must expose structured results and a deterministic agent-readable text renderer.

The CLI will use these renderers.

### 9.1 `read`

Purpose: produce a compact structural view.

Example CLI output:

```text
M.S0001 section columns=2 orientation=landscape
M.P0001 heading level=1 styleId=Heading1 text="Executive Summary"
M.P0002 paragraph list numId=42 level=0 abstractNumId=7 format=decimal level-text="%1." text="Revenue increased"
M.B0001 bookmark name="ClientName" ooxml-id=1 story="main" part=/word/document.xml start=M.P0002 end=M.P0002 complete=True
M.CC0001 content-control kind=plain-text story="main" part=/word/document.xml target=M.P0002 tag="client_name" alias="Client Name" text-length=4
M.F0001 field kind=complex type=REF story="main" part=/word/document.xml target=M.P0002 code="REF ClientName \h" cached-result="Acme" result-text-length=4 nesting-depth=0 bookmark-dependencies="ClientName" safe-edit=flags-only dirty=True complete=True
M.L0001 hyperlink story="main" part=/word/document.xml target=M.P0002 relationship-id=rLink uri="https://example.test/report" uri-scheme=https uri-valid=true tooltip="Open report" target-frame="_blank" history=false external=True broken=False display-text-length=6
M.T0001 table rows=2 columns=3 styleId=TableGrid grid-columns=3 header-row=true
  M.T0001.R01 row cells=3 header=true
  M.T0001.R01.C01 physical-column=1 text="Metric"
  M.T0001.R01.C02 physical-column=2 text="Q3"
  M.T0001.R01.C03 physical-column=3 text="Q4"
M.I0001 image layout=anchor part=/word/media/image1.png content-type=image/png bytes=12345 relationship-id=rImage target=M.P0002 size-emu=914400x457200 description="Revenue chart" wrap=wrapSquare behind-doc=true wrap-dist-top-emu=10 position-h-relative=column position-h-offset-emu=12345 crop-left-percent=10 crop-top-percent=5
```

`read --summary` prints aggregate counts without listing every target:

```text
parts count=50
main-document-part=/word/document.xml
paragraphs count=1510
tables count=41
images count=0
sections count=4
bookmarks count=12
content-controls count=4
fields count=8
hyperlinks count=3
story="main" paragraphs=1510
```

### 9.2 `outline`

Purpose: show only sections, headings, tables, images, bookmarks, content controls,
headers, and footers.

```text
M.P0001 heading level=1 list numId=1 level=0 label="1." text="Executive Summary"
M.T0001 table rows=4 columns=3 styleId=TableGrid grid-columns=3 header-row=true
M.S0001 section columns=2 orientation=landscape
M.I0001 image layout=inline target=M.P0002 part=/word/media/image1.png
M.B0001 bookmark name="ClientName" start=M.P0002 end=M.P0002
M.CC0001 content-control kind=plain-text target=M.P0002 tag="client_name" alias="Client Name"
M.F0001 field kind=complex type=REF target=M.P0002 code="REF ClientName \h" nesting-depth=0 safe-edit=flags-only
M.L0001 hyperlink target=M.P0002 destination="https://example.test/report" broken=False
```

### 9.3 `find`

Purpose: find targetable text.

```text
M.P0004 list numId=1 level=0 label="1." text="Revenue increased by 8.4% compared with the prior quarter."
M.T0001.R02.C01 text="Revenue"
```

Paragraph `find` matches and heading `outline` lines include compact resolved list
metadata when the source paragraph has numbering.

### 9.4 `dump`

Purpose: inspect a single target in detail.

```text
text="Revenue increased by 8.4% compared with the prior quarter."
runs:
  M.P0004.R0001 text="Revenue increased by "
  M.P0004.R0002 markup=inserted-run revision-id=9 author="Reviewer" timestamp-utc=2026-06-01T12:00:00.0000000+00:00 text="8.4%"
  M.P0004.R0003 markup=hyperlink hyperlink-relationship-id=rLink text="source"
```

`dump --runs --json` also exposes the runs as structured `Runs[]` objects. Run IDs
are renderer IDs and are not the same namespace as change IDs from `changes`.
Hyperlink runs expose `HyperlinkRelationshipId` and `HyperlinkAnchor` when present.
Comment body targets such as `C001.C0001` and `comment:<id>` dump metadata only:
comment ID, comments story/part, anchor/reference target, reviewer metadata, and text
length. They do not print comment body text.

### 9.5 `styles`

Purpose: list styles by style ID and display name.

```text
paragraph styleId=Normal name="Normal" default=true
paragraph styleId=Heading1 name="Heading 1" based-on=Normal next=Normal
character styleId=Emphasis name="Emphasis"
table styleId=TableGrid name="Table Grid"
```

### 9.6 `media`

Purpose: list images and allow CLI extraction to a directory.

CLI output:

```text
M.I0001 image layout=inline part=/word/media/image1.png content-type=image/png bytes=12345 relationship-id=rImage target=M.P0002 size-emu=914400x457200 description="Revenue chart"
```

Linked images are never fetched and are omitted from editable image records. VML,
grouped drawings, charts, SmartArt, OLE objects, equations, and generic shapes are
reported as diagnostics and preserved.

### 9.7 `changes`

Purpose: list tracked-change and comment markup. By default the command does not
print private revision or comment text. `--include-comment-text` explicitly includes
bounded comment body snippets.

Plain output begins with summaries and ends with individual records:

```text
inserted-run count=1
summary group=story key="main" type=inserted-run count=1
target-summary target=M.P0004 count=2 types="inserted-run:1,deleted-run:1"
comment-summary comment-id=3 anchor-target=M.P0004 reference-target=M.P0004 count=4 types="comment:1,comment-range-start:1,comment-range-end:1,comment-reference:1"
M.CH0001 inserted-run story="main" part=/word/document.xml target=M.P0004 target-status=targeted target-source=ancestor text-length=8 children=1 revision-id=9 author="Reviewer" timestamp-utc=2026-06-01T12:00:00.0000000+00:00
```

With `--include-comment-text`, comment summaries and comment body records add
`comment-text-length`, `comment-text`, and `comment-text-truncated` fields.

JSON output includes `Summary`, `GroupSummary`, `TargetSummary`, `CommentSummary`,
and `Changes`.

### 9.8 `context`

Purpose: summarize nearby modeled structure around one target without broad document
text. The default `MaxText` is `0`; callers must opt in to text snippets.

```text
before M.P0003 paragraph story="main" text=""
target M.P0004 paragraph story="main" bookmark-names="ClientName" content-controls="M.CC0001" content-control-tags="client_name" fields="M.F0001" field-codes="REF ClientName \h" field-kinds="complex" field-types="REF" hyperlinks="M.L0001" hyperlink-targets="https://example.test/report" comments="3" comment-bodies="C001.C0001" text=""
after M.P0005 paragraph story="main" text=""
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
target heading:"Executive Summary"
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

target heading:"Executive Summary"
target heading:2:"Executive Summary"
target text:"Revenue increased"
target bookmark:"ClientName"
target content-control:"client_name"
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
expect-row-count 4
expect-column-count 3
expect-cell-count 3
expect-content-type image/png
expect-columns 2
expect-orientation landscape
```

For destructive operations, callers should provide the most specific available guard.
The parser accepts unguarded operations where documented, but fresh-agent workflows
should prefer guards before `apply`. This especially applies to:

```text
replace-text
replace-paragraph
delete-block
set-cell
set-table-style
set-table-metadata
set-row-header
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
* `expect-text` is optional but strongly recommended.
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
preserve-runs true
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
copy-paragraph-properties true
style "Normal"
text <<<
This paragraph is inserted before the revenue paragraph.
>>>
end
```

Rules:

* Target must be a paragraph or table.
* Insert a new paragraph before the target.
* `copy-paragraph-properties true` requires a paragraph target and copies the target
  paragraph `w:pPr`, including style and numbering properties.
* `style`, when supplied, overrides the copied or default paragraph style.

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
* `copy-paragraph-properties true` requires a paragraph target and copies the target
  paragraph `w:pPr`, including style and numbering properties.
* `style`, when supplied, overrides the copied or default paragraph style.
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

### 11.6a Bookmark and content-control text operations

Update a plain-text or guarded rich-text content control:

```text
op set-content-control-text
target M.CC0001
expect-text Current value
text Updated value
end
```

Toggle a checkbox content control:

```text
op set-content-control-checkbox
target M.CC0002
checked true
end
```

Select a dropdown or combo box item:

```text
op set-content-control-choice
target M.CC0003
value south
end
```

Update a date content control:

```text
op set-content-control-date
target M.CC0004
value 2026-07-01T00:00:00Z
display-text 2026-07-01
end

op add-repeating-section-item
target M.CC0005
index 1
text Added item
end

op delete-repeating-section-item
target M.CC0006
end
```

Create a bookmark around a guarded paragraph target:

```text
op add-bookmark
target M.P0004
expect-text Paragraph text to bookmark.
name ClientParagraph
end
```

Replace a simple bookmark range while preserving markers:

```text
op replace-bookmark-text
target M.B0001
text Updated value
end
```

Rename a bookmark and matching same-story internal hyperlink anchors:

```text
op rename-bookmark
target M.B0001
name NewBookmarkName
end
```

Remove complete unreferenced bookmark markers while preserving their content:

```text
op delete-bookmark
target M.B0001
end
```

Rules:

* Content-control targets use IDs from `read` or `outline`, such as `M.CC0001`,
  `H001.CC0001`, or `F001.CC0001`.
* `set-content-control-text` supports plain-text controls (`w:sdtPr/w:text`) and
  preserves the `w:sdt` wrapper and properties.
* `set-content-control-text` also supports rich-text controls (`w:sdtPr/w:richText`
  or no specific kind) when `expect-text` matches the current visible text, the
  content container contains only paragraphs, and the replacement does not cross
  protected OOXML boundaries such as fields, bookmarks, nested controls, comments,
  drawings, or existing revision markup. Replacement writes a single paragraph and
  preserves the `w:sdt` wrapper and properties.
* Content-control edit operations reject controls whose `w:lock` value is present
  and not `unlocked`.
* `set-content-control-checkbox` supports checkbox controls (`w:sdtPr/w:checkBox`),
  updates `w:checked`, and updates the displayed state symbol.
* `set-content-control-choice` supports dropdown and combo box controls
  (`w:sdtPr/w:dropDownList` or `w:comboBox`), requires exactly one of `value` or
  `display-text`, verifies a matching `w:listItem`, and updates the displayed content
  while preserving the `w:sdt` wrapper and properties.
* `set-content-control-date` supports date controls (`w:sdtPr/w:date`), updates
  `w:fullDate` to `value`, and updates the displayed content to `display-text` when
  provided or `value` otherwise while preserving the `w:sdt` wrapper and properties.
* `add-repeating-section-item` and `delete-repeating-section-item` are recognized so
  callers receive explicit `E4315` diagnostics. They fail until DocxEdit safely
  models repeating-section subtree insertion/deletion, including IDs, data bindings,
  and section-boundary preservation.
* Bookmark targets use IDs from `read` or `outline`, such as `M.B0001`, `H001.B0001`,
  or `F001.B0001`.
* `add-bookmark` creates a complete bookmark around one modeled paragraph target,
  preserving paragraph properties and visible text. It rejects invalid or duplicate
  names, supports optional `expect-text`, and fails paragraphs containing protected
  OOXML boundaries such as fields, existing bookmarks, comments, content controls,
  drawings, or existing revision markup.
* `replace-bookmark-text` supports complete paragraph-bounded bookmark ranges whose
  contents do not cross protected OOXML boundaries. Same-paragraph multi-run ranges
  and same-container multi-paragraph ranges are supported. Newline-separated
  replacement text becomes multiple paragraphs. Markers are preserved, and unsupported
  ranges fail instead of flattening surrounding OOXML.
* `rename-bookmark` rejects invalid or duplicate new names, changes `w:bookmarkStart`
  `w:name`, and rewrites same-story `w:hyperlink/@w:anchor` values that referenced
  the old name when that old name is unambiguous.
* `delete-bookmark` removes only the `w:bookmarkStart` and matching `w:bookmarkEnd`
  markers for complete bookmarks. It fails if same-story internal hyperlink anchors
  still reference the bookmark name.

### 11.6b Comment body operations

```text
op add-comment
target M.P0004
expect-text <<<
Reviewed paragraph text.
>>>
text Review note
author Reviewer
initials RV
date 2026-06-07T12:00:00Z
end

op set-comment-text
target comment:3
text Updated review note
end

op resolve-comment
target comment:3
end

op reopen-comment
target C001.C0001
end

op delete-comment
target C001.C0001
end

op add-comment-reply
target comment:3
text Reply text
end

op delete-comment-reply
target comment:3.reply:1
end
```

Rules:

* `add-comment` anchors a new comment to a whole modeled paragraph target such as
  `M.P0004`, `H001.P0002`, or `F001.P0002`. It creates `/word/comments.xml`, the
  main-document comments relationship, the comments content-type override, a new
  numeric comment ID, and matching `commentRangeStart`, `commentRangeEnd`, and
  `commentReference` markers when needed.
* `add-comment` accepts optional `expect-text`, `author`, `initials`, and ISO-8601
  `date` fields. Without `author` or `date`, library options supply the author and
  timestamp.
* Comment targets use `comment:<id>` from `changes` output or comment body IDs such
  as `C001.C0001`.
* `set-comment-text` replaces the body with one paragraph and preserves comment
  metadata such as author, initials, timestamp, and OOXML comment ID.
* `resolve-comment` and `reopen-comment` toggle the matching `commentsExtended.xml`
  `w15:done` flag. When a basic comment lacks modern metadata, DocxEdit adds a
  `w15:paraId`, creates `/word/commentsExtended.xml` and the main-document
  relationship/content-type override when needed, and appends a matching
  `w15:commentEx` record. Unsupported body shapes fail with `E4312`.
* `delete-comment` removes the comment body and matching `commentRangeStart`,
  `commentRangeEnd`, and `commentReference` markers from document stories, plus
  matching `commentsExtended.xml` records when present.
* `add-comment-reply` and `delete-comment-reply` are recognized so callers receive
  explicit `E4314` diagnostics. They fail until DocxEdit safely models
  commentsIds/threaded-comments metadata.
* Selected-range comment creation and full threaded comment workflows are out of
  scope for v0.1.

### 11.6c Field flag operations

```text
op set-field-dirty
target M.F0001
dirty true
end

op set-field-lock
target M.F0001
locked true
end

op set-field-code
target M.F0001
expect-code REF OldBookmark \h
code REF NewBookmark \h
end

op set-field-result
target M.F0001
expect-result Old cached result
text New cached result
end

op refresh-field-result
target M.F0001
expect-result Old cached result
end
```

Rules:

* Field targets use IDs from `read` or `outline`, such as `M.F0001`, `H001.F0001`,
  or `F001.F0001`.
* `target all` updates every modeled field in main/header/footer stories.
* `set-field-dirty` updates `w:dirty` and `set-field-lock` updates `w:fldLock` on
  `w:fldSimple` or the complex field begin `w:fldChar`.
* `set-field-code` updates `w:fldSimple/@w:instr`, supports an optional normalized
  `expect-code` guard, preserves the cached result, and marks that field dirty.
* `set-field-result` replaces the cached result runs inside `w:fldSimple`, supports
  an optional exact `expect-result` guard, preserves the field code and field
  boundary, and does not trigger document-level field-update marking when it is the
  only patch operation.
* `refresh-field-result` updates simple `w:fldSimple` REF/PAGEREF/NOTEREF cached
  results from exactly one same-part bookmark whose range is a simple same-paragraph
  range without protected OOXML boundaries. It supports optional normalized
  `expect-code` and exact `expect-result` guards and does not trigger
  document-level field-update marking when it is the only patch operation.
* Complex-field code/result replacement is not supported yet and fails with `E4313`.
* DocxEdit does not recalculate field results beyond the limited REF-style refresh
  above.

### 11.6d Hyperlink operations

Update an existing hyperlink destination:

```text
op set-hyperlink-target
target M.L0001
uri https://example.test/report
tooltip Open report
target-frame _blank
history false
end
```

Update display text:

```text
op set-hyperlink-text
target M.L0001
text Open report
end
```

Insert a hyperlink paragraph after a paragraph or table target:

```text
op insert-hyperlink-after
target M.P0004
text Appendix
anchor AppendixA
end
```

Remove hyperlink markup while preserving its child runs:

```text
op remove-hyperlink
target M.L0001
end
```

Rules:

* Hyperlink targets use IDs from `read` or `outline`, such as `M.L0001`,
  `H001.L0001`, or `F001.L0001`.
* Destination operations require exactly one of `uri` or `anchor`.
* `uri` must be an absolute `http`, `https`, or `mailto` URI and is stored as a
  hyperlink relationship with `TargetMode="External"`. Relative targets, malformed
  URIs, `file`, UNC/file-style targets, and unsafe schemes are rejected.
* `anchor` is stored as `w:anchor` and removes an unused old external relationship.
* Optional `tooltip`, `target-frame`, and `history true|false` update `w:tooltip`,
  `w:tgtFrame`, and `w:history`.
* Updating one hyperlink that shares a relationship with another hyperlink must allocate
  a new relationship ID so the other hyperlink keeps its destination.

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
* `expect-text`, `expect-row-count`, and `expect-column-count` are supported guards.
* Preserve `w:tcPr`.
* Replace cell content with a single paragraph.
* If the cell contains multiple paragraphs, nested tables, images, or fields, fail unless `force true` is supplied.
* `force true` still preserves `w:tcPr`.

### 11.7a `set-table-style`

```text
op set-table-style
target M.T0001
expect-style ExistingStyle
style TableGrid
end
```

Rules:

* Target must be a table.
* `expect-style` is an optional guard against the current `w:tblStyle` value.
* Create `w:tblPr` and `w:tblStyle` when missing.
* Preserve table grid, rows, cells, and existing table properties.

### 11.7b `set-table-metadata`

```text
op set-table-metadata
target M.T0001
expect-caption Existing caption
caption Updated caption
description <<<
Updated accessibility description.
>>>
end
```

Rules:

* Target must be a table.
* At least one of `caption` or `description` is required.
* `expect-caption` and `expect-description` are optional guards against the
  current `w:tblCaption` and `w:tblDescription` values. Missing and empty values
  are equivalent for these guards.
* Non-empty `caption` and `description` values create or update
  `w:tblPr/w:tblCaption` and `w:tblPr/w:tblDescription`.
* Empty heredoc values remove the corresponding metadata element.
* Preserve table grid, rows, cells, style, and other existing table properties.

### 11.7c `set-row-header`

```text
op set-row-header
target M.T0001.R01
expect-header false
header true
end
```

Rules:

* Target must be a table row.
* `header` and `expect-header` use `true` or `false`.
* Setting `header true` creates `w:trPr/w:tblHeader` when missing.
* Setting `header false` removes `w:tblHeader` while preserving other row
  properties.

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
* `expect-row-count` and `expect-column-count` are supported guards.
* Only support rectangular tables without `gridSpan`, `gridBefore`, `gridAfter`,
  or vertical merges in v0.1.
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
* `expect-row-count`, `expect-column-count`, `expect-cell-count`, and
  `expect-contains` are supported guards.
* Do not allow deleting the only row of a table.
* Fail on visual-grid tables with `gridSpan`, `gridBefore`, `gridAfter`, or
  vertical merges in v0.1 unless `force true` is explicitly supported by that
  operation.

### 11.11b Unsupported table-column transforms

Recognized but unsupported operations:

```text
op append-column
target M.T0001
cell New row 1 value
cell New row 2 value
end

op insert-column-before
target M.T0001
column 2
cell New row 1 value
end

op delete-column
target M.T0001
column 2
end
```

Rules:

* `target` is a table ID.
* `column` is a one-based visual column index for insert/delete operations.
* These operations are recognized so callers receive explicit `E4316`
  diagnostics. They fail until DocxEdit safely models table-column transforms
  across `tblGrid`, horizontal spans, omitted cells, nested tables, and vertical
  merge state.

### 11.12 `replace-image`

```text
op replace-image
target M.I0001
asset "assets/revenue-chart.png"
alt "Updated revenue chart"
end
```

Rules:

* Target must be an image.
* Resolve `asset` through `IDocxAssetProvider`.
* CLI resolves asset paths relative to the patch file directory.
* Support PNG and JPEG initially.
* Detect content type from magic bytes, not only extension.
* `expect-content-type` is a supported guard.
* Preserve existing drawing extents when replacing media bytes.
* Replacing a floating image’s bytes is allowed; changing floating layout is not.
* Linked images and non-picture drawing shapes are not valid image targets.
* `preserve-size` is not supported.
* Update content type declarations as needed.

### 11.13 `insert-image-after`

```text
op insert-image-after
target M.P0012
asset "assets/architecture.png"
width 5.5in
alt "Architecture diagram"
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
* `expect-content-type` is a supported guard.
* `caption` is not supported.

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

### 11.15 `set-image-metadata`

```text
op set-image-metadata
target M.I0001
alt "Updated revenue chart"
title "Revenue chart"
name "Revenue picture"
end
```

Rules:

* Update DrawingML `wp:docPr` `descr`, `title`, and `name` without replacing media
  bytes.
* Require at least one of `alt`, `title`, or `name`.
* `expect-content-type` is a supported guard.
* Preserve image bytes, size, layout, and crop metadata.

### 11.16 `set-image-size`

```text
op set-image-size
target M.I0001
width 2in
end
```

Rules:

* Update DrawingML `wp:extent` and picture transform extents without replacing media
  bytes.
* At least one of `width` or `height` is required.
* Width and height accept the same `in`, `cm`, `pt`, `px`, and `emu` units as
  `insert-image-after`.
* If only `width` or `height` is provided, preserve the current aspect ratio when an
  existing extent or media pixel size is available.
* `expect-content-type` is a supported guard.
* Preserve image bytes, crop, and other layout metadata.

### 11.17 `set-image-wrap`

```text
op set-image-wrap
target M.I0001
mode top-bottom
dist-top 1pt
dist-right 2pt
end
```

Rules:

* Update anchored DrawingML `wp:wrap*` mode and anchor wrap-distance attributes
  without replacing media bytes.
* At least one of `mode`, `dist-top`, `dist-bottom`, `dist-left`, or `dist-right` is
  required.
* Supported modes are `none`, `square`, `tight`, `through`, and `top-bottom`, plus
  their `wp:wrap*` element names.
* Distance fields accept the same `in`, `cm`, `pt`, `px`, and `emu` units as image
  dimensions and are written as EMUs.
* Inline images fail with an explicit diagnostic instead of being converted to
  anchored images.
* `expect-content-type` is a supported guard.
* Preserve image bytes, size, crop, and position metadata.

### 11.18 `set-image-position`

```text
op set-image-position
target M.I0001
horizontal-relative page
horizontal-offset -0.25in
vertical-relative paragraph
vertical-align bottom
end
```

Rules:

* Update anchored DrawingML `wp:positionH` and `wp:positionV` relative bases,
  signed offsets, or alignments without replacing media bytes.
* At least one position field is required.
* Supported horizontal `relativeFrom` values are `page`, `margin`, `column`,
  `character`, `leftMargin`, `rightMargin`, `insideMargin`, and `outsideMargin`.
* Supported vertical `relativeFrom` values are `page`, `margin`, `paragraph`, `line`,
  `topMargin`, `bottomMargin`, `insideMargin`, and `outsideMargin`.
* Offsets accept signed `in`, `cm`, `pt`, `px`, and `emu` dimensions and are written
  as EMUs.
* `*-offset` and `*-align` are mutually exclusive for the same axis.
* Inline images fail with an explicit diagnostic instead of being converted to
  anchored images.
* `expect-content-type` is a supported guard.
* Preserve image bytes, size, crop, wrap mode, and wrap distances.

### 11.19 `set-image-crop`

```text
op set-image-crop
target M.I0001
left-percent 12.5
top-percent 5
right-percent 0
end
```

Rules:

* Update DrawingML `a:srcRect` crop percentages without replacing media bytes.
* At least one of `left-percent`, `top-percent`, `right-percent`, or
  `bottom-percent` is required.
* Omitted crop sides keep their current value.
* Zero-valued crop sides are removed from `a:srcRect`; if all sides are zero, remove
  `a:srcRect`.
* Opposing side sums (`left-percent` + `right-percent`, `top-percent` +
  `bottom-percent`) must remain below 100.
* `expect-content-type` is a supported guard.
* Preserve image bytes, size, and layout.

### 11.20 `delete-image`

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
end
```

Rules:

* Target must be a section.
* `expect-columns` and `expect-orientation` are supported guards.
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
  --output output.docx \
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

Current tracked-change generation support:

```text
replace-text          simple text-only matches in one paragraph
replace-paragraph     whole-paragraph text replacement without concurrent style change
insert-before         inserted paragraph text as w:ins
insert-after          inserted paragraph text as w:ins
delete-block          simple paragraph targets as w:del text
set-style             paragraph property revision with w:pPrChange
set-cell              simple single-paragraph cells as w:del/w:ins text
```

The supported tracked text shapes must not contain tabs or line breaks, must not cross
protected OOXML boundaries, must not be inside existing revision markup, and must have
compatible direct run-property shape. `delete-block` tracked output is paragraph-only.
`set-cell` tracked output is limited to simple single-paragraph cells without
`force true`. `set-style` records the previous paragraph properties in `w:pPrChange`.

Known preserve-only operations:

```text
set-content-control-text
set-content-control-checkbox
set-content-control-choice
set-content-control-date
replace-bookmark-text
rename-bookmark
delete-bookmark
add-comment
set-comment-text
resolve-comment
reopen-comment
delete-comment
set-field-dirty
set-field-lock
set-field-code
set-field-result
refresh-field-result
set-hyperlink-target
set-hyperlink-text
insert-hyperlink-after
remove-hyperlink
set-table-style
set-table-metadata
set-row-header
append-row
insert-row-before
insert-row-after
delete-row
replace-image
insert-image-after
set-image-alt
set-image-metadata
set-image-size
set-image-wrap
set-image-position
set-image-crop
delete-image
set-section-columns
set-section-orientation
```

In `Suggest`, these preserve-only operations apply directly and return `W4001`
warnings. The diagnostic message includes the operation's shared catalog support
value, such as `preserve-only`.

In `Require`, they fail with `E6001`. The diagnostic message includes the
operation's shared catalog support value so integrations can distinguish
preserve-only operations from unclassified operations.

Revision metadata:

* Use `DocxEditOptions.Author`.
* Use `DocxEditOptions.TimestampUtc`.
* Generate monotonically increasing `w:id` values by scanning existing revision IDs and incrementing from the max.
* Preserve pre-existing tracked-change and comment markup unless the requested target
  would directly replace a protected boundary; protected-boundary edits fail or fall
  back according to `Require`/`Suggest`.

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
operation index=1 name=replace-text target=M.P0004 success=True
operation index=2 name=append-row target=M.T0001 success=True
  affected id=M.T0001.R03 kind=row action=append parent=M.T0001 row=3 rows-before=2 rows-after=3 columns=2 cells=2
  affected id=M.T0001.R03.C01 kind=cell action=append parent=M.T0001.R03 row=3 column=1 rows-before=2 rows-after=3 columns=2
```

Use `--report <path>` for the full JSON operation report.

### 13.2 `apply`

`Apply` must do everything `Check` does and then write a new `.docx` to the output stream.

It must never modify the input stream.

### 13.3 Internal validation

Because the production library cannot use the Open XML SDK, internal validation is not full schema validation. The dependency policy is explicit: the production library stays dependency-free beyond the .NET platform, so validation is a layered internal invariant checker rather than an ISO/IEC 29500 XSD validator.

`DocxValidateOptions.Profile` controls how much validation is run:

* `Structural` is the default and runs package/XML root checks plus DocxEdit's
  WordprocessingML invariants.
* `Package` limits validation to safe package loading and XML root checks.

A strict schema profile is not exposed until DocxEdit has a real schema validator
bridge.
`DocxValidateOptions.MaxDiagnostics` caps returned validation diagnostics. When
diagnostics are omitted, validation adds `E9199` if omitted diagnostics include
errors or `W9199` when only warnings are omitted.

Patch apply validation must check:

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

`DocxEditor.Validate` and `docxedit validate` must additionally scan safe XML parts
without editing and report stable `E91xx` diagnostics with `PartName` metadata for:

* expected roots for relationship parts and known WordprocessingML parts (`document`, `styles`, `numbering`,
  `settings`, `comments`, `commentsExtended`, headers, footers, footnotes, and
  endnotes);
* paired bookmark and comment range start/end IDs;
* comment body, anchor, and reference ID consistency;
* duplicate bookmark names and duplicate content-control tag/alias values with
  candidate metadata IDs;
* missing paragraph style definitions as warnings when `/word/styles.xml` is
  present;
* missing numbering and abstract numbering definitions as warnings when
  `/word/numbering.xml` is present;
* settings metadata validity, including `w:updateFields`;
* header/footer section references and relationship types;
* section property validity for columns and orientation;
* content-control metadata validity, including `w:id`, `w:lock`, and checkbox
  `w:checked` values;
* modern `commentsExtended.xml` paraId consistency;
* complex field begin/separate/end balance, instruction-text containment, cached-result containment, and dirty/lock flag validity;
* DrawingML `a:blip` relationship references, image target parts, and image media
  content types;
* duplicate DrawingML `wp:docPr` IDs within a Word part;
* DrawingML `wp:extent` positivity and `a:srcRect` crop bounds;
* basic table row/cell shape;
* table visual-grid consistency for `gridSpan`, `gridBefore`, `gridAfter`,
  `vMerge` continuations, and declared `tblGrid` width.

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

Structural validation codes:

```text
E9101 XML part has no root element
E9102 known WordprocessingML part has an unexpected root
E9103 bookmark or comment range start/end IDs are unbalanced
E9104 complex field begin/end markers are unbalanced
E9105 drawing references a missing relationship ID
E9106 table or row is missing required row/cell structure
E9107 duplicate drawing wp:docPr ID within one Word part
E9108 commentsExtended metadata has missing, duplicate, or orphan paraId records
E9109 drawing wp:extent is missing or non-positive
E9110 drawing a:srcRect crop values are malformed or collapse the visible area
E9111 comment body, anchor, or reference IDs are missing, duplicated, or inconsistent
E9112 field instruction text, result containment, or dirty/lock flag metadata is malformed
E9113 drawing image relationship target or media content type is invalid
E9114 table visual-grid metadata is inconsistent
E9115 content-control metadata is malformed
E9118 settings metadata is malformed
E9119 header/footer references are missing or use the wrong relationship type
E9120 section properties are malformed
E9199 validation diagnostics were capped and omitted diagnostics include errors
W9109 duplicate semantic selectors make bookmark/content-control selectors ambiguous
W9116 paragraph style reference is not defined in /word/styles.xml
W9117 numbering or abstract numbering reference is not defined in /word/numbering.xml
W9199 validation diagnostics were capped and omitted diagnostics are warnings only
```

Unsupported or approximated feature diagnostics must be emitted as warnings during read/check/apply when they can affect the requested workflow. Repeated occurrences of the same unsupported feature in the same part should be aggregated into one warning unless the exact target list is useful for fixing the patch.

Current read-model unsupported-feature warning coverage:

```text
W1001 tracked-change markup detected
W1002 hyperlink markup detected
W1003 field markup detected
W1004 comment anchors detected
W1005 bookmark ranges detected
W1006 content controls detected
W1007 floating DrawingML detected
W1008 external image relationship detected
W1009 chart detected
W1010 SmartArt or diagram content detected
W1011 equation detected
W1012 generic shape detected
W1013 altChunk detected
W1014 complex section flow detected
W1015 broken hyperlink relationship detected
W1016 invalid hyperlink URI detected
W1017 missing hyperlink anchor detected
W1018 duplicate hyperlink anchor detected
W1019 linked image detected
W1020 VML drawing detected
W1021 grouped drawing detected
W1022 OLE object detected
W1023 unsupported internal part hyperlink detected
W1024 picture bullet numbering detected
W1025 unsupported numbering format detected
```

These warnings do not mean the library may corrupt the document. They mean the feature is preserved, ignored, approximated, or made read-only according to the operation semantics.

Examples:

```text
E1201 selector matched 0 targets:
  target heading:"Executive summary"

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
docxedit help dump
docxedit help context
docxedit help changes
docxedit help validate
docxedit help check
docxedit help apply

docxedit read input.docx
docxedit read input.docx --summary
docxedit outline input.docx
docxedit find input.docx "some text"
docxedit dump input.docx --id M.P0004 --runs
docxedit context input.docx --id M.P0004
docxedit styles input.docx
docxedit media input.docx
docxedit media input.docx --extract media
docxedit validate input.docx
docxedit validate input.docx --profile package
docxedit changes input.docx

docxedit check input.docx edits.docxpatch
docxedit apply input.docx edits.docxpatch --output output.docx
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
--summary
--headers-footers
--all-stories
--view final|original|markup
--radius <count>
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
target heading:"Executive Summary"
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
alt "Updated revenue chart"
end
```

### 16.6 Set a section to two columns

```text
docxpatch 1

op set-section-columns
target M.S0002
count 2
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
* Surface normalized field code, parsed type, nesting depth, bookmark/hyperlink
  dependencies, and safe-edit status.
* Do not attempt to evaluate fields.
* If `MarkFieldsDirtyWhenEditing` is true, set the document settings `w:updateFields`
  flag so Word can refresh field results on open.
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
public sealed record DocxStyleInfo(string StyleId, string Name, string Type, bool IsDefault)
{
    public string? BasedOnStyleId { get; init; }
    public string? NextStyleId { get; init; }
    public string? LinkedStyleId { get; init; }
    public string? NumberingId { get; init; }
    public int? NumberingLevel { get; init; }
}
```

The public style list includes paragraph, character, and table styles. Numbering styles
may influence resolved paragraph list metadata but are not listed as editable styles.

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

The CLI must be self-explanatory for a fresh coding agent. Help material belongs to
the NuGet library through `DocxHelp`; the CLI is only one renderer of that catalog.

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
  context    Show nearby structure around one target without broad text
  styles     List paragraph, character, and table styles
  media      List embedded images
  validate   Validate package and WordprocessingML invariants
  changes    List tracked-change and comment markup without printing private text

Patch:
  check      Validate a .docxpatch file without writing output
  apply      Apply a .docxpatch file and write a new .docx

Help:
  help dump|context|changes|validate|check|apply|patch

Examples:
  docxedit read report.docx [--view final|original|markup]
  docxedit read report.docx --summary
  docxedit dump report.docx --id M.P0004 --runs
  docxedit context report.docx --id M.P0004
  docxedit media report.docx --extract media
  docxedit validate report.docx
  docxedit changes report.docx
  docxedit check report.docx edits.docxpatch
  docxedit apply report.docx edits.docxpatch --output report.edited.docx
```

Command-specific help must exist for:

```text
docxedit help dump
docxedit help context
docxedit help changes
docxedit help validate
docxedit help check
docxedit help apply
docxedit help patch
```

`docxedit help patch` must include:

* DSL preamble
* operation block syntax
* heredoc syntax
* selector examples
* operation examples
* track-changes explanation and generated per-operation support table
* warning that `expect-hash` is unsupported
* warning that `preserve-size` is unsupported

`docxedit help changes` must describe `Summary`, `GroupSummary`, `TargetSummary`,
`CommentSummary`, target status/source/reason fields, comment anchor fields,
timestamp shape, range `paired-change-id`, and the explicit
`--include-comment-text` / `--max-comment-text` privacy boundary.

`docxedit help dump` must explain that `Runs[]` is structured JSON metadata and that
dump run IDs are separate from `changes` change IDs.

`docxedit help context` must state that default `--max-text` is `0`.

`docxedit help validate` must describe structural package validation, WordprocessingML
invariants, JSON diagnostics, and the fact that diagnostics include stable codes and
part names.

---

## 25. Machine-readable output

The CLI should support `--json` for automation.

Use `System.Text.Json`.

Example shape:

```json
{
  "Success": true,
  "Diagnostics": [],
  "Operations": [
    {
      "Index": 1,
      "OperationName": "replace-text",
      "Target": "M.P0004",
      "Success": true,
      "Diagnostics": []
    }
  ]
}
```

The public library exposes data models directly; JSON is a CLI rendering concern.
Plain-text summaries are not CLI-only: reusable text renderers are exposed through
`DocxTextRenderer`. Current CLI JSON uses `System.Text.Json` default property names,
matching the public PascalCase result property names such as `Success`, `Diagnostics`,
`Summary`, `TargetSummary`, `Runs`, and `Items`.

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
4. CLI can read, outline, find, dump, context, styles, media, validate, changes, check, and apply.
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
19. `changes` exposes private-text-free summaries by default, target/comment
    rollups, comment anchor joins, target localization metadata, range pair links,
    and explicit opt-in bounded comment text snippets.
20. `dump --runs --json` exposes structured run markup metadata.
21. `context` can summarize nearby modeled structure with `MaxText = 0` by default.
22. `check` and `apply` expose operation-level reports in JSON and compact plain text.
23. Command-specific help exists for `dump`, `context`, `changes`, `validate`,
    `check`, `apply`, and `patch`.
24. Golden tests prove deterministic read/probe output.
25. Stream-only tests prove no file-system dependency in the library.
26. Public edit-case validation tools can run at least one smoke case.
27. Private-case tooling rejects tracked or out-of-directory confidential inputs.
28. Optional Office tests can validate generated files on Windows when enabled.

---

## 28. Recommended implementation order

1. Create solution and projects.
2. Implement ZIP package loader/saver.
3. Implement content types and relationships parser.
4. Discover main document, styles, headers, footers, media.
5. Implement logical scanner for paragraphs, tables, images, sections.
6. Implement deterministic ID assignment.
7. Implement read/outline/find/dump/context/styles/media/changes renderers.
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
21. Implement command-specific help and JSON output.
22. Implement CLI diagnostics JSON and strict mode.
23. Add golden tests.
24. Add public edit-case validation tools.
25. Add private-case validation guardrails and agent challenge probes.
26. Add optional Office integration tests.
27. Publish local `Lokad.DocxEdit` NuGet package and verify the CLI consumes the library through project reference only.

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
