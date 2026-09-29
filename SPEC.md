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

The solution (`Lokad.DocxEdit.slnx`) holds the `Lokad.DocxEdit` library and
its local CLI (`src/Lokad.DocxEdit.Cli`, never packaged), the `Lokad.DocxEdit.Tests`
and `Lokad.DocxEdit.OfficeTests` suites, PowerShell tooling (`tools/`),
data-driven edit cases (`edit-cases/`), and these docs (`docs/`).
`artifacts/`, `private-cases/`, `bin/`, and `obj/` are ignored build outputs.

Dependency rules (enforced in build, not just written here): the production
package carries no consumer dependency entries — the only production
`PackageReference` is build-only `Microsoft.SourceLink.GitHub` with
`PrivateAssets="All"`, and `Directory.Build.targets` rejects anything else so
new consumer dependencies are intentional. Test projects may use xUnit and its
runner SDK only; `PackagingTests` asserts the shipped `.nuspec` has no
dependencies. Shared MSBuild settings live in `Directory.Build.props`
(nullable, implicit usings, deterministic compiler builds) with `-tl:off` in
`Directory.Build.rsp` for stable automation logging.

A Release pack must emit both `.nupkg` and `.snupkg` under ignored
`artifacts/nuget/`. Other configurations fail unless
`/p:AllowNonReleasePackage=true` is passed for troubleshooting
(`PackagingTests` covers both). The CLI project is not packaged; it references
only the library and system libraries. It can be published as a single
executable for local use with a runtime identifier and `PublishSingleFile`:

```bash
dotnet publish src/Lokad.DocxEdit.Cli/Lokad.DocxEdit.Cli.csproj \
  -c Release \
  -r win-x64 \
  --self-contained false \
  -p:PublishSingleFile=true
```
## 5. Public API requirements

The public API must be stream-first and command-line agnostic.

### 5.1 Main entry point

Expose a simple façade:

```csharp
namespace Lokad.DocxEdit;

public sealed class DocxEditor
{
    public DocxReadResult Read(
        Stream input,
        DocxReadOptions options,
        CancellationToken cancellationToken);

    public DocxOutlineResult Outline(
        Stream input,
        DocxOutlineOptions options,
        CancellationToken cancellationToken);

    public DocxFindResult Find(
        Stream input,
        string query,
        DocxFindOptions options,
        CancellationToken cancellationToken);

    public DocxDumpResult Dump(
        Stream input,
        string targetId,
        DocxDumpOptions options,
        CancellationToken cancellationToken);

    public DocxContextResult Context(
        Stream input,
        string targetId,
        DocxContextOptions options,
        CancellationToken cancellationToken);

    public DocxStylesResult Styles(
        Stream input,
        DocxStylesOptions options,
        CancellationToken cancellationToken);

    public DocxMediaResult Media(
        Stream input,
        DocxMediaOptions options,
        CancellationToken cancellationToken);

    public DocxMediaExtractResult ExtractMedia(
        Stream input,
        DocxMediaOptions options,
        CancellationToken cancellationToken);

    public DocxValidateResult Validate(
        Stream input,
        DocxValidateOptions options,
        CancellationToken cancellationToken);

    public DocxChangesResult Changes(
        Stream input,
        DocxChangesOptions options,
        CancellationToken cancellationToken);

    public DocxPatch ParsePatch(
        TextReader patchReader,
        CancellationToken cancellationToken);

    public DocxCheckResult Check(
        Stream input,
        TextReader patchReader,
        DocxEditOptions options,
        CancellationToken cancellationToken);

    public DocxApplyResult Apply(
        Stream input,
        TextReader patchReader,
        Stream output,
        DocxEditOptions options,
        CancellationToken cancellationToken);
}
```

Each operation also offers overloads dropping the trailing options and cancellation parameters; no parameter carries a default value. The library must not require file paths. The CLI may provide path-based wrappers. Every long-running public operation must observe the cancellation token while loading the package, parsing XML, scanning document stories, resolving assets, validating, and writing output.

Resource limits bind reads before allocation and parsing: ZIP entry metadata (counts, declared sizes, duplicate names) is validated before any part content is read; `[Content_Types].xml` is read through a bounded copy before parsing; every part is buffered through bounded copies with the actual total enforced alongside the declared total; patch text is read through a bounded character scan (`DocxEditOptions.MaxPatchChars`); image assets are copied through a bound (`MaxSinglePartBytes`, the same quota that bounds parts, since assets become parts); and the edited package total is re-checked against `MaxUncompressedBytes` before publication. XML parsing itself (DTD prohibited, no entity expansion) always operates on quota-bounded buffers.

Cancellation is cooperative at chunk, part, patch-read, operation, and validation boundaries: an in-flight synchronous parse, ZIP inflate, or save completes once started, while loops between those units observe the token and scoped cleanup (temporary files, open streams, partial buffers) still runs through `finally`/`using`.

Stream ownership follows one matrix on every exit (success, failure, or cancellation): the loader accepts input ownership at entry and disposes the input if and only if `LeaveInputOpen` is false, always disposing its internal non-seekable copy; `check`/`apply` dispose the input and (for apply) the output on every exit under the same flags, so parse, load, edit, save, and cancellation failures never leak owned streams; the patch reader stays caller-owned and is never disposed by the library; asset streams opened by a provider are engine-owned once `TryOpen` returns true and are disposed after reading on every outcome.

### 5.2 Asset provider

Patches need to reference image assets. Since the library must work without file-system access, asset references are logical names resolved through an asset provider.

```csharp
namespace Lokad.DocxEdit;

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

    public DocxPackageLimits Quotas { get; init; } = DocxPackageLimits.Default;

    public bool AllowMacroEnabledDocuments { get; init; } = false;
    public bool MarkFieldsDirtyWhenEditing { get; init; } = true;
}
```

Read-only option classes compose the same `DocxPackageLimits` quotas and default `LeaveInputOpen` to true,
including `DocxValidateOptions`.
`DocxReadOptions`, `DocxOutlineOptions`, `DocxFindOptions`, `DocxDumpOptions`, and
`DocxContextOptions` also accept `DocxTextView` (`Final`, `Original`, or `Markup`)
where visible text is rendered. `DocxContextOptions.MaxText` defaults to `0`; callers opt in when context
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

Findings travel as `DocxDiagnostic` records: three positional values (severity, code, human message) plus init-only location metadata. Operation outcomes share the `DocxOperationResult` base (success flag plus causally ordered diagnostics).

`Check` and `Apply` must include per-operation reports with the operation index, name, target text, success flag, diagnostics, affected targets, generated revision IDs, and bounded before/after preview text when previews are enabled.

Read and explore results must expose structured data (IDs, match and outline records, text bounded at build) in addition to what the CLI text renderings show; renderers format records but never replace them.

The records in `src/Lokad.DocxEdit/Public/` are the definition of every shape above; this section states the rules they must satisfy, not a second copy of their members.

`Changes` must be private-text-free by default. It may report revision/comment
metadata, text lengths, child counts, IDs, targets, stories, and parts. It must not
copy revision text into the result. Comment body snippets may appear only when
`DocxChangesOptions.IncludeCommentText` is true, and then must be bounded by
`MaxCommentText`.

### 5.4 Public integration surfaces

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
    public static string RenderPatchTrackChangesSupportTable();
}
```

`DocxCommandCatalog` must include structured commands, options, examples, output
fields, privacy notes, and patch operation specs. The CLI must render its overview
and command-specific help from this catalog. `docxedit help patch` and the
published patch documentation must use `RenderPatchTrackChangesSupportTable()` for
the track-change support matrix.

```csharp
public sealed record DocxPatchOperationInfo
{
    public string Name { get; init; } = string.Empty;
    public string Category { get; init; } = string.Empty;
    public IReadOnlyList<string> RequiredFields { get; init; } = [];
    public IReadOnlyList<IReadOnlyList<string>> RequiredAlternatives { get; init; } = [];
    public IReadOnlyList<string> OptionalFields { get; init; } = [];
    public IReadOnlyList<string> RepeatableFields { get; init; } = [];
    public string Description { get; init; } = string.Empty;
    public string TrackChangesSupportClass { get; init; } = "unsupported";
    public string TrackChangesSupport { get; init; } = "unsupported";
    public string TrackChangesNote { get; init; } = "Suggest applies directly with W4001; Require fails with E6001.";
    public bool GeneratesTrackedChanges { get; }
}
```

`TrackChangesSupportClass` is the stable behavior class consumed by integrations
and by the patch engine. Values include `text-run`, `paragraph-block`,
`paragraph-property`, `table-property`, `row-property`, `cell-property`,
`row-structure`, `section-property`, `relationship-metadata`, `preserve-only`,
and `unsupported`.
`GeneratesTrackedChanges` is true when the support class creates new revision
markup. `TrackChangesSupport` remains the more specific operation-level support
value, such as `tracked-simple`, `tracked-paragraph`,
`tracked-paragraph-insert`, `tracked-paragraph-delete`, `tracked-style`,
`tracked-cell-simple`, `preserve-only`, or `unsupported`. `preserve-only`
operation notes must begin with an operation-specific rationale, then state that
existing tracked-change markup is preserved but no new revision markup is created;
they apply directly under `Suggest` with `W4001` and fail under `Require` with
`E6001`. `unsupported` is reserved for known operations that intentionally fail
with stable diagnostics, or for unknown/unclassified tracked-output behavior.
Supported tracked operations may still reject an unsafe shape with `W4002` or
`E6002`. Those diagnostics include the operation name, target ID, catalog support
value, and exact unsupported-shape reason. `W4002` additionally states that
`Suggest` is applying the edit directly. Track-change diagnostics must also set
machine-readable metadata: `W4001`/`E6001` use
`Feature = "track-changes-no-revision-representation"`, `W4002`/`E6002` use
`Feature = "track-changes-unsupported-target-shape"`, suggest fallbacks use
`Fallback = "direct-edit-preserve-existing-revisions"`, and require failures use
`Fallback = "require-failed"`. `Suggest` and `Require` must reject empty revision
authors before operation execution with `E6003`, `Feature =
"track-changes-revision-metadata"`, and `Fallback = "no-output-written"`.
Generated revision authors are trimmed, and generated revision dates are rendered
as invariant UTC ISO-8601 values. When no author is supplied, the CLI and public
options default to `docxedit`; when no timestamp is supplied, the options default
to the current UTC instant.

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

`ReadSummary` and `ContextMetadataOnly` use `MaxText = 0`. `MaxText` bounds per-item body text in both rendered text and structured/JSON payloads: paragraph/run/cell text, field cached results, and table/image captions/descriptions/titles are truncated (0 drops). IDs, counts, style IDs/names, bookmark names, content-control tags/aliases, field codes/kinds/types, hyperlink URIs/anchors, authors, and revision IDs are retained as structural metadata; length/count properties keep full-text values. `DocxPrivacyPresets.RenderReadSummary` emits counts only; `RenderContextMetadata` and `RenderChangesMarkup` strip item text and comment snippets even when the input result carries them. Prefer these renderers over serializing full result objects for privacy-safe agent output.

---

## 6. Internal package model

Implement an internal package layer, not exposed as a general-purpose Open XML SDK.

```csharp
internal sealed class OoxmlPackage
{
    public IReadOnlyDictionary<string, OoxmlPart> Parts { get; }
    public OoxmlPart ContentTypesPart { get; }
    public string MainDocumentPartName { get; }

    public static OoxmlPackage Load(
        Stream input,
        OoxmlPackageOptions options,
        CancellationToken cancellationToken);

    public OoxmlPart? GetPart(string partName);
    public IReadOnlyList<OoxmlRelationship> GetRelationships(
        string sourcePartName,
        CancellationToken cancellationToken);

    public void Save(Stream output, CancellationToken cancellationToken);
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

Part roles resolve through package relationships, never fixed paths: styles,
numbering, settings, comments, headers, footers, footnotes, and endnotes are
discovered from the main document part relationships (conventional paths such
as `/word/styles.xml` are only used when creating new parts). A part shared by
several relationships of one kind is processed once under the first
relationship index. Footnotes, endnotes, and comment bodies stay out of the
read model (no paragraph/table targets) but are inventoried by changes and
validated; parts without a relationship role receive no role-based checks.

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

IDs are physical and positional: every top-level w:p / w:tbl / w:sectPr in document order owns one ordinal, including blocks wrapped in a single block-level w:ins / w:del / w:moveFrom / w:moveTo container. Table rows own physical ordinals; table cells use visual grid columns (gridBefore plus gridSpan). Text views filter which blocks are reported but never renumber the survivors, so Final omits deleted blocks (leaving gaps), Original omits inserted blocks, and Markup reports everything. The same physical enumeration backs read, dump, context, outline, find, changes, and patch target resolution, so a discovered ID always resolves to the same element. IDs are stable for the same document bytes and same scanner version, but they are not permanently stable across independently modified input versions. Within one patch, explicit paragraph/table/row/cell/section IDs bind to the input snapshot: earlier structural operations never renumber a later explicit ID, a target deleted earlier in the patch fails instead of editing a neighbour, and blocks inserted by the patch itself are not addressable by pre-discovered explicit IDs in the same patch (address them with a result alias via as/@name, a semantic selector, or a follow-up patch; see Result aliases in docs/patch-format.md). Semantic selectors resolve live against the evolving package in operation order. Guards evaluate sequentially against current content.

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

Counters are 1-based with minimum zero-padded widths (4 for entities, 2 for rows and cells, 4 for merge groups, 3 for story parts); larger counters extend the width, and parsers accept any wider all-digit form while rejecting non-digits, zeros, signs, and short runs.

### 8.3 Paragraph model

```csharp
public sealed record DocxParagraphInfo(
    DocxTargetId Id,
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
    public DocxLabelStatus LabelStatus { get; init; } = DocxLabelStatus.NotResolved;
    public IReadOnlyList<string> LabelWarnings { get; init; } = [];
    public int? StartValue { get; init; }
    public string? Suffix { get; init; }
    public bool IsLegal { get; init; }
    public int? RestartAfterLevel { get; init; }
    public string? ParagraphStyleId { get; init; }
    public DocxLabelSource Source { get; init; } = DocxLabelSource.Direct;
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

List labels are computed from the numbered paragraphs visible in the selected
`DocxTextView`. Block-level `w:ins` and `w:moveTo` paragraphs participate in
final and markup counters but not original counters. Block-level `w:del` and
`w:moveFrom` paragraphs participate in original and markup counters but not final
counters. Run-level revision text inside an existing numbered paragraph changes
the paragraph text view but does not change numbering participation. Paragraph
property revisions that store previous `w:numPr` state are preserved and reported
with `W1026`; original-view reconstruction of those previous numbering
properties is not modeled.

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
    DocxTargetId Id,
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
    public required DocxTargetId Id { get; init; }
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
    DocxTargetId Id,
    int RowIndex,
    int ColumnIndex,
    string Text,
    int ColumnSpan,
    DocxVerticalMerge? VerticalMerge,
    bool HasNestedTable)
{
    public int PhysicalColumnIndex { get; init; }
    public int VisualColumnEndIndex { get; init; }
    public DocxTargetId? MergeGroupId { get; init; }
    public DocxTargetId? VerticalMergeRootCellId { get; init; }
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
    DocxTargetId Id,
    string PartName,
    string? ContentType,
    long ByteLength)
{
    public string LayoutKind { get; init; } = "unknown";
    public string? RelationshipId { get; init; }
    public DocxTargetId? ContainingTargetId { get; init; }
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
    DocxTargetId Id,
    string Story,
    int Columns,
    DocxOrientation Orientation);
```

### 8.8 Bookmark and content-control model

`DocxEditor.Read` exposes bookmark and content-control metadata so agents can inspect
selector candidates without raw OOXML.

```csharp
public sealed record DocxBookmarkInfo
{
    public required DocxTargetId Id { get; init; }
    public string Name { get; init; } = string.Empty;
    public string? OoxmlId { get; init; }
    public string Story { get; init; } = string.Empty;
    public string PartName { get; init; } = string.Empty;
    public DocxTargetId? StartTargetId { get; init; }
    public DocxTargetId? EndTargetId { get; init; }
    public bool IsComplete { get; init; }
    public bool IsNameDuplicate { get; init; }
    public IReadOnlyList<string> DuplicateNameBookmarkIds { get; init; } = [];
}

public sealed record DocxContentControlInfo
{
    public required DocxTargetId Id { get; init; }
    public string Story { get; init; } = string.Empty;
    public string PartName { get; init; } = string.Empty;
    public DocxTargetId? TargetId { get; init; }
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
    public string? SafeEditReason { get; init; }
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
`unsupported-group`, or `unsupported-repeating-section`. `SafeEditReason`
explains non-editable statuses such as picture containers, group containers,
repeating-section subtree edits, and lock values without exposing control text.

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
    public required DocxTargetId Id { get; init; }
    public string Story { get; init; } = string.Empty;
    public string PartName { get; init; } = string.Empty;
    public DocxTargetId? TargetId { get; init; }
    public string Kind { get; init; } = "unknown";
    public string? FieldType { get; init; }
    public string Code { get; init; } = string.Empty;
    public IReadOnlyList<string> Arguments { get; init; } = [];
    public IReadOnlyList<string> Switches { get; init; } = [];
    public string CachedResultText { get; init; } = string.Empty;
    public int ResultTextLength { get; init; }
    public int NestingDepth { get; init; }
    public IReadOnlyList<string> BookmarkDependencies { get; init; } = [];
    public IReadOnlyList<string> HyperlinkDependencies { get; init; } = [];
    public DocxRefreshPolicy RefreshPolicy { get; init; } = DocxRefreshPolicy.Unsupported;
    public string? RefreshReason { get; init; }
    public bool CanRefreshDeterministically { get; init; }
    public string SafeEditStatus { get; init; } = "unknown";
    public bool? IsDirty { get; init; }
    public bool? IsLocked { get; init; }
    public bool IsComplete { get; init; }
}
```

Field IDs use the `F` namespace, for example `M.F0001`. They are metadata IDs, not
general text edit targets. `CachedResultText` is the stored visible result currently
in the document, not a recalculated value. `Arguments`, `Switches`, `RefreshPolicy`,
`RefreshReason`, and `CanRefreshDeterministically` summarize parsed field-code
behavior for agent planning. `set-field-dirty` and `set-field-lock` explicitly
accept field metadata IDs.

### 8.10 Hyperlink model

`DocxEditor.Read` exposes hyperlink metadata for external relationship links,
internal anchors, relationship targets, and broken relationship IDs. External
targets include URI scheme validation metadata. Internal anchors include
missing-bookmark and duplicate-bookmark flags.

```csharp
public sealed record DocxHyperlinkInfo
{
    public required DocxTargetId Id { get; init; }
    public string Story { get; init; } = string.Empty;
    public string PartName { get; init; } = string.Empty;
    public DocxTargetId? TargetId { get; init; }
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
    public required DocxChangeId Id { get; init; }
    public string Type { get; init; } = string.Empty;
    public string Story { get; init; } = string.Empty;
    public string PartName { get; init; } = string.Empty;
    public string? ParentType { get; init; }
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
    public string? CommentDurableId { get; init; }
    public bool? CommentIsReply { get; init; }
    public bool? CommentResolved { get; init; }
    public string? CommentAnchorTargetId { get; init; }
    public string? CommentReferenceTargetId { get; init; }
    public string? CommentAnchorStory { get; init; }
    public string? CommentAnchorPartName { get; init; }
    public int? CommentTextLength { get; init; }
    public string? CommentTextSnippet { get; init; }
    public bool CommentTextTruncated { get; init; }
    public DocxTargetStatus TargetStatus { get; init; } = DocxTargetStatus.Targetless;
    public DocxTargetSource TargetSource { get; init; } = DocxTargetSource.None;
    public DocxTargetReason? TargetReason { get; init; }
    public string? NearestTargetId { get; init; }
    public string? TargetNote { get; init; }
    public DocxChangeId? PairedChangeId { get; init; }
}
```

`CommentParaId`, `CommentParentParaId`, `CommentRootParaId`, `CommentIsReply`,
and `CommentResolved` are populated from Word's modern `commentsExtended.xml`
metadata when it is present. `CommentDurableId` is populated from
`commentsIds.xml` when a matching `w16cid:paraId` record is present. They are
safe metadata fields and do not expose comment body text.
`DocxCommentThreadSummary` exposes the same root/reply, durable ID, and resolution
fields, plus the same opt-in text snippet fields, at the thread level:

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
    public string? DurableId { get; init; }
    public bool? IsReply { get; init; }
    public bool? Resolved { get; init; }
    public int? TextLength { get; init; }
    public string? TextSnippet { get; init; }
    public bool TextTruncated { get; init; }
    public int Count { get; init; }
    public IReadOnlyList<DocxChangeSummary> Summary { get; init; } = [];
}
```

`ParentType` is a normalized immediate WordprocessingML parent kind such as
`paragraph`, `run-properties`, `paragraph-properties`, `cell-properties`,
`section-properties`, `comment`, or `body`. `TargetStatus` is `targeted`,
`comment-anchor`, or `targetless`. `TargetSource`
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
    public IReadOnlyList<string> CommentParaIds { get; init; } = [];
    public IReadOnlyList<string> CommentParentParaIds { get; init; } = [];
    public IReadOnlyList<string> CommentRootParaIds { get; init; } = [];
    public IReadOnlyList<string> CommentDurableIds { get; init; } = [];
    public IReadOnlyList<string> CommentReplyIds { get; init; } = [];
    public IReadOnlyList<string> CommentResolvedIds { get; init; } = [];
    public string? Caption { get; init; }
    public string? Description { get; init; }
    public int? RowCount { get; init; }
    public int? ColumnCount { get; init; }
    public int? RowIndex { get; init; }
    public int? ColumnIndex { get; init; }
    public int? ColumnSpan { get; init; }
    public int? VisualColumnEndIndex { get; init; }
    public string? MergeGroupId { get; init; }
    public DocxVerticalMerge? VerticalMerge { get; init; }
    public string? VerticalMergeRootCellId { get; init; }
    public bool HasNestedTable { get; init; }
}
```

Paragraph and cell context items include `CommentIds` and `CommentBodyIds` when
comment anchors or references are attached to that target. They also expose
privacy-safe thread metadata through `CommentParaIds`, `CommentParentParaIds`,
`CommentRootParaIds`, `CommentDurableIds`, `CommentReplyIds`, and
`CommentResolvedIds`. Comment body targets such as `C001.C0001` and
`comment:<id>` return a `Kind = "comment"` context item whose metadata links back
to the anchor target and does not include comment body text.
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

`read --summary` prints aggregate counts without listing every target (with `--json` it emits the same counts-only shape, not the full result object):

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

`outline` uses `DocxOutlineOptions.TextView`. In `markup` view, heading text and
list labels include block-level inserted/deleted headings according to the same
numbering visibility policy as `read`, `find`, `dump`, and `context`.

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
When the target owns tracked markup records, `dump` appends a `changes:` block
with change IDs, types, parent type, revision metadata, and child element counts.
This target-level output is privacy-safe for property revisions such as
`w:pPrChange`, `w:tblPrChange`, `w:trPrChange`, `w:tcPrChange`, and
`w:sectPrChange` because it does not print raw OOXML or revision text.
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
bounded comment body snippets. `--operation-report <path>` accepts a check/apply
JSON report and annotates scanned revisions whose `revision-id` appears in the
report's `GeneratedRevisionIds`.

Plain output begins with summaries and ends with individual records:

```text
inserted-run count=1
summary group=story key="main" type=inserted-run count=1
target-summary target=M.P0004 count=2 types="inserted-run:1,deleted-run:1"
comment-summary comment-id=3 anchor-target=M.P0004 reference-target=M.P0004 count=4 types="comment:1,comment-range-start:1,comment-range-end:1,comment-reference:1"
M.CH0001 inserted-run story="main" part=/word/document.xml parent=paragraph target=M.P0004 target-status=targeted target-source=ancestor operation-index=1 operation-name=replace-text operation-target=M.P0004 text-length=8 children=1 revision-id=9 author="Reviewer" timestamp-utc=2026-06-01T12:00:00.0000000+00:00
```

With `--include-comment-text`, comment summaries and comment body records add
`comment-text-length`, `comment-text`, and `comment-text-truncated` fields.

JSON output includes `Summary`, `GroupSummary`, `TargetSummary`, `CommentSummary`,
and `Changes`. When `--operation-report` is supplied, matching `Changes` records
include `OperationIndex`, `OperationName`, and `OperationTarget`.

### 9.8 `context`

Purpose: summarize nearby modeled structure around one target without broad document
text. The default `MaxText` is `0`; callers must opt in to text snippets.

```text
before M.P0003 paragraph story="main" text=""
target M.P0004 paragraph story="main" bookmark-names="ClientName" content-controls="M.CC0001" content-control-tags="client_name" fields="M.F0001" field-codes="REF ClientName \h" field-kinds="complex" field-types="REF" hyperlinks="M.L0001" hyperlink-targets="https://example.test/report" comments="3,4" comment-bodies="C001.C0001,C001.C0002" comment-para-ids="00PARENT,00REPLY1" comment-parent-para-ids="00PARENT" comment-root-para-ids="00PARENT" comment-durable-ids="DURABLEP,DURABLER" comment-reply-ids="4" text=""
after M.P0005 paragraph story="main" text=""
```

---

### 9.9 capabilities

Purpose: describe supported, conditional, and unsupported edits for one
paragraph, content-control, cell, merge-group, bookmark, table, row, section,
hyperlink, field, or image target under an effective
track-change policy (--track-changes off, preserve, suggest, or require).
Each operation carries an actionable reason, a help topic naming the operation
reference, and a safe alternative operation or target form when one exists.

Example text output:

    target M.P0001 paragraph main
    supported replace-text: Plain paragraph text. ...
    unsupported add-bookmark alternative=add-comment: Paragraph contains a protected hyperlink boundary, so bookmark creation fails with E4311. ...

Support values reuse the execution predicates, so guidance agrees with check:
protected boundaries, lock and kind gates, merge shape, section properties,
orphaned ranges, and tracked-shape validation come from the same checks that
gate patch execution. Verdicts distinguish target-level facts from
patch-dependent conditions: whether a particular text span, replacement,
style name, anchor-text span, or force flag succeeds depends on the exact
patch, and is reported as conditional until checked. An unknown target, or a
target kind without a capability model, fails with E1201.

Capabilities output is metadata-only: it carries no document body text.
Capabilities describe the inspected document state and can change after an
edit; inspection is guidance, while check remains authoritative for a patch.
With JSON output, the Capabilities object exposes the resolved target ID,
kind, story, and per-operation support, reason, help topic, and alternative.

---

### 9.10 template

Purpose: print a guarded patch template for one discovered target of any
capability-covered kind (see 9.9). The template carries the discovered target
ID, current values in guards or active fields, one active block when a
check-clean starter exists, and commented
example blocks for the remaining supported operations. The active block passes
check under the requested track-change policy and changes nothing visible;
commented blocks use placeholder content and are enabled by removing the
leading hash and space from each line. Unsupported operations are omitted and
conditional ones carry their condition, following the capabilities verdicts in
9.9. Templates embed current target values in guards; they are working material
for one target, not broad document text. With JSON output, the result carries
the template text plus the capabilities it was generated from.

---

## 10. Patch DSL: `.docxpatch`

Field sets and worked examples live in docs/patch-format.md; the operation registry is their machine owner. This section states the grammar and resolution contracts.

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

Escape processing applies to quoted single-line values only: bare and unbalanced values
stay verbatim, unknown backslash sequences stay literal, and heredoc content stays raw.

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
* heading selectors match real outline levels (direct outline levels, style inheritance, or built-in Heading 1-9), not digits in style IDs.
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
set-cell-shading
set-table-style
set-table-metadata
set-row-header
delete-row
replace-image
delete-image
```

---

## 11. Patch operations v0.1

Per-operation behavior contracts. Field sets live in docs/patch-format.md (frozen copies of `DocxHelp.Catalog` output); examples there show the same operations in use.

### 11.1 `replace-text`

Replace text inside one paragraph or table cell.

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

* Target must be a paragraph, table cell, or merge group; merge groups resolve to the root cell.
* Within a cell, matches span cell paragraphs in document order without crossing paragraph boundaries; occurrence counts across those paragraphs and expect-text guards the whole cell text.
* `expect-text` is optional but strongly recommended.
* `find` must match exactly once unless `occurrence` selects one match (`occurrence N`) or requests every match (`occurrence all`). An ambiguous match without `occurrence` fails instead of replacing every match.
* If replacement is within a single run, split the run and preserve run properties.
* If replacement spans simple adjacent runs, preserve the first matched run’s properties for the replacement.
* Fail if a direct run-preserving replacement span crosses:

  * field boundaries
  * hyperlink boundaries
  * comment range boundaries
  * structured document tag boundaries
  * tracked revision boundaries

* Spans through plain text beside those boundaries succeed and preserve them.
  Tracked edits for such spans preserve the surrounding markup in place and
  record delete and insert revisions for the span.
  Paragraph rewrites with preserve-runs false keep the whole-paragraph rule:
  they fail when the paragraph contains any of the boundaries above, because
  they rebuild the container.
* Preserve `xml:space="preserve"` when replacement text has leading/trailing spaces or repeated spaces.

Optional field:

```text
occurrence 2
occurrence all
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
* `style` may be a style ID or display name, resolved like `set-style`; unknown, ambiguous, or wrong-kind styles fail before publication.
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
* Insert one new paragraph per text field, in file order, before the target.
* `copy-paragraph-properties true` requires a paragraph target and copies the target
  paragraph `w:pPr`, including style and numbering properties, but excludes copied
  `w:pPrChange` and `w:sectPr`. The copy applies to every inserted paragraph.
* `style`, when supplied, overrides the copied or default paragraph style.
* A single style applies to every inserted paragraph; one style per text field applies positionally; any other style count fails.
* An embedded `style` may be a style ID or display name, resolved like `set-style`; unknown, ambiguous, or wrong-kind styles fail before publication.

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
* Insert one new paragraph per text field, in file order, after the target.
* `copy-paragraph-properties true` requires a paragraph target and copies the target
  paragraph `w:pPr`, including style and numbering properties, but excludes copied
  `w:pPrChange` and `w:sectPr`. The copy applies to every inserted paragraph.
* `style`, when supplied, overrides the copied or default paragraph style.
* A single style applies to every inserted paragraph; one style per text field applies positionally; any other style count fails.
* An embedded `style` may be a style ID or display name, resolved like `set-style`; unknown, ambiguous, or wrong-kind styles fail before publication.
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
* Refuse with `E4305` when deletion would orphan a bookmark, comment-range, or complex-field boundary whose counterpart lies outside the deleted element; delete the range first or retarget. Removing an element that fully contains a range is allowed.
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
  Picture controls fail with `E4310` and image/media guidance. Group controls
  fail with `E4310` and guidance to target an editable child control.
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
  and same-container multi-paragraph ranges are supported. Simple table-spanning
  ranges are supported only when all table cell paragraphs can be edited as visible
  text slots; newline-separated replacement text must provide one line per slot so
  paragraph, table, row, and cell structure is preserved. Markers are preserved, and
  unsupported ranges fail instead of flattening surrounding OOXML.
  Under tracked output, simple same-paragraph bookmark ranges made only of compatible
  runs emit `w:del`/`w:ins` between the preserved bookmark markers. Multi-paragraph or
  otherwise complex bookmark ranges warn with `W4002` under `suggest` or fail with
  `E6002` under `require`.
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
anchor-text paragraph text
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
* `add-comment` accepts optional `expect-text`, `anchor-text`, `occurrence`,
  `author`, `initials`, and ISO-8601 `date` fields. Without `author` or `date`,
  library options supply the author and timestamp.
* When `anchor-text` is present, `add-comment` anchors the comment to exactly one
  normalized text span inside the target paragraph. Repeated anchor text requires
  `occurrence`; missing text fails with `E4203`, ambiguous text fails with `E1202`,
  and protected markup boundaries fail with `E4305` or `E4317`.
* Comment targets use `comment:<id>` from `changes` output or comment body IDs such
  as `C001.C0001`.
* `set-comment-text` replaces the body with one paragraph and preserves comment
  metadata such as author, initials, timestamp, and OOXML comment ID. Under
  tracked output, simple paragraph-only comment bodies emit `w:del`/`w:ins` in
  `comments.xml` while preserving existing comment and paragraph metadata;
  complex comment bodies warn with `W4002` under `suggest` or fail with `E6002`
  under `require`.
* `resolve-comment` and `reopen-comment` toggle the matching `commentsExtended.xml`
  `w15:done` flag. When a basic comment lacks modern metadata, DocxEdit adds a
  `w15:paraId`, creates `/word/commentsExtended.xml` and the main-document
  relationship/content-type override when needed, and appends a matching
  `w15:commentEx` record. Unsupported body shapes fail with `E4312`.
* `delete-comment` removes the comment body and matching `commentRangeStart`,
  `commentRangeEnd`, and `commentReference` markers from document stories, plus
  matching `commentsExtended.xml` records when present.
* `add-comment-reply` creates a simple threaded reply in `comments.xml`, ensures
  parent and reply `w15:paraId` values, creates or updates `commentsExtended.xml`
  with `w15:paraIdParent`, and creates `commentsIds.xml` durable ID metadata for
  the new reply.
* `delete-comment-reply` removes a leaf reply comment plus its matching
  `commentsExtended.xml` and `commentsIds.xml` records. Targets can be an explicit
  reply comment target such as `comment:4` or an ordinal selector such as
  `comment:3.reply:1`. Replies with child replies fail with `E4314`.
* Full threaded comment workflows beyond simple reply add/delete remain out of
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
* `set-field-result` replaces the cached result runs inside `w:fldSimple` and
  simple same-paragraph complex fields whose begin/separate/end result topology is
  validated. It supports an optional exact `expect-result` guard, preserves the
  field code and field boundary, and does not trigger document-level field-update
  marking when it is the only patch operation.
* `refresh-field-result` updates simple `w:fldSimple` REF/PAGEREF/NOTEREF cached
  results from exactly one same-part bookmark whose range is a simple same-paragraph
  range without protected OOXML boundaries. It also refreshes simple `QUOTE`
  fields from literal field-code arguments. It supports optional normalized
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

* Target must be a table cell ID such as `M.T0001.R02.C03` or a merge-group ID
  such as `M.T0001.MG0001`.
* `expect-text`, `expect-row-count`, and `expect-column-count` are supported guards.
* Cell IDs use visual grid coordinates. A target column inside a horizontal
  `w:gridSpan` resolves to the spanning cell.
* Merge-group targets resolve to the root cell of the merge group.
* Vertical-merge continuation cells are rejected; target the root cell instead.
* Preserve `w:tcPr`.
* Replace cell content with a single paragraph.
* If the cell contains multiple paragraphs, nested tables, images, or fields, fail unless `force true` is supplied.
* `force true` still preserves `w:tcPr`.

### 11.7a `set-cell-shading`

```text
op set-cell-shading
target M.T0001.R02.C03
expect-fill none
fill A1B2C3
end
```

Rules:

* Target must be a table cell ID such as `M.T0001.R02.C03` or a merge-group ID
  such as `M.T0001.MG0001`.
* `fill` sets `w:tcPr/w:shd/@w:fill` to a 6-digit hexadecimal color or `auto`.
* `clear true` removes the cell shading element. `fill` and `clear true` are
  mutually exclusive.
* `expect-fill` is an optional guard against the current fill; accepted guard
  values are a 6-digit hexadecimal color, `auto`, or `none`.
* Vertical-merge continuation cells are rejected; target the root cell instead.
* Preserve existing cell properties other than the updated shading element.
* Under tracked output, record the previous cell properties in `w:tcPrChange`.

### 11.7b `set-table-style`

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
* `style` may be a table style ID or display name, resolved like `set-style`; unknown, ambiguous, or wrong-kind styles fail before publication.
* Preserve table grid, rows, cells, and existing table properties.
* Under tracked output, record the previous table properties in `w:tblPrChange`.

### 11.7c `set-table-metadata`

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

### 11.7d `set-row-header`

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
* Under tracked output, record the previous row properties in `w:trPrChange`.

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
* The table must have a consistent visual grid.
* Clone the last row’s row properties and cell properties.
* Number of `cell` fields must equal the physical cell count of the cloned row.
* Preserve table style and grid.
* If the last row contains vertical merge cells, fail because appending would
  need to choose whether to extend or terminate those merge chains.
* In tracked output, mark the inserted row through `w:trPr/w:ins`.
* Tracked output is limited to simple rectangular tables without nested tables or
  existing row-level revision markers.

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

Rules are the same as `append-row`, except the template is the target row. The
insertion boundary must not cross an active vertical merge chain, and the target
row must not itself contain vertical merge cells. In tracked output, mark the
inserted row through `w:trPr/w:ins`; `force true`, visual-grid, and other
complex shapes remain a direct-edit fallback under `Suggest` or an `E6002`
failure under `Require`.

### 11.10 `insert-row-after`

```text
op insert-row-after
target M.T0001.R03
cell <<<New metric>>>
cell <<<Q3 value>>>
cell <<<Q4 value>>>
end
```

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
* Refuse with `E4305` when deletion would orphan a bookmark, comment-range, or complex-field boundary whose counterpart lies outside the deleted row; delete the range first or retarget.
* Direct deletion supports consistent visual-grid tables.
* When deleting a row whose vertical-merge root is followed by a matching
  continuation, promote the next continuation from `continue` to `restart`.
* Fail when vertical-merge promotion would require changing a continuation whose
  column span differs from the deleted root.
* In tracked output, keep the row and mark it through `w:trPr/w:del`.
* Tracked output is limited to simple rectangular tables without nested tables or
  existing row-level revision markers; `force true`, visual-grid, and other
  complex shapes remain a direct-edit fallback under `Suggest` or an `E6002`
  failure under `Require`.

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

op insert-column-after
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

### 11.21 `set-section-columns`

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
* In tracked output, record the previous section properties in
  `w:sectPrChange` while preserving header/footer references and page-size
  metadata.
* If the section already contains `w:sectPrChange`, `Suggest` applies the edit
  directly with `W4002` and `Require` fails with `E6002` so existing section
  property revision markup is not replaced.
* Do not attempt to move content between columns.

### 11.22 `set-section-orientation`

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
* In tracked output, record the previous section properties in
  `w:sectPrChange`; the current `w:pgSz` receives the new orientation and
  swapped dimensions, while the prior size/orientation remains in the change
  record.
* If the section already contains `w:sectPrChange`, `Suggest` applies the edit
  directly with `W4002` and `Require` fails with `E6002` so existing section
  property revision markup is not replaced.

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
replace-paragraph     whole-paragraph text replacement, plus w:pPrChange for style
insert-before         inserted paragraph text as w:ins
insert-after          inserted paragraph text as w:ins
delete-block          simple paragraph targets as w:del text
set-style             paragraph property revision with w:pPrChange
set-content-control-text simple plain-text and guarded paragraph-only rich-text content controls as w:del/w:ins inside the content wrapper
replace-bookmark-text simple same-paragraph bookmark ranges as w:del/w:ins between preserved markers
set-comment-text    simple paragraph-only comment bodies as w:del/w:ins inside comments.xml
set-field-result     simple w:fldSimple cached result text as w:del/w:ins inside the field wrapper
set-hyperlink-text    simple hyperlink display text as w:del/w:ins inside the hyperlink wrapper
insert-hyperlink-after simple inserted hyperlink display text as w:ins inside the hyperlink wrapper
set-cell              simple text-only cells as w:del/w:ins text, including compatible multi-paragraph cells
set-cell-shading      cell shading property changes as w:tcPrChange
set-table-style      table style property changes as w:tblPrChange
set-row-header       repeating-row header property changes as w:trPrChange
append-row           simple rectangular row insertions as w:trPr/w:ins
insert-row-before    simple rectangular row insertions as w:trPr/w:ins
insert-row-after     simple rectangular row insertions as w:trPr/w:ins
delete-row           simple rectangular row deletions as w:trPr/w:del
set-section-columns  section column property changes as w:sectPrChange
set-section-orientation section page orientation changes as w:sectPrChange
```

The supported tracked text shapes must not contain tabs, line breaks, soft hyphens,
symbols, or other non-text run content, must not cross protected OOXML boundaries,
must not be inside existing revision markup, and must have compatible direct
run-property shape. `delete-block` tracked output is paragraph-only. `set-cell`
tracked output is limited to text-only cell paragraphs without `force true`; for
compatible multi-paragraph cells, deleted text remains in its original paragraphs
and inserted replacement text is emitted in the first paragraph.
`set-content-control-text` tracked output is limited to simple plain-text content
controls and guarded paragraph-only rich-text content controls. It preserves
`w:sdt` properties and `w:sdtContent`; compatible rich-text replacements preserve
inner paragraph containers, delete existing paragraph text in place, and insert the
replacement in the first paragraph. Complex content controls remain direct under
`suggest` with `W4002` or fail under `require` with `E6002`. `set-field-result`
tracked output is limited to simple `w:fldSimple`
cached result text and preserves the field instruction; simple complex-field
result replacements remain direct under `suggest` or fail under `require`, and
unsafe complex topologies fail with `E4313`. `replace-bookmark-text` tracked output is limited to simple
same-paragraph bookmark ranges with compatible run-only content; multi-paragraph
and table-spanning bookmark replacements remain direct under `suggest` or fail
under `require`.
`set-comment-text` tracked output is limited to simple paragraph-only comment
bodies; it preserves comment metadata, deletes existing paragraph text in place,
and inserts the replacement in the first paragraph.
`set-hyperlink-text`
preserves the hyperlink wrapper and relationship or anchor while replacing simple
display text with generated revisions. `insert-hyperlink-after` tracked output
preserves relationship or anchor metadata and wraps the inserted display text in
`w:ins`; display text with tabs or line breaks remains direct under `suggest` or
fails under `require`. `set-style` records the previous paragraph properties in
`w:pPrChange`. `set-cell-shading` records previous cell properties in
`w:tcPrChange`, `set-table-style` records previous table properties in
`w:tblPrChange`, and `set-row-header` records previous row properties in
`w:trPrChange`. `append-row`, `insert-row-before`, and `insert-row-after` mark
inserted rows with `w:trPr/w:ins`; `delete-row` marks deleted rows with
`w:trPr/w:del`. `set-section-columns` and `set-section-orientation` record
previous section properties in `w:sectPrChange` while preserving page size and
header/footer references.

Overlap policy:

* Existing tracked-change, move, custom XML revision, comment, bookmark,
  content-control, field, hyperlink, image, table, and section markup outside the
  edit range is preserved.
* Generated tracked edits may be adjacent to existing revision or comment markup,
  but must not replace through it.
* Existing `w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`, move range markers, and
  custom XML revision range markers are protected boundaries for generated text
  edits. `TrackChangesMode.Suggest` falls back with `W4002`; `Require` fails
  with `E6002`.
* Generic paragraph/cell text edits do not cross comment ranges, bookmarks,
  content controls, fields, or hyperlinks. Dedicated operations own those wrapper
  surfaces when the shape is simple enough to preserve.
* Generated property revisions are not nested or replaced. If a target already
  owns the same tracked property revision shape, `Suggest` falls back and
  `Require` fails.

Known preserve-only operations:

```text
set-content-control-checkbox
set-content-control-choice
set-content-control-date
rename-bookmark
delete-bookmark
add-comment
resolve-comment
reopen-comment
delete-comment
set-field-dirty
set-field-lock
set-field-code
refresh-field-result
set-hyperlink-target
remove-hyperlink
set-table-metadata
replace-image
insert-image-after
set-image-alt
set-image-metadata
set-image-size
set-image-wrap
set-image-position
set-image-crop
delete-image
```

In `Suggest`, these preserve-only operations apply directly and return `W4001`
warnings. The diagnostic message includes the operation's shared catalog support
value, such as `preserve-only`.

In `Require`, they fail with `E6001`, except declared annotation operations
(add-comment, resolve-comment, reopen-comment, delete-comment,
add-comment-reply, delete-comment-reply), which stay permitted because a comment
is already review markup. The diagnostic message includes the
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

`Check` must parse and validate a patch against an input document without writing output. It executes the same mutations as `Apply` against the disposable in-memory package (including field-refresh marking) and discards them instead of saving, so dependent edits, repeated guards, topology changes, field refresh, and post-edit validation reach the same outcomes; only publication differs. `GeneratedRevisionIds` stays empty in check reports and is populated only for apply operations that actually create revision markup.

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
docxedit apply: OK
operation index=1 name=replace-text target=M.P0004 success=True generated-revision-ids=1,2
operation index=2 name=append-row target=M.T0001 success=True
  affected id=M.T0001.R03 kind=row action=append parent=M.T0001 row=3 rows-before=2 rows-after=3 columns=2 cells=2
  affected id=M.T0001.R03.C01 kind=cell action=append parent=M.T0001.R03 row=3 column=1 visual-column-end=1 rows-before=2 rows-after=3 columns=2
```

Use `--report <path>` for the full JSON operation report. `GeneratedRevisionIds`
is populated only for apply operations that actually create revision markup.
Feed the same report to `changes --operation-report <path>` to annotate matching
revision records with operation metadata.
Affected table targets may also include visual-grid metadata (`visual-column-end`,
`grid-before`, `grid-after`), merge-group IDs, and nested-table paths so agents can
audit edits against complex table topology without inspecting raw OOXML.

Check reports record the effective track-change policy, author, timestamp, and tool version on success and failure alike, so a report pins the policy it ran under.

### 13.2 `apply`

`Apply` must do everything `Check` does and then write a new `.docx` to the output stream.

It must never modify the input stream.

Apply reports record the same effective policy, author, timestamp, and tool version on success and failure alike.

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
* Edited bookmarks, comment ranges, complex fields, and table grids still pair and balance; only violations introduced by the patch fail, pre-existing ones do not block it.
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
* modern `commentsIds.xml` paraId/durableId consistency;
* complex field begin/separate/end balance, instruction-text containment, cached-result containment, and dirty/lock flag validity;
* tracked revision metadata and child-shape checks for `w:ins`, `w:del`,
  `w:delText`, `w:pPrChange`, `w:tblPrChange`, `w:trPrChange`,
  `w:tcPrChange`, and `w:sectPrChange`;
* DrawingML `a:blip` relationship references, image target parts, and image media
  content types;
* duplicate DrawingML `wp:docPr` IDs within a Word part;
* DrawingML `wp:extent` positivity and `a:srcRect` crop bounds;
* basic table row/cell shape;
* table visual-grid consistency for `gridSpan`, `gridBefore`, `gridAfter`,
  `vMerge` continuations, and declared `tblGrid` width.

The schema-validation limitation is documented in help and docs, not emitted as a
default warning on every successful validation. Emitting such a warning by
default would make CLI `--strict` fail otherwise-valid documents even though no
document problem was found.

---

### 13.4 lint

Lint validates patch shape without loading a document: syntax, required
fields, alternative groups, exclusive fields, and explicit target ID kinds. It reuses the registry
contracts and the same diagnostic codes check would produce for shape errors:
E2010 and E2011 for unknown operations and fields, E2012 and E2013 for bad
literals, E2015 for repeats, E4202 for missing fields and unsatisfied
alternative groups, E4205 for conflicting exclusive fields, and E1201 for wrong-kind explicit IDs. Lint failures
always predict check failures; lint success leaves document-dependent
targets, guards, assets, and shapes to check. The result reports the parsed
operation count with operation indexes and names.

## 14. Diagnostics

Code meanings live in docs/diagnostics.md. This section states the code and record contracts.

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
E9108 commentsExtended metadata has missing, duplicate, orphan, or self-parented paraId records
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
E9121 tracked revision markup is malformed
E9122 commentsIds metadata is malformed
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
W1026 tracked numbering property revision detected
```

These warnings do not mean the library may corrupt the document. They mean the feature is preserved, ignored, approximated, or made read-only according to the operation semantics.

Examples:

```text
E1201 selector matched 0 targets: heading:"Executive summary". Nearby headings: M.P0003 heading level=1, M.P0041 heading level=1.
```

Suggestions are the first three headings in document order. Non-heading selectors suggest `Nearby paragraphs:` instead, or `No nearby paragraph targets are available.` when the body has none.

```text
E1202 selector matched 3 targets: M.P0004, M.T0001.R02.C01, M.P0044. Use a more specific selector or an explicit ID.
```

```text
E3201 guard failed for M.P0004. Expected text does not match current text.
```

Text guards report the mismatch without echoing either text.

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
docxedit outline input.docx --view markup
docxedit find input.docx "some text"
docxedit dump input.docx --id M.P0004 --runs
docxedit context input.docx --id M.P0004
docxedit styles input.docx
docxedit media input.docx
docxedit media input.docx --extract media
docxedit validate input.docx
docxedit validate input.docx --profile package
docxedit changes input.docx
docxedit changes output.docx --operation-report apply-report.json

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
```

Options for read/probe commands:

```text
--runs
--summary
--headers-footers
--view final|original|markup
--radius <count>
--max-text <chars>
--operation-report <path>    # changes only
--diagnostics <path>
--strict
--json
```

`--diagnostics <path>` writes the full diagnostics array as JSON for any command. `--report <path>` remains the operation-level patch report for `check` and `apply`. `--operation-report <path>` on `changes` reads a check/apply report and annotates matching revision records.

`apply` publishes file outputs atomically: the edited document lands in a temporary sibling and replaces `--output` only after successful editing, so a failed run preserves any pre-existing destination file (temporary files are removed on failure, cancellation, or I/O errors). `--output`/`--report`/`--diagnostics` must differ from the input document, the patch file, the operation report source, and each other; collisions fail upfront with exit code 2 before any writer opens. `-` keeps streaming standard input/output directly.

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

For each paragraph, matches are visible-text offset ranges:

```csharp
internal readonly record struct TextRange(int Start, int Length);
```

The replacement engine uses these ranges to translate visible character offsets back
to XML nodes, splitting runs per section 17.2. Tabs, breaks, inserted/deleted runs, and
field runs are distinguished while walking runs, not in the range itself.

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
W5103 Document contains fields and was marked for Word-side field refresh; DocxEdit does not recalculate field results.
```

### 17.5 Revision markup validation

Structural validation and post-edit touched-part validation must reject malformed
tracked revision markup with `E9121` for missing or invalid `w:id`, `w:author`,
or `w:date` on generated-revision containers, orphan `w:delText`, and malformed
paragraph property revisions such as `w:pPrChange` without child `w:pPr`.

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

Current behavior:

* `set-cell` works for visual-grid cell IDs and merge-group IDs unless the
  resolved target is a vertical merge continuation.
* Direct row insertion can clone safe visual-grid row shapes when the insertion
  boundary does not cross an active vertical merge chain.
* Direct row deletion can delete safe visual-grid rows and promote the next
  vertical-merge continuation when deleting a root.
* Generated tracked row insertion/deletion remains limited to simple rectangular
  tables.

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

Do not trust file extension alone. Magic bytes are authoritative: assets
without a recognized PNG/JPEG signature are rejected (`E5203`), provider
MIME/name hints must agree with the magic, and a recognized extension that
conflicts with the magic is rejected. Beyond signatures, structural validation
is bounded and decoder-free: PNG requires a leading IHDR (sane dimensions and
encoding), at least one IDAT chunk, and IEND terminating the file; JPEG
requires a frame header with sane dimensions and scan data closed by an end
marker. Touched image parts are re-validated post-edit.

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

v0.1 does not model first-page/even-page/default header semantics; targets resolve by relationship within each header or footer part.

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

### 23.2 Deterministic output

Determinism is scoped from strongest to weakest claim:

1. Repeatable targeting: the same document bytes scanned with the same scanner version produce the same target IDs (SPEC 8.2). IDs are positional, not permanently stable across edits.
2. Equivalent part payloads: the same input, patch, asset bytes, author, and UTC timestamp produce byte-identical ZIP part payloads across separate executions. Pin `--author` and `--timestamp-utc` (or `Author`/`TimestampUtc`); omitting the timestamp records the current time into revision markup, so payloads then differ by design.
3. Stable rendered output: equal payloads render equal text. Representative fixtures assert exact rendered text in `EditCaseTests` readback values; there are no checked-in golden files.
4. Archive bytes are NOT reproducible: ZIP entry timestamps record the save time, so two separately saved packages differ byte-for-byte even when every part payload matches. Compare part payloads or rendered text, never archives.

`Deterministic` in the build props refers to compiler builds, not documents.

### 23.3 Fixture generation

Provide an internal fixture builder that creates minimal `.docx` packages using only system libraries.

Also allow optional Microsoft Word-generated fixtures for real-world compatibility.

### 23.4 Office integration tests

Create `Lokad.DocxEdit.OfficeTests`.

These tests are optional and run only when:

```text
DOCXEDIT_ENABLE_OFFICE_TESTS=1
```

and only on Windows with Microsoft Word installed.

Use late-bound COM via `Type.GetTypeFromProgID("Word.Application")` and `dynamic`. Do not add a production dependency. Avoid adding a Microsoft Office interop package.

Office tests should:

1. Apply representative generated tracked outputs, including text, paragraph,
   and row revisions, to minimal local `.docx` files.
2. Open generated `.docx` files in Word.
3. Save them to a temp output path.
4. Close Word cleanly.
5. Re-read with DocxEdit and fail if representative tracked markup did not
   survive.
6. Fail if Word cannot open/save the document.
7. Optionally export to PDF for visual inspection in local developer workflows.

Always close documents and quit Word in `finally`.

### 23.5 Edit case validation harness

In addition to unit tests, a JSON case catalog covers end-to-end edit workflows.

Public case layout:

```text
edit-cases/
  cases/
    basic-replace.json
  families/
    tracked-change-matrix.json
  fixtures/
    word-smoke.docx
```

Case manifests (`docs/validation.md`) define synthetic WordprocessingML body XML
(or a fixture), a `.docxpatch` payload, apply options, assets, and expected
readback values. The `tracked-change-matrix.json` family lists the tracked-change
coverage slice and the `TrackChangesMode` variants run for each case.

`EditCaseTests` in the test project runs every public case plus every family
variant data-driven from those manifests: it builds inputs, applies patches
through the CLI entry point in-process, and asserts exit codes, diagnostics,
readback values, and change summaries.

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

Usage lives in docs/cli.md. This section states what help must provide.

The CLI must be self-explanatory for a fresh coding agent. Help material belongs to
the NuGet library through `DocxHelp`; the CLI is only one renderer of that catalog.

`docxedit -h`:

```text
docxedit - read and patch .docx documents

Usage:
  docxedit <command> [options]

Read / explore:
  read         Produce an agent-friendly structural view of a .docx
  outline      Show headings, tables, images, sections, headers, footers
  find         Find text and print stable edit targets
  dump         Dump one target in detail
  context      Show nearby structure around one target without broad text
  capabilities Show supported, conditional, and unsupported edits for one target
  template     Print a guarded patch template for one target
  styles       List paragraph, character, and table styles
  media        List embedded images
  validate     Validate package and WordprocessingML invariants
  changes      List tracked-change and comment markup; comment text is opt-in
  catalog      Print the machine-readable command and patch-operation catalog
  version      Print the docxedit version

Patch:
  lint      Validate patch shape without loading a document
  check     Validate a .docxpatch file without writing output
  apply     Apply a .docxpatch file and write a new .docx

Help:
  help read|outline|find|dump|context|capabilities|template|styles|media|validate|changes|catalog|version|lint|check|apply|patch

Examples:
  docxedit read report.docx [--view final|original|markup]
  docxedit read report.docx --summary
  docxedit outline report.docx --view markup
  docxedit dump report.docx --id M.P0004 --runs
  docxedit context report.docx --id M.P0004
  docxedit media report.docx --extract media
  docxedit validate report.docx
  docxedit changes report.docx
  docxedit check report.docx edits.docxpatch
  docxedit apply report.docx edits.docxpatch --output report.edited.docx
Pipes:
  - stands for stdin/stdout: the input .docx, the patch file, and --output accept -
  - with --output -, stdout carries exactly the edited package: status goes to
    stderr (or --report), --json is rejected, the input document and patch file
    cannot both be -, and failed runs emit no stdout bytes

Exit codes:
  0 success (warnings allowed; inspect diagnostics)
  1 unsuccessful result; text mode prints error diagnostics to stderr
  2 invalid command-line usage
  3 successful result with warnings under --strict (a published output document, if any, is still written)
  4 unhandled exception
```

Command-specific help must exist for:

```text
docxedit help dump
docxedit help context
docxedit help capabilities
docxedit help template
docxedit help changes
docxedit help validate
docxedit help check
docxedit help apply
docxedit help lint
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
`CommentSummary`, parent type, target status/source/reason fields, comment anchor
fields, timestamp shape, range `paired-change-id`, `--operation-report`
operation annotations, and the explicit `--include-comment-text` /
`--max-comment-text` privacy boundary.

`docxedit help dump` must explain that `Runs[]` is structured JSON metadata and that
dump run IDs are separate from `changes` change IDs.

`docxedit help context` must state that default `--max-text` is `0`.

`docxedit help capabilities` must describe support, reason, and alternative verdicts per operation under an effective track-change policy, and must state that check remains authoritative for a patch.

`docxedit help template` must describe the check-clean starter block and the commented example blocks generated from target capabilities.

`docxedit help lint` must describe document-independent shape validation and state that lint failures predict check failures while lint success leaves document-dependent targets, guards, assets, and shapes to check.

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

1. The produced `Lokad.DocxEdit` package has no consumer NuGet dependencies; any production project package reference is private build tooling only.
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
24. Edit-case readback assertions prove deterministic rendered output.
25. Stream-only tests prove no file-system dependency in the library.
26. Public edit-case validation tools can run at least one smoke case.
27. Private-case tooling rejects tracked or out-of-directory confidential inputs.
28. Optional Office tests can validate generated files on Windows when enabled.

---

## 28. Implementation history

The project is implemented; the original step-by-step build order is retired.
Current capability and validation coverage are described in docs/status.md.

The most important invariant throughout the implementation is:

```text
Every edit must specify where to edit and what change to make.
Guards (find, expect-text, expect-*) let an edit additionally assert the
expected current content; check and apply fail when a supplied guard
mismatches. Some operations require a guard, elsewhere guards are optional.
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
