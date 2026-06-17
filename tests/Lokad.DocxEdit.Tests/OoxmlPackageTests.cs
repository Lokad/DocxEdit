using System.IO.Compression;
using System.Text;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Tests;

public static class OoxmlPackageTests
{
    [Fact]
    public static void LoadDiscoversMainDocumentThroughPackageRelationship()
    {
        using MemoryStream stream = CreateMinimalDocx();

        OoxmlPackage package = OoxmlPackage.Load(stream, new OoxmlPackageOptions());

        Assert.Equal("/word/document.xml", package.MainDocumentPartName);
        Assert.NotNull(package.GetPart("/word/document.xml"));
        Assert.NotNull(package.GetPart("word/document.xml"));
    }

    [Fact]
    public static void LoadRejectsUnsafeZipEntryPath()
    {
        using MemoryStream stream = CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "../word/document.xml", "<w:document />");
        });

        InvalidDataException ex = Assert.Throws<InvalidDataException>(() =>
            OoxmlPackage.Load(stream, new OoxmlPackageOptions()));
        Assert.Contains("Unsafe OOXML", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void LoadRejectsDuplicateNormalizedPartNames()
    {
        using MemoryStream stream = CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", PackageRelationshipsXml());
            AddEntry(archive, "word/document.xml", DocumentXml());
            AddEntry(archive, "WORD/document.xml", DocumentXml());
            AddEntry(archive, "word/_rels/document.xml.rels", RelationshipsXml());
        });

        InvalidDataException ex = Assert.Throws<InvalidDataException>(() =>
            OoxmlPackage.Load(stream, new OoxmlPackageOptions()));
        Assert.Contains("duplicate part", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public static void LoadRejectsMissingRequiredPackageRelationships()
    {
        using MemoryStream stream = CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "word/document.xml", DocumentXml());
            AddEntry(archive, "word/_rels/document.xml.rels", RelationshipsXml());
        });

        InvalidDataException ex = Assert.Throws<InvalidDataException>(() =>
            OoxmlPackage.Load(stream, new OoxmlPackageOptions()));
        Assert.Contains("/_rels/.rels", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void LoadPreservesUnknownSafeParts()
    {
        using MemoryStream stream = CreateMinimalDocx(archive =>
            AddEntry(archive, "custom/data.bin", "opaque"));

        OoxmlPackage package = OoxmlPackage.Load(stream, new OoxmlPackageOptions());

        OoxmlPart? part = package.GetPart("/custom/data.bin");
        Assert.NotNull(part);
        Assert.Null(part.ContentType);
        Assert.Equal("opaque", Encoding.UTF8.GetString(part.Bytes));
    }

    [Fact]
    public static void LoadRejectsXmlDtd()
    {
        using MemoryStream stream = CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", """
                <!DOCTYPE Types [
                  <!ELEMENT Types ANY>
                ]>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types" />
                """);
        });

        Assert.ThrowsAny<Exception>(() => OoxmlPackage.Load(stream, new OoxmlPackageOptions()));
    }

    [Fact]
    public static void PublicReadReturnsDiagnosticForXmlDtd()
    {
        using MemoryStream stream = CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", """
                <!DOCTYPE Types [
                  <!ELEMENT Types ANY>
                ]>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types" />
                """);
        });
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal("E0001", diagnostic.Code);
        Assert.Equal(DocxSeverity.Error, diagnostic.Severity);
    }

    [Fact]
    public static void PublicReadReturnsDiagnosticForMalformedHeaderXml()
    {
        using MemoryStream stream = CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", PackageRelationshipsXml());
            AddEntry(archive, "word/document.xml", DocumentXml());
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/header1.xml", """
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:p>
                """);
        });
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream, new DocxReadOptions { IncludeHeadersFooters = true });

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E0001" && diagnostic.Severity == DocxSeverity.Error);
        Assert.NotEmpty(result.PartNames);
    }

    [Fact]
    public static void PublicStylesReturnsDiagnosticForMalformedStylesXml()
    {
        using MemoryStream stream = CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", PackageRelationshipsXml());
            AddEntry(archive, "word/document.xml", DocumentXml());
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/styles.xml", """
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style>
                """);
        });
        var editor = new DocxEditor();

        DocxStylesResult result = editor.Styles(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E0001" && diagnostic.Severity == DocxSeverity.Error);
        Assert.Empty(result.Styles);
    }

    [Fact]
    public static void GetRelationshipsResolvesRelativeAndExternalTargets()
    {
        using MemoryStream stream = CreateMinimalDocx();
        OoxmlPackage package = OoxmlPackage.Load(stream, new OoxmlPackageOptions());

        IReadOnlyList<OoxmlRelationship> relationships = package.GetRelationships("/word/document.xml");

        Assert.Contains(relationships, relationship => relationship.Id == "rImage" && relationship.ResolvedTarget == "/word/media/image1.png");
        Assert.Contains(relationships, relationship => relationship.Id == "rExternal" && relationship.IsExternal && relationship.ResolvedTarget is null);
    }

    [Fact]
    public static void LoadRejectsRelationshipTargetsEscapingPackageRoot()
    {
        using MemoryStream stream = CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", PackageRelationshipsXml());
            AddEntry(archive, "word/document.xml", DocumentXml());
            AddEntry(archive, "word/_rels/document.xml.rels", RelationshipsXml("../../escape.xml"));
        });

        InvalidDataException ex = Assert.Throws<InvalidDataException>(() =>
            OoxmlPackage.Load(stream, new OoxmlPackageOptions()));
        Assert.Contains("escapes", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void LoadRejectsMissingInternalRelationshipTarget()
    {
        using MemoryStream stream = CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", PackageRelationshipsXml());
            AddEntry(archive, "word/document.xml", DocumentXml());
            AddEntry(archive, "word/_rels/document.xml.rels", RelationshipsXml("media/missing.png"));
        });

        InvalidDataException ex = Assert.Throws<InvalidDataException>(() =>
            OoxmlPackage.Load(stream, new OoxmlPackageOptions()));

        Assert.Contains("targets missing part", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void LoadRejectsMacroEnabledDocumentByDefault()
    {
        using MemoryStream stream = CreateMinimalDocx(macroEnabled: true);

        InvalidDataException ex = Assert.Throws<InvalidDataException>(() =>
            OoxmlPackage.Load(stream, new OoxmlPackageOptions()));

        Assert.Contains("Macro-enabled", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void LoadAllowsMacroEnabledDocumentWhenExplicitlyConfigured()
    {
        using MemoryStream stream = CreateMinimalDocx(macroEnabled: true);

        OoxmlPackage package = OoxmlPackage.Load(stream, new OoxmlPackageOptions(AllowMacroEnabledDocuments: true));

        Assert.Equal("/word/document.xml", package.MainDocumentPartName);
    }

    [Fact]
    public static void ApplyValidatesTouchedDocumentPartsBeforeSave()
    {
        using MemoryStream stream = CreateMinimalDocx();
        OoxmlPackage package = OoxmlPackage.Load(stream, new OoxmlPackageOptions());
        package.ReplacePartBytes(
            "/word/document.xml",
            Encoding.UTF8.GetBytes("""
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />
                """));
        var patch = new DocxPatch(true, 1, [], []);

        PatchExecutionResult result = DocxPatchEngine.Apply(package, patch, new DocxEditOptions());

        Assert.False(result.Success);
        Assert.Contains(package.TouchedPartNames, partName => partName == "/word/document.xml");
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9001" &&
            diagnostic.PartName == "/word/document.xml" &&
            diagnostic.Message.Contains("w:body", StringComparison.Ordinal));
    }

    [Fact]
    public static void PublicReadReturnsDiagnosticForInvalidPackage()
    {
        using var stream = new MemoryStream("not a zip"u8.ToArray());
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E0001" && diagnostic.Severity == DocxSeverity.Error);
    }

    [Fact]
    public static void PublicReadReturnsDiagnosticForInvalidMainDocumentNamespace()
    {
        using MemoryStream stream = CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", PackageRelationshipsXml());
            AddEntry(archive, "word/document.xml", "<root />");
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
                """);
        });
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E0001" && diagnostic.Message.Contains("WordprocessingML", StringComparison.Ordinal));
    }

    private static MemoryStream CreateMinimalDocx(Action<ZipArchive>? extra = null, bool macroEnabled = false)
    {
        return CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml(macroEnabled));
            AddEntry(archive, "_rels/.rels", PackageRelationshipsXml());
            AddEntry(archive, "word/document.xml", DocumentXml());
            AddEntry(archive, "word/_rels/document.xml.rels", RelationshipsXml());
            AddEntry(archive, "word/media/image1.png", "png");
            extra?.Invoke(archive);
        });
    }

    private static MemoryStream CreatePackage(Action<ZipArchive> configure)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            configure(archive);
        }

        stream.Position = 0;
        return stream;
    }

    private static void AddEntry(ZipArchive archive, string name, string text)
    {
        ZipArchiveEntry entry = archive.CreateEntry(name);
        using Stream stream = entry.Open();
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static string ContentTypesXml(bool macroEnabled = false)
    {
        string mainDocumentContentType = macroEnabled
            ? "application/vnd.ms-word.document.macroEnabled.main+xml"
            : "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
        return $$"""
            <?xml version="1.0" encoding="utf-8"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/word/document.xml" ContentType="{{mainDocumentContentType}}"/>
            </Types>
            """;
    }

    private static string PackageRelationshipsXml()
    {
        return """
            <?xml version="1.0" encoding="utf-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """;
    }

    private static string RelationshipsXml(string imageTarget = "media/image1.png")
    {
        return $$"""
            <?xml version="1.0" encoding="utf-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{{imageTarget}}"/>
              <Relationship Id="rExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.test/image.png" TargetMode="External"/>
            </Relationships>
            """;
    }

    private static string DocumentXml()
    {
        return """
            <?xml version="1.0" encoding="utf-8"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
              </w:body>
            </w:document>
            """;
    }
}
