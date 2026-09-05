using System.IO.Compression;
using System.Text;

namespace Lokad.DocxEdit.Tests;

public static class OoxmlPackageTests
{
    private const string HelloPatch = """
        docxpatch 1

        op replace-text
        target M.P0001
        find Hello
        with Hello
        end
        """;

    [Fact]
    public static void PublicReadDiscoversMainDocumentThroughPackageRelationship()
    {
        using MemoryStream stream = CreateMinimalDocx();
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.True(result.Success);
        Assert.Equal("/word/document.xml", result.MainDocumentPartName);
        DocxParagraphInfo paragraph = Assert.Single(result.Paragraphs);
        Assert.Equal("M.P0001", paragraph.Id.ToWireValue());
        Assert.Contains("Hello", paragraph.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void PublicReadReturnsDiagnosticForUnsafeZipEntryPath()
    {
        using MemoryStream stream = CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "../word/document.xml", "<w:document />");
        });
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E0001" && diagnostic.Severity == DocxSeverity.Error);
    }

    [Fact]
    public static void PublicReadReturnsDiagnosticForDuplicateNormalizedPartNames()
    {
        using MemoryStream stream = CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", PackageRelationshipsXml());
            AddEntry(archive, "word/document.xml", DocumentXml());
            AddEntry(archive, "WORD/document.xml", DocumentXml());
            AddEntry(archive, "word/_rels/document.xml.rels", RelationshipsXml());
        });
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E0001" && diagnostic.Severity == DocxSeverity.Error);
    }

    [Fact]
    public static void PublicReadReturnsDiagnosticForMissingPackageRelationships()
    {
        using MemoryStream stream = CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "word/document.xml", DocumentXml());
            AddEntry(archive, "word/_rels/document.xml.rels", RelationshipsXml());
        });
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E0001" && diagnostic.Severity == DocxSeverity.Error);
    }

    [Fact]
    public static void PublicApplyPreservesUnknownSafeParts()
    {
        using MemoryStream stream = CreateMinimalDocx(archive =>
            AddEntry(archive, "custom/data.bin", "opaque"));
        using var output = new MemoryStream();
        var editor = new DocxEditor();

        DocxApplyResult result = editor.Apply(stream, new StringReader(HelloPatch), output);

        Assert.True(result.Success);
        output.Position = 0;
        using var archive = new ZipArchive(output, ZipArchiveMode.Read);
        ZipArchiveEntry entry = archive.GetEntry("custom/data.bin")
            ?? throw new InvalidDataException("Missing custom/data.bin.");
        using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
        Assert.Equal("opaque", reader.ReadToEnd());
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
    public static void PublicMediaResolvesRelativeImageTargets()
    {
        using MemoryStream stream = CreateMinimalDocxWithImage();
        var editor = new DocxEditor();

        DocxMediaResult result = editor.Media(stream);

        Assert.True(result.Success);
        DocxImageInfo image = Assert.Single(result.Images);
        Assert.Equal("/word/media/image1.png", image.PartName);
    }

    [Fact]
    public static void PublicReadReturnsDiagnosticForRelationshipTargetEscapingPackageRoot()
    {
        using MemoryStream stream = CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", PackageRelationshipsXml());
            AddEntry(archive, "word/document.xml", DocumentXml());
            AddEntry(archive, "word/_rels/document.xml.rels", RelationshipsXml("../../escape.xml"));
        });
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E0001" && diagnostic.Severity == DocxSeverity.Error);
    }

    [Fact]
    public static void PublicReadReturnsDiagnosticForMissingInternalRelationshipTarget()
    {
        using MemoryStream stream = CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", PackageRelationshipsXml());
            AddEntry(archive, "word/document.xml", DocumentXml());
            AddEntry(archive, "word/_rels/document.xml.rels", RelationshipsXml("media/missing.png"));
        });
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E0001" && diagnostic.Severity == DocxSeverity.Error);
    }

    [Fact]
    public static void PublicCheckRejectsMacroEnabledDocumentByDefault()
    {
        using MemoryStream stream = CreateMinimalDocx(extra: null, macroEnabled: true);
        var editor = new DocxEditor();

        DocxCheckResult result = editor.Check(stream, new StringReader(HelloPatch));

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error
            && diagnostic.Message.Contains("Macro-enabled", StringComparison.Ordinal));
    }

    [Fact]
    public static void PublicCheckAllowsMacroEnabledDocumentWhenExplicitlyConfigured()
    {
        using MemoryStream stream = CreateMinimalDocx(extra: null, macroEnabled: true);
        var editor = new DocxEditor();
        var options = new DocxEditOptions { AllowMacroEnabledDocuments = true };

        DocxCheckResult result = editor.Check(stream, new StringReader(HelloPatch), options);

        Assert.True(result.Success);
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

    private static MemoryStream CreateMinimalDocxWithImage()
    {
        return CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", PackageRelationshipsXml());
            AddEntry(archive, "word/document.xml", DocumentXmlWithImage());
            AddEntry(archive, "word/_rels/document.xml.rels", RelationshipsXml());
            AddEntry(archive, "word/media/image1.png", "png");
        });
    }

    private static MemoryStream CreateMinimalDocx()
    {
        return CreateMinimalDocx(extra: null, macroEnabled: false);
    }

    private static MemoryStream CreateMinimalDocx(Action<ZipArchive>? extra)
    {
        return CreateMinimalDocx(extra, macroEnabled: false);
    }

    private static MemoryStream CreateMinimalDocx(Action<ZipArchive>? extra, bool macroEnabled)
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

    private static string ContentTypesXml()
    {
        return ContentTypesXml(macroEnabled: false);
    }

    private static string ContentTypesXml(bool macroEnabled)
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

    private static string RelationshipsXml()
    {
        return RelationshipsXml("media/image1.png");
    }

    private static string RelationshipsXml(string imageTarget)
    {
        return $$"""
            <?xml version="1.0" encoding="utf-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{{imageTarget}}"/>
              <Relationship Id="rExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.test/image.png" TargetMode="External"/>
            </Relationships>
            """;
    }

    private static string DocumentXmlWithImage()
    {
        return """
            <?xml version="1.0" encoding="utf-8"?>
            <w:document
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <w:body>
                <w:p>
                  <w:r><w:t>Hello</w:t></w:r>
                  <w:r>
                    <w:drawing>
                      <wp:inline>
                        <a:graphic>
                          <a:graphicData>
                            <pic:pic>
                              <pic:blipFill>
                                <a:blip r:embed="rImage"/>
                              </pic:blipFill>
                            </pic:pic>
                          </a:graphicData>
                        </a:graphic>
                      </wp:inline>
                    </w:drawing>
                  </w:r>
                </w:p>
              </w:body>
            </w:document>
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
