using System.IO.Compression;
using System.Text;

namespace Lokad.DocxEdit.Tests;

public static class XmlValueTests
{
    [Fact]
    public static void PublicDumpPreservesRawTextAndEscapesRuns()
    {
        using MemoryStream stream = CreateDocxWithText("a\\b\"c&#13;d\ne\tf");
        var editor = new DocxEditor();

        DocxDumpResult dump = editor.Dump(stream, "M.P0001");
        stream.Position = 0;
        DocxDumpResult runs = editor.Dump(stream, "M.P0001", new DocxDumpOptions { IncludeRuns = true });
        stream.Position = 0;
        DocxReadResult read = editor.Read(stream);

        Assert.Equal("a\\b\"c\rd\ne\tf", dump.Text);
        Assert.Contains("text=\"a\\\\b\\\"c\\rd\\ne\\tf\"", runs.Text, StringComparison.Ordinal);
        Assert.Equal("a\\b\"c\rd\ne\tf", read.Paragraphs[0].Text);
    }

    [Fact]
    public static void PublicRenderStylesEscapesQuotingInNames()
    {
        using MemoryStream stream = CreateDocxWithStyleName("a\\b&quot;c");
        var editor = new DocxEditor();

        DocxStylesResult result = editor.Styles(stream);

        Assert.True(result.Success);
        Assert.Contains("a\\\\b\\\"c", DocxTextRenderer.RenderStyles(result), StringComparison.Ordinal);
    }

    [Fact]
    public static void PublicStyleRenderingEscapesCarriageReturns()
    {
        var result = new DocxStylesResult
        {
            Success = true,
            Diagnostics = [],
            Styles = [new DocxStyleInfo("S1", "Name\rb", "paragraph", false)],
        };

        string output = DocxTextRenderer.RenderStyles(result);

        Assert.Contains("Name\\rb", output, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("12", 12)]
    [InlineData("abc", 1)]
    [InlineData("-3", 1)]
    [InlineData(" 7 ", 7)]
    [InlineData("12.5", 1)]
    [InlineData("", 1)]
    [InlineData(null, 1)]
    public static void PublicSectionColumnsFallBackForInvalidIntegers(string? columns, int expectedColumns)
    {
        using MemoryStream stream = CreateDocxWithColumns(columns);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.True(result.Success);
        Assert.Equal(expectedColumns, Assert.Single(result.Sections).Columns);
    }

    private static MemoryStream CreateDocxWithText(string text)
    {
        return CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", PackageRelationshipsXml());
            AddEntry(archive, "word/document.xml", DocumentXmlWithText(text));
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <?xml version="1.0" encoding="utf-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
                """);
        });
    }

    private static MemoryStream CreateDocxWithColumns(string? columns)
    {
        return CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", PackageRelationshipsXml());
            AddEntry(archive, "word/document.xml", DocumentXmlWithColumns(columns));
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <?xml version="1.0" encoding="utf-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
                """);
        });
    }

    private static MemoryStream CreateDocxWithStyleName(string encodedName)
    {
        return CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", PackageRelationshipsXml());
            AddEntry(archive, "word/document.xml", DocumentXmlWithText("Hello"));
            AddEntry(archive, "word/_rels/document.xml.rels", StylesRelationshipsXml());
            AddEntry(archive, "word/styles.xml", StylesXmlWithName(encodedName));
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
        return """
            <?xml version="1.0" encoding="utf-8"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
              <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
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

    private static string StylesRelationshipsXml()
    {
        return """
            <?xml version="1.0" encoding="utf-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
            </Relationships>
            """;
    }

    private static string DocumentXmlWithText(string text)
    {
        return $$"""
            <?xml version="1.0" encoding="utf-8"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p><w:r><w:t>{{text}}</w:t></w:r></w:p>
              </w:body>
            </w:document>
            """;
    }

    private static string DocumentXmlWithColumns(string? columns)
    {
        string columnsElement = columns is null ? "" : $"<w:cols w:num=\"{columns}\"/>";
        return $$"""
            <?xml version="1.0" encoding="utf-8"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
                <w:sectPr>{{columnsElement}}</w:sectPr>
              </w:body>
            </w:document>
            """;
    }

    private static string StylesXmlWithName(string encodedName)
    {
        return $$"""
            <?xml version="1.0" encoding="utf-8"?>
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:style w:type="paragraph" w:styleId="S1"><w:name w:val="{{encodedName}}"/></w:style>
            </w:styles>
            """;
    }
}
