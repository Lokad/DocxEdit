using System.IO.Compression;
using System.Text;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// PLAN B14: part roles resolve through package relationships, never fixed
// paths. A relocated package (same logical document under unusual part paths)
// reads, validates, and edits like the standard package; a part shared by two
// header relationships is scanned once under the first index; unsupported
// story kinds (footnotes) stay out of the read model without breaking it.
public static class PartRoleTests
{
    [Fact]
    public static void RelocatedPackageReadsLikeStandard()
    {
        using MemoryStream standard = CreateStandardDocx();
        using MemoryStream relocated = CreateRelocatedDocx();
        var options = new DocxReadOptions { IncludeHeadersFooters = true };
        DocxReadResult expected = new DocxEditor().Read(standard, options);
        Assert.True(expected.Success);
        standard.Position = 0;
        DocxReadResult actual = new DocxEditor().Read(relocated, options);
        Assert.True(actual.Success);
        Assert.Equal(
            expected.Paragraphs.Select(static p => p.Id.ToWireValue() + "|" + p.Text).ToArray(),
            actual.Paragraphs.Select(static p => p.Id.ToWireValue() + "|" + p.Text).ToArray());
        Assert.Equal(
            expected.Paragraphs.Select(static p => p.StyleName ?? string.Empty).ToArray(),
            actual.Paragraphs.Select(static p => p.StyleName ?? string.Empty).ToArray());
        Assert.Equal(
            expected.Paragraphs.Select(static p => p.List?.LabelText ?? string.Empty).ToArray(),
            actual.Paragraphs.Select(static p => p.List?.LabelText ?? string.Empty).ToArray());
        Assert.Equal(
            expected.Tables.SelectMany(static table => table.Cells).Select(static c => c.Id.ToWireValue() + "|" + c.Text).ToArray(),
            actual.Tables.SelectMany(static table => table.Cells).Select(static c => c.Id.ToWireValue() + "|" + c.Text).ToArray());
    }

    [Fact]
    public static void RelocatedPackageValidatesLikeStandard()
    {
        using MemoryStream standard = CreateStandardDocx();
        using MemoryStream relocated = CreateRelocatedDocx();
        DocxValidateResult expected = new DocxEditor().Validate(standard);
        DocxValidateResult actual = new DocxEditor().Validate(relocated);
        Assert.Equal(expected.Success, actual.Success);
        Assert.Equal(
            expected.Diagnostics.Select(static d => d.Code).Order().ToArray(),
            actual.Diagnostics.Select(static d => d.Code).Order().ToArray());
    }

    [Fact]
    public static void RelocatedStylesWithWrongRootIsReported()
    {
        using MemoryStream input = CreateMislabeledStylesDocx();
        DocxValidateResult result = new DocxEditor().Validate(input);
        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static d => d.Code == "E9102");
        Assert.Contains(result.Diagnostics, static d => (d.PartName ?? string.Empty).Contains("styles", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public static void RelocatedPackageEditsLikeStandard()
    {
        const string patchText = "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Title\nwith Heading\nend\n";
        using MemoryStream standardInput = CreateStandardDocx();
        using var standardOutput = new MemoryStream();
        DocxApplyResult standardApply = new DocxEditor().Apply(standardInput, new StringReader(patchText), standardOutput);
        Assert.True(standardApply.Success);
        using MemoryStream relocatedInput = CreateRelocatedDocx();
        using var relocatedOutput = new MemoryStream();
        DocxApplyResult relocatedApply = new DocxEditor().Apply(relocatedInput, new StringReader(patchText), relocatedOutput);
        Assert.True(relocatedApply.Success);
        Assert.Contains(relocatedApply.Diagnostics, static d => d.Code == "W5103");
        Assert.Contains(standardApply.Diagnostics, static d => d.Code == "W5103");
        relocatedOutput.Position = 0;
        Assert.Equal("Heading", Assert.Single(new DocxEditor().Read(relocatedOutput).Paragraphs, static p => p.Id.ToWireValue() == "M.P0001").Text);
    }

    [Fact]
    public static void SharedHeaderPartIsScannedOnce()
    {
        using MemoryStream input = CreateSharedHeaderDocx();
        DocxReadResult result = new DocxEditor().Read(input, new DocxReadOptions { IncludeHeadersFooters = true });
        Assert.True(result.Success);
        string[] headerIds = result.Paragraphs
            .Select(static p => p.Id.ToWireValue())
            .Where(static id => id.StartsWith("H", StringComparison.Ordinal))
            .ToArray();
        Assert.Equal(["H001.P0001"], headerIds);
    }

    [Fact]
    public static void UnsupportedFootnotePartDoesNotBreakRead()
    {
        using MemoryStream input = CreateFootnoteDocx();
        DocxReadResult read = new DocxEditor().Read(input);
        Assert.True(read.Success);
        Assert.DoesNotContain(read.Paragraphs, static p => p.Story.Contains("footnote", StringComparison.OrdinalIgnoreCase));
        DocxValidateResult validate = new DocxEditor().Validate(input);
        Assert.True(validate.Success);
    }

    private static MemoryStream CreateStandardDocx()
    {
        return CreateDocx("/word/document.xml", "/word/styles.xml", "/word/numbering.xml", "/word/header1.xml", false);
    }

    private static MemoryStream CreateRelocatedDocx()
    {
        return CreateDocx("/doc/document.xml", "/doc/styles.xml", "/doc/numbering.xml", "/doc/header1.xml", false);
    }

    private static MemoryStream CreateMislabeledStylesDocx()
    {
        return CreateDocx("/doc/document.xml", "/doc/styles.xml", "/doc/numbering.xml", "/doc/header1.xml", true);
    }

    private static MemoryStream CreateDocx(string documentPath, string stylesPath, string numberingPath, string headerPath, bool mislabeledStyles)
    {
        string documentDir = documentPath.Contains('/') ? documentPath[..documentPath.LastIndexOf('/')] : string.Empty;
        string relsPath = documentDir + "/_rels/" + documentPath[(documentDir.Length + 1)..] + ".rels";
        string stylesFile = stylesPath[(stylesPath.LastIndexOf('/') + 1)..];
        string numberingFile = numberingPath[(numberingPath.LastIndexOf('/') + 1)..];
        string headerFile = headerPath[(headerPath.LastIndexOf('/') + 1)..];
        string stylesTarget = stylesFile;
        string numberingTarget = numberingFile;
        string headerTarget = headerFile;
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml(documentPath, stylesPath, numberingPath, headerPath));
            AddEntry(archive, "_rels/.rels", RootRelsXml(documentPath));
            AddEntry(archive, relsPath.TrimStart('/'), DocumentRelsXml(stylesTarget, numberingTarget, headerTarget));
            AddEntry(archive, documentPath.TrimStart('/'), DocumentXml());
            AddEntry(archive, stylesPath.TrimStart('/'), mislabeledStyles
                ? "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>"
                : StylesXml());
            AddEntry(archive, numberingPath.TrimStart('/'), NumberingXml());
            AddEntry(archive, headerPath.TrimStart('/'), HeaderXml());
        }
        stream.Position = 0;
        return stream;
    }

    private static string ContentTypesXml(string documentPath, string stylesPath, string numberingPath, string headerPath)
    {
        return "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Override PartName=\"" + documentPath + "\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
            + "<Override PartName=\"" + stylesPath + "\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>"
            + "<Override PartName=\"" + numberingPath + "\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml\"/>"
            + "<Override PartName=\"" + headerPath + "\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>"
            + "</Types>";
    }

    private static string RootRelsXml(string documentPath)
    {
        return "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rDocument\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"" + documentPath.TrimStart('/') + "\"/>"
            + "</Relationships>";
    }

    private static string DocumentRelsXml(string stylesTarget, string numberingTarget, string headerTarget)
    {
        return "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rHeader\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"" + headerTarget + "\"/>"
            + "<Relationship Id=\"rNumbering\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering\" Target=\"" + numberingTarget + "\"/>"
            + "<Relationship Id=\"rStyles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"" + stylesTarget + "\"/>"
            + "</Relationships>";
    }

    private static string DocumentXml()
    {
        return "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + "<w:body>"
            + "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"9\"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>"
            + "<w:p><w:fldSimple w:instr=\" REF ClientName \\h \"><w:r><w:t>Acme</w:t></w:r></w:fldSimple></w:p>"
            + "<w:tbl><w:tblPr><w:tblStyle w:val=\"TableGrid\"/></w:tblPr><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
            + "</w:body>"
            + "</w:document>";
    }

    private static string StylesXml()
    {
        return "<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + "<w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"Heading 1\"/></w:style>"
            + "</w:styles>";
    }

    private static string NumberingXml()
    {
        return "<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + "<w:abstractNum w:abstractNumId=\"7\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/></w:lvl></w:abstractNum>"
            + "<w:num w:numId=\"9\"><w:abstractNumId w:val=\"7\"/></w:num>"
            + "</w:numbering>";
    }

    private static string HeaderXml()
    {
        return "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + "<w:p><w:r><w:t>Head</w:t></w:r></w:p>"
            + "</w:hdr>";
    }

    private static MemoryStream CreateSharedHeaderDocx()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "<Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>"
                + "</Types>");
            AddEntry(archive, "_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rDocument\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + "</Relationships>");
            AddEntry(archive, "word/_rels/document.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rHeaderEven\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>"
                + "<Relationship Id=\"rHeaderOdd\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>"
                + "</Relationships>");
            AddEntry(archive, "word/document.xml", "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>"
                + "<w:p><w:r><w:t>Body</w:t></w:r></w:p>"
                + "</w:body></w:document>");
            AddEntry(archive, "word/header1.xml", HeaderXml());
        }
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateFootnoteDocx()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "<Override PartName=\"/word/footnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml\"/>"
                + "</Types>");
            AddEntry(archive, "_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rDocument\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + "</Relationships>");
            AddEntry(archive, "word/_rels/document.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rFootnotes\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes\" Target=\"footnotes.xml\"/>"
                + "</Relationships>");
            AddEntry(archive, "word/document.xml", "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>"
                + "<w:p><w:r><w:t>Body</w:t></w:r></w:p>"
                + "</w:body></w:document>");
            AddEntry(archive, "word/footnotes.xml", "<w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                + "<w:footnote w:id=\"1\"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:footnote>"
                + "</w:footnotes>");
        }
        stream.Position = 0;
        return stream;
    }

}
