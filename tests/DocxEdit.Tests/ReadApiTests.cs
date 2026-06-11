using System.IO.Compression;
using System.Text;

namespace DocxEdit.Tests;

public static class ReadApiTests
{
    [Fact]
    public static void ReadExtractsParagraphsTablesAndImages()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.True(result.Success);
        Assert.Equal("/word/document.xml", result.MainDocumentPartName);
        Assert.Equal(2, result.Paragraphs.Count);
        Assert.Equal("M.P0001", result.Paragraphs[0].Id);
        Assert.Equal(1, result.Paragraphs[0].HeadingLevel);
        Assert.Equal("Executive Summary", result.Paragraphs[0].Text);
        Assert.Single(result.Tables);
        Assert.Equal("North", result.Tables[0].Cells[0].Text);
        Assert.Single(result.Images);
        Assert.Equal("/word/media/image1.png", result.Images[0].PartName);
        Assert.Contains("M.P0001 heading level=1", result.Text, StringComparison.Ordinal);
        Assert.Contains("M.T0001 table rows=1 columns=2", result.Text, StringComparison.Ordinal);
        Assert.Contains("M.I0001 image part=/word/media/image1.png", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void FindMatchesParagraphAndTableCellText()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxFindResult result = editor.Find(stream, "Revenue");

        Assert.True(result.Success);
        Assert.Contains(result.Matches, match => match.StartsWith("M.P0002", StringComparison.Ordinal));
        Assert.Contains(result.Matches, match => match.StartsWith("M.T0001.R01.C02", StringComparison.Ordinal));
    }

    [Fact]
    public static void DumpReturnsParagraphOrCellText()
    {
        using MemoryStream paragraphStream = CreateDocx();
        using MemoryStream cellStream = CreateDocx();
        var editor = new DocxEditor();

        DocxDumpResult paragraph = editor.Dump(paragraphStream, "M.P0002");
        DocxDumpResult cell = editor.Dump(cellStream, "M.T0001.R01.C02");

        Assert.True(paragraph.Success);
        Assert.Equal("Revenue increased", paragraph.Text);
        Assert.True(cell.Success);
        Assert.Equal("Revenue", cell.Text);
    }

    [Fact]
    public static void ReadHonorsMaxTextInRenderedOutput()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream, new DocxReadOptions { MaxText = 10 });

        Assert.True(result.Success);
        Assert.Contains("Revenue...", result.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Revenue increased", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void DumpCanIncludeParagraphRuns()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxDumpResult result = editor.Dump(stream, "M.P0002", new DocxDumpOptions { IncludeRuns = true });

        Assert.True(result.Success);
        Assert.NotNull(result.Text);
        Assert.Contains("runs:", result.Text, StringComparison.Ordinal);
        Assert.Contains("M.P0002.R0001 text=\"Revenue\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("M.P0002.R0002 text=\" increased\"", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void MediaListsReferencedImagesOnly()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxMediaResult result = editor.Media(stream);

        DocxImageInfo image = Assert.Single(result.Images);
        Assert.Equal("M.I0001", image.Id);
        Assert.Equal("/word/media/image1.png", image.PartName);
    }

    [Fact]
    public static void StylesListsParagraphCharacterAndTableStyles()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxStylesResult result = editor.Styles(stream);

        Assert.True(result.Success);
        Assert.Equal(3, result.Styles.Count);
        Assert.Contains(result.Styles, style => style.StyleId == "Normal" && style.Type == "paragraph" && style.IsDefault);
        Assert.Contains(result.Styles, style => style.StyleId == "Emphasis" && style.Type == "character");
        Assert.Contains(result.Styles, style => style.StyleId == "TableGrid" && style.Type == "table");
    }

    [Fact]
    public static void ReadIncludesHeaderAndFooterStoriesWhenRequested()
    {
        using MemoryStream defaultStream = CreateDocx();
        using MemoryStream allStoriesStream = CreateDocx();
        var editor = new DocxEditor();

        DocxReadResult defaultResult = editor.Read(defaultStream);
        DocxReadResult allStories = editor.Read(allStoriesStream, new DocxReadOptions { IncludeHeadersFooters = true });

        Assert.DoesNotContain(defaultResult.Paragraphs, paragraph => paragraph.Id.StartsWith("H", StringComparison.Ordinal));
        Assert.Contains(allStories.Paragraphs, paragraph => paragraph.Id == "H001.P0001" && paragraph.Story == "header[1]" && paragraph.Text == "Confidential");
        Assert.Contains(allStories.Paragraphs, paragraph => paragraph.Id == "F001.P0001" && paragraph.Story == "footer[1]" && paragraph.Text == "Page 1");
    }

    [Fact]
    public static void FindCanSearchHeaderAndFooterStoriesWhenRequested()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxFindResult result = editor.Find(stream, "Confidential", new DocxFindOptions { IncludeHeadersFooters = true });

        Assert.True(result.Success);
        Assert.Contains(result.Matches, match => match.StartsWith("H001.P0001", StringComparison.Ordinal));
    }

    private static MemoryStream CreateDocx()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
                  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
                </Types>
                """);
            AddEntry(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
                  <Relationship Id="rStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                  <Relationship Id="rHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                  <Relationship Id="rFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/document.xml", """
                <w:document
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <w:body>
                    <w:p>
                      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
                      <w:r><w:t>Executive Summary</w:t></w:r>
                    </w:p>
                    <w:p>
                      <w:r><w:t>Revenue</w:t></w:r>
                      <w:r><w:t xml:space="preserve"> increased</w:t></w:r>
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
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                  </w:body>
                </w:document>
                """);
            AddEntry(archive, "word/styles.xml", """
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
                    <w:name w:val="Normal"/>
                  </w:style>
                  <w:style w:type="character" w:styleId="Emphasis">
                    <w:name w:val="Emphasis"/>
                  </w:style>
                  <w:style w:type="table" w:styleId="TableGrid">
                    <w:name w:val="Table Grid"/>
                  </w:style>
                  <w:style w:type="numbering" w:styleId="ListNumber">
                    <w:name w:val="List Number"/>
                  </w:style>
                </w:styles>
                """);
            AddEntry(archive, "word/header1.xml", """
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:p><w:r><w:t>Confidential</w:t></w:r></w:p>
                </w:hdr>
                """);
            AddEntry(archive, "word/footer1.xml", """
                <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:p><w:r><w:t>Page 1</w:t></w:r></w:p>
                </w:ftr>
                """);
            AddEntry(archive, "word/media/image1.png", "fake-png");
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
}
