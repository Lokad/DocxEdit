using System.IO.Compression;
using System.Text;
using System.Text.Json;

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
    public static void ReadExtractsMergedAndNestedTableMetadata()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/></w:tcPr>
                          <w:p><w:r><w:t>Wide</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc><w:p><w:r><w:t>East</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:vMerge/></w:tcPr>
                          <w:p><w:r><w:t>Continued</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                          <w:p><w:r><w:t>Outer</w:t></w:r></w:p>
                          <w:tbl>
                            <w:tr><w:tc><w:p><w:r><w:t>Inner</w:t></w:r></w:p></w:tc></w:tr>
                          </w:tbl>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        DocxTableInfo table = Assert.Single(result.Tables);
        Assert.Equal(3, table.ColumnCount);
        DocxTableCellInfo wide = table.Cells.Single(cell => cell.Id == "M.T0001.R01.C01");
        Assert.Equal(2, wide.ColumnSpan);
        Assert.Equal("restart", wide.VerticalMerge);
        Assert.False(wide.HasNestedTable);
        DocxTableCellInfo east = table.Cells.Single(cell => cell.Id == "M.T0001.R01.C03");
        Assert.Equal("East", east.Text);
        DocxTableCellInfo continued = table.Cells.Single(cell => cell.Id == "M.T0001.R02.C01");
        Assert.Equal("continue", continued.VerticalMerge);
        DocxTableCellInfo nested = table.Cells.Single(cell => cell.Id == "M.T0001.R02.C02");
        Assert.True(nested.HasNestedTable);
        Assert.Contains("M.T0001.R01.C01 column-span=2 vertical-merge=restart", result.Text, StringComparison.Ordinal);
        Assert.Contains("M.T0001.R02.C02 nested-table=true", result.Text, StringComparison.Ordinal);
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
    public static void ReadExtractsTabsAndLineBreaks()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:t>A</w:t><w:tab/><w:t>B</w:t><w:br/><w:t>C</w:t></w:r>
                    </w:p>
            """);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.Equal("A\tB\nC", Assert.Single(result.Paragraphs).Text);
        Assert.Contains("A\\tB\\nC", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void ReadExtractsParagraphListMetadata()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p>
                      <w:pPr>
                        <w:numPr>
                          <w:ilvl w:val="1"/>
                          <w:numId w:val="42"/>
                        </w:numPr>
                      </w:pPr>
                      <w:r><w:t>List item</w:t></w:r>
                    </w:p>
            """);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        DocxParagraphInfo paragraph = Assert.Single(result.Paragraphs);
        Assert.NotNull(paragraph.List);
        Assert.Equal("42", paragraph.List.NumberingId);
        Assert.Equal(1, paragraph.List.Level);
        Assert.Contains("M.P0001 paragraph list numId=42 level=1", result.Text, StringComparison.Ordinal);
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
    public static void ReadExtractsSectionColumnsAndOrientation()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        DocxSectionInfo section = Assert.Single(result.Sections);
        Assert.Equal("M.S0001", section.Id);
        Assert.Equal(2, section.Columns);
        Assert.Equal("landscape", section.Orientation);
        Assert.Contains("M.S0001 section columns=2 orientation=landscape", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void OutlineIncludesSections()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxOutlineResult result = editor.Outline(stream);

        Assert.Contains(result.Lines, line => line == "M.S0001 section columns=2 orientation=landscape");
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

    [Fact]
    public static void ChangesListsTrackedMarkupWithoutRevisionText()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:t>Before</w:t></w:r>
                      <w:ins w:id="9" w:author="Alice" w:date="2026-06-01T12:00:00Z">
                        <w:r><w:t>Inserted</w:t></w:r>
                      </w:ins>
                      <w:del w:id="10" w:author="Bob" w:date="2026-06-02T12:00:00Z">
                        <w:r><w:delText>Deleted</w:delText></w:r>
                      </w:del>
                      <w:r>
                        <w:rPr>
                          <w:rPrChange w:id="11" w:author="Carol" w:date="2026-06-03T12:00:00Z"/>
                        </w:rPr>
                        <w:t>After</w:t>
                      </w:r>
                    </w:p>
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:tcPrChange w:id="12" w:author="Dan" w:date="2026-06-04T12:00:00Z"/></w:tcPr>
                          <w:p><w:r><w:t>Cell</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr>
                      <w:sectPrChange w:id="13" w:author="Eve" w:date="2026-06-05T12:00:00Z"/>
                    </w:sectPr>
                    <w:customXmlDelRangeStart w:id="14" w:author="Frank" w:date="2026-06-06T12:00:00Z"/>
                    <w:customXmlDelRangeEnd w:id="14"/>
            """);
        var editor = new DocxEditor();

        DocxChangesResult result = editor.Changes(stream);

        Assert.True(result.Success);
        Assert.Equal(7, result.Changes.Count);
        Assert.Contains(result.Summary, summary => summary.Type == "inserted-run" && summary.Count == 1);
        Assert.Contains(result.Summary, summary => summary.Type == "deleted-run" && summary.Count == 1);
        Assert.Contains(result.Summary, summary => summary.Type == "run-properties-change" && summary.Count == 1);
        Assert.Contains(result.Summary, summary => summary.Type == "cell-properties-change" && summary.Count == 1);
        Assert.Contains(result.Summary, summary => summary.Type == "section-properties-change" && summary.Count == 1);
        Assert.Contains(result.Summary, summary => summary.Type == "custom-xml-delete-range-start" && summary.Count == 1);
        Assert.Contains(result.Summary, summary => summary.Type == "custom-xml-delete-range-end" && summary.Count == 1);

        DocxChangeInfo insertion = Assert.Single(result.Changes, change => change.Type == "inserted-run");
        Assert.Equal("M.CH0001", insertion.Id);
        Assert.Equal("main", insertion.Story);
        Assert.Equal("/word/document.xml", insertion.PartName);
        Assert.Equal("M.P0001", insertion.TargetId);
        Assert.Equal("Alice", insertion.Author);
        Assert.Equal("9", insertion.RevisionId);
        Assert.Equal(8, insertion.TextLength);
        Assert.Equal(DateTimeOffset.Parse("2026-06-01T12:00:00Z").ToUniversalTime(), insertion.TimestampUtc);

        DocxChangeInfo cellChange = Assert.Single(result.Changes, change => change.Type == "cell-properties-change");
        Assert.Equal("M.T0001.R01.C01", cellChange.TargetId);

        string serialized = JsonSerializer.Serialize(result);
        Assert.DoesNotContain("Inserted", serialized, StringComparison.Ordinal);
        Assert.DoesNotContain("Deleted", serialized, StringComparison.Ordinal);
    }

    [Fact]
    public static void ReadWarnsWhenTrackedChangeMarkupIsOnlyPartiallyModeled()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p>
                      <w:ins w:id="1" w:author="A">
                        <w:r><w:t>PrivateInserted</w:t></w:r>
                      </w:ins>
                    </w:p>
            """);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.True(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1001" && diagnostic.Feature == "tracked-changes");
        string serializedDiagnostics = JsonSerializer.Serialize(result.Diagnostics);
        Assert.DoesNotContain("PrivateInserted", serializedDiagnostics, StringComparison.Ordinal);
    }

    [Fact]
    public static void MediaWarnsAboutFloatingAndExternalImages()
    {
        using MemoryStream stream = CreateDocxWithBody(
            """
                    <w:p
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                      <w:r>
                        <w:drawing>
                          <wp:anchor>
                            <a:graphic>
                              <a:graphicData>
                                <a:blip r:embed="rExternalImage"/>
                              </a:graphicData>
                            </a:graphic>
                          </wp:anchor>
                        </w:drawing>
                      </w:r>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rExternalImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.test/image.png" TargetMode="External"/>
                </Relationships>
                """);
        var editor = new DocxEditor();

        DocxMediaResult result = editor.Media(stream);

        Assert.True(result.Success);
        Assert.Empty(result.Images);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1007" && diagnostic.Feature == "floating-image");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1008" && diagnostic.Feature == "external-image");
    }

    [Fact]
    public static void ReadWarnsAboutUnsupportedPreservedObjects()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                        xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                      <w:r>
                        <w:drawing>
                          <c:chart r:id="rChart"/>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:altChunk xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rAltChunk"/>
            """);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.True(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1009" && diagnostic.Feature == "chart");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1013" && diagnostic.Feature == "alt-chunk");
    }

    [Fact]
    public static void ReadWarnsAboutComplexSectionFlow()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p>
                      <w:pPr>
                        <w:sectPr>
                          <w:cols w:num="2"/>
                        </w:sectPr>
                      </w:pPr>
                      <w:r><w:t>First section</w:t></w:r>
                    </w:p>
                    <w:p><w:r><w:t>Second section</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:cols w:num="1"/>
                    </w:sectPr>
            """);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.True(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "W1014" &&
            diagnostic.Feature == "section-flow" &&
            diagnostic.PartName == "/word/document.xml" &&
            diagnostic.Story == "main");
    }

    [Fact]
    public static void ReadUnsupportedFeatureDiagnosticsCarryStableMetadata()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                        xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                      <w:bookmarkStart w:id="1" w:name="Bookmark"/>
                      <w:hyperlink r:id="rHyperlink"><w:r><w:t>Link</w:t></w:r></w:hyperlink>
                      <w:fldSimple w:instr="DATE"><w:r><w:t>Date</w:t></w:r></w:fldSimple>
                      <w:commentRangeStart w:id="1"/>
                      <w:sdt><w:sdtContent><w:r><w:t>Control</w:t></w:r></w:sdtContent></w:sdt>
                      <w:ins w:id="2" w:author="A"><w:r><w:t>Revision</w:t></w:r></w:ins>
                      <w:r>
                        <w:drawing>
                          <wp:anchor>
                            <a:graphic>
                              <a:graphicData>
                                <c:chart r:id="rChart"/>
                              </a:graphicData>
                            </a:graphic>
                          </wp:anchor>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:altChunk xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rAltChunk"/>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rExternalImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.test/image.png" TargetMode="External"/>
                </Relationships>
                """);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        DocxDiagnostic[] warnings = result.Diagnostics
            .Where(diagnostic => diagnostic.Code.StartsWith("W10", StringComparison.Ordinal))
            .ToArray();
        Assert.NotEmpty(warnings);
        Assert.All(warnings, diagnostic =>
        {
            Assert.Equal(DocxSeverity.Warning, diagnostic.Severity);
            Assert.Matches("^W10[0-9]{2}$", diagnostic.Code);
            Assert.Equal("/word/document.xml", diagnostic.PartName);
            Assert.Equal("main", diagnostic.Story);
            Assert.False(string.IsNullOrWhiteSpace(diagnostic.Feature));
            Assert.False(string.IsNullOrWhiteSpace(diagnostic.Fallback));
        });
        Assert.Contains(warnings, diagnostic => diagnostic.Code == "W1001" && diagnostic.Fallback == "final-view");
        Assert.Contains(warnings, diagnostic => diagnostic.Code == "W1008" && diagnostic.Fallback == "omit-from-editable-images");
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
                    <w:sectPr>
                      <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
                      <w:cols w:num="2"/>
                    </w:sectPr>
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

    private static MemoryStream CreateDocxWithBody(string bodyXml, string? documentRelationshipsXml = null)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """);
            AddEntry(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/_rels/document.xml.rels", documentRelationshipsXml ?? """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
                """);
            AddEntry(archive, "word/document.xml", """
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>

                """ + bodyXml + """

                  </w:body>
                </w:document>
                """);
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
