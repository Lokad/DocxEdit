using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml.Linq;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

public static class ReadInspectTests
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
        Assert.Equal("M.P0001", result.Paragraphs[0].Id.ToWireValue());
        Assert.Equal(1, result.Paragraphs[0].HeadingLevel);
        Assert.Equal("Executive Summary", result.Paragraphs[0].Text);
        Assert.Single(result.Tables);
        Assert.Equal("North", result.Tables[0].Cells[0].Text);
        Assert.Single(result.Images);
        Assert.Equal("/word/media/image1.png", result.Images[0].PartName);
        Assert.Contains("M.P0001 heading level=1", result.Text, StringComparison.Ordinal);
        Assert.Contains("M.T0001 table rows=1 columns=2", result.Text, StringComparison.Ordinal);
        Assert.Contains("M.I0001 image layout=inline part=/word/media/image1.png", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void FindBoundsMatchTextByMaxText()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxFindResult full = editor.Find(stream, "Revenue");
        stream.Position = 0;
        DocxFindResult bounded = editor.Find(stream, "Revenue", new DocxFindOptions { MaxText = 0 });

        Assert.NotEmpty(full.Matches);
        Assert.Equal(full.Matches.Count, bounded.Matches.Count);
        Assert.All(full.Matches, match => Assert.False(string.IsNullOrEmpty(match.Text)));
        Assert.All(bounded.Matches, match => Assert.Equal(string.Empty, match.Text));
    }

    [Fact]
    public static void OutlineBoundsHeadingTextByMaxText()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxOutlineResult full = editor.Outline(stream);
        stream.Position = 0;
        DocxOutlineResult bounded = editor.Outline(stream, new DocxOutlineOptions { MaxText = 0 });

        Assert.True(full.Success);
        Assert.True(bounded.Success);
        Assert.NotEmpty(full.Items);
        Assert.Equal(full.Items.Count, bounded.Items.Count);
        Assert.All(bounded.Items, item => Assert.Equal(string.Empty, item.Text));
    }

    [Fact]
    public static void OutlineAndFindRenderersPreserveLineFormat()
    {
        var outline = new DocxOutlineResult
        {
            Success = true,
            Diagnostics = [],
            Items =
            [
                new DocxOutlineItem { TargetId = "M.P0001", Kind = "heading", Text = "Top", HeadingLevel = 1 },
                new DocxOutlineItem { TargetId = "M.S0001", Kind = "section", Columns = 2, Orientation = "landscape" },
                new DocxOutlineItem { TargetId = "M.B0001", Kind = "bookmark", Name = "Intro", StartTargetId = "M.P0001", EndTargetId = "M.P0002" },
                new DocxOutlineItem { TargetId = "M.H001", Kind = "hyperlink", HyperlinkTarget = "M.P0001", Destination = "https://example.com", IsBroken = false },
            ]
        };

        string outlineText = DocxTextRenderer.RenderOutline(outline);
        Assert.Contains("M.P0001 heading level=1 text=\"Top\"", outlineText, StringComparison.Ordinal);
        Assert.Contains("M.S0001 section columns=2 orientation=landscape", outlineText, StringComparison.Ordinal);
        Assert.Contains("M.B0001 bookmark name=\"Intro\" start=M.P0001 end=M.P0002", outlineText, StringComparison.Ordinal);
        Assert.Contains("M.H001 hyperlink target=M.P0001 destination=\"https://example.com\" broken=False", outlineText, StringComparison.Ordinal);

        var found = new DocxFindResult
        {
            Success = true,
            Diagnostics = [],
            Query = "Revenue",
            Matches =
            [
                new DocxFindMatch { TargetId = "M.P0002", Kind = "paragraph", Text = "Revenue increased." },
                new DocxFindMatch { TargetId = "M.T0001.R01.C02", Kind = "cell", ParentId = "M.T0001", Text = "Revenue" },
            ]
        };

        string findText = DocxTextRenderer.RenderFind(found);
        Assert.Contains("M.P0002 text=\"Revenue increased.\"", findText, StringComparison.Ordinal);
        Assert.Contains("M.T0001.R01.C02 text=\"Revenue\"", findText, StringComparison.Ordinal);
    }

    [Fact]
    public static void FindMatchesParagraphAndTableCellText()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxFindResult result = editor.Find(stream, "Revenue");

        Assert.True(result.Success);
        Assert.Contains(result.Matches, match => match.TargetId == "M.P0002" && match.Kind == "paragraph");
        Assert.Contains(result.Matches, match => match.TargetId == "M.T0001.R01.C02" && match.Kind == "cell" && match.ParentId == "M.T0001");
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
    public static void DumpShowsSectionLine()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p><w:r><w:t>Main text</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:pgSz w:w="12240" w:h="15840"/>
                      <w:cols w:num="1"/>
                    </w:sectPr>
            """);
        var editor = new DocxEditor();
        DocxDumpResult result = editor.Dump(stream, "M.S0001");
        Assert.True(result.Success);
        Assert.Contains("M.S0001 section columns=1 orientation=portrait", result.Text, StringComparison.Ordinal);
    }
    [Fact]
    public static void DumpShowsImageLine()
    {
        using MemoryStream stream = CreateDocxWithImage("png", "image/png", "old-png");
        var editor = new DocxEditor();
        DocxDumpResult result = editor.Dump(stream, "M.I0001");
        Assert.True(result.Success);
        Assert.Contains("M.I0001 image", result.Text, StringComparison.Ordinal);
        Assert.Contains("content-type=image/png", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void DumpShowsMergeGroupMembers()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
                          <w:p><w:r><w:t>North</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:vMerge/></w:tcPr>
                          <w:p><w:r><w:t>South</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc><w:p><w:r><w:t>Profit</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        var editor = new DocxEditor();
        DocxDumpResult result = editor.Dump(stream, "M.T0001.MG0001");
        Assert.True(result.Success);
        Assert.Contains("North", result.Text, StringComparison.Ordinal);
        Assert.Contains("South", result.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Revenue", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void DumpShowsRowMembers()
    {
        using MemoryStream stream = CreateDocxWithSimpleTwoByTwoTable();
        var editor = new DocxEditor();
        DocxDumpResult result = editor.Dump(stream, "M.T0001.R01");
        Assert.True(result.Success);
        Assert.Contains("North", result.Text, StringComparison.Ordinal);
        Assert.Contains("Revenue", result.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("South", result.Text, StringComparison.Ordinal);
    }
    [Fact]
    public static void DumpShowsHyperlinkLine()
    {
        using MemoryStream stream = CreateDocxWithBody(
            """
                    <w:p xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                      <w:hyperlink r:id="rLink" w:tooltip="Open example">
                        <w:r><w:t>External</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/report" TargetMode="External"/>
                </Relationships>
                """, null);
        var editor = new DocxEditor();
        DocxDumpResult result = editor.Dump(stream, "M.L0001");
        Assert.True(result.Success);
        Assert.Contains("M.L0001 hyperlink", result.Text, StringComparison.Ordinal);
        Assert.Contains("https://example.test/report", result.Text, StringComparison.Ordinal);
        Assert.Contains("broken=False", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void DumpShowsFieldLine()
    {
        using MemoryStream stream = CreateDocxWithRefField();
        var editor = new DocxEditor();
        DocxDumpResult result = editor.Dump(stream, "M.F0001");
        Assert.True(result.Success);
        Assert.Contains("M.F0001 field", result.Text, StringComparison.Ordinal);
        Assert.Contains("ClientName", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void DumpShowsContentControlLine()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:id w:val="77"/>
                          <w:alias w:val="Client Name"/>
                          <w:tag w:val="client_name"/>
                          <w:text/>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>Acme</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        var editor = new DocxEditor();
        DocxDumpResult result = editor.Dump(stream, "M.CC0001");
        Assert.True(result.Success);
        Assert.Contains("M.CC0001 content-control kind=plain-text", result.Text, StringComparison.Ordinal);
        Assert.Contains("client_name", result.Text, StringComparison.Ordinal);
    }


    [Fact]
    public static void DumpReturnsBookmarkLine()
    {
        using MemoryStream stream = CreateDocxWithSingleBookmark();
        var editor = new DocxEditor();
        DocxDumpResult bookmark = editor.Dump(stream, "M.B0001");
        Assert.True(bookmark.Success);
        Assert.Contains("M.B0001 bookmark", bookmark.Text, StringComparison.Ordinal);
        Assert.Contains("name=", bookmark.Text, StringComparison.Ordinal);
        Assert.Contains("start=M.P0001", bookmark.Text, StringComparison.Ordinal);
        Assert.Contains("complete=True", bookmark.Text, StringComparison.Ordinal);
    }
    [Fact]
    public static void DumpUnknownBookmarkFails()
    {
        using MemoryStream stream = CreateDocxWithSingleBookmark();
        var editor = new DocxEditor();
        DocxDumpResult missing = editor.Dump(stream, "M.B0009");
        Assert.False(missing.Success);
        Assert.Contains(missing.Diagnostics, static diagnostic => diagnostic.Code == "E1201");
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
    public static void DumpExposesTargetLevelPropertyRevisionSummaries()
    {
        const string bodyXml = """
                    <w:p>
                      <w:pPr>
                        <w:pPrChange w:id="1" w:author="Reviewer" w:date="2026-06-01T00:00:00Z">
                          <w:pPr/>
                        </w:pPrChange>
                      </w:pPr>
                      <w:r><w:t>Private paragraph text</w:t></w:r>
                    </w:p>
                    <w:tbl>
                      <w:tblPr>
                        <w:tblPrChange w:id="2" w:author="Reviewer" w:date="2026-06-02T00:00:00Z">
                          <w:tblPr/>
                        </w:tblPrChange>
                      </w:tblPr>
                      <w:tr>
                        <w:trPr>
                          <w:trPrChange w:id="3" w:author="Reviewer" w:date="2026-06-03T00:00:00Z">
                            <w:trPr/>
                          </w:trPrChange>
                        </w:trPr>
                        <w:tc><w:p><w:r><w:t>Private cell text</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr>
                      <w:sectPrChange w:id="4" w:author="Reviewer" w:date="2026-06-04T00:00:00Z">
                        <w:sectPr/>
                      </w:sectPrChange>
                    </w:sectPr>
            """;
        var editor = new DocxEditor();
        using MemoryStream paragraphStream = CreateDocxWithBody(bodyXml);
        using MemoryStream tableStream = CreateDocxWithBody(bodyXml);
        using MemoryStream rowStream = CreateDocxWithBody(bodyXml);
        using MemoryStream sectionStream = CreateDocxWithBody(bodyXml);

        DocxDumpResult paragraph = editor.Dump(paragraphStream, "M.P0001", new DocxDumpOptions { IncludeRuns = true });
        DocxDumpResult table = editor.Dump(tableStream, "M.T0001");
        DocxDumpResult row = editor.Dump(rowStream, "M.T0001.R01");
        DocxDumpResult section = editor.Dump(sectionStream, "M.S0001");

        Assert.True(paragraph.Success);
        Assert.Contains("type=paragraph-properties-change", paragraph.Text, StringComparison.Ordinal);
        Assert.Contains("revision-id=1", paragraph.Text, StringComparison.Ordinal);
        Assert.True(table.Success);
        Assert.Contains("type=table-properties-change", table.Text, StringComparison.Ordinal);
        Assert.Contains("revision-id=2", table.Text, StringComparison.Ordinal);
        Assert.True(row.Success);
        Assert.Contains("changes:", row.Text, StringComparison.Ordinal);
        Assert.Contains("type=row-properties-change", row.Text, StringComparison.Ordinal);
        Assert.Contains("parent=row-properties", row.Text, StringComparison.Ordinal);
        Assert.Contains("revision-id=3", row.Text, StringComparison.Ordinal);
        Assert.Contains("M.T0001.R01.C01", row.Text, StringComparison.Ordinal);
        Assert.True(section.Success);
        Assert.Contains("type=section-properties-change", section.Text, StringComparison.Ordinal);
        Assert.Contains("revision-id=4", section.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Private", section.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void ContextShowsBookmarkNeighborhood()
    {
        using MemoryStream stream = CreateDocxWithSingleBookmark();
        var editor = new DocxEditor();
        DocxContextResult result = editor.Context(stream, "M.B0001");
        Assert.True(result.Success);
        DocxContextItem target = Assert.Single(result.Items, item => item.Id == "M.B0001");
        Assert.Equal("bookmark", target.Kind);
        Assert.Equal("target", target.Relation);
        Assert.Equal("main", target.Story);
        DocxContextItem anchor = Assert.Single(result.Items, item => item.Id == "M.P0001");
        Assert.Equal("anchor", anchor.Relation);
        Assert.Contains("ClientName", anchor.BookmarkNames);
        Assert.Contains(result.Items, item => item.Id == "M.P0002" && item.Relation == "after");
        Assert.Contains("target M.B0001 bookmark", result.Text, StringComparison.Ordinal);
        Assert.Contains("bookmark-names=\"ClientName\"", result.Text, StringComparison.Ordinal);
    }
    [Fact]
    public static void ContextShowsHyperlinkNeighborhood()
    {
        using MemoryStream stream = CreateDocxWithBody(
            """
                    <w:p xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                      <w:hyperlink r:id="rLink" w:tooltip="Open example">
                        <w:r><w:t>External</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/report" TargetMode="External"/>
                </Relationships>
                """, null);
        var editor = new DocxEditor();
        DocxContextResult result = editor.Context(stream, "M.L0001");
        Assert.True(result.Success);
        DocxContextItem target = Assert.Single(result.Items, item => item.Id == "M.L0001");
        Assert.Equal("hyperlink", target.Kind);
        Assert.Equal("target", target.Relation);
        Assert.Equal("main", target.Story);
        Assert.Equal("M.P0001", target.ParentId);
        DocxContextItem anchor = Assert.Single(result.Items, item => item.Id == "M.P0001");
        Assert.Equal("anchor", anchor.Relation);
        Assert.Contains("target M.L0001 hyperlink", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void ContextShowsFieldNeighborhood()
    {
        using MemoryStream stream = CreateDocxWithRefField();
        var editor = new DocxEditor();
        DocxContextResult result = editor.Context(stream, "M.F0001");
        Assert.True(result.Success);
        DocxContextItem target = Assert.Single(result.Items, item => item.Id == "M.F0001");
        Assert.Equal("field", target.Kind);
        Assert.Equal("target", target.Relation);
        Assert.Equal("main", target.Story);
        Assert.Contains(result.Items, item => item.Kind == "paragraph" && item.Relation == "anchor");
        Assert.Contains("target M.F0001 field", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void ContextShowsContentControlNeighborhood()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:id w:val="77"/>
                          <w:alias w:val="Client Name"/>
                          <w:tag w:val="client_name"/>
                          <w:text/>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>Acme</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        var editor = new DocxEditor();
        DocxContextResult result = editor.Context(stream, "M.CC0001");
        Assert.True(result.Success);
        DocxContextItem target = Assert.Single(result.Items, item => item.Id == "M.CC0001");
        Assert.Equal("content-control", target.Kind);
        Assert.Equal("target", target.Relation);
        Assert.Equal("main", target.Story);
        Assert.Equal("M.P0001", target.ParentId);
        DocxContextItem anchor = Assert.Single(result.Items, item => item.Id == "M.P0001");
        Assert.Equal("anchor", anchor.Relation);
        Assert.Contains("target M.CC0001 content-control", result.Text, StringComparison.Ordinal);
    }


    [Fact]
    public static void ContextShowsMergeGroupMembers()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
                          <w:p><w:r><w:t>North</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:vMerge/></w:tcPr>
                          <w:p><w:r><w:t>South</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc><w:p><w:r><w:t>Profit</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        var editor = new DocxEditor();
        DocxContextResult result = editor.Context(stream, "M.T0001.MG0001");
        Assert.True(result.Success);
        DocxContextItem target = Assert.Single(result.Items, item => item.Id == "M.T0001.MG0001");
        Assert.Equal("merge-group", target.Kind);
        Assert.Equal("target", target.Relation);
        Assert.Equal("M.T0001", target.ParentId);
        Assert.Contains(result.Items, item => item.Id == "M.T0001" && item.Relation == "parent");
        Assert.Contains("target M.T0001.MG0001 merge-group", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void ContextSummarizesNearbyStructureWithoutTextByDefault()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxContextResult result = editor.Context(stream, "M.P0002");

        Assert.True(result.Success);
        Assert.Contains(result.Items, item => item.Id == "M.P0001" && item.Relation == "before" && item.Text == string.Empty);
        DocxContextItem target = Assert.Single(result.Items, item => item.Id == "M.P0002");
        Assert.Equal("paragraph", target.Kind);
        Assert.Equal("target", target.Relation);
        Assert.Equal(string.Empty, target.Text);
        Assert.Contains("target M.P0002 paragraph", result.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Revenue", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void ContextCanSummarizeTableCellNeighborhood()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxContextResult result = editor.Context(stream, "M.T0001.R01.C02", new DocxContextOptions { MaxText = 20 });

        Assert.True(result.Success);
        Assert.Contains(result.Items, item => item.Id == "M.T0001" && item.Relation == "parent" && item.Kind == "table");
        DocxContextItem target = Assert.Single(result.Items, item => item.Id == "M.T0001.R01.C02");
        Assert.Equal("cell", target.Kind);
        Assert.Equal("target", target.Relation);
        Assert.Equal("M.T0001", target.ParentId);
        Assert.Equal(1, target.RowIndex);
        Assert.Equal(2, target.ColumnIndex);
        Assert.Equal("Revenue", target.Text);
    }

    [Fact]
    public static void ContextReportsUnknownTarget()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxContextResult result = editor.Context(stream, "M.P9999");

        Assert.False(result.Success);
        Assert.Empty(result.Items);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E1201" && diagnostic.TargetId == "M.P9999");
    }
    [Fact]
    public static void ContextUnknownTargetSuggestsNearestId()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxContextResult result = editor.Context(stream, "M.P002");

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = result.Diagnostics.First(static d => d.Code == "E1201");
        Assert.Contains("Did you mean", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("M.P0002", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void ContextFarMissStaysSilent()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxContextResult result = editor.Context(stream, "ZZZZZ");

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = result.Diagnostics.First(static d => d.Code == "E1201");
        Assert.DoesNotContain("Did you mean", diagnostic.Message, StringComparison.Ordinal);
    }


    [Fact]
    public static void DumpReportsUnknownTarget()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxDumpResult result = editor.Dump(stream, "M.P9999");

        Assert.False(result.Success);
        Assert.Null(result.Text);
        Assert.Empty(result.Runs);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E1201" && diagnostic.TargetId == "M.P9999");
    }
    [Fact]
    public static void DumpUnknownTargetSuggestsNearestId()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxDumpResult result = editor.Dump(stream, "M.P002");

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = result.Diagnostics.First(static d => d.Code == "E1201");
        Assert.Contains("Did you mean", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("M.P0002", diagnostic.Message, StringComparison.Ordinal);
    }


    [Fact]
    public static void DumpResolvesHeaderTarget()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxDumpResult result = editor.Dump(stream, "H001.P0001");

        Assert.True(result.Success);
        Assert.Contains("Confidential", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void ContextResolvesHeaderTargetWithoutOptIn()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxContextResult result = editor.Context(stream, "H001.P0001");

        Assert.True(result.Success);
        DocxContextItem target = Assert.Single(result.Items, item => item.Id == "H001.P0001");
        Assert.Equal("target", target.Relation);
    }

    [Fact]
    public static void MediaListsReferencedImagesOnly()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxMediaResult result = editor.Media(stream);

        DocxImageInfo image = Assert.Single(result.Images);
        Assert.Equal("M.I0001", image.Id.ToWireValue());
        Assert.Equal("/word/media/image1.png", image.PartName);
        Assert.Equal("inline", image.LayoutKind);
        Assert.Equal("rImage", image.RelationshipId);
        Assert.Equal("M.P0002", image.ContainingTargetId?.ToWireValue());
        Assert.Equal(914400, image.WidthEmu);
        Assert.Equal(457200, image.HeightEmu);
        Assert.Equal("Revenue chart", image.Description);
    }

    [Fact]
    public static void ReadModelsAnchoredImageLayoutMetadata()
    {
        using MemoryStream stream = CreateDocxWithImageBody("""
                    <w:p>
                      <w:r><w:t>Floating image</w:t></w:r>
                      <w:r>
                        <w:drawing>
                          <wp:anchor behindDoc="1" relativeHeight="251659264" distT="10" distB="20" distL="30" distR="40" allowOverlap="1">
                            <wp:positionH relativeFrom="column"><wp:posOffset>12345</wp:posOffset></wp:positionH>
                            <wp:positionV relativeFrom="paragraph"><wp:align>top</wp:align></wp:positionV>
                            <wp:extent cx="1000" cy="2000"/>
                            <wp:wrapSquare/>
                            <wp:docPr id="2" name="Floating picture" descr="Floating chart"/>
                            <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
                            <a:graphic>
                              <a:graphicData>
                                  <pic:pic>
                                    <pic:blipFill>
                                      <a:blip r:embed="rImage"/>
                                      <a:srcRect l="10000" t="5000" r="2500" b="0"/>
                                    </pic:blipFill>
                                  </pic:pic>
                                </a:graphicData>
                            </a:graphic>
                          </wp:anchor>
                        </w:drawing>
                      </w:r>
                    </w:p>
            """);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.True(result.Success);
        DocxImageInfo image = Assert.Single(result.Images);
        Assert.Equal("anchor", image.LayoutKind);
        Assert.Equal("M.P0001", image.ContainingTargetId?.ToWireValue());
        Assert.Equal(1000, image.WidthEmu);
        Assert.Equal(2000, image.HeightEmu);
        Assert.Equal("wrapSquare", image.WrapMode);
        Assert.True(image.BehindDoc);
        Assert.Equal(10L, image.WrapDistanceTopEmu);
        Assert.Equal(20L, image.WrapDistanceBottomEmu);
        Assert.Equal(30L, image.WrapDistanceLeftEmu);
        Assert.Equal(40L, image.WrapDistanceRightEmu);
        Assert.Equal(251659264L, image.RelativeHeight);
        Assert.True(image.AllowOverlap);
        Assert.True(image.LockAspectRatio);
        Assert.Equal("column", image.HorizontalPositionRelativeFrom);
        Assert.Equal(12345L, image.HorizontalPositionOffsetEmu);
        Assert.Null(image.HorizontalPositionAlign);
        Assert.Equal("paragraph", image.VerticalPositionRelativeFrom);
        Assert.Null(image.VerticalPositionOffsetEmu);
        Assert.Equal("top", image.VerticalPositionAlign);
        Assert.Equal(10m, image.CropLeftPercent);
        Assert.Equal(5m, image.CropTopPercent);
        Assert.Equal(2.5m, image.CropRightPercent);
        Assert.Equal(0m, image.CropBottomPercent);
        Assert.Equal("Floating chart", image.Description);
        Assert.Contains("layout=anchor", result.Text, StringComparison.Ordinal);
        Assert.Contains("wrap=wrapSquare", result.Text, StringComparison.Ordinal);
        Assert.Contains("behind-doc=true", result.Text, StringComparison.Ordinal);
        Assert.Contains("wrap-dist-top-emu=10", result.Text, StringComparison.Ordinal);
        Assert.Contains("wrap-dist-right-emu=40", result.Text, StringComparison.Ordinal);
        Assert.Contains("relative-height=251659264", result.Text, StringComparison.Ordinal);
        Assert.Contains("allow-overlap=true", result.Text, StringComparison.Ordinal);
        Assert.Contains("lock-aspect=true", result.Text, StringComparison.Ordinal);
        Assert.Contains("position-h-relative=column", result.Text, StringComparison.Ordinal);
        Assert.Contains("position-h-offset-emu=12345", result.Text, StringComparison.Ordinal);
        Assert.Contains("position-v-relative=paragraph", result.Text, StringComparison.Ordinal);
        Assert.Contains("position-v-align=top", result.Text, StringComparison.Ordinal);
        Assert.Contains("crop-left-percent=10", result.Text, StringComparison.Ordinal);
        Assert.Contains("crop-top-percent=5", result.Text, StringComparison.Ordinal);
        Assert.Contains("crop-right-percent=2.5", result.Text, StringComparison.Ordinal);
        Assert.Contains("crop-bottom-percent=0", result.Text, StringComparison.Ordinal);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1007" && diagnostic.Fallback == "modeled-metadata");
    }

    [Fact]
    public static void OutlineIncludesSections()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxOutlineResult result = editor.Outline(stream);

        DocxOutlineItem section = Assert.Single(result.Items, item => item.TargetId == "M.S0001");
        Assert.Equal("section", section.Kind);
        Assert.Equal(2, section.Columns);
        Assert.Equal("landscape", section.Orientation);
    }

    [Fact]
    public static void FindCanSearchHeaderAndFooterStoriesWhenRequested()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxFindResult result = editor.Find(stream, "Confidential", new DocxFindOptions { IncludeHeadersFooters = true });

        Assert.True(result.Success);
        Assert.Contains(result.Matches, match => match.TargetId == "H001.P0001" && match.Kind == "paragraph");
    }

    [Fact]
    public static void ContextAnnotatesTargetsWithBookmarkAndContentControlMetadata()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="1" w:name="ClientName"/>
                      <w:r><w:t>Client </w:t></w:r>
                      <w:sdt>
                        <w:sdtPr>
                          <w:alias w:val="Client Name"/>
                          <w:tag w:val="client_name"/>
                          <w:text/>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>Acme</w:t></w:r></w:sdtContent>
                      </w:sdt>
                      <w:bookmarkEnd w:id="1"/>
                    </w:p>
            """);
        var editor = new DocxEditor();

        DocxContextResult result = editor.Context(stream, "M.P0001");

        Assert.True(result.Success);
        DocxContextItem item = Assert.Single(result.Items);
        Assert.Equal(new[] { "ClientName" }, item.BookmarkNames);
        Assert.Equal(new[] { "M.CC0001" }, item.ContentControlIds);
        Assert.Equal(new[] { "client_name" }, item.ContentControlTags);
        Assert.Equal(new[] { "Client Name" }, item.ContentControlAliases);
        Assert.Contains("bookmark-names=\"ClientName\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("content-controls=\"M.CC0001\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("content-control-tags=\"client_name\"", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void ContextAnnotatesTargetsWithFieldMetadata()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> PAGE </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:t>1</w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                    </w:p>
            """);
        var editor = new DocxEditor();

        DocxContextResult result = editor.Context(stream, "M.P0001");

        Assert.True(result.Success);
        DocxContextItem item = Assert.Single(result.Items);
        Assert.Equal(new[] { "M.F0001" }, item.FieldIds);
        Assert.Equal(new[] { "PAGE" }, item.FieldCodes);
        Assert.Equal(new[] { "complex" }, item.FieldKinds);
        Assert.Equal(new[] { "PAGE" }, item.FieldTypes);
        Assert.Contains("fields=\"M.F0001\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("field-codes=\"PAGE\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("field-kinds=\"complex\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("field-types=\"PAGE\"", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void DumpRunsAnnotatesHyperlinkMarkup()
    {
        using MemoryStream stream = CreateDocxWithBody(
            """
                    <w:p xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                      <w:hyperlink r:id="rLink">
                        <w:r><w:t>External</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/report" TargetMode="External"/>
                </Relationships>
                """, null);
        var editor = new DocxEditor();

        DocxDumpResult result = editor.Dump(stream, "M.P0001", new DocxDumpOptions { IncludeRuns = true });

        Assert.True(result.Success);
        DocxDumpRunInfo run = Assert.Single(result.Runs);
        Assert.Equal("hyperlink", run.MarkupType);
        Assert.Equal("rLink", run.HyperlinkRelationshipId);
        Assert.Contains("markup=hyperlink", result.Text, StringComparison.Ordinal);
        Assert.Contains("hyperlink-relationship-id=rLink", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void ContextAnnotatesTargetsWithHyperlinkMetadata()
    {
        using MemoryStream stream = CreateDocxWithBody(
            """
                    <w:p xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                      <w:hyperlink r:id="rLink">
                        <w:r><w:t>External</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/report" TargetMode="External"/>
                </Relationships>
                """, null);
        var editor = new DocxEditor();

        DocxContextResult result = editor.Context(stream, "M.P0001");

        Assert.True(result.Success);
        DocxContextItem item = Assert.Single(result.Items);
        Assert.Equal(new[] { "M.L0001" }, item.HyperlinkIds);
        Assert.Equal(new[] { "https://example.test/report" }, item.HyperlinkTargets);
        Assert.Contains("hyperlinks=\"M.L0001\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("hyperlink-targets=\"https://example.test/report\"", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void DumpRunsAnnotateTrackedChangeAndCommentMarkup()
    {
        using MemoryStream stream = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:ins w:id="7" w:author="Alice" w:date="2026-06-01T12:00:00Z">
                        <w:r><w:t>Inserted</w:t></w:r>
                      </w:ins>
                      <w:del w:id="8" w:author="Bob" w:date="2026-06-02T12:00:00Z">
                        <w:r><w:delText>Deleted</w:delText></w:r>
                      </w:del>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="3" w:author="Reviewer">
                    <w:p><w:r><w:t>Private comment text</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        var editor = new DocxEditor();

        DocxDumpResult result = editor.Dump(stream, "M.P0001", new DocxDumpOptions
        {
            IncludeRuns = true,
            TextView = DocxTextView.Markup
        });

        Assert.True(result.Success);
        Assert.NotNull(result.Text);
        Assert.Contains("markup=inserted-run revision-id=7 author=\"Alice\" timestamp-utc=2026-06-01T12:00:00.0000000+00:00", result.Text, StringComparison.Ordinal);
        Assert.Contains("markup=deleted-run revision-id=8 author=\"Bob\" timestamp-utc=2026-06-02T12:00:00.0000000+00:00", result.Text, StringComparison.Ordinal);
        Assert.Contains("markup=comment-range-start comment-id=3", result.Text, StringComparison.Ordinal);
        Assert.Contains("markup=comment-range-end comment-id=3", result.Text, StringComparison.Ordinal);
        Assert.Contains("markup=comment-reference comment-id=3", result.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Private comment text", result.Text, StringComparison.Ordinal);
        Assert.Contains(result.Runs, run => run.MarkupType == "inserted-run" && run.RevisionId == "7" && run.Author == "Alice");
        Assert.Contains(result.Runs, run => run.MarkupType == "deleted-run" && run.RevisionId == "8" && run.Author == "Bob");
        Assert.Contains(result.Runs, run => run.MarkupType == "comment-reference" && run.CommentId == "3");
    }

    [Fact]
    public static void DumpCanTargetCommentBodyWithoutCommentText()
    {
        using MemoryStream stream = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="3" w:author="Reviewer" w:initials="RV" w:date="2026-06-07T12:00:00Z">
                    <w:p><w:r><w:t>Private comment text</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        var editor = new DocxEditor();

        DocxDumpResult result = editor.Dump(stream, "C001.C0001");

        Assert.True(result.Success);
        Assert.NotNull(result.Text);
        Assert.Contains("comment-id=3", result.Text, StringComparison.Ordinal);
        Assert.Contains("anchor-target=M.P0001", result.Text, StringComparison.Ordinal);
        Assert.Contains("reference-target=M.P0001", result.Text, StringComparison.Ordinal);
        Assert.Contains("text-length=20", result.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Private comment text", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void ContextAnnotatesTargetsWithCommentMetadata()
    {
        using MemoryStream stream = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="3" w:author="Reviewer">
                    <w:p><w:r><w:t>Private comment text</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        var editor = new DocxEditor();

        DocxContextResult result = editor.Context(stream, "M.P0001");

        Assert.True(result.Success);
        DocxContextItem item = Assert.Single(result.Items);
        Assert.Equal(new[] { "3" }, item.CommentIds);
        Assert.Equal(new[] { "C001.C0001" }, item.CommentBodyIds);
        Assert.Contains("comments=\"3\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("comment-bodies=\"C001.C0001\"", result.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Private comment text", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void ContextAnnotatesCommentThreadMetadataWithoutCommentText()
    {
        using MemoryStream stream = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w:comment w:id="3" w:author="Reviewer">
                    <w:p w14:paraId="00PARENT"><w:r><w:t>Private parent text</w:t></w:r></w:p>
                  </w:comment>
                  <w:comment w:id="4" w:author="Second Reviewer">
                    <w:p w14:paraId="00REPLY1"><w:r><w:t>Private reply text</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """,
            """
                <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w15:commentEx w15:paraId="00PARENT" w15:done="1"/>
                  <w15:commentEx w15:paraId="00REPLY1" w15:paraIdParent="00PARENT" w15:done="0"/>
                </w15:commentsEx>
                """,
            """
                <w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">
                  <w16cid:commentId w16cid:paraId="00PARENT" w16cid:durableId="DURABLEP"/>
                  <w16cid:commentId w16cid:paraId="00REPLY1" w16cid:durableId="DURABLER"/>
                </w16cid:commentsIds>
                """);
        var editor = new DocxEditor();

        DocxContextResult result = editor.Context(stream, "M.P0001");

        Assert.True(result.Success);
        DocxContextItem item = Assert.Single(result.Items);
        Assert.Equal(new[] { "3", "4" }, item.CommentIds);
        Assert.Equal(new[] { "C001.C0001", "C001.C0002" }, item.CommentBodyIds);
        Assert.Equal(new[] { "00PARENT", "00REPLY1" }, item.CommentParaIds);
        Assert.Equal(new[] { "00PARENT" }, item.CommentParentParaIds);
        Assert.Equal(new[] { "00PARENT" }, item.CommentRootParaIds);
        Assert.Equal(new[] { "DURABLEP", "DURABLER" }, item.CommentDurableIds);
        Assert.Equal(new[] { "4" }, item.CommentReplyIds);
        Assert.Equal(new[] { "3" }, item.CommentResolvedIds);
        Assert.Contains("comment-reply-ids=\"4\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("comment-durable-ids=\"DURABLEP,DURABLER\"", result.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Private parent text", result.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Private reply text", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void ContextCanTargetCommentBodyWithoutCommentText()
    {
        using MemoryStream stream = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="3" w:author="Reviewer">
                    <w:p><w:r><w:t>Private comment text</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        var editor = new DocxEditor();

        DocxContextResult result = editor.Context(stream, "comment:3");

        Assert.True(result.Success);
        DocxContextItem item = Assert.Single(result.Items);
        Assert.Equal("C001.C0001", item.Id);
        Assert.Equal("comment", item.Kind);
        Assert.Equal("M.P0001", item.ParentId);
        Assert.Equal(new[] { "3" }, item.CommentIds);
        Assert.Equal(new[] { "C001.C0001" }, item.CommentBodyIds);
        Assert.DoesNotContain("Private comment text", result.Text, StringComparison.Ordinal);
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
                """, null);
        var editor = new DocxEditor();

        DocxMediaResult result = editor.Media(stream);

        Assert.True(result.Success);
        Assert.Empty(result.Images);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1007" && diagnostic.Feature == "floating-image");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1008" && diagnostic.Feature == "external-image");
    }

    private static MemoryStream CreateDocxWithImageBody(string bodyXml)
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

                """ + bodyXml + """

                  </w:body>
                </w:document>
                """);
            AddEntry(archive, "word/media/image1.png", "fake-png");
        }

        stream.Position = 0;
        return stream;
    }
}
