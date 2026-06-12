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
        Assert.Contains("M.I0001 image layout=inline part=/word/media/image1.png", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void ValidateAcceptsBasicDocument()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxValidateResult result = editor.Validate(stream);

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error);
        Assert.Contains("/word/document.xml", result.PartNames);
        Assert.Equal("/word/document.xml", result.MainDocumentPartName);
    }

    [Fact]
    public static void ValidateReportsWordprocessingInvariants()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                        xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                      <w:bookmarkStart w:id="7" w:name="OpenBookmark"/>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r>
                        <w:drawing>
                          <wp:inline>
                            <wp:docPr id="5" name="Picture 5"/>
                            <a:graphic>
                              <a:graphicData>
                                <pic:pic>
                                  <pic:blipFill>
                                    <a:blip r:embed="rMissing"/>
                                  </pic:blipFill>
                                </pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                      <w:r>
                        <w:drawing>
                          <wp:inline>
                            <wp:docPr id="5" name="Duplicate picture 5"/>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:tbl>
                      <w:tr/>
                    </w:tbl>
            """);
        var editor = new DocxEditor();

        DocxValidateResult result = editor.Validate(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9103" &&
            diagnostic.PartName == "/word/document.xml" &&
            diagnostic.Message.Contains("bookmark", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E9104");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E9105");
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9107" &&
            diagnostic.Message.Contains("docPr id '5'", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9106" &&
            diagnostic.Message.Contains("row", StringComparison.Ordinal));
    }

    [Fact]
    public static void ValidateReportsInvalidCommentsExtendedMetadata()
    {
        using MemoryStream stream = CreateDocxWithBodyAndComments(
            """
                    <w:p><w:r><w:t>Body</w:t></w:r></w:p>
            """,
            """
                <w:comments
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w:comment w:id="1" w:author="Reviewer">
                    <w:p w15:paraId="00AAA111"><w:r><w:t>Comment</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
            """,
            """
                <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w15:commentEx w15:paraId="00AAA111" w15:done="0"/>
                  <w15:commentEx w15:paraId="00AAA111" w15:done="1"/>
                  <w15:commentEx w15:paraId="00BBB222" w15:done="0"/>
                  <w15:commentEx w15:done="0"/>
                </w15:commentsEx>
            """);
        var editor = new DocxEditor();

        DocxValidateResult result = editor.Validate(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9108" &&
            diagnostic.PartName == "/word/commentsExtended.xml" &&
            diagnostic.Message.Contains("Duplicate commentsExtended paraId '00AAA111'", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9108" &&
            diagnostic.Message.Contains("00BBB222", StringComparison.Ordinal) &&
            diagnostic.Message.Contains("no matching comment paragraph", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9108" &&
            diagnostic.Message.Contains("missing w15:paraId", StringComparison.Ordinal));
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
        Assert.True(table.HasMergedCells);
        Assert.True(table.HasNestedTables);
        DocxTableCellInfo wide = table.Cells.Single(cell => cell.Id == "M.T0001.R01.C01");
        Assert.Equal(2, wide.ColumnSpan);
        Assert.Equal(2, wide.VisualColumnEndIndex);
        Assert.Equal("M.T0001.MG0001", wide.MergeGroupId);
        Assert.Equal("restart", wide.VerticalMerge);
        Assert.Equal("M.T0001.R01.C01", wide.VerticalMergeRootCellId);
        Assert.False(wide.HasNestedTable);
        DocxTableCellInfo east = table.Cells.Single(cell => cell.Id == "M.T0001.R01.C03");
        Assert.Equal("East", east.Text);
        DocxTableCellInfo continued = table.Cells.Single(cell => cell.Id == "M.T0001.R02.C01");
        Assert.Equal(1, continued.VisualColumnEndIndex);
        Assert.Equal("M.T0001.MG0001", continued.MergeGroupId);
        Assert.Equal("continue", continued.VerticalMerge);
        Assert.Equal("M.T0001.R01.C01", continued.VerticalMergeRootCellId);
        DocxTableCellInfo nested = table.Cells.Single(cell => cell.Id == "M.T0001.R02.C02");
        Assert.True(nested.HasNestedTable);
        Assert.Contains("M.T0001.R01.C01 physical-column=1 column-span=2 visual-column-end=2 merge-group=M.T0001.MG0001 vertical-merge=restart vertical-merge-root=M.T0001.R01.C01", result.Text, StringComparison.Ordinal);
        Assert.Contains("M.T0001.R02.C01 physical-column=1 merge-group=M.T0001.MG0001 vertical-merge=continue vertical-merge-root=M.T0001.R01.C01", result.Text, StringComparison.Ordinal);
        Assert.Contains("M.T0001.R02.C02 physical-column=2 nested-table=true", result.Text, StringComparison.Ordinal);
        Assert.Contains("M.T0001 table rows=2 columns=3 merged=true nested-table=true", result.Text, StringComparison.Ordinal);

        stream.Position = 0;
        DocxContextResult context = editor.Context(stream, "M.T0001.R02.C01", new DocxContextOptions { MaxText = 20 });
        DocxContextItem contextTarget = Assert.Single(context.Items, item => item.Id == "M.T0001.R02.C01");
        Assert.Equal("M.T0001.MG0001", contextTarget.MergeGroupId);
        Assert.Equal("M.T0001.R01.C01", contextTarget.VerticalMergeRootCellId);
        Assert.Contains("merge-group=M.T0001.MG0001 vertical-merge=continue vertical-merge-root=M.T0001.R01.C01", context.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void ReadModelsTableRowsStyleAndGridMetadata()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>
                      <w:tblGrid>
                        <w:gridCol w:w="2000"/>
                        <w:gridCol w:w="2000"/>
                        <w:gridCol w:w="2000"/>
                      </w:tblGrid>
                      <w:tr>
                        <w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>
                        <w:tc><w:p><w:r><w:t>Header 1</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Header 2</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Header 3</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:trPr><w:gridBefore w:val="1"/></w:trPr>
                        <w:tc><w:p><w:r><w:t>Offset 1</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Offset 2</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        DocxTableInfo table = Assert.Single(result.Tables);
        Assert.Equal("TableGrid", table.StyleId);
        Assert.Equal(3, table.GridColumnCount);
        Assert.True(table.HasHeaderRow);
        Assert.True(table.HasMergedCells);
        Assert.Equal(2, table.Rows.Count);
        Assert.True(table.Rows[0].IsHeader);
        Assert.True(table.Rows[0].CantSplit);
        Assert.Equal(1, table.Rows[1].GridBefore);
        Assert.Equal(2, table.Rows[1].CellCount);
        DocxTableCellInfo offset = table.Cells.Single(cell => cell.Text == "Offset 1");
        Assert.Equal(2, offset.ColumnIndex);
        Assert.Equal(1, offset.PhysicalColumnIndex);

        Assert.Contains("M.T0001 table rows=2 columns=3 styleId=TableGrid grid-columns=3 header-row=true merged=true", result.Text, StringComparison.Ordinal);
        Assert.Contains("M.T0001.R01 row cells=3 header=true cant-split=true", result.Text, StringComparison.Ordinal);
        Assert.Contains("M.T0001.R02 row cells=2 grid-before=1", result.Text, StringComparison.Ordinal);
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
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E2001" && diagnostic.TargetId == "M.P9999");
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
    public static void ReadResolvesNumberingDefinitionMetadata()
    {
        using MemoryStream stream = CreateDocxWithStylesAndNumbering(
            """
                    <w:p>
                      <w:pPr>
                        <w:numPr>
                          <w:ilvl w:val="1"/>
                          <w:numId w:val="9"/>
                        </w:numPr>
                      </w:pPr>
                      <w:r><w:t>Nested bullet</w:t></w:r>
                    </w:p>
            """,
            """
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
            """,
            """
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="7">
                    <w:lvl w:ilvl="1">
                      <w:numFmt w:val="bullet"/>
                      <w:lvlText w:val="o"/>
                      <w:pStyle w:val="BulletStyle"/>
                    </w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="9">
                    <w:abstractNumId w:val="7"/>
                  </w:num>
                </w:numbering>
            """);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        DocxListInfo list = Assert.Single(result.Paragraphs).List!;
        Assert.Equal("9", list.NumberingId);
        Assert.Equal(1, list.Level);
        Assert.Equal("7", list.AbstractNumberingId);
        Assert.Equal("bullet", list.Format);
        Assert.Equal("o", list.LevelText);
        Assert.Equal("BulletStyle", list.ParagraphStyleId);
        Assert.Equal("direct", list.Source);
        Assert.Equal("o", list.LabelText);
        Assert.Equal("resolved", list.LabelStatus);
        Assert.Contains("list numId=9 level=1 abstractNumId=7 format=bullet level-text=\"o\" paragraph-style=BulletStyle", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void ReadResolvesStyleLinkedNumbering()
    {
        using MemoryStream stream = CreateDocxWithStylesAndNumbering(
            """
                    <w:p>
                      <w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr>
                      <w:r><w:t>Styled list item</w:t></w:r>
                    </w:p>
            """,
            """
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="paragraph" w:styleId="ListParagraph">
                    <w:name w:val="List Paragraph"/>
                    <w:basedOn w:val="Normal"/>
                    <w:pPr>
                      <w:numPr>
                        <w:ilvl w:val="0"/>
                        <w:numId w:val="11"/>
                      </w:numPr>
                    </w:pPr>
                  </w:style>
                </w:styles>
            """,
            """
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="2">
                    <w:lvl w:ilvl="0">
                      <w:numFmt w:val="decimal"/>
                      <w:lvlText w:val="%1."/>
                    </w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="11">
                    <w:abstractNumId w:val="2"/>
                  </w:num>
                </w:numbering>
            """);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        DocxParagraphInfo paragraph = Assert.Single(result.Paragraphs);
        Assert.Equal("ListParagraph", paragraph.StyleId);
        Assert.Equal("List Paragraph", paragraph.StyleName);
        Assert.NotNull(paragraph.List);
        Assert.Equal("style", paragraph.List.Source);
        Assert.Equal("decimal", paragraph.List.Format);
        Assert.Equal("1.", paragraph.List.LabelText);
        Assert.Contains("styleId=ListParagraph list numId=11 level=0 abstractNumId=2 format=decimal level-text=\"%1.\" source=style", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void ReadExpandsResolvedNumberingLabels()
    {
        using MemoryStream stream = CreateDocxWithStylesAndNumbering(
            """
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr>
                      <w:r><w:t>Top one</w:t></w:r>
                    </w:p>
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="9"/></w:numPr></w:pPr>
                      <w:r><w:t>Nested one</w:t></w:r>
                    </w:p>
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="9"/></w:numPr></w:pPr>
                      <w:r><w:t>Nested two</w:t></w:r>
                    </w:p>
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr>
                      <w:r><w:t>Top two</w:t></w:r>
                    </w:p>
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="9"/></w:numPr></w:pPr>
                      <w:r><w:t>Nested reset</w:t></w:r>
                    </w:p>
            """,
            """
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
            """,
            """
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="7">
                    <w:lvl w:ilvl="0">
                      <w:start w:val="3"/>
                      <w:numFmt w:val="decimal"/>
                      <w:lvlText w:val="%1."/>
                      <w:suff w:val="space"/>
                    </w:lvl>
                    <w:lvl w:ilvl="1">
                      <w:start w:val="2"/>
                      <w:numFmt w:val="lowerLetter"/>
                      <w:lvlText w:val="%1.%2)"/>
                    </w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="9">
                    <w:abstractNumId w:val="7"/>
                  </w:num>
                </w:numbering>
            """);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        string[] labels = result.Paragraphs
            .Select(paragraph => paragraph.List?.LabelText ?? string.Empty)
            .ToArray();
        Assert.Equal(["3.", "3.b)", "3.c)", "4.", "4.b)"], labels);
        Assert.All(result.Paragraphs, paragraph => Assert.Equal("resolved", paragraph.List?.LabelStatus));
        Assert.Equal(3, result.Paragraphs[0].List?.StartValue);
        Assert.Equal("space", result.Paragraphs[0].List?.Suffix);
        Assert.Contains("M.P0002 paragraph list numId=9 level=1 abstractNumId=7 format=lowerLetter level-text=\"%1.%2)\" label=\"3.b)\" start=2", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void ReadUsesNumberingStartOverridePerNumberingInstance()
    {
        using MemoryStream stream = CreateDocxWithStylesAndNumbering(
            """
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr>
                      <w:r><w:t>Override start</w:t></w:r>
                    </w:p>
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="10"/></w:numPr></w:pPr>
                      <w:r><w:t>Default start</w:t></w:r>
                    </w:p>
            """,
            """
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
            """,
            """
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="7">
                    <w:lvl w:ilvl="0">
                      <w:start w:val="1"/>
                      <w:numFmt w:val="decimal"/>
                      <w:lvlText w:val="%1."/>
                    </w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="9">
                    <w:abstractNumId w:val="7"/>
                    <w:lvlOverride w:ilvl="0">
                      <w:startOverride w:val="7"/>
                    </w:lvlOverride>
                  </w:num>
                  <w:num w:numId="10">
                    <w:abstractNumId w:val="7"/>
                  </w:num>
                </w:numbering>
            """);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.Equal("7.", result.Paragraphs[0].List?.LabelText);
        Assert.Equal(7, result.Paragraphs[0].List?.StartValue);
        Assert.Equal("1.", result.Paragraphs[1].List?.LabelText);
        Assert.Equal(1, result.Paragraphs[1].List?.StartValue);
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
        Assert.Equal("inline", image.LayoutKind);
        Assert.Equal("rImage", image.RelationshipId);
        Assert.Equal("M.P0002", image.ContainingTargetId);
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
        Assert.Equal("M.P0001", image.ContainingTargetId);
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
    public static void StylesExposeInheritanceLinksAndNumberingDefaults()
    {
        using MemoryStream stream = CreateDocxWithStylesAndNumbering(
            """
                    <w:p><w:r><w:t>Styled paragraph</w:t></w:r></w:p>
            """,
            """
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="paragraph" w:styleId="Base">
                    <w:name w:val="Base"/>
                  </w:style>
                  <w:style w:type="paragraph" w:styleId="Derived">
                    <w:name w:val="Derived"/>
                    <w:basedOn w:val="Base"/>
                    <w:next w:val="NextStyle"/>
                    <w:link w:val="DerivedChar"/>
                    <w:pPr>
                      <w:numPr>
                        <w:ilvl w:val="2"/>
                        <w:numId w:val="44"/>
                      </w:numPr>
                    </w:pPr>
                  </w:style>
                  <w:style w:type="character" w:styleId="DerivedChar">
                    <w:name w:val="Derived Char"/>
                  </w:style>
                </w:styles>
            """,
            """
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
            """);
        var editor = new DocxEditor();

        DocxStylesResult result = editor.Styles(stream);

        DocxStyleInfo style = result.Styles.Single(style => style.StyleId == "Derived");
        Assert.Equal("Base", style.BasedOnStyleId);
        Assert.Equal("NextStyle", style.NextStyleId);
        Assert.Equal("DerivedChar", style.LinkedStyleId);
        Assert.Equal("44", style.NumberingId);
        Assert.Equal(2, style.NumberingLevel);
        string stylesText = DocxTextRenderer.RenderStyles(result);
        Assert.Contains("styleId=Derived", stylesText, StringComparison.Ordinal);
        Assert.Contains("based-on=Base", stylesText, StringComparison.Ordinal);
        Assert.Contains("numbering numId=44 level=2", stylesText, StringComparison.Ordinal);
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
        Assert.Contains(result.GroupSummary, summary => summary.Group == "story" && summary.Key == "main" && summary.Type == "inserted-run" && summary.Count == 1);
        Assert.Contains(result.GroupSummary, summary => summary.Group == "part" && summary.Key == "/word/document.xml" && summary.Type == "deleted-run" && summary.Count == 1);
        Assert.Contains(result.GroupSummary, summary => summary.Group == "author" && summary.Key == "Alice" && summary.Type == "inserted-run" && summary.Count == 1);
        Assert.Contains(result.GroupSummary, summary => summary.Group == "target" && summary.Key == "M.P0001" && summary.Type == "run-properties-change" && summary.Count == 1);
        DocxChangeTargetSummary paragraphSummary = Assert.Single(result.TargetSummary, summary => summary.TargetId == "M.P0001");
        Assert.Contains(paragraphSummary.Summary, summary => summary.Type == "inserted-run" && summary.Count == 1);

        DocxChangeInfo insertion = Assert.Single(result.Changes, change => change.Type == "inserted-run");
        Assert.Equal("M.CH0001", insertion.Id);
        Assert.Equal("main", insertion.Story);
        Assert.Equal("/word/document.xml", insertion.PartName);
        Assert.Equal("M.P0001", insertion.TargetId);
        Assert.Equal("targeted", insertion.TargetStatus);
        Assert.Equal("ancestor", insertion.TargetSource);
        Assert.Null(insertion.TargetNote);
        Assert.Equal("Alice", insertion.Author);
        Assert.Equal("9", insertion.RevisionId);
        Assert.Equal(8, insertion.TextLength);
        Assert.Equal(DateTimeOffset.Parse("2026-06-01T12:00:00Z").ToUniversalTime(), insertion.TimestampUtc);

        DocxChangeInfo cellChange = Assert.Single(result.Changes, change => change.Type == "cell-properties-change");
        Assert.Equal("M.T0001.R01.C01", cellChange.TargetId);

        DocxChangeInfo customRangeStart = Assert.Single(result.Changes, change => change.Type == "custom-xml-delete-range-start");
        DocxChangeInfo customRangeEnd = Assert.Single(result.Changes, change => change.Type == "custom-xml-delete-range-end");
        Assert.Equal("adjacent-range", customRangeStart.TargetSource);
        Assert.Equal("range-boundary", customRangeStart.TargetReason);
        Assert.Equal(customRangeEnd.Id, customRangeStart.PairedChangeId);
        Assert.Equal(customRangeStart.Id, customRangeEnd.PairedChangeId);

        string serialized = JsonSerializer.Serialize(result);
        Assert.DoesNotContain("Inserted", serialized, StringComparison.Ordinal);
        Assert.DoesNotContain("Deleted", serialized, StringComparison.Ordinal);
    }

    [Fact]
    public static void ChangesListsCommentMarkupWithoutCommentText()
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

        DocxChangesResult result = editor.Changes(stream);

        Assert.True(result.Success);
        Assert.Contains(result.Summary, summary => summary.Type == "comment-range-start" && summary.Count == 1);
        Assert.Contains(result.Summary, summary => summary.Type == "comment-range-end" && summary.Count == 1);
        Assert.Contains(result.Summary, summary => summary.Type == "comment-reference" && summary.Count == 1);
        Assert.Contains(result.Summary, summary => summary.Type == "comment" && summary.Count == 1);

        DocxChangeInfo commentStart = Assert.Single(result.Changes, change => change.Type == "comment-range-start");
        Assert.Equal("M.P0001", commentStart.TargetId);
        Assert.Equal("ancestor", commentStart.TargetSource);
        Assert.Null(commentStart.RevisionId);
        Assert.Equal("3", commentStart.CommentId);
        Assert.Equal("M.P0001", commentStart.CommentAnchorTargetId);
        Assert.Equal("M.P0001", commentStart.CommentReferenceTargetId);
        Assert.Equal("main", commentStart.CommentAnchorStory);
        Assert.Equal("/word/document.xml", commentStart.CommentAnchorPartName);
        Assert.Equal("Reviewer", commentStart.CommentAuthor);
        Assert.Equal("RV", commentStart.CommentInitials);
        Assert.Equal(DateTimeOffset.Parse("2026-06-07T12:00:00Z").ToUniversalTime(), commentStart.CommentTimestampUtc);

        DocxChangeInfo comment = Assert.Single(result.Changes, change => change.Type == "comment");
        Assert.Equal("comments[1]", comment.Story);
        Assert.Equal("/word/comments.xml", comment.PartName);
        Assert.Equal("C001.C0001", comment.TargetId);
        Assert.Equal("M.P0001", comment.CommentAnchorTargetId);
        Assert.Equal("M.P0001", comment.CommentReferenceTargetId);
        Assert.Equal("main", comment.CommentAnchorStory);
        Assert.Equal("/word/document.xml", comment.CommentAnchorPartName);
        Assert.Equal(20, comment.TextLength);
        Assert.Null(comment.CommentTextLength);
        Assert.Null(comment.CommentTextSnippet);
        DocxChangeInfo commentEnd = Assert.Single(result.Changes, change => change.Type == "comment-range-end");
        Assert.Equal(commentEnd.Id, commentStart.PairedChangeId);
        Assert.Equal(commentStart.Id, commentEnd.PairedChangeId);
        Assert.Contains(result.GroupSummary, summary => summary.Group == "target" && summary.Key == "M.P0001" && summary.Type == "comment" && summary.Count == 1);
        DocxCommentThreadSummary commentSummary = Assert.Single(result.CommentSummary);
        Assert.Equal("3", commentSummary.CommentId);
        Assert.Equal("M.P0001", commentSummary.AnchorTargetId);
        Assert.Equal("M.P0001", commentSummary.ReferenceTargetId);
        Assert.Equal("Reviewer", commentSummary.Author);
        Assert.Null(commentSummary.TextLength);
        Assert.Null(commentSummary.TextSnippet);
        Assert.Contains(commentSummary.Summary, summary => summary.Type == "comment" && summary.Count == 1);

        string serialized = JsonSerializer.Serialize(result);
        Assert.DoesNotContain("Private comment text", serialized, StringComparison.Ordinal);
    }

    [Fact]
    public static void ChangesCanIncludeBoundedCommentTextWhenExplicitlyRequested()
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

        DocxChangesResult result = editor.Changes(stream, new DocxChangesOptions
        {
            IncludeCommentText = true,
            MaxCommentText = 7
        });

        Assert.True(result.Success);
        DocxChangeInfo comment = Assert.Single(result.Changes, change => change.Type == "comment");
        Assert.Equal(20, comment.CommentTextLength);
        Assert.Equal("Private", comment.CommentTextSnippet);
        Assert.True(comment.CommentTextTruncated);

        DocxChangeInfo commentStart = Assert.Single(result.Changes, change => change.Type == "comment-range-start");
        Assert.Null(commentStart.CommentTextSnippet);

        DocxCommentThreadSummary commentSummary = Assert.Single(result.CommentSummary);
        Assert.Equal("RV", commentSummary.Initials);
        Assert.Equal(20, commentSummary.TextLength);
        Assert.Equal("Private", commentSummary.TextSnippet);
        Assert.True(commentSummary.TextTruncated);

        string rendered = DocxTextRenderer.RenderChanges(result);
        Assert.Contains("comment-text-length=20", rendered, StringComparison.Ordinal);
        Assert.Contains("comment-text=\"Private\"", rendered, StringComparison.Ordinal);
        Assert.Contains("comment-text-truncated=true", rendered, StringComparison.Ordinal);
        Assert.DoesNotContain("Private comment text", rendered, StringComparison.Ordinal);
    }

    [Fact]
    public static void ChangesExposeCommentResolutionMetadata()
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
                    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w:comment w:id="3" w:author="Reviewer" w:initials="RV" w:date="2026-06-07T12:00:00Z">
                    <w:p w15:paraId="00ABCDEF"><w:r><w:t>Private comment text</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """,
            """
                <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w15:commentEx w15:paraId="00ABCDEF" w15:paraIdParent="00112233" w15:done="1"/>
                </w15:commentsEx>
                """);
        var editor = new DocxEditor();

        DocxChangesResult result = editor.Changes(stream);

        DocxChangeInfo comment = Assert.Single(result.Changes, change => change.Type == "comment");
        Assert.Equal("00ABCDEF", comment.CommentParaId);
        Assert.Equal("00112233", comment.CommentParentParaId);
        Assert.True(comment.CommentResolved);
        Assert.Null(comment.CommentTextSnippet);
        DocxCommentThreadSummary summary = Assert.Single(result.CommentSummary);
        Assert.Equal("00ABCDEF", summary.ParaId);
        Assert.Equal("00112233", summary.ParentParaId);
        Assert.True(summary.Resolved);
        string rendered = DocxTextRenderer.RenderChanges(result);
        Assert.Contains("resolved=true", rendered, StringComparison.Ordinal);
        Assert.Contains("comment-resolved=true", rendered, StringComparison.Ordinal);
        Assert.DoesNotContain("Private comment text", rendered, StringComparison.Ordinal);
    }

    [Fact]
    public static void ChangesClassifiesTargetlessRecords()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:del w:id="44" w:author="Reviewer">
                      <w:r><w:delText>Body level deletion</w:delText></w:r>
                    </w:del>
                    <w:p><w:r><w:t>Nearby modeled paragraph</w:t></w:r></w:p>
            """);
        var editor = new DocxEditor();

        DocxChangesResult result = editor.Changes(stream);

        DocxChangeInfo deletion = Assert.Single(result.Changes);
        Assert.Equal("deleted-run", deletion.Type);
        Assert.Null(deletion.TargetId);
        Assert.Equal("targetless", deletion.TargetStatus);
        Assert.Equal("none", deletion.TargetSource);
        Assert.Equal("body-level-markup", deletion.TargetReason);
        Assert.Equal("M.P0001", deletion.NearestTargetId);
        Assert.Contains("direct child of the document body", deletion.TargetNote, StringComparison.Ordinal);
        Assert.Contains("nearest-target=M.P0001", deletion.TargetNote, StringComparison.Ordinal);
        Assert.Contains(result.TargetSummary, summary => summary.TargetId == "(none)" && summary.Count == 1);
    }

    [Fact]
    public static void ReadModelsBookmarkAndContentControlMetadata()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="1" w:name="ClientName"/>
                      <w:r><w:t>Client </w:t></w:r>
                      <w:sdt>
                        <w:sdtPr>
                          <w:id w:val="77"/>
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

        DocxReadResult result = editor.Read(stream);

        Assert.True(result.Success);
        DocxBookmarkInfo bookmark = Assert.Single(result.Bookmarks);
        Assert.Equal("M.B0001", bookmark.Id);
        Assert.Equal("ClientName", bookmark.Name);
        Assert.Equal("1", bookmark.OoxmlId);
        Assert.Equal("main", bookmark.Story);
        Assert.Equal("/word/document.xml", bookmark.PartName);
        Assert.Equal("M.P0001", bookmark.StartTargetId);
        Assert.Equal("M.P0001", bookmark.EndTargetId);
        Assert.True(bookmark.IsComplete);

        DocxContentControlInfo control = Assert.Single(result.ContentControls);
        Assert.Equal("M.CC0001", control.Id);
        Assert.Equal("plain-text", control.Kind);
        Assert.Equal("77", control.OoxmlId);
        Assert.Equal("client_name", control.Tag);
        Assert.Equal("Client Name", control.Alias);
        Assert.Equal("M.P0001", control.TargetId);
        Assert.Equal(4, control.TextLength);

        Assert.Contains("M.B0001 bookmark name=\"ClientName\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("M.CC0001 content-control kind=plain-text", result.Text, StringComparison.Ordinal);
        Assert.Contains("tag=\"client_name\"", result.Text, StringComparison.Ordinal);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1005" && diagnostic.Fallback == "modeled-metadata");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1006" && diagnostic.Fallback == "modeled-metadata");
    }

    [Fact]
    public static void ReadModelsAdvancedContentControlMetadata()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:checkBox>
                            <w:checked w:val="1"/>
                            <w:checkedState w:val="2612"/>
                            <w:uncheckedState w:val="2610"/>
                          </w:checkBox>
                          <w:tag w:val="accepted"/>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>Checked</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:dropDownList>
                            <w:listItem w:displayText="North" w:value="north"/>
                            <w:listItem w:displayText="South" w:value="south"/>
                          </w:dropDownList>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>North</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:date>
                            <w:dateFormat w:val="yyyy-MM-dd"/>
                            <w:lid w:val="en-US"/>
                            <w:calendar w:val="gregorian"/>
                          </w:date>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>2026-06-12</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        DocxContentControlInfo checkbox = result.ContentControls.Single(control => control.Kind == "checkbox");
        Assert.True(checkbox.Checked);
        Assert.Equal("2612", checkbox.CheckedSymbol);
        Assert.Equal("2610", checkbox.UncheckedSymbol);
        DocxContentControlInfo dropdown = result.ContentControls.Single(control => control.Kind == "dropdown-list");
        Assert.Equal(2, dropdown.ListItems.Count);
        Assert.Equal("North", dropdown.ListItems[0].DisplayText);
        Assert.Equal("north", dropdown.ListItems[0].Value);
        DocxContentControlInfo date = result.ContentControls.Single(control => control.Kind == "date");
        Assert.Equal("yyyy-MM-dd", date.DateFormat);
        Assert.Equal("en-US", date.DateLanguage);
        Assert.Equal("gregorian", date.DateCalendar);
        Assert.Contains("kind=checkbox", result.Text, StringComparison.Ordinal);
        Assert.Contains("checked=true", result.Text, StringComparison.Ordinal);
        Assert.Contains("list-items=2", result.Text, StringComparison.Ordinal);
        Assert.Contains("date-format=\"yyyy-MM-dd\"", result.Text, StringComparison.Ordinal);
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
    public static void ReadModelsSimpleAndComplexFieldMetadata()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p>
                      <w:fldSimple w:instr=" DATE " w:dirty="true" w:fldLock="1">
                        <w:r><w:t>June 12, 2026</w:t></w:r>
                      </w:fldSimple>
                    </w:p>
                    <w:p>
                      <w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>
                      <w:r><w:instrText> REF  ClientName \h </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:t>Client result</w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                    </w:p>
            """);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.True(result.Success);
        Assert.Equal(2, result.Fields.Count);
        DocxFieldInfo simple = result.Fields[0];
        Assert.Equal("M.F0001", simple.Id);
        Assert.Equal("simple", simple.Kind);
        Assert.Equal("DATE", simple.Code);
        Assert.Equal("M.P0001", simple.TargetId);
        Assert.Equal(13, simple.ResultTextLength);
        Assert.True(simple.IsDirty);
        Assert.True(simple.IsLocked);
        Assert.True(simple.IsComplete);

        DocxFieldInfo complex = result.Fields[1];
        Assert.Equal("M.F0002", complex.Id);
        Assert.Equal("complex", complex.Kind);
        Assert.Equal(@"REF ClientName \h", complex.Code);
        Assert.Equal("M.P0002", complex.TargetId);
        Assert.Equal(13, complex.ResultTextLength);
        Assert.True(complex.IsDirty);
        Assert.Null(complex.IsLocked);
        Assert.True(complex.IsComplete);

        Assert.Contains("M.F0001 field kind=simple", result.Text, StringComparison.Ordinal);
        Assert.Contains("code=\"DATE\"", result.Text, StringComparison.Ordinal);
        Assert.Contains(@"code=""REF ClientName \\h""", result.Text, StringComparison.Ordinal);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1003" && diagnostic.Fallback == "modeled-metadata");
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
        Assert.Contains("fields=\"M.F0001\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("field-codes=\"PAGE\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("field-kinds=\"complex\"", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void ReadModelsExternalInternalAndBrokenHyperlinks()
    {
        using MemoryStream stream = CreateDocxWithBody(
            """
                    <w:p xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                      <w:hyperlink r:id="rLink" w:tooltip="Open example" w:tgtFrame="_blank" w:history="1">
                        <w:r><w:t>External</w:t></w:r>
                      </w:hyperlink>
                      <w:r><w:t xml:space="preserve"> </w:t></w:r>
                      <w:hyperlink w:anchor="Section1">
                        <w:r><w:t>Internal</w:t></w:r>
                      </w:hyperlink>
                      <w:r><w:t xml:space="preserve"> </w:t></w:r>
                      <w:hyperlink r:id="rMissing">
                        <w:r><w:t>Broken</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/report" TargetMode="External"/>
                </Relationships>
                """);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.True(result.Success);
        Assert.Equal(3, result.Hyperlinks.Count);
        DocxHyperlinkInfo external = result.Hyperlinks[0];
        Assert.Equal("M.L0001", external.Id);
        Assert.Equal("M.P0001", external.TargetId);
        Assert.Equal("rLink", external.RelationshipId);
        Assert.Equal("https://example.test/report", external.Uri);
        Assert.Equal("https", external.UriScheme);
        Assert.True(external.IsUriValid);
        Assert.Null(external.UriValidationReason);
        Assert.True(external.IsExternal);
        Assert.False(external.IsBroken);
        Assert.Equal("Open example", external.Tooltip);
        Assert.Equal("_blank", external.TargetFrame);
        Assert.True(external.History);
        Assert.Equal(8, external.DisplayTextLength);

        DocxHyperlinkInfo internalLink = result.Hyperlinks[1];
        Assert.Equal("Section1", internalLink.Anchor);
        Assert.True(internalLink.IsAnchorMissing);
        Assert.False(internalLink.IsAnchorDuplicate);
        Assert.False(internalLink.IsExternal);
        Assert.False(internalLink.IsBroken);

        DocxHyperlinkInfo broken = result.Hyperlinks[2];
        Assert.Equal("rMissing", broken.RelationshipId);
        Assert.True(broken.IsBroken);
        Assert.Null(broken.Uri);

        Assert.Contains("M.L0001 hyperlink", result.Text, StringComparison.Ordinal);
        Assert.Contains("uri=\"https://example.test/report\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("uri-scheme=https", result.Text, StringComparison.Ordinal);
        Assert.Contains("uri-valid=true", result.Text, StringComparison.Ordinal);
        Assert.Contains("target-frame=\"_blank\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("history=true", result.Text, StringComparison.Ordinal);
        Assert.Contains("anchor=\"Section1\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("anchor-missing=true", result.Text, StringComparison.Ordinal);
        Assert.Contains("broken=True", result.Text, StringComparison.Ordinal);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1002" && diagnostic.Fallback == "modeled-metadata");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1015" && diagnostic.Fallback == "broken-relationship");
    }

    [Fact]
    public static void ReadValidatesHyperlinkDestinations()
    {
        using MemoryStream stream = CreateDocxWithBody(
            """
                    <w:p>
                      <w:bookmarkStart w:id="1" w:name="Known"/>
                      <w:r><w:t>Anchor</w:t></w:r>
                      <w:bookmarkEnd w:id="1"/>
                    </w:p>
                    <w:p>
                      <w:bookmarkStart w:id="2" w:name="Dup"/>
                      <w:r><w:t>First duplicate</w:t></w:r>
                      <w:bookmarkEnd w:id="2"/>
                    </w:p>
                    <w:p>
                      <w:bookmarkStart w:id="3" w:name="Dup"/>
                      <w:r><w:t>Second duplicate</w:t></w:r>
                      <w:bookmarkEnd w:id="3"/>
                    </w:p>
                    <w:p xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                      <w:hyperlink r:id="rUnsafe"><w:r><w:t>Unsafe</w:t></w:r></w:hyperlink>
                      <w:r><w:t xml:space="preserve"> </w:t></w:r>
                      <w:hyperlink r:id="rRelative"><w:r><w:t>Relative</w:t></w:r></w:hyperlink>
                      <w:r><w:t xml:space="preserve"> </w:t></w:r>
                      <w:hyperlink w:anchor="Missing"><w:r><w:t>Missing</w:t></w:r></w:hyperlink>
                      <w:r><w:t xml:space="preserve"> </w:t></w:r>
                      <w:hyperlink w:anchor="Known"><w:r><w:t>Known</w:t></w:r></w:hyperlink>
                      <w:r><w:t xml:space="preserve"> </w:t></w:r>
                      <w:hyperlink w:anchor="Dup"><w:r><w:t>Duplicate</w:t></w:r></w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rUnsafe" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="javascript:alert(1)" TargetMode="External"/>
                  <Relationship Id="rRelative" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="../relative/report" TargetMode="External"/>
                </Relationships>
                """);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.True(result.Success);
        DocxHyperlinkInfo unsafeLink = result.Hyperlinks.Single(link => link.RelationshipId == "rUnsafe");
        Assert.Equal("javascript", unsafeLink.UriScheme);
        Assert.False(unsafeLink.IsUriValid);
        Assert.Equal("unsupported-uri-scheme", unsafeLink.UriValidationReason);

        DocxHyperlinkInfo relativeLink = result.Hyperlinks.Single(link => link.RelationshipId == "rRelative");
        Assert.Null(relativeLink.UriScheme);
        Assert.False(relativeLink.IsUriValid);
        Assert.Equal("malformed-or-relative-uri", relativeLink.UriValidationReason);

        DocxHyperlinkInfo missing = result.Hyperlinks.Single(link => link.Anchor == "Missing");
        Assert.True(missing.IsAnchorMissing);
        Assert.False(missing.IsAnchorDuplicate);

        DocxHyperlinkInfo known = result.Hyperlinks.Single(link => link.Anchor == "Known");
        Assert.False(known.IsAnchorMissing);
        Assert.False(known.IsAnchorDuplicate);

        DocxHyperlinkInfo duplicate = result.Hyperlinks.Single(link => link.Anchor == "Dup");
        Assert.False(duplicate.IsAnchorMissing);
        Assert.True(duplicate.IsAnchorDuplicate);

        Assert.Contains("uri-valid=false", result.Text, StringComparison.Ordinal);
        Assert.Contains("uri-reason=unsupported-uri-scheme", result.Text, StringComparison.Ordinal);
        Assert.Contains("uri-reason=malformed-or-relative-uri", result.Text, StringComparison.Ordinal);
        Assert.Contains("anchor-missing=true", result.Text, StringComparison.Ordinal);
        Assert.Contains("anchor-duplicate=true", result.Text, StringComparison.Ordinal);
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
                """);
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
                """);
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
    public static void DocxChangeInfoUsesPropertyBasedPublicShape()
    {
        Assert.Contains(typeof(DocxChangeInfo).GetConstructors(), constructor => constructor.GetParameters().Length == 0);
        Assert.DoesNotContain(typeof(DocxChangeInfo).GetConstructors(), constructor => constructor.GetParameters().Length > 0);

        var change = new DocxChangeInfo
        {
            Id = "M.CH0001",
            Type = "inserted-run",
            Story = "main",
            PartName = "/word/document.xml",
            TextLength = 8,
            ChildElementCount = 1
        };

        string serialized = JsonSerializer.Serialize(change);
        Assert.Contains("\"Id\":\"M.CH0001\"", serialized, StringComparison.Ordinal);
        Assert.Contains("\"Type\":\"inserted-run\"", serialized, StringComparison.Ordinal);
        Assert.Contains("\"TextLength\":8", serialized, StringComparison.Ordinal);
    }

    [Fact]
    public static void ChangesLinksBodyLevelRangeMarkersToAdjacentTargets()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:moveFromRangeStart w:id="4" w:author="Alice"/>
                    <w:p><w:r><w:t>Moved paragraph</w:t></w:r></w:p>
                    <w:moveFromRangeEnd w:id="4"/>
                    <w:customXmlDelRangeStart w:id="5" w:author="Bob"/>
                    <w:tbl>
                      <w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:customXmlDelRangeEnd w:id="5"/>
            """);
        var editor = new DocxEditor();

        DocxChangesResult result = editor.Changes(stream);

        Assert.True(result.Success);
        DocxChangeInfo moveStart = Assert.Single(result.Changes, change => change.Type == "move-from-range-start");
        DocxChangeInfo moveEnd = Assert.Single(result.Changes, change => change.Type == "move-from-range-end");
        DocxChangeInfo customStart = Assert.Single(result.Changes, change => change.Type == "custom-xml-delete-range-start");
        DocxChangeInfo customEnd = Assert.Single(result.Changes, change => change.Type == "custom-xml-delete-range-end");

        Assert.Equal("M.P0001", moveStart.TargetId);
        Assert.Equal("M.P0001", moveEnd.TargetId);
        Assert.Equal("M.T0001", customStart.TargetId);
        Assert.Equal("M.T0001", customEnd.TargetId);
        Assert.Equal("adjacent-range", moveStart.TargetSource);
        Assert.Equal("adjacent-range", moveEnd.TargetSource);
        Assert.Equal(moveEnd.Id, moveStart.PairedChangeId);
        Assert.Equal(moveStart.Id, moveEnd.PairedChangeId);
        Assert.Equal(customEnd.Id, customStart.PairedChangeId);
        Assert.Equal(customStart.Id, customEnd.PairedChangeId);
    }

    [Fact]
    public static void ReadCanSwitchTrackedChangeTextViews()
    {
        using MemoryStream finalStream = CreateDocxWithBody(TrackedChangeBodyXml);
        using MemoryStream originalStream = CreateDocxWithBody(TrackedChangeBodyXml);
        using MemoryStream markupStream = CreateDocxWithBody(TrackedChangeBodyXml);
        var editor = new DocxEditor();

        DocxReadResult final = editor.Read(finalStream);
        DocxReadResult original = editor.Read(originalStream, new DocxReadOptions { TextView = DocxTextView.Original });
        DocxReadResult markup = editor.Read(markupStream, new DocxReadOptions { TextView = DocxTextView.Markup });

        Assert.Equal("Before Inserted After", Assert.Single(final.Paragraphs).Text);
        Assert.Equal("Before Deleted After", Assert.Single(original.Paragraphs).Text);
        Assert.Contains("[+Inserted +]", Assert.Single(markup.Paragraphs).Text, StringComparison.Ordinal);
        Assert.Contains("[-Deleted -]", Assert.Single(markup.Paragraphs).Text, StringComparison.Ordinal);
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
    public static void ReadFindAndDumpReportUnsupportedFeatureDiagnosticsConsistently()
    {
        using MemoryStream readStream = CreateDocxWithBody("""
                    <w:p xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                      <w:hyperlink r:id="rLink"><w:r><w:t>Link</w:t></w:r></w:hyperlink>
                      <w:fldSimple w:instr="DATE"><w:r><w:t>Revenue</w:t></w:r></w:fldSimple>
                      <w:commentRangeStart w:id="1"/>
                      <w:ins w:id="2" w:author="A"><w:r><w:t>Inserted</w:t></w:r></w:ins>
                    </w:p>
            """);
        using MemoryStream findStream = CreateDocxWithBody("""
                    <w:p xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                      <w:hyperlink r:id="rLink"><w:r><w:t>Link</w:t></w:r></w:hyperlink>
                      <w:fldSimple w:instr="DATE"><w:r><w:t>Revenue</w:t></w:r></w:fldSimple>
                      <w:commentRangeStart w:id="1"/>
                      <w:ins w:id="2" w:author="A"><w:r><w:t>Inserted</w:t></w:r></w:ins>
                    </w:p>
            """);
        using MemoryStream dumpStream = CreateDocxWithBody("""
                    <w:p xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                      <w:hyperlink r:id="rLink"><w:r><w:t>Link</w:t></w:r></w:hyperlink>
                      <w:fldSimple w:instr="DATE"><w:r><w:t>Revenue</w:t></w:r></w:fldSimple>
                      <w:commentRangeStart w:id="1"/>
                      <w:ins w:id="2" w:author="A"><w:r><w:t>Inserted</w:t></w:r></w:ins>
                    </w:p>
            """);
        var editor = new DocxEditor();

        DocxReadResult read = editor.Read(readStream);
        DocxFindResult find = editor.Find(findStream, "Revenue");
        DocxDumpResult dump = editor.Dump(dumpStream, "M.P0001");

        AssertUnsupportedFeatureDiagnostics(read.Diagnostics);
        AssertUnsupportedFeatureDiagnostics(find.Diagnostics);
        AssertUnsupportedFeatureDiagnostics(dump.Diagnostics);
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
        Assert.Contains(warnings, diagnostic => diagnostic.Code == "W1001" && diagnostic.Fallback == "selected-text-view");
        Assert.Contains(warnings, diagnostic => diagnostic.Code == "W1008" && diagnostic.Fallback == "omit-from-editable-images");
    }

    private const string TrackedChangeBodyXml = """
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:ins w:id="1" w:author="Alice">
                        <w:r><w:t>Inserted </w:t></w:r>
                      </w:ins>
                      <w:del w:id="2" w:author="Bob">
                        <w:r><w:delText>Deleted </w:delText></w:r>
                      </w:del>
                      <w:r><w:t>After</w:t></w:r>
                    </w:p>
            """;

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
                            <wp:extent cx="914400" cy="457200"/>
                            <wp:docPr id="1" name="Picture 1" descr="Revenue chart" title="Chart title"/>
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

    private static MemoryStream CreateDocxWithStylesAndNumbering(
        string bodyXml,
        string stylesXml,
        string numberingXml)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                </Types>
                """);
            AddEntry(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                  <Relationship Id="rNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/document.xml", """
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>

                """ + bodyXml + """

                  </w:body>
                </w:document>
                """);
            AddEntry(archive, "word/styles.xml", stylesXml);
            AddEntry(archive, "word/numbering.xml", numberingXml);
        }

        stream.Position = 0;
        return stream;
    }

    private static void AssertUnsupportedFeatureDiagnostics(IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "W1001" && diagnostic.Feature == "tracked-changes");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "W1002" && diagnostic.Feature == "hyperlink");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "W1003" && diagnostic.Feature == "field");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "W1004" && diagnostic.Feature == "comment");
    }

    private static MemoryStream CreateDocxWithBodyAndComments(
        string bodyXml,
        string commentsXml,
        string? commentsExtendedXml = null)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            string contentTypesXml = commentsExtendedXml is null
                ? """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
                </Types>
                """
                : """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
                  <Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.ms-word.commentsExtended+xml"/>
                </Types>
                """;
            AddEntry(archive, "[Content_Types].xml", contentTypesXml);
            AddEntry(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            string relationshipsXml = commentsExtendedXml is null
                ? """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
                </Relationships>
                """
                : """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
                  <Relationship Id="rCommentsExtended" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>
                </Relationships>
                """;
            AddEntry(archive, "word/_rels/document.xml.rels", relationshipsXml);
            AddEntry(archive, "word/document.xml", """
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>

                """ + bodyXml + """

                  </w:body>
                </w:document>
                """);
            AddEntry(archive, "word/comments.xml", commentsXml);
            if (commentsExtendedXml is not null)
            {
                AddEntry(archive, "word/commentsExtended.xml", commentsExtendedXml);
            }
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
