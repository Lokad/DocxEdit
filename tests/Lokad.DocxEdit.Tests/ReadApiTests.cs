using System.IO.Compression;
using System.Text;
using System.Text.Json;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

public static class ReadApiTests
{






















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
        DocxTableCellInfo wide = table.Cells.Single(cell => cell.Id.ToWireValue() == "M.T0001.R01.C01");
        Assert.Equal(2, wide.ColumnSpan);
        Assert.Equal(2, wide.VisualColumnEndIndex);
        Assert.Equal("M.T0001.MG0001", wide.MergeGroupId?.ToWireValue());
        Assert.Equal(DocxVerticalMerge.Restart, wide.VerticalMerge);
        Assert.Equal("M.T0001.R01.C01", wide.VerticalMergeRootCellId?.ToWireValue());
        Assert.False(wide.HasNestedTable);
        DocxTableCellInfo east = table.Cells.Single(cell => cell.Id.ToWireValue() == "M.T0001.R01.C03");
        Assert.Equal("East", east.Text);
        DocxTableCellInfo continued = table.Cells.Single(cell => cell.Id.ToWireValue() == "M.T0001.R02.C01");
        Assert.Equal(1, continued.VisualColumnEndIndex);
        Assert.Equal("M.T0001.MG0001", continued.MergeGroupId?.ToWireValue());
        Assert.Equal(DocxVerticalMerge.Continue, continued.VerticalMerge);
        Assert.Equal("M.T0001.R01.C01", continued.VerticalMergeRootCellId?.ToWireValue());
        DocxTableCellInfo nested = table.Cells.Single(cell => cell.Id.ToWireValue() == "M.T0001.R02.C02");
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
                      <w:tblPr>
                        <w:tblStyle w:val="TableGrid"/>
                        <w:tblCaption w:val="Revenue summary"/>
                        <w:tblDescription w:val="Quarterly revenue by region"/>
                      </w:tblPr>
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
        Assert.Equal("Revenue summary", table.Caption);
        Assert.Equal("Quarterly revenue by region", table.Description);
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

        Assert.Contains("M.T0001 table rows=2 columns=3 styleId=TableGrid caption=\"Revenue summary\" description=\"Quarterly revenue by region\" grid-columns=3 header-row=true merged=true", result.Text, StringComparison.Ordinal);
        Assert.Contains("M.T0001.R01 row cells=3 header=true cant-split=true", result.Text, StringComparison.Ordinal);
        Assert.Contains("M.T0001.R02 row cells=2 grid-before=1", result.Text, StringComparison.Ordinal);

        stream.Position = 0;
        DocxContextResult context = editor.Context(stream, "M.T0001.R01.C01");
        DocxContextItem parent = Assert.Single(context.Items, item => item.Id == "M.T0001");
        Assert.Equal("Revenue summary", parent.Caption);
        Assert.Equal("Quarterly revenue by region", parent.Description);
        Assert.Contains("caption=\"Revenue summary\" description=\"Quarterly revenue by region\"", context.Text, StringComparison.Ordinal);
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

        DocxListInfo list = Assert.IsType<DocxListInfo>(Assert.Single(result.Paragraphs).List);
        Assert.Equal("9", list.NumberingId);
        Assert.Equal(1, list.Level);
        Assert.Equal("7", list.AbstractNumberingId);
        Assert.Equal("bullet", list.Format);
        Assert.Equal("o", list.LevelText);
        Assert.Equal("BulletStyle", list.ParagraphStyleId);
        Assert.Equal(DocxLabelSource.Direct, list.Source);
        Assert.Equal("o", list.LabelText);
        Assert.Equal(DocxLabelStatus.Resolved, list.LabelStatus);
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
        Assert.Equal(DocxLabelSource.Style, paragraph.List.Source);
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
                      <w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr>
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
        Assert.All(result.Paragraphs, paragraph => Assert.True(paragraph.List?.LabelStatus == DocxLabelStatus.Resolved));
        Assert.Equal(3, result.Paragraphs[0].List?.StartValue);
        Assert.Equal("space", result.Paragraphs[0].List?.Suffix);
        DocxListLabelComponent[] components = Assert.IsType<DocxListInfo>(result.Paragraphs[1].List).LabelComponents.ToArray();
        Assert.Equal(2, components.Length);
        Assert.Equal(new DocxListLabelComponent(0, 3, "3", "decimal"), components[0]);
        Assert.Equal(new DocxListLabelComponent(1, 2, "b", "lowerLetter"), components[1]);
        Assert.Contains("M.P0002 paragraph list numId=9 level=1 abstractNumId=7 format=lowerLetter level-text=\"%1.%2)\" label=\"3.b)\" label-components=\"0:3:decimal:3,1:2:lowerLetter:b\" start=2", result.Text, StringComparison.Ordinal);

        stream.Position = 0;
        DocxFindResult find = editor.Find(stream, "Nested two");
        Assert.Contains(find.Matches, match => match.TargetId == "M.P0003" && match.Kind == "paragraph" && match.Text == "Nested two" && match.List?.LabelText == "3.c)");

        stream.Position = 0;
        DocxOutlineResult outline = editor.Outline(stream);
        Assert.Contains(outline.Items, item => item.TargetId == "M.P0001" && item.Kind == "heading" && item.HeadingLevel == 1 && item.Text == "Top one" && item.List?.LabelText == "3.");
    }

    [Fact]
    public static void ReadNumberingLabelsAreStableAcrossTrackedRunTextViews()
    {
        using MemoryStream stream = CreateDocxWithStylesAndNumbering(
            """
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr>
                      <w:r><w:t>First </w:t></w:r>
                      <w:ins w:id="1" w:author="Alice"><w:r><w:t>inserted</w:t></w:r></w:ins>
                      <w:del w:id="2" w:author="Bob"><w:r><w:delText>deleted</w:delText></w:r></w:del>
                    </w:p>
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr>
                      <w:r><w:t>Second</w:t></w:r>
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
                  <w:num w:numId="9"><w:abstractNumId w:val="7"/></w:num>
                </w:numbering>
            """);
        var editor = new DocxEditor();

        DocxReadResult finalView = editor.Read(stream);
        stream.Position = 0;
        DocxReadResult originalView = editor.Read(stream, new DocxReadOptions { TextView = DocxTextView.Original });
        stream.Position = 0;
        DocxReadResult markupView = editor.Read(stream, new DocxReadOptions { TextView = DocxTextView.Markup });

        Assert.Equal(["1.", "2."], finalView.Paragraphs.Select(paragraph => paragraph.List?.LabelText ?? string.Empty).ToArray());
        Assert.Equal(["1.", "2."], originalView.Paragraphs.Select(paragraph => paragraph.List?.LabelText ?? string.Empty).ToArray());
        Assert.Equal(["1.", "2."], markupView.Paragraphs.Select(paragraph => paragraph.List?.LabelText ?? string.Empty).ToArray());
        Assert.Equal("First inserted", finalView.Paragraphs[0].Text);
        Assert.Equal("First deleted", originalView.Paragraphs[0].Text);
        Assert.Equal("First [+inserted+][-deleted-]", markupView.Paragraphs[0].Text);
    }

    [Fact]
    public static void ReadNumberingLabelsFollowBlockRevisionTextViews()
    {
        using MemoryStream stream = CreateDocxWithStylesAndNumbering(
            """
                    <w:p>
                      <w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr>
                      <w:r><w:t>First</w:t></w:r>
                    </w:p>
                    <w:ins w:id="1" w:author="Alice">
                      <w:p>
                        <w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr>
                        <w:r><w:t>Inserted</w:t></w:r>
                      </w:p>
                    </w:ins>
                    <w:p>
                      <w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr>
                      <w:r><w:t>Second</w:t></w:r>
                    </w:p>
                    <w:del w:id="2" w:author="Bob">
                      <w:p>
                        <w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr>
                        <w:r><w:delText>Deleted</w:delText></w:r>
                      </w:p>
                    </w:del>
                    <w:p>
                      <w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr>
                      <w:r><w:t>Third</w:t></w:r>
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
                  <w:num w:numId="9"><w:abstractNumId w:val="7"/></w:num>
                </w:numbering>
            """);
        var editor = new DocxEditor();

        DocxReadResult finalView = editor.Read(stream);
        stream.Position = 0;
        DocxReadResult originalView = editor.Read(stream, new DocxReadOptions { TextView = DocxTextView.Original });
        stream.Position = 0;
        DocxReadResult markupView = editor.Read(stream, new DocxReadOptions { TextView = DocxTextView.Markup });

        Assert.Equal(["First", "Inserted", "Second", "Third"], finalView.Paragraphs.Select(paragraph => paragraph.Text).ToArray());
        Assert.Equal(["1.", "2.", "3.", "4."], finalView.Paragraphs.Select(paragraph => paragraph.List?.LabelText ?? string.Empty).ToArray());
        Assert.Equal(["First", "Second", "Deleted", "Third"], originalView.Paragraphs.Select(paragraph => paragraph.Text).ToArray());
        Assert.Equal(["1.", "2.", "3.", "4."], originalView.Paragraphs.Select(paragraph => paragraph.List?.LabelText ?? string.Empty).ToArray());
        Assert.Equal(["First", "[+Inserted+]", "Second", "[-Deleted-]", "Third"], markupView.Paragraphs.Select(paragraph => paragraph.Text).ToArray());
        Assert.Equal(["1.", "2.", "3.", "4.", "5."], markupView.Paragraphs.Select(paragraph => paragraph.List?.LabelText ?? string.Empty).ToArray());

        stream.Position = 0;
        DocxFindResult finalFind = editor.Find(stream, "Deleted");
        Assert.Empty(finalFind.Matches);

        stream.Position = 0;
        DocxFindResult originalFind = editor.Find(stream, "Deleted", new DocxFindOptions { TextView = DocxTextView.Original });
        Assert.Contains(originalFind.Matches, match => match.TargetId == "M.P0004" && match.Text == "Deleted" && match.List?.LabelText == "3.");

        stream.Position = 0;
        DocxContextResult originalContext = editor.Context(stream, "M.P0004", new DocxContextOptions { TextView = DocxTextView.Original, Radius = 0, MaxText = 20 });
        DocxContextItem contextTarget = Assert.Single(originalContext.Items);
        Assert.Equal("Deleted", contextTarget.Text);
        Assert.Equal("3.", contextTarget.List?.LabelText);

        stream.Position = 0;
        DocxOutlineResult finalOutline = editor.Outline(stream);
        Assert.DoesNotContain(finalOutline.Items, item => item.Text.Contains("Deleted", StringComparison.Ordinal));
        Assert.Contains(finalOutline.Items, item => item.TargetId == "M.P0002" && item.Kind == "heading" && item.List?.LabelText == "2." && item.Text.Contains("Inserted", StringComparison.Ordinal));

        stream.Position = 0;
        DocxOutlineResult markupOutline = editor.Outline(stream, new DocxOutlineOptions { TextView = DocxTextView.Markup });
        Assert.Contains(markupOutline.Items, item => item.TargetId == "M.P0004" && item.Kind == "heading" && item.List?.LabelText == "4." && item.Text.Contains("[-Deleted-]", StringComparison.Ordinal));
    }

    [Fact]
    public static void ReadNumberingLabelsFollowMultiLevelBlockRevisionTextViews()
    {
        using MemoryStream stream = CreateDocxWithStylesAndNumbering(
            """
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr>
                      <w:r><w:t>Top one</w:t></w:r>
                    </w:p>
                    <w:ins w:id="1" w:author="Alice">
                      <w:p>
                        <w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="9"/></w:numPr></w:pPr>
                        <w:r><w:t>Inserted child</w:t></w:r>
                      </w:p>
                    </w:ins>
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="9"/></w:numPr></w:pPr>
                      <w:r><w:t>Normal child</w:t></w:r>
                    </w:p>
                    <w:del w:id="2" w:author="Bob">
                      <w:p>
                        <w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="9"/></w:numPr></w:pPr>
                        <w:r><w:delText>Deleted child</w:delText></w:r>
                      </w:p>
                    </w:del>
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr>
                      <w:r><w:t>Top two</w:t></w:r>
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
                    <w:lvl w:ilvl="1">
                      <w:start w:val="1"/>
                      <w:numFmt w:val="decimal"/>
                      <w:lvlText w:val="%1.%2."/>
                    </w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="9"><w:abstractNumId w:val="7"/></w:num>
                </w:numbering>
            """);
        var editor = new DocxEditor();

        DocxReadResult finalView = editor.Read(stream);
        stream.Position = 0;
        DocxReadResult originalView = editor.Read(stream, new DocxReadOptions { TextView = DocxTextView.Original });
        stream.Position = 0;
        DocxReadResult markupView = editor.Read(stream, new DocxReadOptions { TextView = DocxTextView.Markup });

        Assert.Equal(["1.", "1.1.", "1.2.", "2."], finalView.Paragraphs.Select(paragraph => paragraph.List?.LabelText ?? string.Empty).ToArray());
        Assert.Equal(["1.", "1.1.", "1.2.", "2."], originalView.Paragraphs.Select(paragraph => paragraph.List?.LabelText ?? string.Empty).ToArray());
        Assert.Equal(["1.", "1.1.", "1.2.", "1.3.", "2."], markupView.Paragraphs.Select(paragraph => paragraph.List?.LabelText ?? string.Empty).ToArray());
        Assert.Equal("Inserted child", finalView.Paragraphs[1].Text);
        Assert.Equal("Deleted child", originalView.Paragraphs[2].Text);
        Assert.Equal("[-Deleted child-]", markupView.Paragraphs[3].Text);
    }

    [Fact]
    public static void ReadNumberingStartOverridesFollowBlockRevisionTextViews()
    {
        using MemoryStream stream = CreateDocxWithStylesAndNumbering(
            """
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr>
                      <w:r><w:t>Default one</w:t></w:r>
                    </w:p>
                    <w:ins w:id="1" w:author="Alice">
                      <w:p>
                        <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="10"/></w:numPr></w:pPr>
                        <w:r><w:t>Inserted override</w:t></w:r>
                      </w:p>
                    </w:ins>
                    <w:del w:id="2" w:author="Bob">
                      <w:p>
                        <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="10"/></w:numPr></w:pPr>
                        <w:r><w:delText>Deleted override</w:delText></w:r>
                      </w:p>
                    </w:del>
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr>
                      <w:r><w:t>Default two</w:t></w:r>
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
                  <w:num w:numId="9"><w:abstractNumId w:val="7"/></w:num>
                  <w:num w:numId="10">
                    <w:abstractNumId w:val="7"/>
                    <w:lvlOverride w:ilvl="0">
                      <w:startOverride w:val="7"/>
                    </w:lvlOverride>
                  </w:num>
                </w:numbering>
            """);
        var editor = new DocxEditor();

        DocxReadResult finalView = editor.Read(stream);
        stream.Position = 0;
        DocxReadResult originalView = editor.Read(stream, new DocxReadOptions { TextView = DocxTextView.Original });
        stream.Position = 0;
        DocxReadResult markupView = editor.Read(stream, new DocxReadOptions { TextView = DocxTextView.Markup });

        Assert.Equal(["1.", "7.", "2."], finalView.Paragraphs.Select(paragraph => paragraph.List?.LabelText ?? string.Empty).ToArray());
        Assert.Equal(["1.", "7.", "2."], originalView.Paragraphs.Select(paragraph => paragraph.List?.LabelText ?? string.Empty).ToArray());
        Assert.Equal(["1.", "7.", "8.", "2."], markupView.Paragraphs.Select(paragraph => paragraph.List?.LabelText ?? string.Empty).ToArray());
        Assert.Equal(7, finalView.Paragraphs[1].List?.StartValue);
        Assert.Equal(7, originalView.Paragraphs[1].List?.StartValue);
        Assert.Equal(7, markupView.Paragraphs[1].List?.StartValue);
    }

    [Fact]
    public static void ReadWarnsWhenNumberingPropertyRevisionAffectsOriginalViewLabels()
    {
        using MemoryStream stream = CreateDocxWithStylesAndNumbering(
            """
                    <w:p>
                      <w:pPr>
                        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr>
                        <w:pPrChange w:id="3" w:author="Alice">
                          <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="10"/></w:numPr></w:pPr>
                        </w:pPrChange>
                      </w:pPr>
                      <w:r><w:t>Changed numbering</w:t></w:r>
                    </w:p>
            """,
            """
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
            """,
            """
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="7">
                    <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="9"><w:abstractNumId w:val="7"/></w:num>
                  <w:num w:numId="10"><w:abstractNumId w:val="7"/></w:num>
                </w:numbering>
            """);

        DocxReadResult result = new DocxEditor().Read(stream, new DocxReadOptions { TextView = DocxTextView.Original });

        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "W1026" &&
            diagnostic.Feature == "numbering" &&
            diagnostic.Fallback == "tracked-numbering-property-revision");
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
    public static void ReadExtractsSectionColumnsAndOrientation()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        DocxSectionInfo section = Assert.Single(result.Sections);
        Assert.Equal("M.S0001", section.Id.ToWireValue());
        Assert.Equal(2, section.Columns);
        Assert.Equal(DocxOrientation.Landscape, section.Orientation);
        Assert.Contains("M.S0001 section columns=2 orientation=landscape", result.Text, StringComparison.Ordinal);
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

        Assert.DoesNotContain(defaultResult.Paragraphs, paragraph => paragraph.Id.ToWireValue().StartsWith("H", StringComparison.Ordinal));
        Assert.Contains(allStories.Paragraphs, paragraph => paragraph.Id.ToWireValue() == "H001.P0001" && paragraph.Story == "header[1]" && paragraph.Text == "Confidential");
        Assert.Contains(allStories.Paragraphs, paragraph => paragraph.Id.ToWireValue() == "F001.P0001" && paragraph.Story == "footer[1]" && paragraph.Text == "Page 1");
    }

    [Fact]
    public static void ReadResolvesHeaderAndFooterNumberingWhenRequested()
    {
        using MemoryStream stream = CreateDocxWithBody(
            """
                    <w:p><w:r><w:t>Main body</w:t></w:r></w:p>
                    <w:sectPr xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                      <w:headerReference w:type="default" r:id="rHeader"/>
                      <w:footerReference w:type="default" r:id="rFooter"/>
                    </w:sectPr>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                  <Relationship Id="rNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                  <Relationship Id="rHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                  <Relationship Id="rFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
                </Relationships>
                """,
            archive =>
            {
                AddEntry(archive, "word/styles.xml", """
                    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
                    """);
                AddEntry(archive, "word/numbering.xml", """
                    <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                      <w:abstractNum w:abstractNumId="7">
                        <w:lvl w:ilvl="0">
                          <w:start w:val="1"/>
                          <w:numFmt w:val="decimal"/>
                          <w:lvlText w:val="%1."/>
                        </w:lvl>
                      </w:abstractNum>
                      <w:num w:numId="9"><w:abstractNumId w:val="7"/></w:num>
                    </w:numbering>
                    """);
                AddEntry(archive, "word/header1.xml", """
                    <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                      <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr><w:r><w:t>Header item</w:t></w:r></w:p>
                    </w:hdr>
                    """);
                AddEntry(archive, "word/footer1.xml", """
                    <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                      <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr><w:r><w:t>Footer item</w:t></w:r></w:p>
                    </w:ftr>
                    """);
            });

        DocxReadResult result = new DocxEditor().Read(stream, new DocxReadOptions { IncludeHeadersFooters = true });

        DocxParagraphInfo header = Assert.Single(result.Paragraphs, paragraph => paragraph.Id.ToWireValue() == "H001.P0001");
        DocxParagraphInfo footer = Assert.Single(result.Paragraphs, paragraph => paragraph.Id.ToWireValue() == "F001.P0001");
        Assert.Equal("1.", header.List?.LabelText);
        Assert.Equal("1.", footer.List?.LabelText);
        Assert.Contains("H001.P0001 paragraph list numId=9 level=0 abstractNumId=7 format=decimal level-text=\"%1.\" label=\"1.\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("F001.P0001 paragraph list numId=9 level=0 abstractNumId=7 format=decimal level-text=\"%1.\" label=\"1.\"", result.Text, StringComparison.Ordinal);
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
        Assert.Equal(new DocxChangeId('M', 0, 1), insertion.Id);
        Assert.Equal("main", insertion.Story);
        Assert.Equal("/word/document.xml", insertion.PartName);
        Assert.Equal("paragraph", insertion.ParentType);
        Assert.Equal("M.P0001", insertion.TargetId);
        Assert.Equal(DocxTargetStatus.Targeted, insertion.TargetStatus);
        Assert.Equal(DocxTargetSource.Ancestor, insertion.TargetSource);
        Assert.Null(insertion.TargetNote);
        Assert.Equal("Alice", insertion.Author);
        Assert.Equal("9", insertion.RevisionId);
        Assert.Equal(8, insertion.TextLength);
        Assert.Equal(DateTimeOffset.Parse("2026-06-01T12:00:00Z").ToUniversalTime(), insertion.TimestampUtc);

        DocxChangeInfo cellChange = Assert.Single(result.Changes, change => change.Type == "cell-properties-change");
        Assert.Equal("cell-properties", cellChange.ParentType);
        Assert.Equal("M.T0001.R01.C01", cellChange.TargetId);

        DocxChangeInfo customRangeStart = Assert.Single(result.Changes, change => change.Type == "custom-xml-delete-range-start");
        DocxChangeInfo customRangeEnd = Assert.Single(result.Changes, change => change.Type == "custom-xml-delete-range-end");
        Assert.Equal(DocxTargetSource.AdjacentRange, customRangeStart.TargetSource);
        Assert.True(customRangeStart.TargetReason == DocxTargetReason.RangeBoundary);
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
        Assert.Equal(DocxTargetSource.Ancestor, commentStart.TargetSource);
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
                """, null);
        var editor = new DocxEditor();

        DocxChangesResult result = editor.Changes(stream);

        DocxChangeInfo comment = Assert.Single(result.Changes, change => change.Type == "comment");
        Assert.Equal("00ABCDEF", comment.CommentParaId);
        Assert.Equal("00112233", comment.CommentParentParaId);
        Assert.Equal("00112233", comment.CommentRootParaId);
        Assert.True(comment.CommentIsReply);
        Assert.True(comment.CommentResolved);
        Assert.Null(comment.CommentTextSnippet);
        DocxCommentThreadSummary summary = Assert.Single(result.CommentSummary);
        Assert.Equal("00ABCDEF", summary.ParaId);
        Assert.Equal("00112233", summary.ParentParaId);
        Assert.Equal("00112233", summary.RootParaId);
        Assert.True(summary.IsReply);
        Assert.True(summary.Resolved);
        string rendered = DocxTextRenderer.RenderChanges(result);
        Assert.Contains("root-para-id=00112233", rendered, StringComparison.Ordinal);
        Assert.Contains("comment-root-para-id=00112233", rendered, StringComparison.Ordinal);
        Assert.Contains("comment-is-reply=true", rendered, StringComparison.Ordinal);
        Assert.Contains("resolved=true", rendered, StringComparison.Ordinal);
        Assert.Contains("comment-resolved=true", rendered, StringComparison.Ordinal);
        Assert.DoesNotContain("Private comment text", rendered, StringComparison.Ordinal);
    }

    [Fact]
    public static void ChangesExposeCommentDurableIdsFromCommentsIds()
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
                  <w15:commentEx w15:paraId="00ABCDEF" w15:done="0"/>
                </w15:commentsEx>
                """,
            """
                <w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">
                  <w16cid:commentId w16cid:paraId="00ABCDEF" w16cid:durableId="7F0A11BC"/>
                </w16cid:commentsIds>
                """);
        var editor = new DocxEditor();

        DocxChangesResult result = editor.Changes(stream);

        DocxChangeInfo comment = Assert.Single(result.Changes, change => change.Type == "comment");
        Assert.Equal("00ABCDEF", comment.CommentParaId);
        Assert.Equal("7F0A11BC", comment.CommentDurableId);
        Assert.Null(comment.CommentTextSnippet);
        DocxCommentThreadSummary summary = Assert.Single(result.CommentSummary);
        Assert.Equal("00ABCDEF", summary.ParaId);
        Assert.Equal("7F0A11BC", summary.DurableId);
        string rendered = DocxTextRenderer.RenderChanges(result);
        Assert.Contains("durable-id=7F0A11BC", rendered, StringComparison.Ordinal);
        Assert.Contains("comment-durable-id=7F0A11BC", rendered, StringComparison.Ordinal);
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
        Assert.Equal(DocxTargetStatus.Targetless, deletion.TargetStatus);
        Assert.Equal(DocxTargetSource.None, deletion.TargetSource);
        Assert.True(deletion.TargetReason == DocxTargetReason.BodyLevelMarkup);
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
        Assert.Equal("M.B0001", bookmark.Id.ToWireValue());
        Assert.Equal("ClientName", bookmark.Name);
        Assert.Equal("1", bookmark.OoxmlId);
        Assert.Equal("main", bookmark.Story);
        Assert.Equal("/word/document.xml", bookmark.PartName);
        Assert.Equal("M.P0001", bookmark.StartTargetId?.ToWireValue());
        Assert.Equal("M.P0001", bookmark.EndTargetId?.ToWireValue());
        Assert.True(bookmark.IsComplete);

        DocxContentControlInfo control = Assert.Single(result.ContentControls);
        Assert.Equal("M.CC0001", control.Id.ToWireValue());
        Assert.Equal("plain-text", control.Kind);
        Assert.Equal("77", control.OoxmlId);
        Assert.Equal("client_name", control.Tag);
        Assert.Equal("Client Name", control.Alias);
        Assert.Equal("M.P0001", control.TargetId?.ToWireValue());
        Assert.Equal(4, control.TextLength);

        Assert.Contains("M.B0001 bookmark name=\"ClientName\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("M.CC0001 content-control kind=plain-text", result.Text, StringComparison.Ordinal);
        Assert.Contains("tag=\"client_name\"", result.Text, StringComparison.Ordinal);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1005" && diagnostic.Fallback == "modeled-metadata");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1006" && diagnostic.Fallback == "modeled-metadata");
    }

    [Fact]
    public static void ReadSurfacesDuplicateBookmarkAndContentControlSelectorMetadata()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="1" w:name="Shared"/>
                      <w:r><w:t>First</w:t></w:r>
                      <w:bookmarkEnd w:id="1"/>
                    </w:p>
                    <w:p>
                      <w:bookmarkStart w:id="2" w:name="Shared"/>
                      <w:r><w:t>Second</w:t></w:r>
                      <w:bookmarkEnd w:id="2"/>
                    </w:p>
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:tag w:val="shared_tag"/>
                          <w:alias w:val="Shared Alias"/>
                          <w:text/>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>One</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:tag w:val="shared_tag"/>
                          <w:alias w:val="Shared Alias"/>
                          <w:text/>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>Two</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.True(result.Success);
        foreach (DocxBookmarkInfo bookmark in result.Bookmarks)
        {
            Assert.True(bookmark.IsNameDuplicate);
            Assert.Equal(new[] { "M.B0001", "M.B0002" }, bookmark.DuplicateNameBookmarkIds);
        }

        foreach (DocxContentControlInfo control in result.ContentControls)
        {
            Assert.True(control.IsTagDuplicate);
            Assert.Equal(new[] { "M.CC0001", "M.CC0002" }, control.DuplicateTagControlIds);
            Assert.True(control.IsAliasDuplicate);
            Assert.Equal(new[] { "M.CC0001", "M.CC0002" }, control.DuplicateAliasControlIds);
        }

        Assert.Contains("name-duplicate=true duplicate-name-bookmark-ids=\"M.B0001,M.B0002\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("tag-duplicate=true duplicate-tag-control-ids=\"M.CC0001,M.CC0002\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("alias-duplicate=true duplicate-alias-control-ids=\"M.CC0001,M.CC0002\"", result.Text, StringComparison.Ordinal);

        stream.Position = 0;
        DocxOutlineResult outline = editor.Outline(stream);

        Assert.True(outline.Success);
        DocxOutlineItem dupBookmark = Assert.Single(outline.Items, item => item.TargetId == "M.B0001");
        Assert.True(dupBookmark.IsNameDuplicate);
        Assert.Equal(new[] { "M.B0001", "M.B0002" }, dupBookmark.DuplicateNameBookmarkIds);
        DocxOutlineItem dupControl = Assert.Single(outline.Items, item => item.TargetId == "M.CC0001");
        Assert.True(dupControl.IsTagDuplicate);
        Assert.Equal(new[] { "M.CC0001", "M.CC0002" }, dupControl.DuplicateTagControlIds);
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
                            <w:fullDate w:val="2026-06-12T00:00:00Z"/>
                          </w:date>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>2026-06-12</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:repeatingSection w:sectionTitle="Line items"/>
                          <w:placeholder><w:docPart w:val="DefaultPlaceholder"/></w:placeholder>
                          <w:showingPlcHdr/>
                          <w:dataBinding w:xpath="/root/item" w:storeItemID="{11111111-1111-1111-1111-111111111111}" w:prefixMappings="xmlns:ns='urn:test'"/>
                        </w:sdtPr>
                        <w:sdtContent>
                          <w:sdt>
                            <w:sdtPr><w:repeatingSectionItem/></w:sdtPr>
                            <w:sdtContent>
                              <w:sdt>
                                <w:sdtPr>
                                  <w:text/>
                                  <w:lock w:val="contentLocked"/>
                                </w:sdtPr>
                                <w:sdtContent><w:r><w:t>Nested</w:t></w:r></w:sdtContent>
                              </w:sdt>
                            </w:sdtContent>
                          </w:sdt>
                        </w:sdtContent>
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
        Assert.Equal("2026-06-12T00:00:00Z", date.DateValue);
        DocxContentControlInfo repeating = result.ContentControls.Single(control => control.Kind == "repeating-section");
        Assert.Equal("DefaultPlaceholder", repeating.PlaceholderDocPart);
        Assert.True(repeating.IsShowingPlaceholderText);
        Assert.Equal("/root/item", repeating.DataBindingXPath);
        Assert.Equal("{11111111-1111-1111-1111-111111111111}", repeating.DataBindingStoreItemId);
        Assert.Equal("xmlns:ns='urn:test'", repeating.DataBindingPrefixMappings);
        Assert.Equal("Line items", repeating.RepeatingSectionTitle);
        Assert.Equal(1, repeating.RepeatingSectionItemCount);
        Assert.Null(repeating.ParentContentControlId);
        Assert.Equal(new[] { "M.CC0005" }, repeating.ChildContentControlIds);
        Assert.Equal("unsupported-repeating-section", repeating.SafeEditStatus);
        DocxContentControlInfo nested = result.ContentControls.Single(control => control.Id.ToWireValue() == "M.CC0006");
        Assert.Equal("M.CC0005", nested.ParentContentControlId?.ToWireValue());
        Assert.Equal("locked", nested.SafeEditStatus);
        Assert.Contains("kind=checkbox", result.Text, StringComparison.Ordinal);
        Assert.Contains("checked=true", result.Text, StringComparison.Ordinal);
        Assert.Contains("list-items=2", result.Text, StringComparison.Ordinal);
        Assert.Contains("date-format=\"yyyy-MM-dd\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("date-value=\"2026-06-12T00:00:00Z\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("placeholder-doc-part=\"DefaultPlaceholder\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("data-binding-xpath=\"/root/item\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("repeating-section-items=1", result.Text, StringComparison.Ordinal);
        Assert.Contains("parent-control=M.CC0005", result.Text, StringComparison.Ordinal);
        Assert.Contains("safe-edit=locked", result.Text, StringComparison.Ordinal);
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
        Assert.Equal("M.F0001", simple.Id.ToWireValue());
        Assert.Equal("simple", simple.Kind);
        Assert.Equal("DATE", simple.FieldType);
        Assert.Equal("DATE", simple.Code);
        Assert.Equal("M.P0001", simple.TargetId?.ToWireValue());
        Assert.Equal("June 12, 2026", simple.CachedResultText);
        Assert.Equal(13, simple.ResultTextLength);
        Assert.Equal(0, simple.NestingDepth);
        Assert.Empty(simple.Arguments);
        Assert.Empty(simple.Switches);
        Assert.Empty(simple.BookmarkDependencies);
        Assert.Empty(simple.HyperlinkDependencies);
        Assert.Equal(DocxRefreshPolicy.DateTimeEvaluation, simple.RefreshPolicy);
        Assert.False(simple.CanRefreshDeterministically);
        Assert.Contains("date/time", simple.RefreshReason, StringComparison.Ordinal);
        Assert.Equal("locked", simple.SafeEditStatus);
        Assert.True(simple.IsDirty);
        Assert.True(simple.IsLocked);
        Assert.True(simple.IsComplete);

        DocxFieldInfo complex = result.Fields[1];
        Assert.Equal("M.F0002", complex.Id.ToWireValue());
        Assert.Equal("complex", complex.Kind);
        Assert.Equal("REF", complex.FieldType);
        Assert.Equal(@"REF ClientName \h", complex.Code);
        Assert.Equal("M.P0002", complex.TargetId?.ToWireValue());
        Assert.Equal("Client result", complex.CachedResultText);
        Assert.Equal(13, complex.ResultTextLength);
        Assert.Equal(0, complex.NestingDepth);
        Assert.Equal(new[] { "ClientName" }, complex.Arguments);
        Assert.Equal(new[] { "\\H" }, complex.Switches);
        Assert.Equal(new[] { "ClientName" }, complex.BookmarkDependencies);
        Assert.Empty(complex.HyperlinkDependencies);
        Assert.Equal(DocxRefreshPolicy.SamePartBookmark, complex.RefreshPolicy);
        Assert.True(complex.CanRefreshDeterministically);
        Assert.Equal("flags-only", complex.SafeEditStatus);
        Assert.True(complex.IsDirty);
        Assert.Null(complex.IsLocked);
        Assert.True(complex.IsComplete);

        Assert.Contains("M.F0001 field kind=simple type=DATE", result.Text, StringComparison.Ordinal);
        Assert.Contains("cached-result=\"June 12, 2026\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("safe-edit=locked", result.Text, StringComparison.Ordinal);
        Assert.Contains(@"code=""REF ClientName \\h""", result.Text, StringComparison.Ordinal);
        Assert.Contains("arguments=\"ClientName\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("switches=\"\\\\H\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("refresh-policy=same-part-bookmark", result.Text, StringComparison.Ordinal);
        Assert.Contains("bookmark-dependencies=\"ClientName\"", result.Text, StringComparison.Ordinal);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1003" && diagnostic.Fallback == "modeled-metadata");
    }

    [Fact]
    public static void ReadModelsNestedComplexFieldMetadataAndValidatesBalance()
    {
        const string nestedFieldBody = """
                    <w:p>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> IF </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:t xml:space="preserve">Prefix </w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> DATE </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:t>June 13</w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                      <w:r><w:t xml:space="preserve"> Suffix</w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                    </w:p>
            """;
        var editor = new DocxEditor();

        using MemoryStream readStream = CreateDocxWithBody(nestedFieldBody);
        DocxReadResult read = editor.Read(readStream);

        Assert.True(read.Success);
        Assert.Equal(2, read.Fields.Count);
        DocxFieldInfo inner = Assert.Single(read.Fields, field => field.Code == "DATE");
        Assert.Equal("complex", inner.Kind);
        Assert.Equal("DATE", inner.FieldType);
        Assert.Equal("M.P0001", inner.TargetId?.ToWireValue());
        Assert.Equal("June 13", inner.CachedResultText);
        Assert.Equal(7, inner.ResultTextLength);
        Assert.Equal(1, inner.NestingDepth);
        Assert.True(inner.IsComplete);

        DocxFieldInfo outer = Assert.Single(read.Fields, field => field.Code == "IF");
        Assert.Equal("complex", outer.Kind);
        Assert.Equal("IF", outer.FieldType);
        Assert.Equal("M.P0001", outer.TargetId?.ToWireValue());
        Assert.Equal("Prefix June 13 Suffix", outer.CachedResultText);
        Assert.Equal(21, outer.ResultTextLength);
        Assert.Equal(0, outer.NestingDepth);
        Assert.True(outer.IsComplete);

        using MemoryStream validateStream = CreateDocxWithBody(nestedFieldBody);
        DocxValidateResult validate = editor.Validate(validateStream);

        Assert.True(validate.Success);
        Assert.DoesNotContain(validate.Diagnostics, diagnostic => diagnostic.Code is "E9104" or "E9112");
    }


    [Fact]
    public static void ReadModelsHyperlinkFieldDependencies()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> HYPERLINK "https://example.test/report" \l "Section1" </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:t>Report link</w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                    </w:p>
            """);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        DocxFieldInfo field = Assert.Single(result.Fields);
        Assert.Equal("HYPERLINK", field.FieldType);
        Assert.Equal(new[] { "Section1" }, field.BookmarkDependencies);
        Assert.Equal(new[] { "https://example.test/report" }, field.HyperlinkDependencies);
        Assert.Equal("flags-only", field.SafeEditStatus);
        Assert.Contains("hyperlink-dependencies=\"https://example.test/report\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("bookmark-dependencies=\"Section1\"", result.Text, StringComparison.Ordinal);
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
                """, null);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.True(result.Success);
        Assert.Equal(3, result.Hyperlinks.Count);
        DocxHyperlinkInfo external = result.Hyperlinks[0];
        Assert.Equal("M.L0001", external.Id.ToWireValue());
        Assert.Equal("M.P0001", external.TargetId?.ToWireValue());
        Assert.Equal("rLink", external.RelationshipId);
        Assert.Equal("/word/_rels/document.xml.rels", external.RelationshipPartName);
        Assert.Equal("External", external.RelationshipTargetMode);
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
        Assert.Null(internalLink.RelationshipPartName);
        Assert.Null(internalLink.RelationshipTargetMode);

        DocxHyperlinkInfo broken = result.Hyperlinks[2];
        Assert.Equal("rMissing", broken.RelationshipId);
        Assert.Equal("/word/_rels/document.xml.rels", broken.RelationshipPartName);
        Assert.Null(broken.RelationshipTargetMode);
        Assert.True(broken.IsBroken);
        Assert.Null(broken.Uri);

        Assert.Contains("M.L0001 hyperlink", result.Text, StringComparison.Ordinal);
        Assert.Contains("uri=\"https://example.test/report\"", result.Text, StringComparison.Ordinal);
        Assert.Contains("relationship-part=/word/_rels/document.xml.rels", result.Text, StringComparison.Ordinal);
        Assert.Contains("target-mode=External", result.Text, StringComparison.Ordinal);
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
                      <w:hyperlink r:id="rMalformed"><w:r><w:t>Malformed</w:t></w:r></w:hyperlink>
                      <w:r><w:t xml:space="preserve"> </w:t></w:r>
                      <w:hyperlink r:id="rFile"><w:r><w:t>File</w:t></w:r></w:hyperlink>
                      <w:r><w:t xml:space="preserve"> </w:t></w:r>
                      <w:hyperlink r:id="rMailto"><w:r><w:t>Mail</w:t></w:r></w:hyperlink>
                      <w:r><w:t xml:space="preserve"> </w:t></w:r>
                      <w:hyperlink r:id="rPart"><w:r><w:t>Part</w:t></w:r></w:hyperlink>
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
                  <Relationship Id="rMalformed" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://[::1" TargetMode="External"/>
                  <Relationship Id="rFile" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="file:///C:/secret/report.docx" TargetMode="External"/>
                  <Relationship Id="rMailto" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="mailto:reviewer@example.test" TargetMode="External"/>
                  <Relationship Id="rPart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="document.xml"/>
                </Relationships>
                """, null);
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
        Assert.Equal("relative-uri", relativeLink.UriValidationReason);

        DocxHyperlinkInfo malformedLink = result.Hyperlinks.Single(link => link.RelationshipId == "rMalformed");
        Assert.Null(malformedLink.UriScheme);
        Assert.False(malformedLink.IsUriValid);
        Assert.Equal("malformed-uri", malformedLink.UriValidationReason);

        DocxHyperlinkInfo fileLink = result.Hyperlinks.Single(link => link.RelationshipId == "rFile");
        Assert.Equal("file", fileLink.UriScheme);
        Assert.False(fileLink.IsUriValid);
        Assert.Equal("unsupported-uri-scheme", fileLink.UriValidationReason);

        DocxHyperlinkInfo mailtoLink = result.Hyperlinks.Single(link => link.RelationshipId == "rMailto");
        Assert.Equal("mailto", mailtoLink.UriScheme);
        Assert.True(mailtoLink.IsUriValid);
        Assert.Null(mailtoLink.UriValidationReason);

        DocxHyperlinkInfo partLink = result.Hyperlinks.Single(link => link.RelationshipId == "rPart");
        Assert.False(partLink.IsExternal);
        Assert.Equal("/word/document.xml", partLink.TargetPartName);
        Assert.Null(partLink.UriValidationReason);

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
        Assert.Contains("uri-reason=relative-uri", result.Text, StringComparison.Ordinal);
        Assert.Contains("uri-reason=malformed-uri", result.Text, StringComparison.Ordinal);
        Assert.Contains("anchor-missing=true", result.Text, StringComparison.Ordinal);
        Assert.Contains("anchor-duplicate=true", result.Text, StringComparison.Ordinal);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1016" && diagnostic.Fallback == "invalid-uri");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1017" && diagnostic.Fallback == "missing-anchor");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1018" && diagnostic.Fallback == "duplicate-anchor");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1023" && diagnostic.Fallback == "unsupported-internal-part-link");
    }








    [Fact]
    public static void DocxChangeInfoUsesPropertyBasedPublicShape()
    {
        Assert.Contains(typeof(DocxChangeInfo).GetConstructors(), constructor => constructor.GetParameters().Length == 0);
        Assert.DoesNotContain(typeof(DocxChangeInfo).GetConstructors(), constructor => constructor.GetParameters().Length > 0);

        var change = new DocxChangeInfo
        {
            Id = new DocxChangeId('M', 0, 1),
            Type = "inserted-run",
            Story = "main",
            PartName = "/word/document.xml",
            ParentType = "paragraph",
            TextLength = 8,
            ChildElementCount = 1
        };

        string serialized = JsonSerializer.Serialize(change);
        Assert.Contains("\"Id\":\"M.CH0001\"", serialized, StringComparison.Ordinal);
        Assert.Contains("\"Type\":\"inserted-run\"", serialized, StringComparison.Ordinal);
        Assert.Contains("\"ParentType\":\"paragraph\"", serialized, StringComparison.Ordinal);
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
        Assert.Equal(DocxTargetSource.AdjacentRange, moveStart.TargetSource);
        Assert.Equal(DocxTargetSource.AdjacentRange, moveEnd.TargetSource);
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











}
