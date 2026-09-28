using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml.Linq;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

public static class PatchTableTests
{

    [Fact]
    public static void ApplySetCellCanEditHeaderTableAndUseItAsBlockAnchor()
    {
        using MemoryStream input = CreateDocxWithHeaderFooterContent(
            """
                  <w:tbl>
                    <w:tr><w:tc><w:p><w:r><w:t>Old</w:t></w:r></w:p></w:tc></w:tr>
                  </w:tbl>
            """,
            """
                  <w:p><w:r><w:t>Footer text</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-cell
            target H001.T0001.R01.C01
            expect-text Old
            text New
            end

            op insert-after
            target H001.T0001
            text After header table
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output, new DocxReadOptions { IncludeHeadersFooters = true });
        DocxTableInfo table = Assert.Single(read.Tables, table => table.Story == "header[1]");
        Assert.Equal("New", Assert.Single(table.Cells).Text);
        Assert.Contains(read.Paragraphs, paragraph => paragraph.Id.ToWireValue() == "H001.P0001" && paragraph.Text == "After header table");
    }

    [Fact]
    public static void ApplyRowOperationsCanEditFooterTable()
    {
        using MemoryStream input = CreateDocxWithHeaderFooterContent(
            """
                  <w:p><w:r><w:t>Header text</w:t></w:r></w:p>
            """,
            """
                  <w:tbl>
                    <w:tr><w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc></w:tr>
                    <w:tr><w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc></w:tr>
                  </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op append-row
            target F001.T0001
            cell East
            end

            op insert-row-before
            target F001.T0001.R02
            cell Central
            end

            op delete-row
            target F001.T0001.R01
            expect-contains North
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { IncludeHeadersFooters = true }).Tables, table => table.Story == "footer[1]");
        Assert.Equal(3, table.RowCount);
        Assert.Equal("Central", table.Cells.Single(cell => cell.RowIndex == 1).Text);
        Assert.Equal("South", table.Cells.Single(cell => cell.RowIndex == 2).Text);
        Assert.Equal("East", table.Cells.Single(cell => cell.RowIndex == 3).Text);
    }

    [Fact]
    public static void ApplySetCellPreservesCellPropertiesAndDoesNotMutateInput()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>
                          <w:p>
                            <w:pPr><w:pStyle w:val="TableBody"/></w:pPr>
                            <w:r><w:t>Old</w:t></w:r>
                          </w:p>
                        </w:tc>
                        <w:tc><w:p><w:r><w:t>Other</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-cell
            target M.T0001.R01.C01
            expect-text Old
            text New
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        input.Position = 0;
        output.Position = 0;
        Assert.Equal("Old", new DocxEditor().Read(input).Tables[0].Cells[0].Text);
        Assert.Equal("New", new DocxEditor().Read(output).Tables[0].Cells[0].Text);

        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:tcW", xml, StringComparison.Ordinal);
        Assert.Contains("w:val=\"TableBody\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetCellUsesVisualColumnAfterGridBefore()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:trPr><w:gridBefore w:val="1"/></w:trPr>
                        <w:tc><w:p><w:r><w:t>Old</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-cell
            target M.T0001.R01.C02
            expect-text Old
            text New
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        DocxPatchAffectedTarget affected = Assert.Single(Assert.Single(result.Operations).AffectedTargets);
        Assert.Equal(2, affected.ColumnIndex);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        DocxTableCellInfo cell = Assert.Single(table.Cells);
        Assert.Equal("M.T0001.R01.C02", cell.Id.ToWireValue());
        Assert.Equal("New", cell.Text);
    }

    [Fact]
    public static void ApplySetCellCanTargetHorizontalMergeGroup()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
                          <w:p><w:r><w:t>Old merged</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-cell
            target M.T0001.MG0001
            expect-text Old merged
            text New merged
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        DocxPatchAffectedTarget affected = Assert.Single(Assert.Single(result.Operations).AffectedTargets);
        Assert.Equal("M.T0001.MG0001", affected.Id.ToWireValue());
        Assert.Equal(1, affected.ColumnIndex);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        DocxTableCellInfo cell = Assert.Single(table.Cells);
        Assert.Equal("M.T0001.MG0001", cell.MergeGroupId?.ToWireValue());
        Assert.Equal(2, cell.VisualColumnEndIndex);
        Assert.Equal("New merged", cell.Text);
        output.Position = 0;
        Assert.Contains("<w:gridSpan w:val=\"2\"", ReadDocumentXml(output), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetCellCanTargetVisualColumnInsideHorizontalSpan()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
                          <w:p><w:r><w:t>Old merged</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-cell
            target M.T0001.R01.C02
            expect-text Old merged
            text New merged
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        DocxPatchAffectedTarget affected = Assert.Single(Assert.Single(result.Operations).AffectedTargets);
        Assert.Equal(2, affected.ColumnIndex);
        output.Position = 0;
        Assert.Equal("New merged", Assert.Single(Assert.Single(new DocxEditor().Read(output).Tables).Cells).Text);
    }

    [Fact]
    public static void CheckSetCellRejectsComplexCellWithoutForce()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:p><w:r><w:t>One</w:t></w:r></w:p>
                          <w:p><w:r><w:t>Two</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-cell
            target M.T0001.R01.C01
            text Replacement
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E4302");
    }

    [Fact]
    public static void ApplyAppendRowClonesRowShapeAndCellProperties()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>
                      <w:tr>
                        <w:trPr><w:trHeight w:val="240"/></w:trPr>
                        <w:tc>
                          <w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>
                          <w:p>
                            <w:pPr><w:pStyle w:val="TableRegion"/></w:pPr>
                            <w:r><w:t>North</w:t></w:r>
                          </w:p>
                        </w:tc>
                        <w:tc>
                          <w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>
                          <w:p><w:r><w:t>Revenue</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op append-row
            target M.T0001
            cell South
            cell Profit
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal(2, table.RowCount);
        Assert.Equal("South", table.Cells.Single(cell => cell.RowIndex == 2 && cell.ColumnIndex == 1).Text);
        Assert.Equal("Profit", table.Cells.Single(cell => cell.RowIndex == 2 && cell.ColumnIndex == 2).Text);

        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Equal(2, CountOccurrences(xml, "<w:trHeight"));
        Assert.Equal(4, CountOccurrences(xml, "<w:tcW"));
        Assert.Equal(2, CountOccurrences(xml, "w:val=\"TableRegion\""));
    }

    [Fact]
    public static void ApplySetTableStyleAndRowHeaderFlag()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>
            """, """
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-table-style
            target M.T0001
            style TableGrid
            end

            op set-row-header
            target M.T0001.R01
            expect-header false
            header true
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal("TableGrid", table.StyleId);
        Assert.True(table.HasHeaderRow);
        Assert.True(table.Rows[0].IsHeader);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:tblStyle w:val=\"TableGrid\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:tblHeader", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetCellShadingSetsAndClearsFill()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>
                      <w:tr>
                        <w:tc>
                          <w:tcPr>
                            <w:gridSpan w:val="2"/>
                            <w:shd w:val="clear" w:fill="FF0000"/>
                          </w:tcPr>
                          <w:p><w:r><w:t>Wide</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:shd w:val="clear" w:fill="00FF00"/></w:tcPr>
                          <w:p><w:r><w:t>Clear me</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc><w:p><w:r><w:t>Plain</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-cell-shading
            target M.T0001.MG0001
            expect-fill ff0000
            fill a1b2c3
            end

            op set-cell-shading
            target M.T0001.R02.C01
            expect-fill 00FF00
            clear true
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        Assert.Equal("M.T0001.MG0001", result.Operations[0].Target);
        Assert.Single(result.Operations[0].AffectedTargets);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:shd w:val=\"clear\" w:fill=\"A1B2C3\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("w:fill=\"00FF00\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:gridSpan w:val=\"2\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetTableMetadataUpdatesAndClearsCaptionDescription()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tblPr>
                        <w:tblCaption w:val="Old caption"/>
                        <w:tblDescription w:val="Old description"/>
                      </w:tblPr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-table-metadata
            target M.T0001
            expect-caption Old caption
            expect-description Old description
            caption Revenue table
            description <<<
            >>>
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal("Revenue table", table.Caption);
        Assert.Null(table.Description);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:tblCaption w:val=\"Revenue table\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("tblDescription", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckTablePropertyGuardsRejectMismatches()
    {
        // Operations after the first failure are skipped rather than simulated,
        // so every guard mismatch gets its own patch here.
        const string bodyXml = """
                    <w:tbl>
                      <w:tblPr>
                        <w:tblStyle w:val="ExistingStyle"/>
                        <w:tblCaption w:val="Existing caption"/>
                      </w:tblPr>
                      <w:tr>
                        <w:trPr><w:tblHeader/></w:trPr>
                        <w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """;
        using MemoryStream shadingInput = CreateDocxWithBody(bodyXml);
        using var shadingPatch = new StringReader("""
            docxpatch 1

            op set-cell-shading
            target M.T0001.R01.C01
            expect-fill FFFFFF
            fill A1B2C3
            end
            """);
        DocxCheckResult shading = new DocxEditor().Check(shadingInput, shadingPatch);
        Assert.False(shading.Success);
        Assert.Contains(shading.Diagnostics, diagnostic => diagnostic.Code == "E3201" && diagnostic.TargetId == "M.T0001.R01.C01");
        using MemoryStream styleInput = CreateDocxWithBody(bodyXml);
        using var stylePatch = new StringReader("""
            docxpatch 1

            op set-table-style
            target M.T0001
            expect-style OtherStyle
            style TableGrid
            end
            """);
        DocxCheckResult style = new DocxEditor().Check(styleInput, stylePatch);
        Assert.False(style.Success);
        Assert.Contains(style.Diagnostics, diagnostic => diagnostic.Code == "E3201" && diagnostic.TargetId == "M.T0001");
        using MemoryStream headerInput = CreateDocxWithBody(bodyXml);
        using var headerPatch = new StringReader("""
            docxpatch 1

            op set-row-header
            target M.T0001.R01
            expect-header false
            header false
            end
            """);
        DocxCheckResult header = new DocxEditor().Check(headerInput, headerPatch);
        Assert.False(header.Success);
        Assert.Contains(header.Diagnostics, diagnostic => diagnostic.Code == "E3201" && diagnostic.TargetId == "M.T0001.R01");
        using MemoryStream captionInput = CreateDocxWithBody(bodyXml);
        using var captionPatch = new StringReader("""
            docxpatch 1

            op set-table-metadata
            target M.T0001
            expect-caption Other caption
            caption Updated caption
            end
            """);
        DocxCheckResult caption = new DocxEditor().Check(captionInput, captionPatch);
        Assert.False(caption.Success);
        Assert.Contains(caption.Diagnostics, diagnostic =>
            diagnostic.Code == "E3201" &&
            diagnostic.TargetId == "M.T0001" &&
            diagnostic.Message.Contains("caption", StringComparison.Ordinal));
    }

    [Fact]
    public static void CheckAppendRowRejectsCellCountMismatch()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op append-row
            target M.T0001
            cell South
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E4303");
    }

    [Fact]
    public static void CheckAppendRowRejectsFailedShapeGuards()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op append-row
            target M.T0001
            expect-row-count 2
            expect-column-count 2
            cell South
            cell Profit
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E3201" && diagnostic.TargetId == "M.T0001");
    }

    [Fact]
    public static void ApplyInsertRowBeforeClonesTargetRowShape()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:trPr><w:trHeight w:val="480"/></w:trPr>
                        <w:tc>
                          <w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>
                          <w:p><w:r><w:t>South</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                          <w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>
                          <w:p><w:r><w:t>Profit</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-row-before
            target M.T0001.R02
            cell East
            cell Margin
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal(3, table.RowCount);
        Assert.Equal("East", table.Cells.Single(cell => cell.RowIndex == 2 && cell.ColumnIndex == 1).Text);
        Assert.Equal("South", table.Cells.Single(cell => cell.RowIndex == 3 && cell.ColumnIndex == 1).Text);

        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Equal(2, CountOccurrences(xml, "<w:trHeight"));
        Assert.Equal(4, CountOccurrences(xml, "<w:tcW"));
    }

    [Fact]
    public static void ApplyInsertRowAfterPlacesRowAfterTarget()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-row-after
            target M.T0001.R01
            cell East
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal("North", table.Cells.Single(cell => cell.RowIndex == 1 && cell.ColumnIndex == 1).Text);
        Assert.Equal("East", table.Cells.Single(cell => cell.RowIndex == 2 && cell.ColumnIndex == 1).Text);
        Assert.Equal("South", table.Cells.Single(cell => cell.RowIndex == 3 && cell.ColumnIndex == 1).Text);
    }

    [Fact]
    public static void ApplyDeleteRowRemovesSelectedRow()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>East</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-row
            target M.T0001.R02
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal(2, table.RowCount);
        Assert.Equal("North", table.Cells.Single(cell => cell.RowIndex == 1 && cell.ColumnIndex == 1).Text);
        Assert.Equal("East", table.Cells.Single(cell => cell.RowIndex == 2 && cell.ColumnIndex == 1).Text);
    }

    [Theory]
    [InlineData("append-row", "M.T0001")]
    [InlineData("insert-row-before", "M.T0001.R02")]
    [InlineData("insert-row-after", "M.T0001.R01")]
    public static void CheckRowInsertionRejectsVerticalMergeBoundaries(string operationName, string target)
    {
        string cellFields = """
            cell Inserted
            cell Value
            """;
        using MemoryStream input = CreateDocxWithBody("""
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
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            target {target}
            {cellFields}
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E4301" &&
            diagnostic.Message.Contains("vertical merge", StringComparison.Ordinal));
    }

    [Fact]
    public static void ApplyDeleteRowPromotesVerticalMergeContinuation()
    {
        using MemoryStream input = CreateDocxWithBody("""
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
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>East</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Cost</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-row
            target M.T0001.R01
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal(2, table.RowCount);
        DocxTableCellInfo promoted = table.Cells.Single(cell => cell.RowIndex == 1 && cell.ColumnIndex == 1);
        Assert.Equal(DocxVerticalMerge.Restart, promoted.VerticalMerge);
        Assert.Equal("M.T0001.R01.C01", promoted.VerticalMergeRootCellId?.ToWireValue());
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:vMerge w:val=\"restart\"", xml, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("horizontal-merge", """
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
                          <w:p><w:r><w:t>North revenue</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Profit</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
        """)]
    [InlineData("grid-before", """
                    <w:tbl>
                      <w:tr>
                        <w:trPr><w:gridBefore w:val="1"/></w:trPr>
                        <w:tc><w:p><w:r><w:t>Indented</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Profit</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
        """)]
    [InlineData("grid-after", """
                    <w:tbl>
                      <w:tr>
                        <w:trPr><w:gridAfter w:val="1"/></w:trPr>
                        <w:tc><w:p><w:r><w:t>Trailing omitted</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Profit</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
        """)]
    public static void ApplyAppendRowAllowsSafeVisualGridTables(string caseName, string tableXml)
    {
        _ = caseName;
        using MemoryStream input = CreateDocxWithBody(tableXml);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op append-row
            target M.T0001
            cell East
            cell Margin
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal(3, table.RowCount);
        Assert.Contains(table.Cells, cell => cell.RowIndex == 3 && cell.ColumnIndex == 1 && cell.Text == "East");
        Assert.Contains(table.Cells, cell => cell.RowIndex == 3 && cell.Text == "Margin");
    }

    [Fact]
    public static void ApplyInsertRowBeforePreservesGridBeforeShape()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:trPr><w:gridBefore w:val="1"/></w:trPr>
                        <w:tc><w:p><w:r><w:t>Indented</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Profit</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-row-before
            target M.T0001.R01
            cell Inserted
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        DocxTableCellInfo inserted = table.Cells.Single(cell => cell.RowIndex == 1);
        Assert.Equal("M.T0001.R01.C02", inserted.Id.ToWireValue());
        Assert.Equal(2, inserted.ColumnIndex);
        Assert.Equal("Inserted", inserted.Text);
        output.Position = 0;
        Assert.Equal(2, CountOccurrences(ReadDocumentXml(output), "<w:gridBefore"));
    }

    [Fact]
    public static void ApplyInsertRowBeforePreservesHorizontalSpanShape()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
                          <w:p><w:r><w:t>Total</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Profit</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-row-before
            target M.T0001.R01
            cell Inserted total
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        DocxTableCellInfo inserted = table.Cells.Single(cell => cell.RowIndex == 1);
        Assert.Equal(1, inserted.ColumnIndex);
        Assert.Equal(2, inserted.ColumnSpan);
        Assert.Equal(2, inserted.VisualColumnEndIndex);
        Assert.Equal("Inserted total", inserted.Text);
        output.Position = 0;
        Assert.Equal(2, CountOccurrences(ReadDocumentXml(output), "<w:gridSpan w:val=\"2\""));
    }

    [Theory]
    [InlineData("append-column", """
        target M.T0001
        cell East
        cell West
        """)]
    [InlineData("insert-column-before", """
        target M.T0001
        column 1
        cell East
        cell West
        """)]
    [InlineData("insert-column-after", """
        target M.T0001
        column 2
        cell East
        cell West
        """)]
    [InlineData("delete-column", """
        target M.T0001
        column 1
        expect-contains North
        """)]
    public static void CheckColumnOperationsFailWithExplicitUnsupportedDiagnostic(string operationName, string operationFields)
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Profit</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            {operationFields}
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E4316" &&
            diagnostic.TargetId == "M.T0001" &&
            diagnostic.Message.Contains("column edits", StringComparison.Ordinal));
        Assert.False(Assert.Single(result.Operations).Success);
    }

    [Fact]
    public static void CheckAppendRowReportsAffectedRowAndCells()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op append-row
            target M.T0001
            cell South
            cell Profit
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success);
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        DocxPatchAffectedTarget row = Assert.Single(report.AffectedTargets, target => target.Kind == "row");
        Assert.Equal("M.T0001.R02", row.Id.ToWireValue());
        Assert.Equal("append", row.Action);
        Assert.Equal(1, row.RowCountBefore);
        Assert.Equal(2, row.RowCountAfter);
        Assert.Equal(2, row.CellCount);
        Assert.Contains(report.AffectedTargets, target => target.Id.ToWireValue() == "M.T0001.R02.C01" && target.Kind == "cell" && target.ColumnIndex == 1);
        Assert.Contains(report.AffectedTargets, target => target.Id.ToWireValue() == "M.T0001.R02.C02" && target.Kind == "cell" && target.ColumnIndex == 2);
    }

    [Fact]
    public static void CheckInsertRowReportsVisualGridAffectedCells()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:trPr><w:gridBefore w:val="1"/></w:trPr>
                        <w:tc><w:p><w:r><w:t>Indented</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Profit</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op insert-row-before
            target M.T0001.R01
            cell Inserted
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success);
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        DocxPatchAffectedTarget row = Assert.Single(report.AffectedTargets, target => target.Kind == "row");
        Assert.Equal(1, row.GridBefore);
        Assert.Equal(0, row.GridAfter);
        DocxPatchAffectedTarget cell = Assert.Single(report.AffectedTargets, target => target.Kind == "cell");
        Assert.Equal("M.T0001.R01.C02", cell.Id.ToWireValue());
        Assert.Equal(2, cell.ColumnIndex);
        Assert.Equal(2, cell.VisualColumnEndIndex);
    }

    [Fact]
    public static void CheckCellPropertyReportsMergeGroupAndNestedTablePath()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
                          <w:p><w:r><w:t>Wide</w:t></w:r></w:p>
                          <w:tbl>
                            <w:tr><w:tc><w:p><w:r><w:t>Nested</w:t></w:r></w:p></w:tc></w:tr>
                          </w:tbl>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-cell-shading
            target M.T0001.MG0001
            fill A1B2C3
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success);
        DocxPatchAffectedTarget affected = Assert.Single(Assert.Single(result.Operations).AffectedTargets);
        Assert.Equal("M.T0001.MG0001", affected.Id.ToWireValue());
        Assert.Equal(1, affected.ColumnIndex);
        Assert.Equal(2, affected.VisualColumnEndIndex);
        Assert.Equal("M.T0001.MG0001", affected.MergeGroupId?.ToWireValue());
        Assert.Equal("M.T0001.R01.C01.T0001", affected.NestedTablePath);
    }

    [Fact]
    public static void ApplyDeleteRowReportsAffectedRowAndCells()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Profit</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-row
            target M.T0001.R02
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        DocxPatchAffectedTarget row = Assert.Single(report.AffectedTargets, target => target.Kind == "row");
        Assert.Equal("M.T0001.R02", row.Id.ToWireValue());
        Assert.Equal("delete", row.Action);
        Assert.Equal(2, row.RowCountBefore);
        Assert.Equal(1, row.RowCountAfter);
        Assert.Equal(2, row.CellCount);
        Assert.Contains(report.AffectedTargets, target => target.Id.ToWireValue() == "M.T0001.R02.C01" && target.Kind == "cell" && target.Action == "delete");
        Assert.Contains(report.AffectedTargets, target => target.Id.ToWireValue() == "M.T0001.R02.C02" && target.Kind == "cell" && target.Action == "delete");
    }

    [Fact]
    public static void OperationSummaryRendersAffectedTableTargets()
    {
        var operations = new[]
        {
            new DocxPatchOperationReport(1, "append-row", "M.T0001", true, [])
            {
                AffectedTargets =
                [
                    new(new DocxTargetId('M', 0, DocxTargetKind.Row, 1, 2, 0), "row", "append")
                    {
                        ParentId = new DocxTargetId('M', 0, DocxTargetKind.Table, 1, 0, 0),
                        RowIndex = 2,
                        RowCountBefore = 1,
                        RowCountAfter = 2,
                        ColumnCount = 2,
                        CellCount = 2
                    },
                    new(new DocxTargetId('M', 0, DocxTargetKind.Cell, 1, 2, 1), "cell", "append")
                    {
                        ParentId = new DocxTargetId('M', 0, DocxTargetKind.Row, 1, 2, 0),
                        RowIndex = 2,
                        ColumnIndex = 1,
                        VisualColumnEndIndex = 2,
                        GridBefore = 1,
                        MergeGroupId = new DocxTargetId('M', 0, DocxTargetKind.MergeGroup, 1, 1, 0),
                        NestedTablePath = "M.T0001.R02.C01.T0001",
                        RowCountBefore = 1,
                        RowCountAfter = 2,
                        ColumnCount = 2
                    }
                ]
            }
        };

        string text = DocxTextRenderer.RenderOperationSummary(operations);

        Assert.Contains("operation index=1 name=append-row target=M.T0001 success=True", text, StringComparison.Ordinal);
        Assert.Contains("affected id=M.T0001.R02 kind=row action=append parent=M.T0001 row=2 rows-before=1 rows-after=2 columns=2 cells=2", text, StringComparison.Ordinal);
        Assert.Contains("affected id=M.T0001.R02.C01 kind=cell action=append parent=M.T0001.R02 row=2 column=1 visual-column-end=2 grid-before=1 merge-group=M.T0001.MG0001 nested-table-path=M.T0001.R02.C01.T0001 rows-before=1 rows-after=2 columns=2", text, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckDeleteRowRejectsLastRow()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Only</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op delete-row
            target M.T0001.R01
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E4304");
    }

    [Fact]
    public static void CheckDeleteRowRejectsFailedExpectContainsGuard()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op delete-row
            target M.T0001.R01
            expect-contains South
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E3201" && diagnostic.TargetId == "M.T0001.R01");
        Assert.False(Assert.Single(result.Operations).Success);
    }

    [Fact]
    public static void ApplySetCellWithEmptyTextClearsCell()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-cell
            target M.T0001.R01.C01
            expect-text North
            text ""
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal("", table.Cells.Single(cell => cell.Id.ToWireValue() == "M.T0001.R01.C01").Text);
        Assert.Equal("South", table.Cells.Single(cell => cell.Id.ToWireValue() == "M.T0001.R01.C02").Text);
    }


    [Fact]
    public static void ApplySetCellIdenticalTextIsNoOp()
    {
        using MemoryStream probe = CreateDocxWithSimpleTwoByTwoTable();
        string before = ReadDocumentXml(probe);
        using MemoryStream input = CreateDocxWithSimpleTwoByTwoTable();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-cell
            target M.T0001.R01.C01
            text North
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Contains(result.Diagnostics, static d => d.Code == "I0001");
        output.Position = 0;
        Assert.Equal(before, ReadDocumentXml(output));
    }

}
