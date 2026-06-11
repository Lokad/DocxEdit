using System.IO.Compression;
using System.Text;

namespace DocxEdit.Tests;

public static class PatchApplyTests
{
    [Fact]
    public static void CheckReplaceTextSucceedsForMatchingParagraph()
    {
        using MemoryStream input = CreateDocx("Revenue increased by 8.4%.");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            expect-text <<<
            Revenue increased by 8.4%.
            >>>
            find <<<
            8.4%
            >>>
            with <<<
            9.1%
            >>>
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success);
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.True(report.Success);
    }

    [Fact]
    public static void CheckReplaceTextFailsWhenGuardDoesNotMatch()
    {
        using MemoryStream input = CreateDocx("Revenue increased by 9.1%.");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            expect-text <<<
            Revenue increased by 8.4%.
            >>>
            find <<<
            8.4%
            >>>
            with <<<
            9.1%
            >>>
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E3201");
        Assert.False(Assert.Single(result.Operations).Success);
    }

    [Fact]
    public static void ApplyReplaceTextWritesNewDocumentWithoutMutatingInput()
    {
        using MemoryStream input = CreateDocx("Revenue increased by 8.4%.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            expect-text <<<
            Revenue increased by 8.4%.
            >>>
            find <<<
            8.4%
            >>>
            with <<<
            9.1%
            >>>
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        input.Position = 0;
        output.Position = 0;
        Assert.Equal("Revenue increased by 8.4%.", Assert.Single(new DocxEditor().Read(input).Paragraphs).Text);
        Assert.Equal("Revenue increased by 9.1%.", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void ApplyReplaceTextWorksAcrossAdjacentRuns()
    {
        using MemoryStream input = CreateDocxWithRuns("Revenue ", "increased");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            expect-text <<<
            Revenue increased
            >>>
            find <<<
            Revenue increased
            >>>
            with <<<
            Revenue rose
            >>>
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("Revenue rose", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void ApplyReplaceTextFailsWhenFindTextIsMissing()
    {
        using MemoryStream input = CreateDocx("Revenue increased by 8.4%.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find missing
            with value
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E4203");
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public static void ApplySetCellPreservesCellPropertiesAndDoesNotMutateInput()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>
                          <w:p><w:r><w:t>Old</w:t></w:r></w:p>
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
        Assert.Contains("<w:tcW", ReadDocumentXml(output), StringComparison.Ordinal);
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
                          <w:p><w:r><w:t>North</w:t></w:r></w:p>
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

    private static MemoryStream CreateDocx(string paragraphText)
    {
        return CreateDocxWithRuns(paragraphText);
    }

    private static MemoryStream CreateDocxWithRuns(params string[] runTexts)
    {
        var stream = new MemoryStream();
        string runs = string.Concat(runTexts.Select(text => $"<w:r><w:t>{text}</w:t></w:r>"));
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
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
                """);
            AddEntry(archive, "word/document.xml", $$"""
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>{{runs}}</w:p>
                  </w:body>
                </w:document>
                """);
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithBody(string bodyXml)
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
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
                """);
            string documentXml = """
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>

                """ + bodyXml + """

                  </w:body>
                </w:document>
                """;
            AddEntry(archive, "word/document.xml", documentXml);
        }

        stream.Position = 0;
        return stream;
    }

    private static string ReadDocumentXml(Stream docx)
    {
        using var archive = new ZipArchive(docx, ZipArchiveMode.Read, leaveOpen: true);
        ZipArchiveEntry entry = archive.GetEntry("word/document.xml")
            ?? throw new InvalidDataException("Missing word/document.xml.");
        using Stream stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    }

    private static int CountOccurrences(string text, string value)
    {
        int count = 0;
        int index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += value.Length;
        }

        return count;
    }

    private static void AddEntry(ZipArchive archive, string name, string text)
    {
        ZipArchiveEntry entry = archive.CreateEntry(name);
        using Stream stream = entry.Open();
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        stream.Write(bytes, 0, bytes.Length);
    }
}
