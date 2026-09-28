using System.IO.Compression;
using System.Text;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// PLAN B01: discovered target IDs must identify the same elements when edited.
// Physical IDs include revision-wrapped blocks; resolvers use the same order.
public static class TargetIdentityTests
{
    [Fact]
    public static void ReadAssignsInsertedParagraphPhysicalId()
    {
        using MemoryStream input = CreateDocxWithBody(
            "<w:p><w:r><w:t>First</w:t></w:r></w:p>" +
            "<w:ins w:id=\"1\" w:author=\"Alice\" w:date=\"2026-01-01T00:00:00Z\"><w:p><w:r><w:t>Inserted</w:t></w:r></w:p></w:ins>" +
            "<w:p><w:r><w:t>Last</w:t></w:r></w:p>");
        DocxReadResult read = new DocxEditor().Read(input);
        Assert.True(read.Success);
        Assert.Equal(3, read.Paragraphs.Count);
        Assert.Equal("M.P0001", read.Paragraphs[0].Id.ToWireValue());
        Assert.Equal("First", read.Paragraphs[0].Text);
        Assert.Equal("M.P0002", read.Paragraphs[1].Id.ToWireValue());
        Assert.Equal("Inserted", read.Paragraphs[1].Text);
        Assert.Equal("M.P0003", read.Paragraphs[2].Id.ToWireValue());
        Assert.Equal("Last", read.Paragraphs[2].Text);
    }

    [Fact]
    public static void ReplaceParagraphTargetsInsertedBlockNotNeighbour()
    {
        using MemoryStream input = CreateDocxWithBody(
            "<w:p><w:r><w:t>First</w:t></w:r></w:p>" +
            "<w:ins w:id=\"1\" w:author=\"Alice\" w:date=\"2026-01-01T00:00:00Z\"><w:p><w:r><w:t>Inserted</w:t></w:r></w:p></w:ins>" +
            "<w:p><w:r><w:t>Last</w:t></w:r></w:p>");
        using var output = new MemoryStream();
        using var patch = new StringReader(
            "docxpatch 1\n\nop replace-paragraph\ntarget M.P0002\ntext Replaced\nend\n");
        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output);
        Assert.True(apply.Success);
        output.Position = 0;
        DocxReadResult reread = new DocxEditor().Read(output);
        Assert.True(reread.Success);
        Assert.Equal(3, reread.Paragraphs.Count);
        Assert.Equal("First", reread.Paragraphs[0].Text);
        Assert.Equal("Replaced", reread.Paragraphs[1].Text);
        Assert.Equal("Last", reread.Paragraphs[2].Text);
    }

    [Fact]
    public static void FinalViewOmitsDeletedBlockButKeepsPhysicalIds()
    {
        using MemoryStream input = CreateDocxWithBody(
            "<w:p><w:r><w:t>First</w:t></w:r></w:p>" +
            "<w:del w:id=\"2\" w:author=\"Bob\" w:date=\"2026-01-01T00:00:00Z\"><w:p><w:r><w:delText>Deleted</w:delText></w:r></w:p></w:del>" +
            "<w:p><w:r><w:t>Last</w:t></w:r></w:p>");
        DocxReadResult finalView = new DocxEditor().Read(input);
        Assert.True(finalView.Success);
        Assert.Equal(2, finalView.Paragraphs.Count);
        Assert.Equal("M.P0001", finalView.Paragraphs[0].Id.ToWireValue());
        Assert.Equal("M.P0003", finalView.Paragraphs[1].Id.ToWireValue());
        Assert.Equal("Last", finalView.Paragraphs[1].Text);
        DocxReadResult markupView = new DocxEditor().Read(input, new DocxReadOptions { TextView = DocxTextView.Markup });
        Assert.True(markupView.Success);
        Assert.Equal(3, markupView.Paragraphs.Count);
        Assert.Equal("M.P0002", markupView.Paragraphs[1].Id.ToWireValue());
    }

    [Fact]
    public static void TableCellIdsUseVisualColumnsAcrossReadAndChanges()
    {
        string tableXml =
            "<w:tbl>" +
            "<w:tblPr><w:tblStyle w:val=\"TableGrid\"/></w:tblPr>" +
            "<w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/></w:tblGrid>" +
            "<w:tr>" +
            "<w:tc><w:tcPr><w:gridSpan w:val=\"2\"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>" +
            "<w:tc><w:p><w:ins w:id=\"1\" w:author=\"Alice\" w:date=\"2026-01-01T00:00:00Z\"><w:r><w:t>B</w:t></w:r></w:ins></w:p></w:tc>" +
            "</w:tr>" +
            "</w:tbl>";
        using MemoryStream input = CreateDocxWithBody(tableXml);
        DocxReadResult read = new DocxEditor().Read(input);
        Assert.True(read.Success);
        var table = Assert.Single(read.Tables);
        Assert.Equal(2, table.Cells.Count);
        Assert.Equal("M.T0001.R01.C01", table.Cells[0].Id.ToWireValue());
        Assert.Equal("M.T0001.R01.C03", table.Cells[1].Id.ToWireValue());
        input.Position = 0;
        DocxChangesResult changes = new DocxEditor().Changes(input);
        Assert.True(changes.Success);
        var change = Assert.Single(changes.Changes);
        Assert.Equal("M.T0001.R01.C03", change.TargetId);
    }

    // D01: explicit IDs bind to the input snapshot for one patch. An insertion
    // must not make a later explicit ID slide onto the inserted block.
    [Fact]
    public static void InsertDoesNotRetargetLaterExplicitParagraph()
    {
        using MemoryStream input = CreateDocxWithBody(
            "<w:p><w:r><w:t>Alpha Alpha</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Beta</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Gamma</w:t></w:r></w:p>");
        using var output = new MemoryStream();
        using var patch = new StringReader(
            "docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext Inserted\nend\n\nop replace-paragraph\ntarget M.P0002\ntext Changed\nend\n");
        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output);
        Assert.True(apply.Success, string.Join("|", apply.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        string[] texts = new DocxEditor().Read(output).Paragraphs.Select(static paragraph => paragraph.Text).ToArray();
        Assert.Equal(new[] { "Alpha Alpha", "Inserted", "Changed", "Gamma" }, texts);
    }

    [Fact]
    public static void InsertDoesNotRetargetLaterExplicitParagraphInCheck()
    {
        using MemoryStream input = CreateDocxWithBody(
            "<w:p><w:r><w:t>Alpha Alpha</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Beta</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Gamma</w:t></w:r></w:p>");
        using var patch = new StringReader(
            "docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext Inserted\nend\n\nop replace-paragraph\ntarget M.P0002\ntext Changed\nend\n");
        DocxCheckResult check = new DocxEditor().Check(input, patch);
        Assert.True(check.Success, string.Join("|", check.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
    }

    [Fact]
    public static void DeleteDoesNotRetargetLaterExplicitParagraph()
    {
        using MemoryStream input = CreateDocxWithBody(
            "<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Beta</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Gamma</w:t></w:r></w:p>");
        using var output = new MemoryStream();
        using var patch = new StringReader(
            "docxpatch 1\n\nop delete-block\ntarget M.P0002\nend\n\nop replace-paragraph\ntarget M.P0003\ntext Changed\nend\n");
        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output);
        Assert.True(apply.Success, string.Join("|", apply.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        string[] texts = new DocxEditor().Read(output).Paragraphs.Select(static paragraph => paragraph.Text).ToArray();
        Assert.Equal(new[] { "Alpha", "Changed" }, texts);
    }

    [Fact]
    public static void DeletedExplicitTargetFailsInsteadOfEditingNeighbour()
    {
        using MemoryStream input = CreateDocxWithBody(
            "<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Beta</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Gamma</w:t></w:r></w:p>");
        using var output = new MemoryStream();
        using var patch = new StringReader(
            "docxpatch 1\n\nop delete-block\ntarget M.P0002\nend\n\nop replace-paragraph\ntarget M.P0002\ntext Changed\nend\n");
        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output);
        Assert.False(apply.Success);
        Assert.Contains(apply.Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }

    [Fact]
    public static void ShiftedOrdinalForInsertedBlockIsNotAddressableInSamePatch()
    {
        using MemoryStream input = CreateDocxWithBody(
            "<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Beta</w:t></w:r></w:p>");
        using var output = new MemoryStream();
        using var patch = new StringReader(
            "docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext Inserted\nend\n\nop replace-paragraph\ntarget M.P0003\ntext Changed\nend\n");
        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output);
        Assert.False(apply.Success);
        Assert.Contains(apply.Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }

    [Fact]
    public static void InsertRowDoesNotRetargetLaterExplicitRow()
    {
        string tableXml =
            "<w:tbl>" +
            "<w:tblPr><w:tblStyle w:val=\"TableGrid\"/></w:tblPr>" +
            "<w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>" +
            "<w:tr><w:tc><w:p><w:r><w:t>R1C1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>R1C2</w:t></w:r></w:p></w:tc></w:tr>" +
            "<w:tr><w:tc><w:p><w:r><w:t>R2C1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>R2C2</w:t></w:r></w:p></w:tc></w:tr>" +
            "</w:tbl>";
        using MemoryStream input = CreateDocxWithBody(tableXml);
        using var output = new MemoryStream();
        using var patch = new StringReader(
            "docxpatch 1\n\nop insert-row-after\ntarget M.T0001.R01\ncell New1\ncell New2\nend\n\nop set-cell\ntarget M.T0001.R02.C01\ntext Changed\nend\n");
        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output);
        Assert.True(apply.Success, string.Join("|", apply.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.True(read.Success);
        var table = Assert.Single(read.Tables);
        DocxTableCellInfo edited = table.Cells.Single(cell => cell.Id.ToWireValue() == "M.T0001.R03.C01");
        Assert.Equal("Changed", edited.Text);
        DocxTableCellInfo inserted = table.Cells.Single(cell => cell.Id.ToWireValue() == "M.T0001.R02.C01");
        Assert.Equal("New1", inserted.Text);
    }

    [Fact]
    public static void SemanticSelectorStillResolvesLiveAfterInsert()
    {
        using MemoryStream input = CreateDocxWithBody(
            "<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Beta</w:t></w:r></w:p>");
        using var output = new MemoryStream();
        using var patch = new StringReader(
            "docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext Beta\nend\n\nop replace-text\ntarget text:\"Beta\"\nfind Beta\nwith Changed\nend\n");
        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output);
        Assert.False(apply.Success);
        Assert.Contains(apply.Diagnostics, static diagnostic => diagnostic.Code == "E1202");
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

}
