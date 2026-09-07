using System.IO.Compression;
using System.Text;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// PLAN B04: destructive edits must not save new cross-element corruption.
// Deleting an element that would orphan a bookmark/comment-range/field boundary
// fails before publication (E4305); whole-range removal still works;
// pre-existing orphans do not block unrelated edits; the post-edit net catches
// whatever slips past op-level guards (E9103 and friends, check/apply parity).
public static class StructuralGuardTests
{
    [Fact]
    public static void DeleteBlockRefusesOrphanedBookmarkSpan()
    {
        string patchText = "docxpatch 1\n\nop delete-block\ntarget M.P0001\nend\n";
        using MemoryStream checkInput = CreateBookmarkSpanDocx();
        DocxCheckResult check = new DocxEditor().Check(checkInput, new StringReader(patchText));
        Assert.False(check.Success);
        Assert.Contains(check.Diagnostics, static d => d.Code == "E4305");
        Assert.Contains(check.Diagnostics, static d => d.Message.Contains("ClientName", StringComparison.Ordinal));
        using MemoryStream applyInput = CreateBookmarkSpanDocx();
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, new StringReader(patchText), output);
        Assert.False(apply.Success);
        Assert.Contains(apply.Diagnostics, static d => d.Code == "E4305");
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public static void DeleteBlockAllowsWholeBookmarkRemoval()
    {
        string patchText = "docxpatch 1\n\nop delete-block\ntarget M.P0001\nend\n";
        using MemoryStream input = CreateWholeBookmarkDocx();
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(input, new StringReader(patchText), output);
        Assert.True(apply.Success);
        output.Position = 0;
        DocxReadResult reread = new DocxEditor().Read(output);
        Assert.True(reread.Success);
        Assert.Empty(reread.Bookmarks);
        Assert.Equal("Second", Assert.Single(reread.Paragraphs).Text);
    }

    [Fact]
    public static void DeleteBlockRefusesOrphanedCommentRange()
    {
        string patchText = "docxpatch 1\n\nop delete-block\ntarget M.P0001\nend\n";
        using MemoryStream checkInput = CreateCommentSpanDocx();
        DocxCheckResult check = new DocxEditor().Check(checkInput, new StringReader(patchText));
        Assert.False(check.Success);
        Assert.Contains(check.Diagnostics, static d => d.Code == "E4305");
        using MemoryStream applyInput = CreateCommentSpanDocx();
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, new StringReader(patchText), output);
        Assert.False(apply.Success);
        Assert.Contains(apply.Diagnostics, static d => d.Code == "E4305");
    }

    [Fact]
    public static void DeleteRowRefusesOrphanedBookmark()
    {
        string patchText = "docxpatch 1\n\nop delete-row\ntarget M.T0001.R01\nend\n";
        using MemoryStream checkInput = CreateRowSpanDocx();
        DocxCheckResult check = new DocxEditor().Check(checkInput, new StringReader(patchText));
        Assert.False(check.Success);
        Assert.Contains(check.Diagnostics, static d => d.Code == "E4305");
        using MemoryStream applyInput = CreateRowSpanDocx();
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, new StringReader(patchText), output);
        Assert.False(apply.Success);
        Assert.Contains(apply.Diagnostics, static d => d.Code == "E4305");
    }

    [Fact]
    public static void TrackedDeleteAlsoRefusesOrphaningSpan()
    {
        string patchText = "docxpatch 1\n\nop delete-block\ntarget M.P0001\nend\n";
        var options = new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest };
        using MemoryStream checkInput = CreateBookmarkSpanDocx();
        DocxCheckResult check = new DocxEditor().Check(checkInput, new StringReader(patchText), options);
        Assert.False(check.Success);
        Assert.Contains(check.Diagnostics, static d => d.Code == "E4305");
        using MemoryStream applyInput = CreateBookmarkSpanDocx();
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, new StringReader(patchText), output, options);
        Assert.False(apply.Success);
        Assert.Contains(apply.Diagnostics, static d => d.Code == "E4305");
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public static void PreExistingOrphanDoesNotBlockUnrelatedEdit()
    {
        string patchText = "docxpatch 1\n\nop replace-text\ntarget M.P0002\nfind Alpha\nwith Beta\nend\n";
        using MemoryStream checkInput = CreatePreOrphanDocx();
        DocxCheckResult check = new DocxEditor().Check(checkInput, new StringReader(patchText));
        Assert.True(check.Success);
        using MemoryStream applyInput = CreatePreOrphanDocx();
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, new StringReader(patchText), output);
        Assert.True(apply.Success);
    }

    [Fact]
    public static void SetCellForceOrphanCaughtByPostEditNet()
    {
        string patchText = "docxpatch 1\n\nop set-cell\ntarget M.T0001.R01.C01\ntext New\nforce true\nend\n";
        using MemoryStream checkInput = CreateCellSpanDocx();
        DocxCheckResult check = new DocxEditor().Check(checkInput, new StringReader(patchText));
        Assert.False(check.Success);
        Assert.Contains(check.Diagnostics, static d => d.Code == "E9103");
        using MemoryStream applyInput = CreateCellSpanDocx();
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, new StringReader(patchText), output);
        Assert.False(apply.Success);
        Assert.Contains(apply.Diagnostics, static d => d.Code == "E9103");
        Assert.Equal(0, output.Length);
    }

    private static MemoryStream CreateBookmarkSpanDocx()
    {
        return CreateDocxWithBody(
            "<w:p><w:bookmarkStart w:id=\"1\" w:name=\"ClientName\"/><w:r><w:t>First</w:t></w:r></w:p>"
            + "<w:p><w:bookmarkEnd w:id=\"1\"/><w:r><w:t>Second</w:t></w:r></w:p>");
    }

    private static MemoryStream CreateWholeBookmarkDocx()
    {
        return CreateDocxWithBody(
            "<w:p><w:bookmarkStart w:id=\"1\" w:name=\"ClientName\"/><w:r><w:t>First</w:t></w:r><w:bookmarkEnd w:id=\"1\"/></w:p>"
            + "<w:p><w:r><w:t>Second</w:t></w:r></w:p>");
    }

    private static MemoryStream CreateCommentSpanDocx()
    {
        return CreateDocxWithBody(
            "<w:p><w:commentRangeStart w:id=\"3\"/><w:r><w:t>First</w:t></w:r></w:p>"
            + "<w:p><w:commentRangeEnd w:id=\"3\"/><w:r><w:t>Second</w:t></w:r></w:p>");
    }

    private static MemoryStream CreateRowSpanDocx()
    {
        return CreateDocxWithBody(
            "<w:tbl><w:tblPr><w:tblStyle w:val=\"TableGrid\"/></w:tblPr><w:tblGrid><w:gridCol/></w:tblGrid>"
            + "<w:tr><w:tc><w:p><w:bookmarkStart w:id=\"2\" w:name=\"RowSpan\"/><w:r><w:t>Top</w:t></w:r></w:p></w:tc></w:tr>"
            + "<w:tr><w:tc><w:p><w:bookmarkEnd w:id=\"2\"/><w:r><w:t>Bottom</w:t></w:r></w:p></w:tc></w:tr>"
            + "</w:tbl>");
    }

    private static MemoryStream CreatePreOrphanDocx()
    {
        return CreateDocxWithBody(
            "<w:p><w:bookmarkStart w:id=\"9\" w:name=\"Orphan\"/><w:r><w:t>Broken</w:t></w:r></w:p>"
            + "<w:p><w:r><w:t>Alpha.</w:t></w:r></w:p>");
    }

    private static MemoryStream CreateCellSpanDocx()
    {
        return CreateDocxWithBody(
            "<w:tbl><w:tblPr><w:tblStyle w:val=\"TableGrid\"/></w:tblPr><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>"
            + "<w:tr>"
            + "<w:tc><w:p><w:bookmarkStart w:id=\"5\" w:name=\"CellSpan\"/><w:r><w:t>Left</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:p><w:bookmarkEnd w:id=\"5\"/><w:r><w:t>Right</w:t></w:r></w:p></w:tc>"
            + "</w:tr>"
            + "</w:tbl>");
    }

    private static MemoryStream CreateDocxWithBody(string bodyXml)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", RootRelsXml());
            AddEntry(archive, "word/_rels/document.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\" />");
            string documentXml = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>"
                + bodyXml
                + "</w:body></w:document>";
            AddEntry(archive, "word/document.xml", documentXml);
        }
        stream.Position = 0;
        return stream;
    }

    private static string ContentTypesXml()
    {
        return "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
            + "</Types>";
    }

    private static string RootRelsXml()
    {
        return "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rDocument\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
            + "</Relationships>";
    }

}
