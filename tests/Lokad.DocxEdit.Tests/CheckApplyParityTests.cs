using System.IO.Compression;
using System.Text;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// PLAN B03: check must simulate the same mutations as apply (on the disposable
// in-memory package) so dependent edits, repeated guards, field refresh, and
// post-edit validation reach equivalent outcomes. Only publication differs:
// apply saves; check discards. GeneratedRevisionIds stay apply-only by design.
public static class CheckApplyParityTests
{
    [Fact]
    public static void DependentEditsAgreeBetweenCheckAndApply()
    {
        using MemoryStream checkInput = CreateDocx("Alpha.");
        using var checkPatch = new StringReader(
            "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n\nop replace-text\ntarget M.P0001\nfind Beta\nwith Gamma\nend\n");
        DocxCheckResult check = new DocxEditor().Check(checkInput, checkPatch);
        using MemoryStream applyInput = CreateDocx("Alpha.");
        using var applyPatch = new StringReader(
            "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n\nop replace-text\ntarget M.P0001\nfind Beta\nwith Gamma\nend\n");
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, applyPatch, output);
        Assert.Equal(apply.Success, check.Success);
        Assert.True(check.Success);
        output.Position = 0;
        Assert.Equal("Gamma.", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void RepeatedGuardAgreesBetweenCheckAndApply()
    {
        string patchText = "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Gamma\nend\n";
        using MemoryStream checkInput = CreateDocx("Alpha.");
        DocxCheckResult check = new DocxEditor().Check(checkInput, new StringReader(patchText));
        using MemoryStream applyInput = CreateDocx("Alpha.");
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, new StringReader(patchText), output);
        Assert.Equal(apply.Success, check.Success);
        Assert.False(check.Success);
    }

    [Fact]
    public static void FieldRefreshWarningAgreesBetweenCheckAndApply()
    {
        using MemoryStream checkInput = CreateDocxWithRefField();
        using var checkPatch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n");
        DocxCheckResult check = new DocxEditor().Check(checkInput, checkPatch);
        using MemoryStream applyInput = CreateDocxWithRefField();
        using var applyPatch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n");
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, applyPatch, output);
        Assert.Equal(apply.Success, check.Success);
        Assert.True(check.Success);
        Assert.Contains(check.Diagnostics, static d => d.Code == "W5103");
        Assert.Contains(apply.Diagnostics, static d => d.Code == "W5103");
    }

    [Fact]
    public static void GeneratedRevisionIdsStayApplyOnly()
    {
        var options = new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest };
        using MemoryStream checkInput = CreateDocx("Alpha.");
        using var checkPatch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n");
        DocxCheckResult check = new DocxEditor().Check(checkInput, checkPatch, options);
        using MemoryStream applyInput = CreateDocx("Alpha.");
        using var applyPatch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n");
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, applyPatch, output, options);
        Assert.Equal(apply.Success, check.Success);
        Assert.True(check.Success);
        Assert.All(check.Operations, static op => Assert.Empty(op.GeneratedRevisionIds));
        Assert.Contains(apply.Operations, static op => op.GeneratedRevisionIds.Count > 0);
    }

    [Fact]
    public static void SettingsAllocationPreservesExternalRelationships()
    {
        using MemoryStream input = CreateDocxWithExternalHyperlink();
        using var output = new MemoryStream();
        using var patch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n");
        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output);
        Assert.True(apply.Success, string.Join("|", apply.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        string rels = ReadEntry(output, "word/_rels/document.xml.rels");
        Assert.Contains("Target=\"https://example.test/keep\"", rels, StringComparison.Ordinal);
        Assert.Contains("Target=\"settings.xml\"", rels, StringComparison.Ordinal);
    }

    private static MemoryStream CreateDocxWithExternalHyperlink()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", RootRelsXml());
            AddEntry(archive, "word/_rels/document.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.test/keep\" TargetMode=\"External\"/>"
                + "</Relationships>");
            string documentXml = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>"
                + "<w:p><w:r><w:t>Alpha.</w:t></w:r></w:p>"
                + "</w:body></w:document>";
            AddEntry(archive, "word/document.xml", documentXml);
        }
        stream.Position = 0;
        return stream;
    }


    private static MemoryStream CreateDocx(string paragraphText)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", RootRelsXml());
            AddEntry(archive, "word/_rels/document.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\" />");
            string documentXml = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>"
                + "<w:p><w:r><w:t>" + paragraphText + "</w:t></w:r></w:p>"
                + "</w:body></w:document>";
            AddEntry(archive, "word/document.xml", documentXml);
        }
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithRefField()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", RootRelsXml());
            AddEntry(archive, "word/_rels/document.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\" />");
            string documentXml = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>"
                + "<w:p><w:r><w:t>Alpha.</w:t></w:r></w:p>"
                + "<w:p><w:fldSimple w:instr=\" REF ClientName \\h \"><w:r><w:t>Acme</w:t></w:r></w:fldSimple></w:p>"
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
