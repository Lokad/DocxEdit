using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D14: ambiguous-match diagnostics carry machine-readable counts and candidate IDs.
[Collection("ConsoleCli")]
public static class DiagnosticMatchTests
{
    [Fact]
    public static void AmbiguousTextSelectorCarriesCandidates()
    {
        using MemoryStream input = CreateDocxWithBody("<w:p><w:r><w:t>Revenue north</w:t></w:r></w:p>\n<w:p><w:r><w:t>Revenue south</w:t></w:r></w:p>\n");
        using var patch = new StringReader("docxpatch 1\n\nop replace-text\ntarget text:Revenue\nfind Revenue\nwith Sales\nend\n");
        DocxCheckResult result = new DocxEditor().Check(input, patch);
        Assert.False(result.Success);
        DocxDiagnostic diagnostic = result.Diagnostics.First(static d => d.Code == "E1202");
        Assert.Equal(2, diagnostic.MatchCount);
        Assert.Equal(new[] { "M.P0001", "M.P0002" }, diagnostic.CandidateIds);
    }
    [Fact]
    public static void UnmatchedTextSelectorCarriesCandidates()
    {
        using MemoryStream input = CreateDocxWithBody("<w:p><w:r><w:t>Revenue north</w:t></w:r></w:p>\n<w:p><w:r><w:t>Revenue south</w:t></w:r></w:p>\n");
        using var patch = new StringReader("docxpatch 1\n\nop replace-text\ntarget text:Expense\nfind Expense\nwith Sales\nend\n");
        DocxCheckResult result = new DocxEditor().Check(input, patch);
        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, static d => d.Code == "E1201");
        Assert.Equal(0, diagnostic.MatchCount);
        Assert.Equal(new[] { "M.P0001", "M.P0002" }, diagnostic.CandidateIds);
    }


    [Fact]
    public static void AmbiguousBookmarkNameCarriesCandidates()
    {
        using MemoryStream input = CreateDocxWithBody("<w:p><w:bookmarkStart w:id=\"1\" w:name=\"ClientName\"/><w:r><w:t>First</w:t></w:r></w:p>\n<w:p><w:bookmarkStart w:id=\"2\" w:name=\"ClientName\"/><w:r><w:t>Second</w:t></w:r></w:p>\n");
        using var patch = new StringReader("docxpatch 1\n\nop replace-bookmark-text\ntarget bookmark:\"ClientName\"\ntext New\nend\n");
        DocxCheckResult result = new DocxEditor().Check(input, patch);
        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, static d => d.Code == "E1202");
        Assert.Equal(2, diagnostic.MatchCount);
        Assert.Equal(new[] { "M.B0001", "M.B0002" }, diagnostic.CandidateIds);
    }

    [Fact]
    public static void AmbiguousContentControlNameCarriesCandidates()
    {
        using MemoryStream input = CreateDocxWithBody("<w:p><w:sdt><w:sdtPr><w:text/><w:tag w:val=\"amount\"/></w:sdtPr><w:sdtContent><w:r><w:t>First</w:t></w:r></w:sdtContent></w:sdt></w:p>\n<w:p><w:sdt><w:sdtPr><w:text/><w:tag w:val=\"amount\"/></w:sdtPr><w:sdtContent><w:r><w:t>Second</w:t></w:r></w:sdtContent></w:sdt></w:p>\n");
        using var patch = new StringReader("docxpatch 1\n\nop set-content-control-text\ntarget content-control:\"amount\"\ntext New\nend\n");
        DocxCheckResult result = new DocxEditor().Check(input, patch);
        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, static d => d.Code == "E1202");
        Assert.Equal(2, diagnostic.MatchCount);
        Assert.Equal(new[] { "M.CC0001", "M.CC0002" }, diagnostic.CandidateIds);
    }

    [Fact]
    public static void MultipleOccurrencesCarryCountWithoutCandidates()
    {
        using MemoryStream input = CreateDocx("Alpha Alpha");
        using var patch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Omega\nend\n");
        DocxCheckResult result = new DocxEditor().Check(input, patch);
        Assert.False(result.Success);
        DocxDiagnostic diagnostic = result.Diagnostics.First(static d => d.Code == "E1202");
        Assert.Equal(2, diagnostic.MatchCount);
        Assert.Null(diagnostic.CandidateIds);
    }

    [Fact]
    public static void AmbiguousAnchorTextCarriesCountWithoutCandidates()
    {
        using MemoryStream input = CreateDocx("Alpha beta beta.");
        using var patch = new StringReader("docxpatch 1\n\nop add-comment\ntarget M.P0001\nanchor-text beta\ntext Review beta\nauthor Reviewer\nend\n");
        DocxCheckResult result = new DocxEditor().Check(input, patch);
        Assert.False(result.Success);
        DocxDiagnostic diagnostic = result.Diagnostics.First(static d => d.Code == "E1202");
        Assert.Equal(2, diagnostic.MatchCount);
        Assert.Null(diagnostic.CandidateIds);
    }

    [Fact]
    public static void CliAmbiguityPrintsMatchCount()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = Path.Combine(temp.Path, "input.docx");
        File.WriteAllBytes(input, CreateDocxWithBody("<w:p><w:r><w:t>Revenue north</w:t></w:r></w:p>\n<w:p><w:r><w:t>Revenue south</w:t></w:r></w:p>\n").ToArray());
        string patch = Path.Combine(temp.Path, "edits.docxpatch");
        File.WriteAllText(patch, "docxpatch 1" + Environment.NewLine + Environment.NewLine + "op replace-text" + Environment.NewLine + "target text:Revenue" + Environment.NewLine + "find Revenue" + Environment.NewLine + "with Sales" + Environment.NewLine + "end" + Environment.NewLine);
        CliTests.CliResult result = CliTests.RunCli("check", input, patch);
        Assert.Equal(1, result.ExitCode);
        Assert.Contains("matches=2", result.Error, StringComparison.Ordinal);
    }
}
