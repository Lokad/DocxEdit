using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D06: guarded patch templates pair D17 capabilities with executable,
// check-clean starters for one discovered target.
public static class PatchTemplateTests
{
    private static DocxTemplateResult GetTemplate(MemoryStream input, string targetId, TrackChangesMode mode)
    {
        input.Position = 0;
        DocxTemplateResult result = new DocxEditor().GetTemplate(input, targetId, new DocxTemplateOptions { TrackChanges = mode });
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        Assert.Equal(targetId, result.TargetId);
        Assert.NotNull(result.Capabilities);
        Assert.False(string.IsNullOrWhiteSpace(result.Template));
        return result;
    }

    private static int CountActiveOps(string template)
    {
        return template.Split((char)10).Count(static line => line.StartsWith("op ", StringComparison.Ordinal));
    }

    private static void AssertNoActiveSetCell(string template)
    {
        foreach (string line in template.Split((char)10))
        {
            Assert.False(line.StartsWith("op set-cell", StringComparison.Ordinal), line);
        }
    }

    private static DocxCheckResult RunCheck(MemoryStream input, string patchText)
    {
        input.Position = 0;
        using var patch = new StringReader(patchText);
        return new DocxEditor().Check(input, patch);
    }

    [Fact]
    public static void ParagraphTemplateStarterPassesCheck()
    {
        DocxTemplateResult template;
        using (MemoryStream input = CreateDocx("Alpha Beta"))
        {
            template = GetTemplate(input, "M.P0001", TrackChangesMode.Off);
        }

        Assert.Equal(1, CountActiveOps(template.Template));
        Assert.Contains("op replace-text", template.Template, StringComparison.Ordinal);
        Assert.Contains("Operation help:", template.Template, StringComparison.Ordinal);

        using MemoryStream checkInput = CreateDocx("Alpha Beta");
        DocxCheckResult result = RunCheck(checkInput, template.Template);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void EmptyParagraphTemplateUsesReplaceParagraph()
    {
        const string body = """
                <w:p><w:r><w:t>Kept</w:t></w:r></w:p>
                <w:p></w:p>
            """;
        DocxTemplateResult template;
        using (MemoryStream input = CreateDocxWithBody(body))
        {
            template = GetTemplate(input, "M.P0002", TrackChangesMode.Off);
        }

        Assert.Equal(1, CountActiveOps(template.Template));
        Assert.Contains("op replace-paragraph", template.Template, StringComparison.Ordinal);

        using MemoryStream checkInput = CreateDocxWithBody(body);
        DocxCheckResult result = RunCheck(checkInput, template.Template);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void LockedControlTemplateHasNoActiveBlock()
    {
        const string body = """
                    <w:p>
                      <w:sdt>
                        <w:sdtPr><w:text/><w:lock w:val="sdtContentLocked"/></w:sdtPr>
                        <w:sdtContent><w:r><w:t>Old Client</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """;
        using MemoryStream input = CreateDocxWithBody(body);
        DocxTemplateResult template = GetTemplate(input, "M.CC0001", TrackChangesMode.Off);

        Assert.Equal(0, CountActiveOps(template.Template));
        Assert.Contains("No check-clean starter", template.Template, StringComparison.Ordinal);
    }

    [Fact]
    public static void PlainControlTemplateStarterPassesCheck()
    {
        const string body = """
                    <w:p>
                      <w:sdt>
                        <w:sdtPr><w:text/><w:tag w:val="client"/></w:sdtPr>
                        <w:sdtContent><w:r><w:t>Old Client</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """;
        DocxTemplateResult template;
        using (MemoryStream input = CreateDocxWithBody(body))
        {
            template = GetTemplate(input, "M.CC0001", TrackChangesMode.Off);
        }

        Assert.Equal(1, CountActiveOps(template.Template));
        Assert.Contains("op set-content-control-text", template.Template, StringComparison.Ordinal);

        using MemoryStream checkInput = CreateDocxWithBody(body);
        DocxCheckResult result = RunCheck(checkInput, template.Template);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void ContinuationCellTemplateOmitsSetCell()
    {
        const string body = """
                <w:tbl>
                  <w:tr>
                    <w:tc>
                      <w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
                      <w:p><w:r><w:t>North</w:t></w:r></w:p>
                    </w:tc>
                  </w:tr>
                  <w:tr>
                    <w:tc>
                      <w:tcPr><w:vMerge/></w:tcPr>
                      <w:p><w:r><w:t>South</w:t></w:r></w:p>
                    </w:tc>
                  </w:tr>
                </w:tbl>
            """;
        using MemoryStream input = CreateDocxWithBody(body);
        DocxTemplateResult template = GetTemplate(input, "M.T0001.R02.C01", TrackChangesMode.Off);

        Assert.Equal(0, CountActiveOps(template.Template));
        AssertNoActiveSetCell(template.Template);
    }

    [Fact]
    public static void ComplexCellTemplateUsesForceAndPassesCheck()
    {
        const string body = """
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:p><w:r><w:t>First</w:t></w:r></w:p>
                          <w:p><w:r><w:t>Second</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """;
        DocxTemplateResult template;
        using (MemoryStream input = CreateDocxWithBody(body))
        {
            template = GetTemplate(input, "M.T0001.R01.C01", TrackChangesMode.Off);
        }

        Assert.Equal(1, CountActiveOps(template.Template));
        Assert.Contains("force true", template.Template, StringComparison.Ordinal);

        using MemoryStream checkInput = CreateDocxWithBody(body);
        DocxCheckResult result = RunCheck(checkInput, template.Template);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void RequireComplexCellTemplateOmitsSetCell()
    {
        const string body = """
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:p><w:fldSimple w:instr=" REF Mark "><w:r><w:t>Old</w:t></w:r></w:fldSimple></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """;
        using MemoryStream input = CreateDocxWithBody(body);
        DocxTemplateResult template = GetTemplate(input, "M.T0001.R01.C01", TrackChangesMode.Require);

        Assert.Equal(0, CountActiveOps(template.Template));
        AssertNoActiveSetCell(template.Template);
        Assert.Contains("op set-cell-shading", template.Template, StringComparison.Ordinal);
    }

    [Fact]
    public static void BookmarkTemplateStarterPassesCheck()
    {
        DocxTemplateResult template;
        using (MemoryStream input = CreateDocxWithSingleBookmark())
        {
            template = GetTemplate(input, "M.B0001", TrackChangesMode.Off);
        }
        Assert.Equal(1, CountActiveOps(template.Template));
        Assert.Contains("op replace-bookmark-text", template.Template, StringComparison.Ordinal);
        using MemoryStream checkInput = CreateDocxWithSingleBookmark();
        DocxCheckResult result = RunCheck(checkInput, template.Template);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }
    [Fact]
    public static void IncompleteBookmarkTemplateHasNoActiveBlock()
    {
        const string body = """
                    <w:p>
                      <w:bookmarkStart w:id="5" w:name="Orphan"/>
                      <w:r><w:t>text</w:t></w:r>
                    </w:p>
            """;
        using MemoryStream input = CreateDocxWithBody(body);
        DocxTemplateResult template = GetTemplate(input, "M.B0001", TrackChangesMode.Off);
        Assert.Equal(0, CountActiveOps(template.Template));
        Assert.Contains("No check-clean starter", template.Template, StringComparison.Ordinal);
        Assert.Contains("# op rename-bookmark", template.Template, StringComparison.Ordinal);
    }

    [Fact]
    public static void UnknownTargetTemplateFailsWithE1201()
    {
        using MemoryStream input = CreateDocx("Alpha");
        input.Position = 0;
        DocxTemplateResult result = new DocxEditor().GetTemplate(input, "M.P0009");
        Assert.False(result.Success);
        Assert.True(string.IsNullOrEmpty(result.Template));
        Assert.Null(result.Capabilities);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }
}
