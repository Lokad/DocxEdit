using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D17: cell and merge-group capabilities reuse the merge-shape,
// simplicity, and tracked-shape predicates from the table engine.
public static class CellCapabilitiesTests
{
    private const string MergeBody = """
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
        """;

    private static DocxOperationCapability FindOperation(DocxTargetCapabilities capabilities, string operation)
    {
        return capabilities.Operations.First(candidate => string.Equals(candidate.Operation, operation, StringComparison.Ordinal));
    }

    private static DocxTargetCapabilities GetCapabilities(MemoryStream input, string targetId, TrackChangesMode mode)
    {
        input.Position = 0;
        DocxCapabilitiesResult result = new DocxEditor().GetCapabilities(input, targetId, new DocxCapabilitiesOptions { TrackChanges = mode });
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        Assert.Equal(targetId, result.TargetId);
        Assert.NotNull(result.Capabilities);
        return result.Capabilities!;
    }

    private static DocxCheckResult RunCheck(MemoryStream input, string patchText, TrackChangesMode mode)
    {
        input.Position = 0;
        using var patch = new StringReader(patchText);
        return new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = mode });
    }

    [Fact]
    public static void SimpleCellSupportsSetAndShading()
    {
        using MemoryStream input = CreateDocxWithSimpleTwoByTwoTable();
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.T0001.R01.C01", TrackChangesMode.Off);

        Assert.Equal("M.T0001.R01.C01", capabilities.TargetId);
        Assert.Equal("cell", capabilities.Kind);
        Assert.Equal("main", capabilities.Story);
        string[] expectedOrder = ["set-cell", "set-cell-shading", "replace-text", "add-comment"];
        Assert.Equal(expectedOrder, capabilities.Operations.Select(static operation => operation.Operation).ToArray());
        Assert.Equal("supported", FindOperation(capabilities, "set-cell").Support);
        Assert.Equal("supported", FindOperation(capabilities, "set-cell-shading").Support);

        using MemoryStream checkInput = CreateDocxWithSimpleTwoByTwoTable();
        DocxCheckResult result = RunCheck(checkInput, "docxpatch 1\n\nop set-cell\ntarget M.T0001.R01.C01\ntext West\nend\n", TrackChangesMode.Off);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void ContinuationCellRefusesWithRootAlternative()
    {
        TrackChangesMode[] modes = [TrackChangesMode.Off, TrackChangesMode.Suggest, TrackChangesMode.Require];
        foreach (TrackChangesMode mode in modes)
        {
            using MemoryStream input = CreateDocxWithBody(MergeBody);
            DocxTargetCapabilities capabilities = GetCapabilities(input, "M.T0001.R02.C01", mode);
            Assert.Equal("cell", capabilities.Kind);
            DocxOperationCapability setCell = FindOperation(capabilities, "set-cell");
            Assert.Equal("unsupported", setCell.Support);
            Assert.Equal("M.T0001.R01.C01", setCell.Alternative);
            DocxOperationCapability shading = FindOperation(capabilities, "set-cell-shading");
            Assert.Equal("unsupported", shading.Support);
            Assert.Equal("M.T0001.R01.C01", shading.Alternative);
        }

        using MemoryStream checkInput = CreateDocxWithBody(MergeBody);
        DocxCheckResult refused = RunCheck(checkInput, "docxpatch 1\n\nop set-cell\ntarget M.T0001.R02.C01\ntext East\nend\n", TrackChangesMode.Off);
        Assert.False(refused.Success);
        Assert.Contains(refused.Diagnostics, static diagnostic => diagnostic.Code == "E4301");
    }

    [Fact]
    public static void MergeGroupResolvesToRootCell()
    {
        using MemoryStream input = CreateDocxWithBody(MergeBody);
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.T0001.MG0001", TrackChangesMode.Off);

        Assert.Equal("M.T0001.MG0001", capabilities.TargetId);
        Assert.Equal("merge-group", capabilities.Kind);
        Assert.Equal("supported", FindOperation(capabilities, "set-cell").Support);
        Assert.Equal("supported", FindOperation(capabilities, "set-cell-shading").Support);

        using MemoryStream missingInput = CreateDocxWithBody(MergeBody);
        missingInput.Position = 0;
        DocxCapabilitiesResult missing = new DocxEditor().GetCapabilities(missingInput, "M.T0001.MG0009");
        Assert.False(missing.Success);
        Assert.Null(missing.Capabilities);
        Assert.Contains(missing.Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }

    [Fact]
    public static void MultiParagraphCellNeedsForce()
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
        using MemoryStream input = CreateDocxWithBody(body);
        Assert.Equal("conditional", FindOperation(GetCapabilities(input, "M.T0001.R01.C01", TrackChangesMode.Off), "set-cell").Support);

        using MemoryStream refusedInput = CreateDocxWithBody(body);
        DocxCheckResult refused = RunCheck(refusedInput, "docxpatch 1\n\nop set-cell\ntarget M.T0001.R01.C01\ntext Solo\nend\n", TrackChangesMode.Off);
        Assert.False(refused.Success);
        Assert.Contains(refused.Diagnostics, static diagnostic => diagnostic.Code == "E4302");

        using MemoryStream forcedInput = CreateDocxWithBody(body);
        DocxCheckResult forced = RunCheck(forcedInput, "docxpatch 1\n\nop set-cell\ntarget M.T0001.R01.C01\ntext Solo\nforce true\nend\n", TrackChangesMode.Off);
        Assert.True(forced.Success, string.Join("|", forced.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));

        using MemoryStream shadingInput = CreateDocxWithBody(body);
        Assert.Equal("supported", FindOperation(GetCapabilities(shadingInput, "M.T0001.R01.C01", TrackChangesMode.Off), "set-cell-shading").Support);
    }

    [Fact]
    public static void FieldCellRefusesRequireButAllowsSuggestFallback()
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
        using MemoryStream requireInput = CreateDocxWithBody(body);
        Assert.Equal("unsupported", FindOperation(GetCapabilities(requireInput, "M.T0001.R01.C01", TrackChangesMode.Require), "set-cell").Support);

        using MemoryStream suggestInput = CreateDocxWithBody(body);
        Assert.Equal("conditional", FindOperation(GetCapabilities(suggestInput, "M.T0001.R01.C01", TrackChangesMode.Suggest), "set-cell").Support);

        using MemoryStream refusedInput = CreateDocxWithBody(body);
        DocxCheckResult refused = RunCheck(refusedInput, "docxpatch 1\n\nop set-cell\ntarget M.T0001.R01.C01\ntext New\nend\n", TrackChangesMode.Require);
        Assert.False(refused.Success);
        Assert.Contains(refused.Diagnostics, static diagnostic => diagnostic.Code == "E6002");

        using MemoryStream fallbackInput = CreateDocxWithBody(body);
        DocxCheckResult fallback = RunCheck(fallbackInput, "docxpatch 1\n\nop set-cell\ntarget M.T0001.R01.C01\ntext New\nforce true\nend\n", TrackChangesMode.Suggest);
        Assert.True(fallback.Success, string.Join("|", fallback.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        Assert.Contains(fallback.Diagnostics, static diagnostic => diagnostic.Code == "W4002");
    }
    [Fact]
    public static void SimpleCellReplaceTextIsConditional()
    {
        using MemoryStream input = CreateDocxWithSimpleTwoByTwoTable();
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.T0001.R01.C01", TrackChangesMode.Off);

        Assert.Equal("conditional", FindOperation(capabilities, "replace-text").Support);
    }

    [Fact]
    public static void ContinuationCellReplaceTextIsUnsupported()
    {
        using MemoryStream input = CreateDocxWithBody(MergeBody);
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.T0001.R02.C01", TrackChangesMode.Off);

        DocxOperationCapability replaceText = FindOperation(capabilities, "replace-text");
        Assert.Equal("unsupported", replaceText.Support);
        Assert.Equal("M.T0001.R01.C01", replaceText.Alternative);
    }
}
