using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D17: section capabilities reuse the guard and tracked-revision predicates
// from the section engine.
public static class SectionCapabilitiesTests
{
    private const string SimpleBody = """
                    <w:p><w:r><w:t>Main text</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:pgSz w:w="12240" w:h="15840"/>
                      <w:cols w:num="1"/>
                    </w:sectPr>
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
    public static void SimpleSectionReportsCurrentState()
    {
        using MemoryStream input = CreateDocxWithBody(SimpleBody);
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.S0001", TrackChangesMode.Off);

        Assert.Equal("M.S0001", capabilities.TargetId);
        Assert.Equal("section", capabilities.Kind);
        Assert.Equal("main", capabilities.Story);
        string[] expectedOrder = ["set-section-columns", "set-section-orientation"];
        Assert.Equal(expectedOrder, capabilities.Operations.Select(static operation => operation.Operation).ToArray());
        Assert.Equal("conditional", FindOperation(capabilities, "set-section-columns").Support);
        Assert.Equal("conditional", FindOperation(capabilities, "set-section-orientation").Support);
        Assert.Contains("1 column(s) in portrait", FindOperation(capabilities, "set-section-columns").Reason, StringComparison.Ordinal);

        using MemoryStream columnsInput = CreateDocxWithBody(SimpleBody);
        DocxCheckResult columns = RunCheck(columnsInput, "docxpatch 1\n\nop set-section-columns\ntarget M.S0001\ncount 2\nend\n", TrackChangesMode.Off);
        Assert.True(columns.Success, string.Join("|", columns.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));

        using MemoryStream orientationInput = CreateDocxWithBody(SimpleBody);
        DocxCheckResult orientation = RunCheck(orientationInput, "docxpatch 1\n\nop set-section-orientation\ntarget M.S0001\norientation landscape\nend\n", TrackChangesMode.Off);
        Assert.True(orientation.Success, string.Join("|", orientation.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void ExistingRevisionFallsBackOrFails()
    {
        using MemoryStream suggestInput = CreateDocxWithExistingSectionPropertyRevision();
        DocxTargetCapabilities suggest = GetCapabilities(suggestInput, "M.S0001", TrackChangesMode.Suggest);
        Assert.Equal("conditional", FindOperation(suggest, "set-section-columns").Support);

        using MemoryStream suggestCheck = CreateDocxWithExistingSectionPropertyRevision();
        DocxCheckResult fallback = RunCheck(suggestCheck, "docxpatch 1\n\nop set-section-columns\ntarget M.S0001\ncount 2\nend\n", TrackChangesMode.Suggest);
        Assert.True(fallback.Success, string.Join("|", fallback.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        Assert.Contains(fallback.Diagnostics, static diagnostic => diagnostic.Code == "W4002");

        using MemoryStream requireInput = CreateDocxWithExistingSectionPropertyRevision();
        DocxTargetCapabilities require = GetCapabilities(requireInput, "M.S0001", TrackChangesMode.Require);
        Assert.Equal("conditional", FindOperation(require, "set-section-columns").Support);

        using MemoryStream requireCheck = CreateDocxWithExistingSectionPropertyRevision();
        DocxCheckResult refused = RunCheck(requireCheck, "docxpatch 1\n\nop set-section-columns\ntarget M.S0001\ncount 2\nend\n", TrackChangesMode.Require);
        Assert.False(refused.Success);
        Assert.Contains(refused.Diagnostics, static diagnostic => diagnostic.Code == "E6002");
    }

    [Fact]
    public static void UnknownSectionFails()
    {
        using MemoryStream input = CreateDocxWithBody(SimpleBody);
        input.Position = 0;
        DocxCapabilitiesResult missing = new DocxEditor().GetCapabilities(input, "M.S0009");
        Assert.False(missing.Success);
        Assert.Null(missing.Capabilities);
        Assert.Contains(missing.Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }
}