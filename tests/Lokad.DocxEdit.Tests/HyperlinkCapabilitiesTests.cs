using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D17: hyperlink capabilities reuse the protection and track-support
// predicates from the hyperlink engine.
public static class HyperlinkCapabilitiesTests
{
    private const string LinkBody = """
                    <w:p>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rLink">
                        <w:r><w:t>Old link</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
        """;

    private const string LinkRelationships = """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/old" TargetMode="External" />
                </Relationships>
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
    public static void SimpleHyperlinkSupportsTargetTextAndRemove()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(LinkBody, LinkRelationships);
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.L0001", TrackChangesMode.Off);

        Assert.Equal("M.L0001", capabilities.TargetId);
        Assert.Equal("hyperlink", capabilities.Kind);
        Assert.Equal("main", capabilities.Story);
        string[] expectedOrder = ["set-hyperlink-target", "set-hyperlink-text", "remove-hyperlink"];
        Assert.Equal(expectedOrder, capabilities.Operations.Select(static operation => operation.Operation).ToArray());
        Assert.Equal("supported", FindOperation(capabilities, "set-hyperlink-target").Support);
        Assert.Equal("supported", FindOperation(capabilities, "set-hyperlink-text").Support);
        Assert.Equal("supported", FindOperation(capabilities, "remove-hyperlink").Support);
        Assert.Contains("an external URI", FindOperation(capabilities, "set-hyperlink-target").Reason, StringComparison.Ordinal);

        using MemoryStream textInput = CreateDocxWithBodyAndRelationships(LinkBody, LinkRelationships);
        DocxCheckResult text = RunCheck(textInput, "docxpatch 1\n\nop set-hyperlink-text\ntarget M.L0001\ntext New link\nend\n", TrackChangesMode.Off);
        Assert.True(text.Success, string.Join("|", text.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));

        using MemoryStream removeInput = CreateDocxWithBodyAndRelationships(LinkBody, LinkRelationships);
        DocxCheckResult removed = RunCheck(removeInput, "docxpatch 1\n\nop remove-hyperlink\ntarget M.L0001\nend\n", TrackChangesMode.Off);
        Assert.True(removed.Success, string.Join("|", removed.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void RequireModeRefusesMetadataOps()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(LinkBody, LinkRelationships);
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.L0001", TrackChangesMode.Require);

        Assert.Equal("unsupported", FindOperation(capabilities, "set-hyperlink-target").Support);
        Assert.Equal("conditional", FindOperation(capabilities, "set-hyperlink-text").Support);
        Assert.Equal("unsupported", FindOperation(capabilities, "remove-hyperlink").Support);

        using MemoryStream targetCheck = CreateDocxWithBodyAndRelationships(LinkBody, LinkRelationships);
        DocxCheckResult target = RunCheck(targetCheck, "docxpatch 1\n\nop set-hyperlink-target\ntarget M.L0001\nuri https://example.test/new\nend\n", TrackChangesMode.Require);
        Assert.False(target.Success);
        Assert.Contains(target.Diagnostics, static diagnostic => diagnostic.Code == "E6001");

        using MemoryStream removeCheck = CreateDocxWithBodyAndRelationships(LinkBody, LinkRelationships);
        DocxCheckResult removed = RunCheck(removeCheck, "docxpatch 1\n\nop remove-hyperlink\ntarget M.L0001\nend\n", TrackChangesMode.Require);
        Assert.False(removed.Success);
        Assert.Contains(removed.Diagnostics, static diagnostic => diagnostic.Code == "E6001");
    }

    [Fact]
    public static void SuggestModeTextIsConditional()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(LinkBody, LinkRelationships);
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.L0001", TrackChangesMode.Suggest);

        Assert.Equal("conditional", FindOperation(capabilities, "set-hyperlink-text").Support);

        using MemoryStream checkInput = CreateDocxWithBodyAndRelationships(LinkBody, LinkRelationships);
        DocxCheckResult text = RunCheck(checkInput, "docxpatch 1\n\nop set-hyperlink-text\ntarget M.L0001\ntext New link\nend\n", TrackChangesMode.Suggest);
        Assert.True(text.Success, string.Join("|", text.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void UnknownHyperlinkFails()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(LinkBody, LinkRelationships);
        input.Position = 0;
        DocxCapabilitiesResult missing = new DocxEditor().GetCapabilities(input, "M.L0009");
        Assert.False(missing.Success);
        Assert.Null(missing.Capabilities);
        Assert.Contains(missing.Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }
}