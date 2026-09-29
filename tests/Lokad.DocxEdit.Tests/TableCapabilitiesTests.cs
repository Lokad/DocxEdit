using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D17: table capabilities reuse the grid-shape, orphan, and track-support
// predicates from the table and text engines.
public static class TableCapabilitiesTests
{
    private const string RaggedBody = """
                    <w:tbl>
                      <w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc></w:tr>
                      <w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
        """;

    private const string OrphanBody = """
                    <w:p><w:r><w:t>Before</w:t></w:r></w:p>
                    <w:tbl>
                      <w:tr><w:tc><w:p><w:bookmarkStart w:id="8" w:name="Cross"/><w:r><w:t>A1</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:p><w:r><w:t>After</w:t></w:r><w:bookmarkEnd w:id="8"/></w:p>
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
    public static void SimpleTableSupportsStructureAndProperties()
    {
        using MemoryStream input = CreateDocxWithSimpleTwoByTwoTable();
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.T0001", TrackChangesMode.Off);

        Assert.Equal("M.T0001", capabilities.TargetId);
        Assert.Equal("table", capabilities.Kind);
        Assert.Equal("main", capabilities.Story);
        string[] expectedOrder = ["insert-before", "insert-after", "insert-hyperlink-after", "delete-block", "set-table-style", "set-table-metadata", "append-row"];
        Assert.Equal(expectedOrder, capabilities.Operations.Select(static operation => operation.Operation).ToArray());
        Assert.Equal("supported", FindOperation(capabilities, "insert-before").Support);
        Assert.Equal("supported", FindOperation(capabilities, "insert-after").Support);
        Assert.Equal("supported", FindOperation(capabilities, "insert-hyperlink-after").Support);
        Assert.Equal("supported", FindOperation(capabilities, "delete-block").Support);
        Assert.Equal("conditional", FindOperation(capabilities, "set-table-style").Support);
        Assert.Equal("supported", FindOperation(capabilities, "set-table-metadata").Support);
        Assert.Equal("conditional", FindOperation(capabilities, "append-row").Support);

        using MemoryStream checkInput = CreateDocxWithSimpleTwoByTwoTable();
        DocxCheckResult result = RunCheck(checkInput, "docxpatch 1\n\nop append-row\ntarget M.T0001\ncell East\ncell West\nend\n", TrackChangesMode.Off);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void RequireModeRefusesDeleteAndMetadata()
    {
        using MemoryStream input = CreateDocxWithSimpleTwoByTwoTable();
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.T0001", TrackChangesMode.Require);

        Assert.Equal("supported", FindOperation(capabilities, "insert-before").Support);
        Assert.Equal("unsupported", FindOperation(capabilities, "delete-block").Support);
        Assert.Equal("conditional", FindOperation(capabilities, "set-table-style").Support);
        Assert.Equal("unsupported", FindOperation(capabilities, "set-table-metadata").Support);
        Assert.Equal("conditional", FindOperation(capabilities, "append-row").Support);

        using MemoryStream deleteCheck = CreateDocxWithSimpleTwoByTwoTable();
        DocxCheckResult deleted = RunCheck(deleteCheck, "docxpatch 1\n\nop delete-block\ntarget M.T0001\nend\n", TrackChangesMode.Require);
        Assert.False(deleted.Success);
        Assert.Contains(deleted.Diagnostics, static diagnostic => diagnostic.Code == "E6002");

        using MemoryStream metadataCheck = CreateDocxWithSimpleTwoByTwoTable();
        DocxCheckResult metadata = RunCheck(metadataCheck, "docxpatch 1\n\nop set-table-metadata\ntarget M.T0001\ncaption Ledger\nend\n", TrackChangesMode.Require);
        Assert.False(metadata.Success);
        Assert.Contains(metadata.Diagnostics, static diagnostic => diagnostic.Code == "E6001");

        using MemoryStream appendCheck = CreateDocxWithSimpleTwoByTwoTable();
        DocxCheckResult appended = RunCheck(appendCheck, "docxpatch 1\n\nop append-row\ntarget M.T0001\ncell East\ncell West\nend\n", TrackChangesMode.Require);
        Assert.True(appended.Success, string.Join("|", appended.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void UnknownTableFails()
    {
        using MemoryStream input = CreateDocxWithSimpleTwoByTwoTable();
        input.Position = 0;
        DocxCapabilitiesResult missing = new DocxEditor().GetCapabilities(input, "M.T0009");
        Assert.False(missing.Success);
        Assert.Null(missing.Capabilities);
        Assert.Contains(missing.Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }

    [Fact]
    public static void RaggedTableRefusesAppendRow()
    {
        using MemoryStream input = CreateDocxWithBody(RaggedBody);
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.T0001", TrackChangesMode.Off);

        Assert.Equal("unsupported", FindOperation(capabilities, "append-row").Support);
        Assert.Equal("supported", FindOperation(capabilities, "insert-after").Support);

        using MemoryStream checkInput = CreateDocxWithBody(RaggedBody);
        DocxCheckResult refused = RunCheck(checkInput, "docxpatch 1\n\nop append-row\ntarget M.T0001\ncell Solo\nend\n", TrackChangesMode.Off);
        Assert.False(refused.Success);
        Assert.Contains(refused.Diagnostics, static diagnostic => diagnostic.Code == "E4301");
    }

    [Fact]
    public static void OrphanedTableDeleteRefuses()
    {
        using MemoryStream input = CreateDocxWithBody(OrphanBody);
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.T0001", TrackChangesMode.Off);

        Assert.Equal("unsupported", FindOperation(capabilities, "delete-block").Support);

        using MemoryStream checkInput = CreateDocxWithBody(OrphanBody);
        DocxCheckResult refused = RunCheck(checkInput, "docxpatch 1\n\nop delete-block\ntarget M.T0001\nend\n", TrackChangesMode.Off);
        Assert.False(refused.Success);
        Assert.Contains(refused.Diagnostics, static diagnostic => diagnostic.Code == "E4305");
    }
}