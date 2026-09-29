using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D17: row capabilities reuse the insertion-boundary, deletion, grid-shape,
// orphan, and track-support predicates from the table and text engines.
public static class RowCapabilitiesTests
{
    private const string MergeRowsBody = """
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

    private const string SingleRowBody = """
                    <w:tbl>
                      <w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
        """;

    private const string OrphanRowBody = """
                    <w:tbl>
                      <w:tr><w:tc><w:p><w:bookmarkStart w:id="9" w:name="Cross"/><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>
                      <w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:p><w:r><w:t>After</w:t></w:r><w:bookmarkEnd w:id="9"/></w:p>
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
    public static void SimpleRowSupportsInsertsDeleteAndHeader()
    {
        using MemoryStream input = CreateDocxWithSimpleTwoByTwoTable();
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.T0001.R01", TrackChangesMode.Off);

        Assert.Equal("M.T0001.R01", capabilities.TargetId);
        Assert.Equal("row", capabilities.Kind);
        Assert.Equal("main", capabilities.Story);
        string[] expectedOrder = ["insert-row-before", "insert-row-after", "delete-row", "set-row-header"];
        Assert.Equal(expectedOrder, capabilities.Operations.Select(static operation => operation.Operation).ToArray());
        Assert.Equal("conditional", FindOperation(capabilities, "insert-row-before").Support);
        Assert.Equal("conditional", FindOperation(capabilities, "insert-row-after").Support);
        Assert.Equal("supported", FindOperation(capabilities, "delete-row").Support);
        Assert.Equal("supported", FindOperation(capabilities, "set-row-header").Support);

        using MemoryStream deleteInput = CreateDocxWithSimpleTwoByTwoTable();
        DocxCheckResult deleted = RunCheck(deleteInput, "docxpatch 1\n\nop delete-row\ntarget M.T0001.R01\nend\n", TrackChangesMode.Off);
        Assert.True(deleted.Success, string.Join("|", deleted.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));

        using MemoryStream headerInput = CreateDocxWithSimpleTwoByTwoTable();
        DocxCheckResult header = RunCheck(headerInput, "docxpatch 1\n\nop set-row-header\ntarget M.T0001.R01\nheader true\nend\n", TrackChangesMode.Off);
        Assert.True(header.Success, string.Join("|", header.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void TrackedDeleteIsConditional()
    {
        using MemoryStream input = CreateDocxWithSimpleTwoByTwoTable();
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.T0001.R02", TrackChangesMode.Require);

        Assert.Equal("conditional", FindOperation(capabilities, "delete-row").Support);

        using MemoryStream checkInput = CreateDocxWithSimpleTwoByTwoTable();
        DocxCheckResult deleted = RunCheck(checkInput, "docxpatch 1\n\nop delete-row\ntarget M.T0001.R02\nend\n", TrackChangesMode.Require);
        Assert.True(deleted.Success, string.Join("|", deleted.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void UnknownRowFails()
    {
        using MemoryStream input = CreateDocxWithSimpleTwoByTwoTable();
        input.Position = 0;
        DocxCapabilitiesResult missing = new DocxEditor().GetCapabilities(input, "M.T0001.R09");
        Assert.False(missing.Success);
        Assert.Null(missing.Capabilities);
        Assert.Contains(missing.Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }

    [Fact]
    public static void LastRowDeleteRefuses()
    {
        using MemoryStream input = CreateDocxWithBody(SingleRowBody);
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.T0001.R01", TrackChangesMode.Off);

        Assert.Equal("unsupported", FindOperation(capabilities, "delete-row").Support);

        using MemoryStream checkInput = CreateDocxWithBody(SingleRowBody);
        DocxCheckResult refused = RunCheck(checkInput, "docxpatch 1\n\nop delete-row\ntarget M.T0001.R01\nend\n", TrackChangesMode.Off);
        Assert.False(refused.Success);
        Assert.Contains(refused.Diagnostics, static diagnostic => diagnostic.Code == "E4304");
    }

    [Fact]
    public static void MergeBoundaryRefusesInsertButPromotesRootDelete()
    {
        using MemoryStream input = CreateDocxWithBody(MergeRowsBody);
        DocxTargetCapabilities after = GetCapabilities(input, "M.T0001.R01", TrackChangesMode.Off);
        Assert.Equal("unsupported", FindOperation(after, "insert-row-after").Support);

        using MemoryStream deleteInput = CreateDocxWithBody(MergeRowsBody);
        DocxTargetCapabilities promoted = GetCapabilities(deleteInput, "M.T0001.R01", TrackChangesMode.Off);
        Assert.Equal("supported", FindOperation(promoted, "delete-row").Support);

        using MemoryStream insertCheck = CreateDocxWithBody(MergeRowsBody);
        DocxCheckResult insertRefused = RunCheck(insertCheck, "docxpatch 1\n\nop insert-row-after\ntarget M.T0001.R01\ncell X\ncell Y\nend\n", TrackChangesMode.Off);
        Assert.False(insertRefused.Success);
        Assert.Contains(insertRefused.Diagnostics, static diagnostic => diagnostic.Code == "E4301");

        using MemoryStream deleteCheck = CreateDocxWithBody(MergeRowsBody);
        DocxCheckResult promotedDelete = RunCheck(deleteCheck, "docxpatch 1\n\nop delete-row\ntarget M.T0001.R01\nend\n", TrackChangesMode.Off);
        Assert.True(promotedDelete.Success, string.Join("|", promotedDelete.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));

    }

    [Fact]
    public static void OrphanedRowDeleteRefusesDirect()
    {
        using MemoryStream input = CreateDocxWithBody(OrphanRowBody);
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.T0001.R01", TrackChangesMode.Off);

        Assert.Equal("unsupported", FindOperation(capabilities, "delete-row").Support);

        using MemoryStream checkInput = CreateDocxWithBody(OrphanRowBody);
        DocxCheckResult refused = RunCheck(checkInput, "docxpatch 1\n\nop delete-row\ntarget M.T0001.R01\nend\n", TrackChangesMode.Off);
        Assert.False(refused.Success);
        Assert.Contains(refused.Diagnostics, static diagnostic => diagnostic.Code == "E4305");
    }
}
