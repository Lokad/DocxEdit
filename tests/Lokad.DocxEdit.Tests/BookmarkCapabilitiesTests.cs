using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D17: bookmark capabilities reuse the range-shape, protection,
// hyperlink-reference, and track-support predicates from the bookmark engine.
public static class BookmarkCapabilitiesTests
{
    private const string ReferencedBody = """
                    <w:p>
                      <w:bookmarkStart w:id="4" w:name="ClientName"/>
                      <w:r><w:t>Old Client</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                    </w:p>
                    <w:p>
                      <w:hyperlink w:anchor="ClientName"><w:r><w:t>jump</w:t></w:r></w:hyperlink>
                    </w:p>
        """;

    private const string IncompleteBody = """
                    <w:p>
                      <w:bookmarkStart w:id="5" w:name="Orphan"/>
                      <w:r><w:t>text</w:t></w:r>
                    </w:p>
        """;

    private const string MultiParagraphBody = """
                    <w:p>
                      <w:bookmarkStart w:id="6" w:name="Span"/>
                      <w:r><w:t>First</w:t></w:r>
                    </w:p>
                    <w:p>
                      <w:r><w:t>Second</w:t></w:r>
                      <w:bookmarkEnd w:id="6"/>
                    </w:p>
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
    public static void SimpleBookmarkSupportsReplaceRenameDelete()
    {
        using MemoryStream input = CreateDocxWithSingleBookmark();
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.B0001", TrackChangesMode.Off);

        Assert.Equal("M.B0001", capabilities.TargetId);
        Assert.Equal("bookmark", capabilities.Kind);
        Assert.Equal("main", capabilities.Story);
        string[] expectedOrder = ["replace-bookmark-text", "rename-bookmark", "delete-bookmark"];
        Assert.Equal(expectedOrder, capabilities.Operations.Select(static operation => operation.Operation).ToArray());
        Assert.Equal("supported", FindOperation(capabilities, "replace-bookmark-text").Support);
        Assert.Equal("supported", FindOperation(capabilities, "rename-bookmark").Support);
        Assert.Equal("supported", FindOperation(capabilities, "delete-bookmark").Support);
        Assert.Contains("expect-name", FindOperation(capabilities, "delete-bookmark").Reason, StringComparison.Ordinal);

        using MemoryStream checkInput = CreateDocxWithSingleBookmark();
        DocxCheckResult result = RunCheck(checkInput, "docxpatch 1\n\nop replace-bookmark-text\ntarget M.B0001\ntext New Client\nend\n", TrackChangesMode.Off);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void TrackedReplaceIsConditionalButRenameDeleteRefuseRequire()
    {
        using MemoryStream suggestInput = CreateDocxWithSingleBookmark();
        Assert.Equal("conditional", FindOperation(GetCapabilities(suggestInput, "M.B0001", TrackChangesMode.Suggest), "replace-bookmark-text").Support);

        using MemoryStream requireInput = CreateDocxWithSingleBookmark();
        DocxTargetCapabilities require = GetCapabilities(requireInput, "M.B0001", TrackChangesMode.Require);
        Assert.Equal("conditional", FindOperation(require, "replace-bookmark-text").Support);
        Assert.Equal("unsupported", FindOperation(require, "rename-bookmark").Support);
        Assert.Equal("unsupported", FindOperation(require, "delete-bookmark").Support);

        using MemoryStream suggestCheck = CreateDocxWithSingleBookmark();
        DocxCheckResult suggest = RunCheck(suggestCheck, "docxpatch 1\n\nop replace-bookmark-text\ntarget M.B0001\ntext New Client\nend\n", TrackChangesMode.Suggest);
        Assert.True(suggest.Success, string.Join("|", suggest.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));

        using MemoryStream renameCheck = CreateDocxWithSingleBookmark();
        DocxCheckResult rename = RunCheck(renameCheck, "docxpatch 1\n\nop rename-bookmark\ntarget M.B0001\nname FreshName\nend\n", TrackChangesMode.Require);
        Assert.False(rename.Success);
        Assert.Contains(rename.Diagnostics, static diagnostic => diagnostic.Code == "E6001");

        using MemoryStream deleteCheck = CreateDocxWithSingleBookmark();
        DocxCheckResult delete = RunCheck(deleteCheck, "docxpatch 1\n\nop delete-bookmark\ntarget M.B0001\nend\n", TrackChangesMode.Require);
        Assert.False(delete.Success);
        Assert.Contains(delete.Diagnostics, static diagnostic => diagnostic.Code == "E6001");
    }

    [Fact]
    public static void UnknownBookmarkFails()
    {
        using MemoryStream input = CreateDocxWithSingleBookmark();
        input.Position = 0;
        DocxCapabilitiesResult missing = new DocxEditor().GetCapabilities(input, "M.B0009");
        Assert.False(missing.Success);
        Assert.Null(missing.Capabilities);
        Assert.Contains(missing.Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }

    [Fact]
    public static void IncompleteBookmarkRefusesReplaceAndDelete()
    {
        using MemoryStream input = CreateDocxWithBody(IncompleteBody);
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.B0001", TrackChangesMode.Off);

        Assert.Equal("unsupported", FindOperation(capabilities, "replace-bookmark-text").Support);
        Assert.Equal("unsupported", FindOperation(capabilities, "delete-bookmark").Support);
        Assert.Equal("supported", FindOperation(capabilities, "rename-bookmark").Support);

        using MemoryStream checkInput = CreateDocxWithBody(IncompleteBody);
        DocxCheckResult refused = RunCheck(checkInput, "docxpatch 1\n\nop delete-bookmark\ntarget M.B0001\nend\n", TrackChangesMode.Off);
        Assert.False(refused.Success);
        Assert.Contains(refused.Diagnostics, static diagnostic => diagnostic.Code == "E4311");
    }

    [Fact]
    public static void ReferencedBookmarkRefusesDelete()
    {
        using MemoryStream input = CreateDocxWithBody(ReferencedBody);
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.B0001", TrackChangesMode.Off);

        Assert.Equal("supported", FindOperation(capabilities, "replace-bookmark-text").Support);
        Assert.Equal("unsupported", FindOperation(capabilities, "delete-bookmark").Support);

        using MemoryStream checkInput = CreateDocxWithBody(ReferencedBody);
        DocxCheckResult refused = RunCheck(checkInput, "docxpatch 1\n\nop delete-bookmark\ntarget M.B0001\nend\n", TrackChangesMode.Off);
        Assert.False(refused.Success);
        Assert.Contains(refused.Diagnostics, static diagnostic => diagnostic.Code == "E4311");
    }

    [Fact]
    public static void MultiParagraphBookmarkReplaceIsConditional()
    {
        using MemoryStream input = CreateDocxWithBody(MultiParagraphBody);
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.B0001", TrackChangesMode.Off);

        Assert.Equal("conditional", FindOperation(capabilities, "replace-bookmark-text").Support);
    }
}