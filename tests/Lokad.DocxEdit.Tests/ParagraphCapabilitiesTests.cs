using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D17: paragraph editing capabilities reuse the execution predicates, so
// reported support stays consistent with check outcomes.
public static class ParagraphCapabilitiesTests
{
    private static DocxOperationCapability FindOperation(DocxTargetCapabilities capabilities, string operation)
    {
        return capabilities.Operations.First(candidate => string.Equals(candidate.Operation, operation, StringComparison.Ordinal));
    }

    private static DocxTargetCapabilities GetParagraphCapabilities(MemoryStream input, string targetId, TrackChangesMode mode)
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
    public static void SimpleParagraphSupportsEveryOperationInDirectMode()
    {
        using MemoryStream input = CreateDocx("Alpha Beta");
        DocxTargetCapabilities capabilities = GetParagraphCapabilities(input, "M.P0001", TrackChangesMode.Off);

        Assert.Equal("M.P0001", capabilities.TargetId);
        Assert.Equal("paragraph", capabilities.Kind);
        Assert.Equal("main", capabilities.Story);
        string[] expectedOrder = ["replace-text", "replace-paragraph", "insert-before", "insert-after", "delete-block", "set-style", "add-bookmark", "add-comment"];
        Assert.Equal(expectedOrder, capabilities.Operations.Select(static operation => operation.Operation).ToArray());
        foreach (DocxOperationCapability operation in capabilities.Operations)
        {
            Assert.Equal("supported", operation.Support);
            Assert.Equal(operation.Operation, operation.HelpTopic);
            Assert.False(string.IsNullOrWhiteSpace(operation.Reason));
        }
    }

    [Fact]
    public static void HyperlinkParagraphRefusesTextEditsInEveryMode()
    {
        TrackChangesMode[] modes = [TrackChangesMode.Off, TrackChangesMode.Suggest, TrackChangesMode.Require];
        foreach (TrackChangesMode mode in modes)
        {
            using MemoryStream input = CreateDocxWithBodyAndRelationships(
                """
                        <w:p>
                          <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rLink">
                            <w:r><w:t>Old link</w:t></w:r>
                          </w:hyperlink>
                        </w:p>
                """,
                """
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                      <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/old" TargetMode="External" />
                    </Relationships>
                """);
            DocxTargetCapabilities capabilities = GetParagraphCapabilities(input, "M.P0001", mode);

            Assert.Equal("conditional", FindOperation(capabilities, "replace-text").Support);
            Assert.Contains("hyperlink", FindOperation(capabilities, "replace-text").Reason, StringComparison.Ordinal);
            Assert.Equal("insert-after", FindOperation(capabilities, "replace-text").Alternative);
            Assert.Equal("unsupported", FindOperation(capabilities, "replace-paragraph").Support);
            Assert.Equal("conditional", FindOperation(capabilities, "add-comment").Support);
            string expectedDelete = mode == TrackChangesMode.Require ? "unsupported" : "supported";             Assert.Equal(expectedDelete, FindOperation(capabilities, "delete-block").Support);
            Assert.Equal("supported", FindOperation(capabilities, "set-style").Support);
            Assert.Equal("supported", FindOperation(capabilities, "insert-before").Support);
        }

        using MemoryStream refuseInput = CreateDocxWithBodyAndRelationships(
            """
                    <w:p>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rLink">
                        <w:r><w:t>Old link</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/old" TargetMode="External" />
                </Relationships>
            """);
        DocxCheckResult refused = RunCheck(refuseInput, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Old\nwith New\nend\n", TrackChangesMode.Off);
        Assert.False(refused.Success);
        Assert.Contains(refused.Diagnostics, static diagnostic => diagnostic.Code == "E4305");

        using MemoryStream deleteInput = CreateDocxWithBodyAndRelationships(
            """
                    <w:p>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rLink">
                        <w:r><w:t>Old link</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/old" TargetMode="External" />
                </Relationships>
            """);
        DocxCheckResult deleted = RunCheck(deleteInput, "docxpatch 1\n\nop delete-block\ntarget M.P0001\nend\n", TrackChangesMode.Off);
        Assert.True(deleted.Success, string.Join("|", deleted.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void MixedFormattingMakesTrackedReplaceConditional()
    {
        const string body = """
                <w:p>
                  <w:r><w:rPr><w:b /></w:rPr><w:t>Alpha </w:t></w:r>
                  <w:r><w:t>Beta</w:t></w:r>
                </w:p>
            """;
        using MemoryStream offInput = CreateDocxWithBody(body);
        Assert.Equal("supported", FindOperation(GetParagraphCapabilities(offInput, "M.P0001", TrackChangesMode.Off), "replace-text").Support);

        using MemoryStream suggestInput = CreateDocxWithBody(body);
        Assert.Equal("conditional", FindOperation(GetParagraphCapabilities(suggestInput, "M.P0001", TrackChangesMode.Suggest), "replace-text").Support);

        using MemoryStream requireInput = CreateDocxWithBody(body);
        Assert.Equal("unsupported", FindOperation(GetParagraphCapabilities(requireInput, "M.P0001", TrackChangesMode.Require), "replace-text").Support);

        using MemoryStream checkInput = CreateDocxWithBody(body);
        DocxCheckResult refused = RunCheck(checkInput, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Omega\nend\n", TrackChangesMode.Require);
        Assert.True(refused.Success, string.Join("|", refused.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void OrphanedBookmarkFailsDeleteInEveryMode()
    {
        const string body = """
                <w:p><w:bookmarkStart w:id="7" w:name="SpanMark" /><w:r><w:t>First</w:t></w:r></w:p>
                <w:p><w:r><w:t>Second</w:t></w:r><w:bookmarkEnd w:id="7" /></w:p>
            """;
        TrackChangesMode[] modes = [TrackChangesMode.Off, TrackChangesMode.Suggest, TrackChangesMode.Require];
        foreach (TrackChangesMode mode in modes)
        {
            using MemoryStream input = CreateDocxWithBody(body);
            DocxTargetCapabilities capabilities = GetParagraphCapabilities(input, "M.P0001", mode);
            DocxOperationCapability delete = FindOperation(capabilities, "delete-block");
            Assert.Equal("unsupported", delete.Support);
            Assert.Contains("orphan", delete.Reason, StringComparison.Ordinal);
        }

        using MemoryStream checkInput = CreateDocxWithBody(body);
        DocxCheckResult refused = RunCheck(checkInput, "docxpatch 1\n\nop delete-block\ntarget M.P0001\nend\n", TrackChangesMode.Off);
        Assert.False(refused.Success);
        Assert.Contains(refused.Diagnostics, static diagnostic => diagnostic.Code == "E4305");
    }

    [Fact]
    public static void RequireRefusesBookmarkMetadataButPermitsComments()
    {
        using MemoryStream requireInput = CreateDocx("Alpha Beta");
        DocxTargetCapabilities requireCapabilities = GetParagraphCapabilities(requireInput, "M.P0001", TrackChangesMode.Require);
        DocxOperationCapability bookmark = FindOperation(requireCapabilities, "add-bookmark");
        Assert.Equal("unsupported", bookmark.Support);
        Assert.Contains("E6001", bookmark.Reason, StringComparison.Ordinal);
        Assert.Equal("supported", FindOperation(requireCapabilities, "add-comment").Support);

        using MemoryStream offInput = CreateDocx("Alpha Beta");
        Assert.Equal("supported", FindOperation(GetParagraphCapabilities(offInput, "M.P0001", TrackChangesMode.Off), "add-bookmark").Support);
    }

    [Fact]
    public static void SupportedCapabilitiesAgreeWithCheck()
    {
        const string styles = """
              <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal" /></w:style>
              <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2" /></w:style>
            """;
        const string body = """
                <w:p><w:r><w:t>Styled paragraph</w:t></w:r></w:p>
            """;
        (string Operation, string Patch)[] cases =
        [
            ("replace-text", "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Styled\nwith Restyled\nend\n"),
            ("replace-paragraph", "docxpatch 1\n\nop replace-paragraph\ntarget M.P0001\ntext Fresh paragraph\nend\n"),
            ("insert-before", "docxpatch 1\n\nop insert-before\ntarget M.P0001\ntext Neighbor before\nend\n"),
            ("insert-after", "docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext Neighbor after\nend\n"),
            ("delete-block", "docxpatch 1\n\nop delete-block\ntarget M.P0001\nend\n"),
            ("set-style", "docxpatch 1\n\nop set-style\ntarget M.P0001\nstyle Heading 2\nend\n"),
            ("add-comment", "docxpatch 1\n\nop add-comment\ntarget M.P0001\ntext Review note\nend\n"),
            ("add-bookmark", "docxpatch 1\n\nop add-bookmark\ntarget M.P0001\nname CapMark\nend\n"),
        ];

        using (MemoryStream input = CreateDocxWithStylesAndBody(styles, body))
        {
            DocxTargetCapabilities capabilities = GetParagraphCapabilities(input, "M.P0001", TrackChangesMode.Off);
            foreach ((string Operation, string Patch) in cases)
            {
                Assert.Equal("supported", FindOperation(capabilities, Operation).Support);
            }
        }

        foreach ((string Operation, string Patch) in cases)
        {
            using MemoryStream input = CreateDocxWithStylesAndBody(styles, body);
            DocxCheckResult result = RunCheck(input, Patch, TrackChangesMode.Off);
            Assert.True(result.Success, Operation + ":" + string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        }

        foreach ((string Operation, string Patch) in cases)
        {
            using MemoryStream input = CreateDocxWithStylesAndBody(styles, body);
            DocxCheckResult result = RunCheck(input, Patch, TrackChangesMode.Require);
            if (Operation == "add-bookmark")
            {
                Assert.False(result.Success);
                Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E6001");
            }
            else
            {
                Assert.True(result.Success, Operation + ":" + string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
            }
        }
    }

    [Fact]
    public static void UnknownAndNonParagraphTargetsFailWithE1201()
    {
        using MemoryStream missingInput = CreateDocx("Alpha");
        missingInput.Position = 0;
        DocxCapabilitiesResult missing = new DocxEditor().GetCapabilities(missingInput, "M.P0009");
        Assert.False(missing.Success);
        Assert.Null(missing.Capabilities);
        Assert.Equal("M.P0009", missing.TargetId);
        Assert.Contains(missing.Diagnostics, static diagnostic => diagnostic.Code == "E1201");

        using MemoryStream tableInput = CreateDocxWithSimpleTwoByTwoTable();
        tableInput.Position = 0;
        DocxCapabilitiesResult table = new DocxEditor().GetCapabilities(tableInput, "M.T0001");
        Assert.False(table.Success);
        Assert.Null(table.Capabilities);
        Assert.Contains(table.Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }
}
