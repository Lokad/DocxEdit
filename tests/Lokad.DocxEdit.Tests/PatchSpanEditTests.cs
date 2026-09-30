using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D05: direct run-preserving edits scope protected-boundary checks to the
// actual match spans; paragraph rewrites and tracked output keep the
// whole-paragraph gate.
public static class PatchSpanEditTests
{
    private const string HyperlinkBody = """
                <w:p>
                  <w:r><w:t>See </w:t></w:r>
                  <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rLink">
                    <w:r><w:t>link</w:t></w:r>
                  </w:hyperlink>
                  <w:r><w:t> now</w:t></w:r>
                </w:p>
        """;

    private const string HyperlinkRels = """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/old" TargetMode="External" />
            </Relationships>
        """;

    private const string ComplexFieldBody = """
                <w:p>
                  <w:r><w:t>Before</w:t></w:r>
                  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                  <w:r><w:instrText xml:space="preserve"> REF ClientName \h </w:instrText></w:r>
                  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                  <w:r><w:t>Field</w:t></w:r>
                  <w:r><w:fldChar w:fldCharType="end"/></w:r>
                  <w:r><w:t>After</w:t></w:r>
                </w:p>
        """;

    private static DocxCheckResult RunCheck(MemoryStream input, string patchText, TrackChangesMode mode)
    {
        input.Position = 0;
        using var patch = new StringReader(patchText);
        return new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = mode });
    }

    [Fact]
    public static void PlainSpanBesideHyperlinkEditsInPlace()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(HyperlinkBody, HyperlinkRels);
        DocxCheckResult result = RunCheck(input, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind See\nwith Look\nend\n", TrackChangesMode.Off);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));

        using MemoryStream applyInput = CreateDocxWithBodyAndRelationships(HyperlinkBody, HyperlinkRels);
        using var output = new MemoryStream();
        using var applyPatch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind See\nwith Look\nend\n");
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, applyPatch, output);
        Assert.True(apply.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal("Look link now", Assert.Single(read.Paragraphs).Text);
        DocxHyperlinkInfo hyperlink = Assert.Single(read.Hyperlinks);
        Assert.Equal("https://example.test/old", hyperlink.Uri);
    }

    [Fact]
    public static void SpanInsideHyperlinkStillRefused()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(HyperlinkBody, HyperlinkRels);
        DocxCheckResult result = RunCheck(input, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind link\nwith URL\nend\n", TrackChangesMode.Off);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal("E4305", diagnostic.Code);
        Assert.Contains("hyperlink", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void PlainSpanBesideContentControlEditsInPlace()
    {
        const string body = """
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:sdt>
                        <w:sdtPr><w:tag w:val="note"/></w:sdtPr>
                        <w:sdtContent><w:r><w:t>Inside</w:t></w:r></w:sdtContent>
                      </w:sdt>
                      <w:r><w:t> after</w:t></w:r>
                    </w:p>
            """;
        using MemoryStream input = CreateDocxWithBody(body);
        DocxCheckResult edited = RunCheck(input, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Before\nwith After\nend\n", TrackChangesMode.Off);

        Assert.True(edited.Success, string.Join("|", edited.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));

        using MemoryStream blockedInput = CreateDocxWithBody(body);
        DocxCheckResult blocked = RunCheck(blockedInput, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Inside\nwith Outside\nend\n", TrackChangesMode.Off);

        Assert.False(blocked.Success);
        Assert.Contains(blocked.Diagnostics, static diagnostic => diagnostic.Code == "E4305");
    }

    [Fact]
    public static void DirectEditAfterTrackedEditSucceedsOnUntouchedSpan()
    {
        using MemoryStream first = CreateDocx("Alpha Alpha");
        using var firstPatch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Omega\noccurrence 1\nend\n");
        using var tracked = new MemoryStream();
        DocxApplyResult applied = new DocxEditor().Apply(first, firstPatch, tracked, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });
        Assert.True(applied.Success);

        tracked.Position = 0;
        using var reviewed = new MemoryStream();
        tracked.CopyTo(reviewed);
        reviewed.Position = 0;
        DocxCheckResult second = RunCheck(reviewed, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Delta\nend\n", TrackChangesMode.Off);

        Assert.True(second.Success, string.Join("|", second.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void DirectEditOfInsertedSpanStillRefused()
    {
        using MemoryStream first = CreateDocx("Alpha Alpha");
        using var firstPatch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Omega\noccurrence 1\nend\n");
        using var tracked = new MemoryStream();
        DocxApplyResult applied = new DocxEditor().Apply(first, firstPatch, tracked, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });
        Assert.True(applied.Success);

        tracked.Position = 0;
        using var reviewed = new MemoryStream();
        tracked.CopyTo(reviewed);
        reviewed.Position = 0;
        DocxCheckResult overlap = RunCheck(reviewed, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Omega\nwith Omicron\nend\n", TrackChangesMode.Off);

        Assert.False(overlap.Success);
        DocxDiagnostic diagnostic = Assert.Single(overlap.Diagnostics);
        Assert.Equal("E4305", diagnostic.Code);
        Assert.Contains("tracked-insertion", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void TrackedFollowUpEditSucceedsOnUntouchedSpan()
    {
        using MemoryStream first = CreateDocx("Alpha Alpha");
        using var firstPatch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Omega\noccurrence 1\nend\n");
        using var tracked = new MemoryStream();
        DocxApplyResult applied = new DocxEditor().Apply(first, firstPatch, tracked, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });
        Assert.True(applied.Success);

        tracked.Position = 0;
        using var reviewed = new MemoryStream();
        tracked.CopyTo(reviewed);
        reviewed.Position = 0;
        DocxCheckResult second = RunCheck(reviewed, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Delta\nend\n", TrackChangesMode.Require);

        Assert.True(second.Success, string.Join("|", second.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));

        tracked.Position = 0;
        using var reviewedApply = new MemoryStream();
        tracked.CopyTo(reviewedApply);
        reviewedApply.Position = 0;
        using var output = new MemoryStream();
        using var applyPatch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Delta\nend\n");
        DocxApplyResult secondApply = new DocxEditor().Apply(reviewedApply, applyPatch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });
        Assert.True(secondApply.Success);
        Assert.Equal(["3", "4"], Assert.Single(secondApply.Operations).GeneratedRevisionIds);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:delText>Alpha</w:delText>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>Delta</w:t>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>Omega</w:t>", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.True(new DocxEditor().Validate(output).Success);
    }

    [Fact]
    public static void TrackedOverlapRefusesWithRequire()
    {
        using MemoryStream first = CreateDocx("Alpha Alpha");
        using var firstPatch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Omega\noccurrence 1\nend\n");
        using var tracked = new MemoryStream();
        DocxApplyResult applied = new DocxEditor().Apply(first, firstPatch, tracked, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });
        Assert.True(applied.Success);

        tracked.Position = 0;
        using var reviewed = new MemoryStream();
        tracked.CopyTo(reviewed);
        reviewed.Position = 0;
        DocxCheckResult overlap = RunCheck(reviewed, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Omega\nwith Omicron\nend\n", TrackChangesMode.Require);

        Assert.False(overlap.Success);
        DocxDiagnostic diagnostic = Assert.Single(overlap.Diagnostics);
        Assert.Equal("E6002", diagnostic.Code);
        Assert.Contains("tracked-insertion", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void RewriteWithoutPreserveRunsKeepsWholeParagraphGate()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(HyperlinkBody, HyperlinkRels);
        DocxCheckResult result = RunCheck(input, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind See\nwith Look\npreserve-runs false\nend\n", TrackChangesMode.Off);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E4305");
    }

    [Fact]
    public static void BookmarkMarkersDoNotBlockNearbySpan()
    {
        using MemoryStream input = CreateDocxWithSingleBookmark();
        DocxCheckResult result = RunCheck(input, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Before\nwith Pre\nend\n", TrackChangesMode.Off);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));

        using MemoryStream readInput = CreateDocxWithSingleBookmark();
        using var output = new MemoryStream();
        using var applyPatch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Before\nwith Pre\nend\n");
        Assert.True(new DocxEditor().Apply(readInput, applyPatch, output).Success);
        output.Position = 0;
        Assert.Single(new DocxEditor().Read(output).Bookmarks);
    }

    [Fact]
    public static void SamePatchMixesSuccessFailureAndSkip()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(HyperlinkBody, HyperlinkRels);
        DocxCheckResult result = RunCheck(input, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind See\nwith Look\nend\n\nop replace-text\ntarget M.P0001\nfind link\nwith URL\nend\n\nop replace-text\ntarget M.P0001\nfind now\nwith soon\nend\n", TrackChangesMode.Off);

        Assert.False(result.Success);
        Assert.Equal(3, result.Operations.Count);
        Assert.True(result.Operations[0].Success);
        Assert.False(result.Operations[1].Success);
        Assert.Contains(result.Operations[1].Diagnostics, static diagnostic => diagnostic.Code == "E4305");
        Assert.Contains(result.Operations[2].Diagnostics, static diagnostic => diagnostic.Code == "I0002");
    }

    [Fact]
    public static void SamePatchDisjointPlainSpansBothSucceed()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(HyperlinkBody, HyperlinkRels);
        DocxCheckResult result = RunCheck(input, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind See\nwith Look\nend\n\nop replace-text\ntarget M.P0001\nfind now\nwith soon\nend\n", TrackChangesMode.Off);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }
    [Fact]
    public static void RequirePlainSpanBesideHyperlinkSucceeds()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(HyperlinkBody, HyperlinkRels);
        DocxCheckResult result = RunCheck(input, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind See\nwith Look\nend\n", TrackChangesMode.Require);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));

        using MemoryStream applyInput = CreateDocxWithBodyAndRelationships(HyperlinkBody, HyperlinkRels);
        using var output = new MemoryStream();
        using var applyPatch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind See\nwith Look\nend\n");
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, applyPatch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });
        Assert.True(apply.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:delText>See</w:delText>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>Look</w:t>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:hyperlink", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Equal("Look link now", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }
    [Fact]
    public static void SamePatchDisjointTrackedReplacements()
    {
        using MemoryStream checkInput = CreateDocx("Alpha Alpha");
        DocxCheckResult check = RunCheck(checkInput, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Omega\noccurrence 1\nend\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Delta\nend\n", TrackChangesMode.Require);

        Assert.True(check.Success, string.Join("|", check.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));

        using MemoryStream applyInput = CreateDocx("Alpha Alpha");
        using var output = new MemoryStream();
        using var applyPatch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Omega\noccurrence 1\nend\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Delta\nend\n");
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, applyPatch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.True(apply.Success, string.Join("|", apply.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        output.Position = 0;
        Assert.Equal("Omega Delta", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }
    [Fact]
    public static void SamePatchDisjointTrackedReplacementsUnderSuggest()
    {
        using MemoryStream input = CreateDocx("Alpha Alpha");
        DocxCheckResult check = RunCheck(input, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Omega\noccurrence 1\nend\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Delta\nend\n", TrackChangesMode.Suggest);

        Assert.True(check.Success, string.Join("|", check.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void SamePatchOverlappingTrackedReplacementRefuses()
    {
        using MemoryStream input = CreateDocx("Alpha Alpha");
        DocxCheckResult check = RunCheck(input, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Omega\noccurrence 1\nend\n\nop replace-text\ntarget M.P0001\nfind Omega\nwith Omicron\nend\n", TrackChangesMode.Require);

        Assert.False(check.Success);
        Assert.Contains(check.Diagnostics, static diagnostic => diagnostic.Code == "E4305" || diagnostic.Code == "E6002");
    }

    [Fact]
    public static void BookmarkCrossingSpanRefusesWithE4305()
    {
        using MemoryStream input = CreateDocxWithSingleBookmark();
        DocxCheckResult crossing = RunCheck(input, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Before Old\nwith Pre\nend\n", TrackChangesMode.Off);

        Assert.False(crossing.Success);
        DocxDiagnostic diagnostic = Assert.Single(crossing.Diagnostics);
        Assert.Equal("E4305", diagnostic.Code);
        Assert.Contains("bookmark", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void BookmarkInteriorSpanRefusesWithE4305()
    {
        using MemoryStream input = CreateDocxWithSingleBookmark();
        DocxCheckResult interior = RunCheck(input, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Old Client\nwith New Client\nend\n", TrackChangesMode.Off);

        Assert.False(interior.Success);
        Assert.Contains(interior.Diagnostics, static diagnostic => diagnostic.Code == "E4305");
    }

    [Fact]
    public static void BookmarkCrossingSpanRefusesUnderRequire()
    {
        using MemoryStream input = CreateDocxWithSingleBookmark();
        DocxCheckResult crossing = RunCheck(input, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Before Old\nwith Pre\nend\n", TrackChangesMode.Require);

        Assert.False(crossing.Success);
        DocxDiagnostic diagnostic = Assert.Single(crossing.Diagnostics);
        Assert.Equal("E6002", diagnostic.Code);
        Assert.Contains("bookmark", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void FieldCrossingSpanRefusesWithE4305()
    {
        using MemoryStream input = CreateDocxWithBody(ComplexFieldBody);
        DocxCheckResult crossing = RunCheck(input, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind BeforeField\nwith Changed\nend\n", TrackChangesMode.Off);

        Assert.False(crossing.Success);
        DocxDiagnostic diagnostic = Assert.Single(crossing.Diagnostics);
        Assert.Equal("E4305", diagnostic.Code);
        Assert.Contains("field", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void FieldAdjacentSpanSucceedsAndPreservesField()
    {
        using MemoryStream input = CreateDocxWithBody(ComplexFieldBody);
        using var output = new MemoryStream();
        using var patch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Before\nwith Pre\nend\n");
        Assert.True(new DocxEditor().Apply(input, patch, output).Success);
        output.Position = 0;
        Assert.Equal("PreFieldAfter", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:instrText", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>Field</w:t>", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CommentInteriorSpanRefusesWithE4305()
    {
        using MemoryStream input = CreateDocxWithCommentAnchoredParagraph();
        DocxCheckResult crossing = RunCheck(input, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Comment\nwith Remark\nend\n", TrackChangesMode.Off);

        Assert.False(crossing.Success);
        DocxDiagnostic diagnostic = Assert.Single(crossing.Diagnostics);
        Assert.Equal("E4305", diagnostic.Code);
        Assert.Contains("comment", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void MoveRangeCrossingSpanRefusesWithE4305()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:t>Keep </w:t></w:r>
                      <w:moveToRangeStart w:id="1" w:name="mv"/>
                      <w:r><w:t>Moved</w:t></w:r>
                      <w:moveToRangeEnd w:id="1"/>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
            """);
        DocxCheckResult crossing = RunCheck(input, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Keep Moved\nwith Changed\nend\n", TrackChangesMode.Off);

        Assert.False(crossing.Success);
        DocxDiagnostic diagnostic = Assert.Single(crossing.Diagnostics);
        Assert.Equal("E4305", diagnostic.Code);
        Assert.Contains("tracked-move-to-range", diagnostic.Message, StringComparison.Ordinal);
    }
}
