using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D14: reports identify resolved targets instead of echoing only the
// selector, and never present simulated output as committed output.
public static class PatchReportTests
{
    [Fact]
    public static void CheckReplaceTextReportsResolvedParagraph()
    {
        using MemoryStream input = CreateDocx("Alpha Beta");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            with Omega
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success);
        DocxPatchAffectedTarget affected = Assert.Single(Assert.Single(result.Operations).AffectedTargets);
        Assert.Equal("M.P0001", affected.Id.ToWireValue());
        Assert.Equal("paragraph", affected.Kind);
        Assert.Equal("update", affected.Action);
    }

    [Fact]
    public static void CheckReplaceTextWithSemanticSelectorReportsResolvedParagraph()
    {
        using MemoryStream input = CreateDocxWithBody("""
              <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
              <w:p><w:r><w:t>Beta</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target text:"Beta"
            find Beta
            with Gamma
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchAffectedTarget affected = Assert.Single(Assert.Single(result.Operations).AffectedTargets);
        Assert.Equal("M.P0002", affected.Id.ToWireValue());
    }

    [Fact]
    public static void CheckAndApplyAgreeOnAffectedTargets()
    {
        const string patchText = "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Omega\nend\n";
        using MemoryStream checkInput = CreateDocx("Alpha Beta");
        IReadOnlyList<string> checkAffected = Assert.Single(new DocxEditor().Check(checkInput, new StringReader(patchText)).Operations).AffectedTargets.Select(static target => target.Id.ToWireValue()).ToArray();
        using MemoryStream applyInput = CreateDocx("Alpha Beta");
        using var output = new MemoryStream();
        IReadOnlyList<string> applyAffected = Assert.Single(new DocxEditor().Apply(applyInput, new StringReader(patchText), output).Operations).AffectedTargets.Select(static target => target.Id.ToWireValue()).ToArray();
        Assert.Equal(checkAffected, applyAffected);
        Assert.Equal(new[] { "M.P0001" }, applyAffected);
    }

    [Fact]
    public static void ApplyInsertAfterReportsAnchorAsInsert()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text Inserted
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        DocxPatchAffectedTarget affected = Assert.Single(Assert.Single(result.Operations).AffectedTargets);
        Assert.Equal("M.P0001", affected.Id.ToWireValue());
        Assert.Equal("insert", affected.Action);
    }

    [Fact]
    public static void ApplyDeleteBlockReportsDeletedParagraph()
    {
        using MemoryStream input = CreateDocxWithBody("""
              <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
              <w:p><w:r><w:t>Beta</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-block
            target M.P0002
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        DocxPatchAffectedTarget affected = Assert.Single(Assert.Single(result.Operations).AffectedTargets);
        Assert.Equal("M.P0002", affected.Id.ToWireValue());
        Assert.Equal("delete", affected.Action);
    }

    [Fact]
    public static void ApplyInsertAfterOnTableReportsTableAnchor()
    {
        using MemoryStream input = CreateDocxWithBody("""
              <w:tbl>
                <w:tblGrid><w:gridCol/></w:tblGrid>
                <w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>
              </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.T0001
            text Inserted
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        DocxPatchAffectedTarget affected = Assert.Single(Assert.Single(result.Operations).AffectedTargets);
        Assert.Equal("M.T0001", affected.Id.ToWireValue());
        Assert.Equal("table", affected.Kind);
    }

    [Fact]
    public static void FailedApplyClearsGeneratedRevisionIds()
    {
        var options = new DocxEditOptions { TrackChanges = TrackChangesMode.Require };
        using MemoryStream input = CreateDocx("Alpha Beta");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            with Omega
            end

            op replace-text
            target M.P9999
            find Beta
            with Gamma
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, options);

        Assert.False(result.Success);
        Assert.All(result.Operations, static operation => Assert.Empty(operation.GeneratedRevisionIds));
    }

    [Fact]
    public static void NoOpReportsNoAffectedTargets()
    {
        using MemoryStream input = CreateDocx("Anchor");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Anchor
            with Anchor
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        Assert.Contains(result.Diagnostics, static d => d.Code == "I0001");
        Assert.Empty(Assert.Single(result.Operations).AffectedTargets);
    }

    [Fact]
    public static void PreviewStaysOffByDefault()
    {
        using MemoryStream input = CreateDocx("Alpha Beta");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            with Omega
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Null(report.PreviewBefore);
        Assert.Null(report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }

    [Fact]
    public static void PreviewReportsBeforeAndAfterText()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocx("Alpha Beta");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            with Omega
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success);
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal("Alpha Beta", report.PreviewBefore);
        Assert.Equal("Omega Beta", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }

    [Fact]
    public static void PreviewTruncatesToBudget()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 4 };
        using MemoryStream input = CreateDocx("Alpha Beta Gamma");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            with Omega
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success);
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal("Alph", report.PreviewBefore);
        Assert.Equal("Omeg", report.PreviewAfter);
        Assert.True(report.PreviewTruncated);
    }

    [Fact]
    public static void PreviewReportsPropertyChange()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/></w:style>
            """, """
              <w:p><w:r><w:t>Title</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-style
            target M.P0001
            style Heading 2
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success);
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Null(report.PreviewBefore);
        Assert.Equal("Heading2", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }

    [Fact]
    public static void PreviewReportsCellChange()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocxWithSimpleTwoByTwoTable();
        using var patch = new StringReader("""
            docxpatch 1

            op set-cell
            target M.T0001.R01.C01
            text Changed
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success);
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal("North", report.PreviewBefore);
        Assert.Equal("Changed", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }

    [Fact]
    public static void PreviewAgreesBetweenCheckAndApply()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        const string patchText = "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Omega\nend\n";
        using MemoryStream checkInput = CreateDocx("Alpha Beta");
        DocxPatchOperationReport check = Assert.Single(new DocxEditor().Check(checkInput, new StringReader(patchText), options).Operations);
        using MemoryStream applyInput = CreateDocx("Alpha Beta");
        using var output = new MemoryStream();
        DocxPatchOperationReport apply = Assert.Single(new DocxEditor().Apply(applyInput, new StringReader(patchText), output, options).Operations);
        Assert.Equal(check.PreviewBefore, apply.PreviewBefore);
        Assert.Equal(check.PreviewAfter, apply.PreviewAfter);
        Assert.Equal(check.PreviewTruncated, apply.PreviewTruncated);
    }

    [Fact]
    public static void PreviewShowsIdenticalTextForNoOp()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocx("Anchor");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Anchor
            with Anchor
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, options);

        Assert.True(result.Success);
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal("Anchor", report.PreviewBefore);
        Assert.Equal("Anchor", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }

    [Fact]
    public static void OperationsAfterFailureAreSkipped()
    {
        using MemoryStream input = CreateDocx("Alpha Beta");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P9999
            find Alpha
            with Omega
            end

            op replace-text
            target M.P0001
            find Beta
            with Gamma
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.False(result.Success);
        Assert.Equal(2, result.Operations.Count);
        Assert.False(result.Operations[0].Success);
        Assert.False(result.Operations[1].Success);
        DocxDiagnostic skipped = Assert.Single(result.Operations[1].Diagnostics);
        Assert.Equal("I0002", skipped.Code);
        Assert.Equal(DocxSeverity.Info, skipped.Severity);
        Assert.Empty(result.Operations[1].AffectedTargets);
        Assert.Empty(result.Operations[1].GeneratedRevisionIds);
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public static void SkippedOperationsAgreeBetweenCheckAndApply()
    {
        const string patchText = "docxpatch 1\n\nop replace-text\ntarget M.P9999\nfind Alpha\nwith Omega\nend\n\nop replace-text\ntarget M.P0001\nfind Beta\nwith Gamma\nend\n";
        using MemoryStream checkInput = CreateDocx("Alpha Beta");
        DocxCheckResult check = new DocxEditor().Check(checkInput, new StringReader(patchText));
        using MemoryStream applyInput = CreateDocx("Alpha Beta");
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, new StringReader(patchText), output);
        Assert.False(check.Success);
        Assert.False(apply.Success);
        Assert.Equal(check.Operations[1].Success, apply.Operations[1].Success);
        Assert.Contains(check.Operations[1].Diagnostics, static d => d.Code == "I0002");
        Assert.Contains(apply.Operations[1].Diagnostics, static d => d.Code == "I0002");
    }
    [Fact]
    public static void ApplyInsertAfterReportsCreatedParagraph()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text Inserted
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        Assert.Equal(new[] { "M.P0002" }, Assert.Single(result.Operations).CreatedTargetIds);
        Assert.Contains("created-target-ids=M.P0002", DocxTextRenderer.RenderOperationSummary(result.Operations), StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckInsertBeforeReportsCreatedParagraph()
    {
        using MemoryStream input = CreateDocxWithBody("""
              <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
              <w:p><w:r><w:t>Beta</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op insert-before
            target M.P0002
            text Inserted
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Equal(new[] { "M.P0002" }, Assert.Single(result.Operations).CreatedTargetIds);
    }

    [Fact]
    public static void TwoSequentialInsertsReportEachCreatedId()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text First
            end

            op insert-after
            target M.P0001
            text Second
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Equal(2, result.Operations.Count);
        Assert.Equal(new[] { "M.P0002" }, result.Operations[0].CreatedTargetIds);
        Assert.Equal(new[] { "M.P0002" }, result.Operations[1].CreatedTargetIds);
        output.Position = 0;
        Assert.Equal(
            new[] { "Alpha", "Second", "First" },
            new DocxEditor().Read(output).Paragraphs.Select(static p => p.Text).ToArray());
    }

    [Fact]
    public static void AddCommentReportsCreatedCommentId()
    {
        using MemoryStream input = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="3" w:author="Reviewer" w:initials="RV" w:date="2026-06-07T12:00:00Z">
                    <w:p><w:r><w:t>Old comment</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        using var patch = new StringReader("""
            docxpatch 1

            op add-comment
            target M.P0001
            text Fresh note
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Equal(new[] { "comment:4" }, Assert.Single(result.Operations).CreatedTargetIds);
    }

    [Fact]
    public static void CheckAndApplyAgreeOnCreatedTargetIds()
    {
        const string patchText = "docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext Inserted\nend\n";
        using MemoryStream checkInput = CreateDocx("Alpha");
        IReadOnlyList<string> checkCreated = Assert.Single(new DocxEditor().Check(checkInput, new StringReader(patchText)).Operations).CreatedTargetIds;
        using MemoryStream applyInput = CreateDocx("Alpha");
        using var output = new MemoryStream();
        IReadOnlyList<string> applyCreated = Assert.Single(new DocxEditor().Apply(applyInput, new StringReader(patchText), output).Operations).CreatedTargetIds;
        Assert.Equal(checkCreated, applyCreated);
        Assert.Equal(new[] { "M.P0002" }, applyCreated);
    }

    [Fact]
    public static void NonCreatingOperationReportsEmptyCreatedIds()
    {
        using MemoryStream input = CreateDocx("Alpha Beta");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            with Omega
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success);
        Assert.Empty(Assert.Single(result.Operations).CreatedTargetIds);
    }
    [Fact]
    public static void CheckInsertAfterMultipleTextsReportsAllCreatedIds()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var patch = new StringReader("docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext First\ntext Second\nend\n");

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Equal(new[] { "M.P0002", "M.P0003" }, Assert.Single(result.Operations).CreatedTargetIds);
    }

    [Fact]
    public static void CheckInsertBeforeMultipleTextsReportsAllCreatedIds()
    {
        using MemoryStream input = CreateDocxWithBody("""
              <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
              <w:p><w:r><w:t>Beta</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("docxpatch 1\n\nop insert-before\ntarget M.P0002\ntext X\ntext Y\nend\n");

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Equal(new[] { "M.P0002", "M.P0003" }, Assert.Single(result.Operations).CreatedTargetIds);
    }
}
