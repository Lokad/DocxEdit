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
    public static void PreviewReportsShadingChange()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocxWithSimpleTwoByTwoTable();
        using var patch = new StringReader("""
            docxpatch 1

            op set-cell-shading
            target M.T0001.R01.C01
            fill 4472C4
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Null(report.PreviewBefore);
        Assert.Equal("4472C4", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }

    [Fact]
    public static void PreviewReportsTableStyleChange()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>
            """, """
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-table-style
            target M.T0001
            style TableGrid
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Null(report.PreviewBefore);
        Assert.Equal("TableGrid", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }

    [Fact]
    public static void PreviewReportsRowHeaderChange()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocxWithSimpleTwoByTwoTable();
        using var patch = new StringReader("""
            docxpatch 1

            op set-row-header
            target M.T0001.R01
            expect-header false
            header true
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal("false", report.PreviewBefore);
        Assert.Equal("true", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }

    [Fact]
    public static void PreviewReportsControlTextChange()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:id w:val="77"/>
                          <w:alias w:val="Client Name"/>
                          <w:tag w:val="client_name"/>
                          <w:text/>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>Acme</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-content-control-text
            target M.CC0001
            text Acme Corp
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal("Acme", report.PreviewBefore);
        Assert.Equal("Acme Corp", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }


    [Fact]
    public static void PreviewReportsImageAltChange()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-alt
            target M.I0001
            expect-alt Old chart
            alt Updated chart
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal("Old chart", report.PreviewBefore);
        Assert.Equal("Updated chart", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }

    [Fact]
    public static void PreviewReportsFieldResultChange()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocxWithRefField();
        using var patch = new StringReader("""
            docxpatch 1

            op set-field-result
            target M.F0001
            expect-result Old cached result
            text New cached result
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal("Old cached result", report.PreviewBefore);
        Assert.Equal("New cached result", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }

    [Fact]
    public static void PreviewReportsFieldCodeChange()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocxWithRefField();
        using var patch = new StringReader("""
            docxpatch 1

            op set-field-code
            target M.F0001
            expect-code REF ClientName \h
            code REF NewBookmark \h
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal(" REF ClientName \\h ", report.PreviewBefore);
        Assert.Equal("REF NewBookmark \\h", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }

    [Fact]
    public static void PreviewReportsCommentTextChange()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocxWithCommentAnchoredParagraph();
        using var patch = new StringReader("""
            docxpatch 1

            op set-comment-text
            target comment:3
            expect-text Comment body
            text Updated comment
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal("Comment body", report.PreviewBefore);
        Assert.Equal("Updated comment", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }

    [Fact]
    public static void PreviewReportsBookmarkTextChange()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocxWithSingleBookmark();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-bookmark-text
            target M.B0001
            expect-text Old Client
            text New Client
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal("Old Client", report.PreviewBefore);
        Assert.Equal("New Client", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }

    [Fact]
    public static void PreviewReportsTableMetadataChange()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tblPr>
                        <w:tblCaption w:val="Old caption"/>
                        <w:tblDescription w:val="Old description"/>
                      </w:tblPr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-table-metadata
            target M.T0001
            expect-caption Old caption
            expect-description Old description
            caption Revenue table
            description Quarterly figures
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal("caption=Old caption; description=Old description", report.PreviewBefore);
        Assert.Equal("caption=Revenue table; description=Quarterly figures", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }

    [Fact]
    public static void PreviewReportsDeleteBlock()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
                    <w:p><w:r><w:t>Beta</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op delete-block
            target M.P0001
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal("Alpha", report.PreviewBefore);
        Assert.Null(report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }

    [Fact]
    public static void PreviewReportsSectionColumnsChange()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>Main text</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:pgSz w:w="12240" w:h="15840"/>
                      <w:cols w:num="1"/>
                    </w:sectPr>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-section-columns
            target M.S0001
            count 2
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal("1", report.PreviewBefore);
        Assert.Equal("2", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }

    [Fact]
    public static void PreviewReportsHyperlinkTargetChange()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream stream = CreateDocxWithBody(
            """
                    <w:p xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                      <w:hyperlink r:id="rLink">
                        <w:r><w:t>External</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/report" TargetMode="External"/>
                </Relationships>
                """, null);
        using var patch = new StringReader("""
            docxpatch 1

            op set-hyperlink-target
            target M.L0001
            uri https://example.test/new
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(stream, patch, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal("https://example.test/report", report.PreviewBefore);
        Assert.Equal("https://example.test/new", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }

    [Fact]
    public static void PreviewReportsFieldDirtyChange()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocxWithRefField();
        using var patch = new StringReader("""
            docxpatch 1

            op set-field-dirty
            target M.F0001
            dirty true
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal("false", report.PreviewBefore);
        Assert.Equal("true", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }

    [Fact]
    public static void PreviewReportsFieldLockChange()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocxWithRefField();
        using var patch = new StringReader("""
            docxpatch 1

            op set-field-lock
            target M.F0001
            locked true
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal("0", report.PreviewBefore);
        Assert.Equal("true", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }

    [Fact]
    public static void PreviewReportsSectionOrientationChange()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>Main text</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:pgSz w:w="12240" w:h="15840"/>
                      <w:cols w:num="1"/>
                    </w:sectPr>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-section-orientation
            target M.S0001
            orientation landscape
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal("portrait", report.PreviewBefore);
        Assert.Equal("landscape", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }

    [Fact]
    public static void PreviewAgreesBetweenCheckAndApplyForEveryPreviewOperation()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        (string Name, Func<MemoryStream> Fixture, string Patch)[] cases =
        [
            ("replace-text", static () => CreateDocx("Alpha Beta"), "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Omega\nend\n"),
            ("replace-paragraph", static () => CreateDocx("Alpha"), "docxpatch 1\n\nop replace-paragraph\ntarget M.P0001\ntext Omega\nend\n"),
            ("set-cell", static () => CreateDocxWithSimpleTwoByTwoTable(), "docxpatch 1\n\nop set-cell\ntarget M.T0001.R01.C01\ntext Changed\nend\n"),
            ("set-style", static () => CreateDocxWithStylesAndBody("<w:style w:type=\"paragraph\" w:styleId=\"Heading2\"><w:name w:val=\"Heading 2\"/></w:style>", "<w:p><w:r><w:t>Title</w:t></w:r></w:p>"), "docxpatch 1\n\nop set-style\ntarget M.P0001\nstyle Heading 2\nend\n"),
            ("set-hyperlink-text", static () => CreateDocxWithBody("<w:p xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:hyperlink r:id=\"rLink\"><w:r><w:t>External</w:t></w:r></w:hyperlink></w:p>", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rLink\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.test/report\" TargetMode=\"External\"/></Relationships>", null), "docxpatch 1\n\nop set-hyperlink-text\ntarget M.L0001\ntext New label\nend\n"),
            ("set-cell-shading", static () => CreateDocxWithSimpleTwoByTwoTable(), "docxpatch 1\n\nop set-cell-shading\ntarget M.T0001.R01.C01\nfill 4472C4\nend\n"),
            ("set-table-style", static () => CreateDocxWithStylesAndBody("<w:style w:type=\"table\" w:styleId=\"TableGrid\"><w:name w:val=\"Table Grid\"/></w:style>", "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"), "docxpatch 1\n\nop set-table-style\ntarget M.T0001\nstyle TableGrid\nend\n"),
            ("set-row-header", static () => CreateDocxWithSimpleTwoByTwoTable(), "docxpatch 1\n\nop set-row-header\ntarget M.T0001.R01\nheader true\nend\n"),
            ("set-content-control-text", static () => CreateDocxWithBody("<w:p><w:sdt><w:sdtPr><w:id w:val=\"77\"/><w:alias w:val=\"Client Name\"/><w:tag w:val=\"client_name\"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Acme</w:t></w:r></w:sdtContent></w:sdt></w:p>"), "docxpatch 1\n\nop set-content-control-text\ntarget M.CC0001\ntext Acme Corp\nend\n"),
            ("set-image-alt", static () => CreateDocxWithImage("png", "image/png", "old-png"), "docxpatch 1\n\nop set-image-alt\ntarget M.I0001\nexpect-alt Old chart\nalt Updated chart\nend\n"),
            ("set-field-result", static () => CreateDocxWithRefField(), "docxpatch 1\n\nop set-field-result\ntarget M.F0001\nexpect-result Old cached result\ntext New cached result\nend\n"),
            ("set-field-code", static () => CreateDocxWithRefField(), "docxpatch 1\n\nop set-field-code\ntarget M.F0001\ncode REF NewBookmark \\h\nend\n"),
            ("set-comment-text", static () => CreateDocxWithCommentAnchoredParagraph(), "docxpatch 1\n\nop set-comment-text\ntarget comment:3\ntext Updated comment\nend\n"),
            ("replace-bookmark-text", static () => CreateDocxWithSingleBookmark(), "docxpatch 1\n\nop replace-bookmark-text\ntarget M.B0001\ntext New Client\nend\n"),
            ("set-table-metadata", static () => CreateDocxWithBody("<w:tbl><w:tblPr><w:tblCaption w:val=\"Old caption\"/><w:tblDescription w:val=\"Old description\"/></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"), "docxpatch 1\n\nop set-table-metadata\ntarget M.T0001\ncaption Revenue table\ndescription Quarterly figures\nend\n"),
            ("delete-block", static () => CreateDocxWithBody("<w:p><w:r><w:t>Alpha</w:t></w:r></w:p><w:p><w:r><w:t>Beta</w:t></w:r></w:p>"), "docxpatch 1\n\nop delete-block\ntarget M.P0001\nend\n"),
            ("set-section-columns", static () => CreateDocxWithBody("<w:p><w:r><w:t>Main text</w:t></w:r></w:p><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:cols w:num=\"1\"/></w:sectPr>"), "docxpatch 1\n\nop set-section-columns\ntarget M.S0001\ncount 2\nend\n"),
            ("set-section-orientation", static () => CreateDocxWithBody("<w:p><w:r><w:t>Main text</w:t></w:r></w:p><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:cols w:num=\"1\"/></w:sectPr>"), "docxpatch 1\n\nop set-section-orientation\ntarget M.S0001\norientation landscape\nend\n"),
            ("set-hyperlink-target", static () => CreateDocxWithBody("<w:p xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:hyperlink r:id=\"rLink\"><w:r><w:t>External</w:t></w:r></w:hyperlink></w:p>", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rLink\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.test/report\" TargetMode=\"External\"/></Relationships>", null), "docxpatch 1\n\nop set-hyperlink-target\ntarget M.L0001\nuri https://example.test/new\nend\n"),
            ("set-field-dirty", static () => CreateDocxWithRefField(), "docxpatch 1\n\nop set-field-dirty\ntarget M.F0001\ndirty true\nend\n"),
            ("set-field-lock", static () => CreateDocxWithRefField(), "docxpatch 1\n\nop set-field-lock\ntarget M.F0001\nlocked true\nend\n"),
        ];
        foreach ((string name, Func<MemoryStream> fixture, string patchText) in cases)
        {
            using MemoryStream checkInput = fixture();
            DocxCheckResult checkResult = new DocxEditor().Check(checkInput, new StringReader(patchText), options);
            Assert.True(checkResult.Success, name);
            using MemoryStream applyInput = fixture();
            using var output = new MemoryStream();
            DocxApplyResult applyResult = new DocxEditor().Apply(applyInput, new StringReader(patchText), output, options);
            Assert.True(applyResult.Success, name);
            DocxPatchOperationReport check = Assert.Single(checkResult.Operations);
            DocxPatchOperationReport apply = Assert.Single(applyResult.Operations);
            Assert.True(string.Equals(check.PreviewBefore, apply.PreviewBefore, StringComparison.Ordinal), name);
            Assert.True(string.Equals(check.PreviewAfter, apply.PreviewAfter, StringComparison.Ordinal), name);
            Assert.True(check.PreviewTruncated == apply.PreviewTruncated, name);
        }
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
    public static void TwoSequentialInsertsReportFinalIds()
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
        Assert.Equal(new[] { "M.P0003" }, result.Operations[0].CreatedTargetIds);
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
    [Fact]
    public static void InsertImageAfterReportsCreatedParagraph()
    {
        using MemoryStream input = CreateDocx("Alpha");
        var assets = new MemoryAssetProvider("chart.png", CreatePngBytes(4, 3), null, "chart.png");
        var options = new DocxEditOptions { AssetProvider = assets };
        using var patch = new StringReader("docxpatch 1\n\nop insert-image-after\ntarget M.P0001\nasset chart.png\nend\n");

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Equal(new[] { "M.P0002" }, Assert.Single(result.Operations).CreatedTargetIds);
    }

    [Fact]
    public static void AddBookmarkReportsCreatedBookmarkId()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var patch = new StringReader("docxpatch 1\n\nop add-bookmark\ntarget M.P0001\nname Mark\nend\n");

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Equal(new[] { "M.B0001" }, Assert.Single(result.Operations).CreatedTargetIds);
    }

    [Fact]
    public static void AddBookmarkBeforeExistingReportsFirstId()
    {
        using MemoryStream input = CreateDocxWithBody("""
              <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
              <w:p><w:bookmarkStart w:id="1" w:name="A"/><w:r><w:t>Beta</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>
            """);
        using var patch = new StringReader("docxpatch 1\n\nop add-bookmark\ntarget M.P0001\nname B\nend\n");

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Equal(new[] { "M.B0001" }, Assert.Single(result.Operations).CreatedTargetIds);
    }

    [Fact]
    public static void AddCommentReplyReportsCreatedCommentId()
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

            op add-comment-reply
            target comment:3
            text Thanks
            author Reviewer
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Equal(new[] { "comment:4" }, Assert.Single(result.Operations).CreatedTargetIds);
    }

    [Fact]
    public static void InsertAfterWrappedParagraphReportsFinalId()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>First</w:t></w:r></w:p>
                    <w:ins w:id="1" w:author="Reviewer" w:date="2026-06-07T12:00:00Z"><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:ins>
                    <w:p><w:r><w:t>Third</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0003
            text New
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Equal(new[] { "M.P0004" }, Assert.Single(result.Operations).CreatedTargetIds);
        output.Position = 0;
        List<string> ids = new DocxEditor().Read(output).Paragraphs.Select(static paragraph => paragraph.Id.ToWireValue()).ToList();
        Assert.Equal(new[] { "M.P0001", "M.P0002", "M.P0003", "M.P0004" }, ids);
    }

    [Fact]
    public static void AliasEditAfterInsertReportsFinalId()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>First</w:t></w:r></w:p>
                    <w:ins w:id="1" w:author="Reviewer" w:date="2026-06-07T12:00:00Z"><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:ins>
                    <w:p><w:r><w:t>Third</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0003
            as newSection
            text New
            end

            op replace-text
            target @newSection
            find New
            with Changed
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Equal(new[] { "M.P0004" }, result.Operations[0].CreatedTargetIds);
        Assert.Contains("M.P0004", result.Operations[1].AffectedTargets.Select(static target => target.Id.ToWireValue()));
    }

    [Fact]
    public static void AddFirstBookmarkReportsFirstId()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>First</w:t></w:r></w:p>
                    <w:p>
                      <w:bookmarkStart w:id="7" w:name="Second"/>
                      <w:r><w:t>Second text</w:t></w:r>
                      <w:bookmarkEnd w:id="7"/>
                    </w:p>
                    <w:p>
                      <w:bookmarkStart w:id="8" w:name="Third"/>
                      <w:r><w:t>Third text</w:t></w:r>
                      <w:bookmarkEnd w:id="8"/>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op add-bookmark
            target M.P0001
            name Intro
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Equal(new[] { "M.B0001" }, Assert.Single(result.Operations).CreatedTargetIds);
    }

    [Fact]
    public static void SetHyperlinkTextReportsHyperlinkTarget()
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
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/old" TargetMode="External"/>
                </Relationships>
                """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-hyperlink-text
            target M.L0001
            text New link
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchAffectedTarget affected = Assert.Single(Assert.Single(result.Operations).AffectedTargets);
        Assert.Equal("M.L0001", affected.Id.ToWireValue());
        Assert.Equal("hyperlink", affected.Kind);
        Assert.Equal("update", affected.Action);
    }

    [Fact]
    public static void RemoveHyperlinkReportsHyperlinkDelete()
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
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/old" TargetMode="External"/>
                </Relationships>
                """);
        using var patch = new StringReader("""
            docxpatch 1

            op remove-hyperlink
            target M.L0001
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchAffectedTarget affected = Assert.Single(Assert.Single(result.Operations).AffectedTargets);
        Assert.Equal("M.L0001", affected.Id.ToWireValue());
        Assert.Equal("hyperlink", affected.Kind);
        Assert.Equal("delete", affected.Action);
    }

    [Fact]
    public static void SetImageAltReportsImageTarget()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-alt
            target M.I0001
            alt Updated chart
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchAffectedTarget affected = Assert.Single(Assert.Single(result.Operations).AffectedTargets);
        Assert.Equal("M.I0001", affected.Id.ToWireValue());
        Assert.Equal("image", affected.Kind);
        Assert.Equal("update", affected.Action);
    }

    [Fact]
    public static void AffectedIdsRebaseToFinalCoordinatesAfterShift()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            with Beta
            end

            op insert-before
            target M.P0001
            text Leading
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Equal("M.P0002", Assert.Single(result.Operations[0].AffectedTargets).Id.ToWireValue());
        output.Position = 0;
        List<string> ids = new DocxEditor().Read(output).Paragraphs.Select(static paragraph => paragraph.Id.ToWireValue()).ToList();
        Assert.Equal(new[] { "M.P0001", "M.P0002" }, ids);
    }

    [Fact]
    public static void CreatedIdsRebaseToFinalCoordinatesAfterShift()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text Second
            end

            op insert-before
            target M.P0001
            text First
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Equal(new[] { "M.P0003" }, result.Operations[0].CreatedTargetIds);
        Assert.Equal(new[] { "M.P0001" }, result.Operations[1].CreatedTargetIds);
        output.Position = 0;
        List<string> texts = new DocxEditor().Read(output).Paragraphs.Select(static paragraph => paragraph.Text).ToList();
        Assert.Equal(new[] { "First", "Alpha", "Second" }, texts);
    }

    [Fact]
    public static void CreatedThenDeletedReportsNoSurvivingId()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("docxpatch 1\n\nop insert-after\ntarget M.P0001\nas newSection\ntext Second\nend\n\nop delete-block\ntarget @newSection\nend\n");
        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
        Assert.True(result.Success);
        Assert.Empty(result.Operations[0].CreatedTargetIds);
        output.Position = 0;
        Assert.Equal("Alpha", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }
    [Fact]
    public static void CreatedSurvivesNeighborDeletedReportsLiveId()
    {
        const string body = "<w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p>";
        using MemoryStream input = CreateDocxWithBody(body);
        string patchText = "docxpatch 1\n\nop insert-after\ntarget M.P0001\nas added\ntext Inserted\nend\n\nop delete-block\ntarget M.P0001\nend\n";
        DocxCheckResult check = new DocxEditor().Check(input, new StringReader(patchText));
        Assert.True(check.Success);
        Assert.Equal("M.P0001", Assert.Single(check.Operations[0].CreatedTargetIds));
        using MemoryStream applyInput = CreateDocxWithBody(body);
        using var output = new MemoryStream();
        using var applyPatch = new StringReader(patchText);
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, applyPatch, output);
        Assert.True(apply.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal("Inserted", read.Paragraphs[0].Text);
        Assert.Equal("Second", read.Paragraphs[1].Text);
    }


    [Fact]
    public static void SemanticSelectorPreviewSurvivesMatchChange()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-paragraph
            target text:"Alpha"
            text Beta
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal("Alpha", report.PreviewBefore);
        Assert.Equal("Beta", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }

    [Fact]
    public static void SemanticSelectorOnCreatedParagraphReportsAffectedAndPreview()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocxWithBody("<w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p>");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text Inserted
            as added
            end
            op replace-paragraph
            target text:"Inserted"
            text CHANGED
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Equal(2, result.Operations.Count);
        DocxPatchOperationReport second = result.Operations[1];
        DocxPatchAffectedTarget affected = Assert.Single(second.AffectedTargets);
        Assert.Equal("M.P0002", affected.Id.ToWireValue());
        Assert.Equal("paragraph", affected.Kind);
        Assert.Equal("update", affected.Action);
        Assert.Equal("op1-0", affected.CreatedMark);
        Assert.Equal("Inserted", second.PreviewBefore);
        Assert.Equal("CHANGED", second.PreviewAfter);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal("CHANGED", read.Paragraphs[1].Text);
        Assert.Equal("M.P0002", read.Paragraphs[1].Id.ToWireValue());
    }

    [Fact]
    public static void SemanticSelectorWithInnerQuotesResolvesEscapedText()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocxWithBody("<w:p><w:r><w:t>Say \"hi\" there</w:t></w:r></w:p>");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-paragraph
            target text:"Say \"hi\" there"
            text CHANGED
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        DocxPatchAffectedTarget affected = Assert.Single(report.AffectedTargets);
        Assert.Equal("M.P0001", affected.Id.ToWireValue());
        Assert.Equal("paragraph", affected.Kind);
        Assert.Equal("update", affected.Action);
        Assert.Equal("Say \"hi\" there", report.PreviewBefore);
        Assert.Equal("CHANGED", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal("CHANGED", read.Paragraphs[0].Text);
        Assert.Equal("M.P0001", read.Paragraphs[0].Id.ToWireValue());
    }

    [Fact]
    public static void DeleteBlockPreviewKeepsBeforeValue()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-block
            target M.P0001
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal("Alpha", report.PreviewBefore);
        Assert.Null(report.PreviewAfter);
    }

    [Fact]
    public static void CellSubstringPreviewReportsBeforeAndAfter()
    {
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr><w:tc><w:p><w:r><w:t>Hello World</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.T0001.R01.C01
            find World
            with There
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal("Hello World", report.PreviewBefore);
        Assert.Equal("Hello There", report.PreviewAfter);
        Assert.False(report.PreviewTruncated);
    }
    [Fact]
    public static void DeleteFirstReportsHistoricalIdWithSurvivingNeighbor()
    {
        const string body = "<w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p>";
        using MemoryStream input = CreateDocxWithBody(body);
        string patchText = "docxpatch 1\n\nop delete-block\ntarget M.P0001\nend\n";
        DocxCheckResult check = new DocxEditor().Check(input, new StringReader(patchText));
        Assert.True(check.Success);
        DocxPatchAffectedTarget affected = Assert.Single(check.Operations[0].AffectedTargets);
        Assert.Equal("M.P0001", affected.Id.ToWireValue());
        Assert.Equal("delete", affected.Action);
        using MemoryStream applyInput = CreateDocxWithBody(body);
        using var output = new MemoryStream();
        using var applyPatch = new StringReader(patchText);
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, applyPatch, output);
        Assert.True(apply.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal("Second", Assert.Single(read.Paragraphs).Text);
    }

}
