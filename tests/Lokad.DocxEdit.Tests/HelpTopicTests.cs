using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D06: focused per-operation help renders from the shared catalog, and the
// patch overview separates recognized-but-unimplemented operations.
[Collection("ConsoleCli")]
public static class HelpTopicTests
{
    [Fact]
    public static void PatchOperationTopicRendersFieldsAndTrackSupport()
    {
        string topic = DocxHelp.RenderTopic("replace-text");

        Assert.Contains("op replace-text", topic, StringComparison.Ordinal);
        Assert.Contains("Required fields:", topic, StringComparison.Ordinal);
        Assert.Contains("target", topic, StringComparison.Ordinal);
        Assert.Contains("find", topic, StringComparison.Ordinal);
        Assert.Contains("with", topic, StringComparison.Ordinal);
        Assert.Contains("occurrence", topic, StringComparison.Ordinal);
        Assert.Contains("Track-change support:", topic, StringComparison.Ordinal);
    }

    [Fact]
    public static void PatchOperationTopicMarksRepeatableFields()
    {
        string topic = DocxHelp.RenderTopic("append-row");

        Assert.Contains("cell+", topic, StringComparison.Ordinal);
    }

    [Fact]
    public static void PatchOperationTopicCoversUnimplementedOperation()
    {
        string topic = DocxHelp.RenderTopic("append-column");

        Assert.Contains("op append-column", topic, StringComparison.Ordinal);
        Assert.Contains("E4316", topic, StringComparison.Ordinal);
    }
    [Fact]
    public static void CapabilitiesAndTemplateTopicsCoverAllTargetKinds()
    {
        string capabilities = DocxHelp.RenderTopic("capabilities");
        string template = DocxHelp.RenderTopic("template");

        foreach (string kind in new[] { "paragraph", "content-control", "cell", "merge-group", "bookmark", "table", "row", "section", "hyperlink", "field", "image" })
        {
            Assert.Contains(kind, capabilities, StringComparison.Ordinal);
            Assert.Contains(kind, template, StringComparison.Ordinal);
        }
    }


    [Fact]
    public static void CliHelpPatchOperationMatchesLibraryTopic()
    {
        CliTests.CliResult result = CliTests.RunCli("help", "replace-text");

        Assert.Equal(0, result.ExitCode);
        Assert.Equal(DocxHelp.RenderTopic("replace-text"), result.Output);
    }

    [Fact]
    public static void CliHelpUnknownTopicFails()
    {
        CliTests.CliResult result = CliTests.RunCli("help", "bogus-op");

        Assert.Equal(2, result.ExitCode);
        Assert.Contains("Unknown help topic", result.Error, StringComparison.Ordinal);
    }

    [Fact]
    public static void PatchHelpSeparatesUnimplementedOperations()
    {
        string help = DocxHelp.RenderTopic("patch");
        string[] unimplemented = ["append-column", "insert-column-before", "insert-column-after", "delete-column", "add-repeating-section-item", "delete-repeating-section-item"];

        Assert.Contains("Recognized but unimplemented operations", help, StringComparison.Ordinal);
        int section = help.IndexOf("Recognized but unimplemented operations", StringComparison.Ordinal);
        string supported = help.Substring(0, section);
        string unimplementedSection = help.Substring(section);
        foreach (string operation in unimplemented)
        {
            Assert.Contains(operation, unimplementedSection, StringComparison.Ordinal);
            Assert.DoesNotContain("  " + operation + " ", supported, StringComparison.Ordinal);
        }
    }

    [Fact]
    public static void PatchHelpDocumentsResultAliases()
    {
        string help = DocxHelp.RenderTopic("patch");

        Assert.Contains("Result aliases", help, StringComparison.Ordinal);
        Assert.Contains("@name", help, StringComparison.Ordinal);
        Assert.Contains("created IDs", help, StringComparison.Ordinal);
    }

    [Fact]
    public static void PatchHelpDocumentsWhitespaceContract()
    {
        string help = DocxHelp.RenderTopic("patch");

        Assert.Contains("whitespace normalized", help, StringComparison.Ordinal);
        Assert.Contains("expect-text guards compare exact", help, StringComparison.Ordinal);
    }

    [Fact]
    public static void ChangesHelpDocumentsEmptyMarkupLine()
    {
        string help = DocxHelp.RenderTopic("changes");

        Assert.Contains("No changes line", help, StringComparison.Ordinal);
    }

    [Fact]
    public static void PatchOperationTopicReportsEmptyAllowedFields()
    {
        Assert.Contains("Empty values allowed: with", DocxHelp.RenderTopic("replace-text"), StringComparison.Ordinal);
        Assert.DoesNotContain("Empty values allowed", DocxHelp.RenderTopic("set-style"), StringComparison.Ordinal);
    }

    [Fact]
    public static void PatchOperationExamplesParseAndLintClean()
    {
        var editor = new DocxEditor();
        int covered = 0;
        foreach (DocxPatchOperationInfo operation in DocxHelp.Catalog.PatchOperations)
        {
            foreach (string example in operation.Examples)
            {
                DocxPatch parsed = editor.ParsePatch(new StringReader(example));
                Assert.True(parsed.Success, operation.Name + ":" + string.Join("|", parsed.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
                DocxLintResult lint = editor.Lint(new StringReader(example));
                Assert.True(lint.Success, operation.Name + ":" + string.Join("|", lint.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
                covered++;
            }
        }

        Assert.True(covered >= 22, "Expected at least 22 curated examples, found " + covered + ".");
    }

    [Fact]
    public static void SupportedOperationsHaveMinimalAndGuardedExamples()
    {
        foreach (DocxPatchOperationInfo operation in DocxHelp.Catalog.PatchOperations)
        {
            if (operation.TrackChangesSupportClass == "unsupported")
            {
                continue;
            }

            Assert.Contains(operation.Examples, static example => example.Contains("# Minimal", StringComparison.Ordinal));
            if (operation.OptionalFields.Any(static field => field.StartsWith("expect-", StringComparison.Ordinal)))
            {
                Assert.Contains(operation.Examples, static example => example.Contains("# Guarded", StringComparison.Ordinal));
            }
        }
    }

    [Fact]
    public static void PatchOperationTopicRendersMinimalAndGuardedExamples()
    {
        string topic = DocxHelp.RenderTopic("replace-text");

        Assert.Contains("# Minimal replace-text.", topic, StringComparison.Ordinal);
        Assert.Contains("# Guarded replace-text:", topic, StringComparison.Ordinal);
        Assert.Contains("op replace-text", topic, StringComparison.Ordinal);
    }

    [Fact]
    public static void AliasCompositionExampleChecksClean()
    {
        Assert.True(DocxHelp.TryGetPatchOperation("insert-after", out DocxPatchOperationInfo insertAfter));
        string example = Assert.Single(insertAfter.Examples, static candidate => candidate.Contains("# insert-after with a result alias", StringComparison.Ordinal));

        using MemoryStream input = CreateDocx("Alpha");
        using var patch = new StringReader(example);
        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        using var output = new MemoryStream();
        using var applyPatch = new StringReader(example);
        DocxApplyResult applied = new DocxEditor().Apply(input, applyPatch, output);
        Assert.True(applied.Success, string.Join("|", applied.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        output.Position = 0;
        Assert.Contains("Final note", Assert.Single(new DocxEditor().Read(output).Paragraphs, static paragraph => paragraph.Text.Contains("Final", StringComparison.Ordinal)).Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void MinimalReplaceTextExampleChecksClean()
    {
        Assert.True(DocxHelp.TryGetPatchOperation("replace-text", out DocxPatchOperationInfo replaceText));
        string example = Assert.Single(replaceText.Examples, static candidate => candidate.Contains("# Minimal replace-text.", StringComparison.Ordinal));

        using MemoryStream input = CreateDocx("Alpha");
        using var patch = new StringReader(example);
        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }
    [Fact]
    public static void PatchOperationTopicReportsFieldDefaults()
    {
        Assert.Contains("Defaults: preserve-runs=true", DocxHelp.RenderTopic("replace-text"), StringComparison.Ordinal);
        Assert.Contains("Defaults: force=false", DocxHelp.RenderTopic("set-cell"), StringComparison.Ordinal);
        Assert.Contains("Defaults: clear=false", DocxHelp.RenderTopic("set-cell-shading"), StringComparison.Ordinal);
        Assert.Contains("Defaults: copy-paragraph-properties=false", DocxHelp.RenderTopic("insert-after"), StringComparison.Ordinal);
        Assert.DoesNotContain("Defaults:", DocxHelp.RenderTopic("delete-block"), StringComparison.Ordinal);
    }

    [Fact]
    public static void FieldDefaultsAreMachineReadable()
    {
        Assert.True(DocxHelp.TryGetPatchOperation("replace-text", out DocxPatchOperationInfo replaceText));
        Assert.Contains("preserve-runs=true", replaceText.FieldDefaults);
        Assert.True(DocxHelp.TryGetPatchOperation("delete-block", out DocxPatchOperationInfo deleteBlock));
        Assert.Empty(deleteBlock.FieldDefaults);
    }

    [Fact]
    public static void OmittedPreserveRunsBehavesAsTrue()
    {
        using MemoryStream omittedInput = CreateDocxWithRuns("Alpha ", "Beta");
        using var omittedPatch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Omega\nend\n");
        using var omittedOutput = new MemoryStream();
        DocxApplyResult omitted = new DocxEditor().Apply(omittedInput, omittedPatch, omittedOutput);
        Assert.True(omitted.Success);

        using MemoryStream explicitInput = CreateDocxWithRuns("Alpha ", "Beta");
        using var explicitPatch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Omega\npreserve-runs true\nend\n");
        using var explicitOutput = new MemoryStream();
        DocxApplyResult explicitResult = new DocxEditor().Apply(explicitInput, explicitPatch, explicitOutput);
        Assert.True(explicitResult.Success);

        Assert.Equal(ReadEntryBytes(omittedOutput, "word/document.xml"), ReadEntryBytes(explicitOutput, "word/document.xml"));
    }
    [Fact]
    public static void AllowedValuesAreMachineReadable()
    {
        Assert.True(DocxHelp.TryGetPatchOperation("set-section-orientation", out DocxPatchOperationInfo orientation));
        Assert.Contains("orientation=portrait|landscape", orientation.AllowedValues);
        Assert.True(DocxHelp.TryGetPatchOperation("set-image-wrap", out DocxPatchOperationInfo wrap));
        Assert.Contains("mode=none|wrapNone|square|wrapSquare|tight|wrapTight|through|wrapThrough|top-bottom|topAndBottom|wrapTopAndBottom", wrap.AllowedValues);
        Assert.True(DocxHelp.TryGetPatchOperation("replace-image", out DocxPatchOperationInfo replaceImage));
        Assert.Contains("expect-content-type=image/png|image/jpeg", replaceImage.AllowedValues);
    }

    [Fact]
    public static void HelpRendersAllowedValues()
    {
        string topic = DocxHelp.RenderTopic("set-section-orientation");
        Assert.Contains("Allowed values:", topic, StringComparison.Ordinal);
        Assert.Contains("orientation=portrait|landscape", topic, StringComparison.Ordinal);
    }

    [Fact]
    public static void ClosedValuesStillRejectedByExecution()
    {
        using MemoryStream orientationInput = CreateDocx("Alpha");
        using var orientationPatch = new StringReader("docxpatch 1\n\nop set-section-orientation\ntarget M.S0001\norientation diagonal\nend\n");
        DocxCheckResult orientationResult = new DocxEditor().Check(orientationInput, orientationPatch);
        Assert.False(orientationResult.Success);
        Assert.Contains(orientationResult.Diagnostics, static diagnostic => diagnostic.Code == "E6202");
        using MemoryStream wrapInput = CreateDocxWithAnchoredImage();
        using var wrapPatch = new StringReader("docxpatch 1\n\nop set-image-wrap\ntarget M.I0001\nmode sideways\nend\n");
        DocxCheckResult wrapResult = new DocxEditor().Check(wrapInput, wrapPatch);
        Assert.False(wrapResult.Success);
        Assert.Contains(wrapResult.Diagnostics, static diagnostic => diagnostic.Code == "E5209");
        using MemoryStream contentTypeInput = CreateDocxWithImage("png", "image/png", "old-png");
        using var contentTypePatch = new StringReader("docxpatch 1\n\nop replace-image\ntarget M.I0001\nasset chart.png\nexpect-content-type image/gif\nend\n");
        DocxCheckResult contentTypeResult = new DocxEditor().Check(contentTypeInput, contentTypePatch);
        Assert.False(contentTypeResult.Success);
        Assert.Contains(contentTypeResult.Diagnostics, static diagnostic => diagnostic.Code == "E4205");
    }
    [Fact]
    public static void UnitHintsAreMachineReadable()
    {
        Assert.True(DocxHelp.TryGetPatchOperation("set-image-size", out DocxPatchOperationInfo size));
        Assert.Contains("width=positive-dimension:emu|in|cm|pt|px", size.UnitHints);
        Assert.Contains("height=positive-dimension:emu|in|cm|pt|px", size.UnitHints);
        Assert.True(DocxHelp.TryGetPatchOperation("insert-image-after", out DocxPatchOperationInfo insert));
        Assert.Contains("width=dimension:emu|in|cm|pt|px", insert.UnitHints);
        Assert.True(DocxHelp.TryGetPatchOperation("set-image-wrap", out DocxPatchOperationInfo wrap));
        Assert.Contains("dist-top=dimension:emu|in|cm|pt|px", wrap.UnitHints);
        Assert.True(DocxHelp.TryGetPatchOperation("set-image-position", out DocxPatchOperationInfo position));
        Assert.Contains("horizontal-offset=signed-dimension:emu|in|cm|pt|px", position.UnitHints);
        Assert.True(DocxHelp.TryGetPatchOperation("set-image-crop", out DocxPatchOperationInfo crop));
        Assert.Contains("left-percent=percent:0-100", crop.UnitHints);
        Assert.True(DocxHelp.TryGetPatchOperation("replace-text", out DocxPatchOperationInfo replaceText));
        Assert.Empty(replaceText.UnitHints);
    }

    [Fact]
    public static void HelpRendersUnitHints()
    {
        string topic = DocxHelp.RenderTopic("set-image-crop");
        Assert.Contains("Units:", topic, StringComparison.Ordinal);
        Assert.Contains("left-percent=percent:0-100", topic, StringComparison.Ordinal);
        Assert.DoesNotContain("Units:", DocxHelp.RenderTopic("replace-text"), StringComparison.Ordinal);
    }

    [Fact]
    public static void BadDimensionsStillRejectedByExecution()
    {
        using MemoryStream sizeInput = CreateDocxWithAnchoredImage();
        using var sizePatch = new StringReader("docxpatch 1\n\nop set-image-size\ntarget M.I0001\nwidth enormous\nend\n");
        DocxCheckResult sizeResult = new DocxEditor().Check(sizeInput, sizePatch);
        Assert.False(sizeResult.Success);
        Assert.Contains(sizeResult.Diagnostics, static diagnostic => diagnostic.Code == "E5206");
        using MemoryStream wrapInput = CreateDocxWithAnchoredImage();
        using var wrapPatch = new StringReader("docxpatch 1\n\nop set-image-wrap\ntarget M.I0001\nmode square\ndist-top lots\nend\n");
        DocxCheckResult wrapResult = new DocxEditor().Check(wrapInput, wrapPatch);
        Assert.False(wrapResult.Success);
        Assert.Contains(wrapResult.Diagnostics, static diagnostic => diagnostic.Code == "E5209");
        using MemoryStream positionInput = CreateDocxWithAnchoredImage();
        using var positionPatch = new StringReader("docxpatch 1\n\nop set-image-position\ntarget M.I0001\nhorizontal-offset far\nend\n");
        DocxCheckResult positionResult = new DocxEditor().Check(positionInput, positionPatch);
        Assert.False(positionResult.Success);
        Assert.Contains(positionResult.Diagnostics, static diagnostic => diagnostic.Code == "E5210");
        using MemoryStream cropInput = CreateDocxWithAnchoredImage();
        using var cropPatch = new StringReader("docxpatch 1\n\nop set-image-crop\ntarget M.I0001\nleft-percent 200\nend\n");
        DocxCheckResult cropResult = new DocxEditor().Check(cropInput, cropPatch);
        Assert.False(cropResult.Success);
        Assert.Contains(cropResult.Diagnostics, static diagnostic => diagnostic.Code == "E5208");
    }
    [Fact]
    public static void PositionEnumsAreMachineReadable()
    {
        Assert.True(DocxHelp.TryGetPatchOperation("set-image-position", out DocxPatchOperationInfo position));
        Assert.Contains("horizontal-relative=page|margin|column|character|leftMargin|rightMargin|insideMargin|outsideMargin", position.AllowedValues);
        Assert.Contains("horizontal-align=left|center|right|inside|outside", position.AllowedValues);
        Assert.Contains("vertical-relative=page|margin|paragraph|line|topMargin|bottomMargin|insideMargin|outsideMargin", position.AllowedValues);
        Assert.Contains("vertical-align=top|center|bottom|inside|outside", position.AllowedValues);
    }

    [Fact]
    public static void HelpRendersPositionEnums()
    {
        string topic = DocxHelp.RenderTopic("set-image-position");
        Assert.Contains("Allowed values:", topic, StringComparison.Ordinal);
        Assert.Contains("horizontal-align=left|center|right|inside|outside", topic, StringComparison.Ordinal);
    }

    [Fact]
    public static void BadPositionEnumsStillRejectedByExecution()
    {
        using MemoryStream alignInput = CreateDocxWithAnchoredImage();
        using var alignPatch = new StringReader("docxpatch 1\n\nop set-image-position\ntarget M.I0001\nhorizontal-align sideways\nend\n");
        DocxCheckResult alignResult = new DocxEditor().Check(alignInput, alignPatch);
        Assert.False(alignResult.Success);
        Assert.Contains(alignResult.Diagnostics, static diagnostic => diagnostic.Code == "E5210");
        using MemoryStream relativeInput = CreateDocxWithAnchoredImage();
        using var relativePatch = new StringReader("docxpatch 1\n\nop set-image-position\ntarget M.I0001\nvertical-relative nowhere\nend\n");
        DocxCheckResult relativeResult = new DocxEditor().Check(relativeInput, relativePatch);
        Assert.False(relativeResult.Success);
        Assert.Contains(relativeResult.Diagnostics, static diagnostic => diagnostic.Code == "E5210");
    }
    [Fact]
    public static void EngineFailuresCarryHelpTopic()
    {
        using MemoryStream input = CreateDocxWithAnchoredImage();
        using var patch = new StringReader("docxpatch 1\n\nop set-image-size\ntarget M.I0001\nwidth enormous\nend\n");
        DocxCheckResult result = new DocxEditor().Check(input, patch);
        Assert.False(result.Success);
        DocxDiagnostic failure = Assert.Single(result.Diagnostics, static diagnostic => diagnostic.Code == "E5206");
        Assert.Equal("set-image-size", failure.HelpTopic);
    }

    [Fact]
    public static void LintFailuresCarryHelpTopic()
    {
        using var patch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nend\n");
        DocxLintResult lint = new DocxEditor().Lint(patch);
        Assert.False(lint.Success);
        DocxDiagnostic failure = Assert.Single(lint.Diagnostics, static diagnostic => diagnostic.Code == "E4202");
        Assert.Equal("replace-text", failure.HelpTopic);
    }

    [Fact]
    public static void ParserFindingsCarryHelpTopic()
    {
        using var unknownFieldPatch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nbogus-field value\nend\n");
        DocxPatch unknownField = new DocxEditor().ParsePatch(unknownFieldPatch);
        Assert.False(unknownField.Success);
        DocxDiagnostic fieldFailure = Assert.Single(unknownField.Diagnostics, static diagnostic => diagnostic.Code == "E2011");
        Assert.Equal("replace-text", fieldFailure.HelpTopic);
        using var noPreamblePatch = new StringReader("op replace-text\ntarget M.P0001\nend\n");
        DocxPatch noPreamble = new DocxEditor().ParsePatch(noPreamblePatch);
        Assert.False(noPreamble.Success);
        Assert.Null(Assert.Single(noPreamble.Diagnostics).HelpTopic);
        using var unknownOpPatch = new StringReader("docxpatch 1\n\nop frobnicate\ntarget M.P0001\nend\n");
        DocxPatch unknownOp = new DocxEditor().ParsePatch(unknownOpPatch);
        Assert.False(unknownOp.Success);
        Assert.Null(Assert.Single(unknownOp.Diagnostics).HelpTopic);
    }

    [Fact]
    public static void CliFailuresPrintHelpPointer()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = Path.Combine(temp.Path, "input.docx");
        File.WriteAllBytes(input, CreateDocxWithAnchoredImage().ToArray());
        string patch = Path.Combine(temp.Path, "edits.docxpatch");
        File.WriteAllText(patch, "docxpatch 1" + Environment.NewLine + Environment.NewLine + "op set-image-size" + Environment.NewLine + "target M.I0001" + Environment.NewLine + "width enormous" + Environment.NewLine + "end" + Environment.NewLine);
        CliTests.CliResult result = CliTests.RunCli("check", input, patch);
        Assert.Equal(1, result.ExitCode);
        Assert.Contains("help=set-image-size", result.Error, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliJsonReportCarriesHelpTopic()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = Path.Combine(temp.Path, "input.docx");
        File.WriteAllBytes(input, CreateDocxWithAnchoredImage().ToArray());
        string patch = Path.Combine(temp.Path, "edits.docxpatch");
        File.WriteAllText(patch, "docxpatch 1" + Environment.NewLine + Environment.NewLine + "op set-image-size" + Environment.NewLine + "target M.I0001" + Environment.NewLine + "width enormous" + Environment.NewLine + "end" + Environment.NewLine);
        string report = Path.Combine(temp.Path, "report.json");
        CliTests.CliResult result = CliTests.RunCli("check", input, patch, "--report", report);
        Assert.Equal(1, result.ExitCode);
        Assert.Contains("HelpTopic", File.ReadAllText(report), StringComparison.Ordinal);
        Assert.Contains("set-image-size", File.ReadAllText(report), StringComparison.Ordinal);
    }
    [Fact]
    public static void MinimalReplaceBookmarkTextExampleChecksClean()
    {
        Assert.True(DocxHelp.TryGetPatchOperation("replace-bookmark-text", out DocxPatchOperationInfo bookmark));
        string example = Assert.Single(bookmark.Examples, static candidate => candidate.Contains("# Minimal replace-bookmark-text.", StringComparison.Ordinal));

        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:bookmarkStart w:id="4" w:name="ClientName"/>
                      <w:r><w:t>Old Client</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
            """);
        using var patch = new StringReader(example);
        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void MinimalSetCommentTextExampleChecksClean()
    {
        Assert.True(DocxHelp.TryGetPatchOperation("set-comment-text", out DocxPatchOperationInfo comment));
        string example = Assert.Single(comment.Examples, static candidate => candidate.Contains("# Minimal set-comment-text.", StringComparison.Ordinal));

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
        using var patch = new StringReader(example);
        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void MinimalSetHyperlinkTextExampleChecksClean()
    {
        Assert.True(DocxHelp.TryGetPatchOperation("set-hyperlink-text", out DocxPatchOperationInfo hyperlink));
        string example = Assert.Single(hyperlink.Examples, static candidate => candidate.Contains("# Minimal set-hyperlink-text.", StringComparison.Ordinal));

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
        using var patch = new StringReader(example);
        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void MinimalAppendRowExampleChecksClean()
    {
        Assert.True(DocxHelp.TryGetPatchOperation("append-row", out DocxPatchOperationInfo appendRow));
        string example = Assert.Single(appendRow.Examples, static candidate => candidate.Contains("# Minimal append-row.", StringComparison.Ordinal));

        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader(example);
        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }
    [Fact]
    public static void HelpOverviewSeparatesLongCommandNames()
    {
        string overview = DocxHelp.RenderOverview();

        Assert.Contains("capabilities Show supported", overview, StringComparison.Ordinal);
        Assert.Contains("template     Print a guarded patch template", overview, StringComparison.Ordinal);
        Assert.Contains("lint      Validate patch shape", overview, StringComparison.Ordinal);
        Assert.DoesNotContain("capabilitiesShow", overview, StringComparison.Ordinal);
    }

    [Fact]
    public static void LintHelpSeparatesLongFieldNames()
    {
        string topic = DocxHelp.RenderTopic("lint");

        Assert.Contains("OperationCount Number", topic, StringComparison.Ordinal);
    }

    [Fact]
    public static void SpecHelpRequirementsCoverNewCommands()
    {
        string root = FindRepoRoot();
        string spec = File.ReadAllText(Path.Combine(root, "SPEC.md"));

        Assert.Contains("docxedit help capabilities", spec, StringComparison.Ordinal);
        Assert.Contains("docxedit help template", spec, StringComparison.Ordinal);
        Assert.Contains("docxedit help lint", spec, StringComparison.Ordinal);
        Assert.Contains("capabilities|template", spec, StringComparison.Ordinal);
    }
    [Fact]
    public static void SectionExamplesCheckClean()
    {
        const string body = """
                    <w:p><w:r><w:t>Main text</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:pgSz w:w="12240" w:h="15840"/>
                      <w:cols w:num="1"/>
                    </w:sectPr>
            """;
        Assert.True(DocxHelp.TryGetPatchOperation("set-section-columns", out DocxPatchOperationInfo columns));
        Assert.True(DocxHelp.TryGetPatchOperation("set-section-orientation", out DocxPatchOperationInfo orientation));
        foreach (string example in columns.Examples.Concat(orientation.Examples))
        {
            using MemoryStream input = CreateDocxWithBody(body);
            using var patch = new StringReader(example);
            DocxCheckResult result = new DocxEditor().Check(input, patch);
            Assert.True(result.Success, example + ":" + string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        }
    }
    [Fact]
    public static void RowExamplesCheckClean()
    {
        Assert.True(DocxHelp.TryGetPatchOperation("insert-row-before", out DocxPatchOperationInfo insertBefore));
        Assert.True(DocxHelp.TryGetPatchOperation("insert-row-after", out DocxPatchOperationInfo insertAfter));
        Assert.True(DocxHelp.TryGetPatchOperation("delete-row", out DocxPatchOperationInfo delete));
        Assert.True(DocxHelp.TryGetPatchOperation("set-row-header", out DocxPatchOperationInfo header));
        foreach (string example in insertBefore.Examples.Concat(insertAfter.Examples).Concat(delete.Examples).Concat(header.Examples))
        {
            using MemoryStream input = CreateDocxWithSimpleTwoByTwoTable();
            using var patch = new StringReader(example);
            DocxCheckResult result = new DocxEditor().Check(input, patch);
            Assert.True(result.Success, example + ":" + string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        }
    }

    [Fact]
    public static void FieldExamplesCheckClean()
    {
        Assert.True(DocxHelp.TryGetPatchOperation("set-field-dirty", out DocxPatchOperationInfo dirty));
        Assert.True(DocxHelp.TryGetPatchOperation("set-field-lock", out DocxPatchOperationInfo locked));
        Assert.True(DocxHelp.TryGetPatchOperation("set-field-code", out DocxPatchOperationInfo code));
        Assert.True(DocxHelp.TryGetPatchOperation("set-field-result", out DocxPatchOperationInfo result));
        Assert.True(DocxHelp.TryGetPatchOperation("refresh-field-result", out DocxPatchOperationInfo refresh));
        foreach (string example in dirty.Examples.Concat(locked.Examples).Concat(code.Examples).Concat(result.Examples).Concat(refresh.Examples))
        {
            using MemoryStream input = CreateDocxWithRefField();
            using var patch = new StringReader(example);
            DocxCheckResult check = new DocxEditor().Check(input, patch);
            Assert.True(check.Success, example + ":" + string.Join("|", check.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        }
    }

    [Fact]
    public static void CommentThreadExamplesCheckClean()
    {
        Assert.True(DocxHelp.TryGetPatchOperation("resolve-comment", out DocxPatchOperationInfo resolve));
        Assert.True(DocxHelp.TryGetPatchOperation("reopen-comment", out DocxPatchOperationInfo reopen));
        Assert.True(DocxHelp.TryGetPatchOperation("delete-comment", out DocxPatchOperationInfo delete));
        Assert.True(DocxHelp.TryGetPatchOperation("add-comment-reply", out DocxPatchOperationInfo reply));
        Assert.True(DocxHelp.TryGetPatchOperation("delete-comment-reply", out DocxPatchOperationInfo deleteReply));
        foreach (string example in resolve.Examples.Concat(reopen.Examples).Concat(delete.Examples).Concat(reply.Examples))
        {
            using MemoryStream input = CreateDocxWithCommentAnchoredParagraph();
            using var patch = new StringReader(example);
            DocxCheckResult check = new DocxEditor().Check(input, patch);
            Assert.True(check.Success, example + ":" + string.Join("|", check.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        }
        string replyExample = Assert.Single(reply.Examples);
        using MemoryStream replyInput = CreateDocxWithCommentAnchoredParagraph();
        using var threaded = new MemoryStream();
        using (var replyPatch = new StringReader(replyExample))
        {
            DocxApplyResult applied = new DocxEditor().Apply(replyInput, replyPatch, threaded);
            Assert.True(applied.Success, string.Join("|", applied.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        }
        string deleteExample = Assert.Single(deleteReply.Examples);
        threaded.Position = 0;
        using var deletePatch = new StringReader(deleteExample);
        DocxCheckResult deleted = new DocxEditor().Check(threaded, deletePatch);
        Assert.True(deleted.Success, deleteExample + ":" + string.Join("|", deleted.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void HyperlinkManagementExamplesCheckClean()
    {
        const string body = """
                    <w:p>
                      <w:bookmarkStart w:id="1" w:name="Destination"/>
                      <w:bookmarkEnd w:id="1"/>
                      <w:hyperlink w:anchor="Destination"><w:r><w:t>Old Link</w:t></w:r></w:hyperlink>
                    </w:p>
            """;
        Assert.True(DocxHelp.TryGetPatchOperation("set-hyperlink-target", out DocxPatchOperationInfo target));
        Assert.True(DocxHelp.TryGetPatchOperation("remove-hyperlink", out DocxPatchOperationInfo remove));
        Assert.True(DocxHelp.TryGetPatchOperation("insert-hyperlink-after", out DocxPatchOperationInfo insert));
        foreach (string example in target.Examples.Concat(remove.Examples).Concat(insert.Examples))
        {
            using MemoryStream input = CreateDocxWithBody(body);
            using var patch = new StringReader(example);
            DocxCheckResult check = new DocxEditor().Check(input, patch);
            Assert.True(check.Success, example + ":" + string.Join("|", check.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        }
    }
    [Fact]
    public static void TablePropertyExamplesCheckClean()
    {
        const string styles = """
              <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>
            """;
        const string body = """
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblCaption w:val="Ledger"/></w:tblPr>
                      <w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
            """;
        Assert.True(DocxHelp.TryGetPatchOperation("set-table-style", out DocxPatchOperationInfo style));
        Assert.True(DocxHelp.TryGetPatchOperation("set-table-metadata", out DocxPatchOperationInfo metadata));
        foreach (string example in style.Examples.Concat(metadata.Examples))
        {
            using MemoryStream input = CreateDocxWithStylesAndBody(styles, body);
            using var patch = new StringReader(example);
            DocxCheckResult check = new DocxEditor().Check(input, patch);
            Assert.True(check.Success, example + ":" + string.Join("|", check.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        }
    }

    [Fact]
    public static void ContentControlKindExamplesCheckClean()
    {
        const string checkboxBody = """
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:checkBox>
                            <w:checked w:val="0"/>
                            <w:checkedState w:val="2612"/>
                            <w:uncheckedState w:val="2610"/>
                          </w:checkBox>
                          <w:tag w:val="accepted"/>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>Unchecked</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """;
        const string choiceBody = """
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:dropDownList>
                            <w:listItem w:displayText="North" w:value="north"/>
                            <w:listItem w:displayText="South" w:value="south"/>
                          </w:dropDownList>
                          <w:tag w:val="region"/>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>North</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """;
        const string dateBody = """
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:date>
                            <w:dateFormat w:val="yyyy-MM-dd"/>
                            <w:fullDate w:val="2026-06-12T00:00:00Z"/>
                          </w:date>
                          <w:tag w:val="deadline"/>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>2026-06-12</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """;
        Assert.True(DocxHelp.TryGetPatchOperation("set-content-control-checkbox", out DocxPatchOperationInfo checkbox));
        Assert.True(DocxHelp.TryGetPatchOperation("set-content-control-choice", out DocxPatchOperationInfo choice));
        Assert.True(DocxHelp.TryGetPatchOperation("set-content-control-date", out DocxPatchOperationInfo date));
        foreach ((string example, string body) in new[] { (Assert.Single(checkbox.Examples), checkboxBody), (Assert.Single(choice.Examples), choiceBody), (Assert.Single(date.Examples), dateBody) })
        {
            using MemoryStream input = CreateDocxWithBody(body);
            using var patch = new StringReader(example);
            DocxCheckResult check = new DocxEditor().Check(input, patch);
            Assert.True(check.Success, example + ":" + string.Join("|", check.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        }
    }
    [Fact]
    public static void BookmarkRenameDeleteExamplesCheckClean()
    {
        Assert.True(DocxHelp.TryGetPatchOperation("rename-bookmark", out DocxPatchOperationInfo rename));
        Assert.True(DocxHelp.TryGetPatchOperation("delete-bookmark", out DocxPatchOperationInfo delete));
        foreach (string example in rename.Examples.Concat(delete.Examples))
        {
            using MemoryStream input = CreateDocxWithSingleBookmark();
            using var patch = new StringReader(example);
            DocxCheckResult check = new DocxEditor().Check(input, patch);
            Assert.True(check.Success, example + ":" + string.Join("|", check.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        }
    }

    [Fact]
    public static void ImageExamplesCheckClean()
    {
        var assets = new MemoryAssetProvider("chart.png", CreatePngBytes(4, 3), null, "chart.png");
        var assetOptions = new DocxEditOptions { AssetProvider = assets };
        Assert.True(DocxHelp.TryGetPatchOperation("replace-image", out DocxPatchOperationInfo replace));
        Assert.True(DocxHelp.TryGetPatchOperation("insert-image-after", out DocxPatchOperationInfo insert));
        Assert.True(DocxHelp.TryGetPatchOperation("set-image-alt", out DocxPatchOperationInfo alt));
        Assert.True(DocxHelp.TryGetPatchOperation("set-image-metadata", out DocxPatchOperationInfo metadata));
        Assert.True(DocxHelp.TryGetPatchOperation("set-image-size", out DocxPatchOperationInfo size));
        Assert.True(DocxHelp.TryGetPatchOperation("set-image-crop", out DocxPatchOperationInfo crop));
        Assert.True(DocxHelp.TryGetPatchOperation("delete-image", out DocxPatchOperationInfo delete));
        foreach (string example in replace.Examples.Concat(alt.Examples).Concat(metadata.Examples).Concat(size.Examples).Concat(crop.Examples).Concat(delete.Examples))
        {
            using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
            using var patch = new StringReader(example);
            DocxCheckResult check = new DocxEditor().Check(input, patch, assetOptions);
            Assert.True(check.Success, example + ":" + string.Join("|", check.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        }
        foreach (string example in insert.Examples)
        {
            using MemoryStream input = CreateDocx("Alpha");
            using var patch = new StringReader(example);
            DocxCheckResult check = new DocxEditor().Check(input, patch, assetOptions);
            Assert.True(check.Success, example + ":" + string.Join("|", check.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        }
        Assert.True(DocxHelp.TryGetPatchOperation("set-image-wrap", out DocxPatchOperationInfo wrap));
        Assert.True(DocxHelp.TryGetPatchOperation("set-image-position", out DocxPatchOperationInfo position));
        foreach (string example in wrap.Examples.Concat(position.Examples))
        {
            using MemoryStream input = CreateDocxWithAnchoredImage();
            using var patch = new StringReader(example);
            DocxCheckResult check = new DocxEditor().Check(input, patch);
            Assert.True(check.Success, example + ":" + string.Join("|", check.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        }
    }

    [Fact]
    public static void EverySupportedOperationHasAnExample()
    {
        int covered = 0;
        foreach (DocxPatchOperationInfo operation in DocxHelp.Catalog.PatchOperations)
        {
            if (string.Equals(operation.TrackChangesSupportClass, "unsupported", StringComparison.Ordinal))
            {
                continue;
            }
            Assert.NotEmpty(operation.Examples);
            covered++;
        }
        Assert.True(covered >= 50, "Expected at least 50 supported operations with examples, found " + covered + ".");
    }

    [Fact]
    public static void FindHelpDocumentsMatchingSemantics()
    {
        string topic = DocxHelp.RenderTopic("find");
        Assert.Contains("Matching:", topic, StringComparison.Ordinal);
        Assert.Contains("case-insensitive", topic, StringComparison.Ordinal);
        Assert.Contains("--headers-footers", topic, StringComparison.Ordinal);
    }

}
