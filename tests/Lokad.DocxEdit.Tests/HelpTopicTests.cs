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
    public static void PatchOperationTopicRendersMinimalAndGuardedExamples()
    {
        string topic = DocxHelp.RenderTopic("replace-text");

        Assert.Contains("# Minimal replace-text.", topic, StringComparison.Ordinal);
        Assert.Contains("# Guarded replace-text:", topic, StringComparison.Ordinal);
        Assert.Contains("op replace-text", topic, StringComparison.Ordinal);
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
}
