using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D10: explicit target ID kinds are validated without a document; lint
// failures always predict the matching check failures.
public static class TargetKindLintTests
{
    private static string BuildPatch(string operation, string target, string fields)
    {
        return "docxpatch 1\n\nop " + operation + "\ntarget " + target + "\n" + fields + "end\n";
    }

    private static DocxLintResult LintPatch(string patchText)
    {
        using var patch = new StringReader(patchText);
        return new DocxEditor().Lint(patch);
    }

    private static DocxCheckResult CheckPatch(MemoryStream input, string patchText)
    {
        input.Position = 0;
        using var patch = new StringReader(patchText);
        return new DocxEditor().Check(input, patch);
    }

    [Theory]
    [InlineData("replace-text", "M.T0001", "find Alpha\nwith Omega\n")]
    [InlineData("insert-after", "M.T0001.R01.C01", "text X\n")]
    [InlineData("delete-block", "M.CC0001", "")]
    [InlineData("set-style", "M.T0001", "style Normal\n")]
    [InlineData("add-comment", "M.T0001", "text Note\n")]
    [InlineData("add-bookmark", "M.T0001", "name B1\n")]
    [InlineData("set-content-control-text", "M.P0001", "text X\n")]
    [InlineData("set-content-control-checkbox", "M.P0001", "checked true\n")]
    [InlineData("replace-bookmark-text", "M.P0001", "text X\n")]
    [InlineData("rename-bookmark", "M.P0001", "name New\n")]
    [InlineData("set-field-result", "M.P0001", "text X\n")]
    [InlineData("set-hyperlink-text", "M.P0001", "text X\n")]
    [InlineData("set-image-alt", "M.P0001", "alt X\n")]
    [InlineData("set-cell", "M.P0001", "text X\n")]
    [InlineData("set-cell-shading", "M.T0001.R01", "fill 4472C4\n")]
    [InlineData("set-table-style", "M.P0001", "style TableGrid\n")]
    [InlineData("set-row-header", "M.T0001", "header true\n")]
    [InlineData("append-row", "M.P0001", "cell X\n")]
    [InlineData("insert-row-before", "M.T0001", "cell X\n")]
    [InlineData("delete-row", "M.T0001", "")]
    [InlineData("set-section-columns", "M.P0001", "count 2\n")]
    [InlineData("insert-image-after", "M.T0001", "asset x.png\n")]
    [InlineData("insert-hyperlink-after", "M.CC0001", "text X\nuri https://example.test/x\n")]
    [InlineData("delete-bookmark", "M.P0001", "")]
    [InlineData("set-table-metadata", "M.P0001", "caption X\n")]
    [InlineData("set-content-control-choice", "M.P0001", "value north\n")]
    [InlineData("set-hyperlink-target", "M.P0001", "uri https://example.test/x\n")]
    [InlineData("set-image-size", "M.P0001", "width 100\n")]
    [InlineData("set-section-orientation", "M.P0001", "orientation portrait\n")]
    public static void WrongKindTargetFailsLintAndCheck(string operation, string target, string fields)
    {
        string patchText = BuildPatch(operation, target, fields);
        DocxLintResult lint = LintPatch(patchText);

        Assert.False(lint.Success);
        DocxDiagnostic lintDiagnostic = Assert.Single(lint.Diagnostics, static diagnostic => diagnostic.Code == "E1201");
        Assert.Equal(target, lintDiagnostic.TargetId);
        Assert.Equal(4, lintDiagnostic.Line);
        Assert.Equal(1, lintDiagnostic.Column);

        using MemoryStream input = CreateDocx("Alpha");
        DocxCheckResult check = CheckPatch(input, patchText);

        Assert.False(check.Success);
        Assert.Contains(check.Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }    [Fact]
    public static void ParagraphWrongKindMessageMatchesCheck()
    {
        string patchText = BuildPatch("replace-text", "M.T0001", "find Alpha\nwith Omega\n");
        DocxLintResult lint = LintPatch(patchText);

        Assert.False(lint.Success);
        string lintMessage = Assert.Single(lint.Diagnostics, static diagnostic => diagnostic.Code == "E1201").Message;

        using MemoryStream input = CreateDocx("Alpha");
        DocxCheckResult check = CheckPatch(input, patchText);

        Assert.False(check.Success);
        string checkMessage = Assert.Single(check.Diagnostics, static diagnostic => diagnostic.Code == "E1201").Message;
        Assert.Equal(checkMessage, lintMessage);
    }

    [Theory]
    [InlineData("set-comment-text", "comment:3", "text X\n")]
    [InlineData("set-field-dirty", "all", "dirty true\n")]
    [InlineData("replace-text", "text:\"Alpha\"", "find Alpha\nwith Omega\n")]
    public static void NonIdTargetsSkipKindValidation(string operation, string target, string fields)
    {
        DocxLintResult lint = LintPatch(BuildPatch(operation, target, fields));

        Assert.True(lint.Success, string.Join("|", lint.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Theory]
    [InlineData("replace-text", "M.T0001.R01.C01", "find Alpha\nwith Omega\n")]
    [InlineData("set-cell", "M.T0001.R01.C01", "text X\n")]
    [InlineData("set-cell", "M.T0001.MG0001", "text X\n")]
    [InlineData("insert-after", "M.T0001", "text X\n")]
    [InlineData("set-cell-shading", "M.T0001.R01.C01", "fill 4472C4\n")]
    public static void AcceptedKindsLintClean(string operation, string target, string fields)
    {
        DocxLintResult lint = LintPatch(BuildPatch(operation, target, fields));

        Assert.True(lint.Success, string.Join("|", lint.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void AcceptedTargetsAreMachineReadable()
    {
        Assert.True(DocxHelp.TryGetPatchOperation("replace-text", out DocxPatchOperationInfo replaceText));
        Assert.Equal(["paragraph", "cell", "merge-group"], replaceText.AcceptedTargets);
        Assert.True(DocxHelp.TryGetPatchOperation("set-cell", out DocxPatchOperationInfo setCell));
        Assert.Equal(["cell", "merge-group"], setCell.AcceptedTargets);
        Assert.True(DocxHelp.TryGetPatchOperation("insert-after", out DocxPatchOperationInfo insertAfter));
        Assert.Equal(["paragraph", "table"], insertAfter.AcceptedTargets);
        Assert.True(DocxHelp.TryGetPatchOperation("add-comment-reply", out DocxPatchOperationInfo reply));
        Assert.Empty(reply.AcceptedTargets);
    }
}
