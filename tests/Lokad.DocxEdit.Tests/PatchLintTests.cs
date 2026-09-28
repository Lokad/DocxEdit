using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D10: document-independent shape validation predicts check failures for
// missing fields, alternative groups, and exclusive fields.
public static class PatchLintTests
{
    private static DocxLintResult Lint(string patchText)
    {
        using var patch = new StringReader(patchText);
        return new DocxEditor().Lint(patch);
    }

    private static DocxCheckResult RunCheck(MemoryStream input, string patchText)
    {
        input.Position = 0;
        using var patch = new StringReader(patchText);
        return new DocxEditor().Check(input, patch);
    }

    [Fact]
    public static void MissingRequiredFieldFailsLintAndCheck()
    {
        const string patchText = """
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            end
            """;
        DocxLintResult lint = Lint(patchText);

        Assert.False(lint.Success);
        Assert.Equal(1, lint.OperationCount);
        DocxDiagnostic diagnostic = Assert.Single(lint.Diagnostics);
        Assert.Equal("E4202", diagnostic.Code);
        Assert.Equal(1, diagnostic.OperationIndex);
        Assert.Contains("missing required field", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("with", diagnostic.Message, StringComparison.Ordinal);

        using MemoryStream input = CreateDocx("Alpha Beta");
        DocxCheckResult check = RunCheck(input, patchText);
        Assert.False(check.Success);
        Assert.Contains(check.Diagnostics, checkDiagnostic => checkDiagnostic.Code == "E4202");
    }

    [Fact]
    public static void EmptyRequiredFieldFailsWhileAllowEmptyPasses()
    {
        DocxLintResult emptyFind = Lint("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind <<<\n>>>\nwith X\nend\n");

        Assert.False(emptyFind.Success);
        Assert.Contains(emptyFind.Diagnostics, static diagnostic => diagnostic.Code == "E4202");

        DocxLintResult emptyWith = Lint("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith <<<\n>>>\nend\n");

        Assert.True(emptyWith.Success, string.Join("|", emptyWith.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        Assert.Equal(1, emptyWith.OperationCount);
    }

    [Fact]
    public static void MissingAlternativeGroupFailsLintAndCheck()
    {
        const string patchText = """
            docxpatch 1

            op set-content-control-choice
            target M.CC0001
            end
            """;
        DocxLintResult lint = Lint(patchText);

        Assert.False(lint.Success);
        Assert.Contains(lint.Diagnostics, static diagnostic => diagnostic.Code == "E4202");
    }

    [Fact]
    public static void ExclusiveFieldsFailLintAndCheck()
    {
        const string patchText = """
            docxpatch 1

            op set-content-control-choice
            target M.CC0001
            value north
            display-text North
            end
            """;
        DocxLintResult lint = Lint(patchText);

        Assert.False(lint.Success);
        DocxDiagnostic diagnostic = Assert.Single(lint.Diagnostics);
        Assert.Equal("E4205", diagnostic.Code);
        Assert.Equal(1, diagnostic.OperationIndex);
    }

    [Fact]
    public static void ShadingFillAndClearTrueFailLintAndCheck()
    {
        const string patchText = """
            docxpatch 1

            op set-cell-shading
            target M.T0001.R01.C01
            fill 4472C4
            clear true
            end
            """;
        DocxLintResult lint = Lint(patchText);

        Assert.False(lint.Success);
        Assert.Contains(lint.Diagnostics, static diagnostic => diagnostic.Code == "E4205");

        using MemoryStream input = CreateDocxWithSimpleTwoByTwoTable();
        DocxCheckResult check = RunCheck(input, patchText);
        Assert.False(check.Success);
    }

    [Fact]
    public static void ClearFalseCountsAsAbsent()
    {
        const string patchText = """
            docxpatch 1

            op set-cell-shading
            target M.T0001.R01.C01
            clear false
            end
            """;
        DocxLintResult lint = Lint(patchText);

        Assert.False(lint.Success);
        Assert.Contains(lint.Diagnostics, static diagnostic => diagnostic.Code == "E4202");

        using MemoryStream input = CreateDocxWithSimpleTwoByTwoTable();
        DocxCheckResult check = RunCheck(input, patchText);
        Assert.False(check.Success);
        Assert.Contains(check.Diagnostics, static diagnostic => diagnostic.Code == "E4202");
    }

    [Fact]
    public static void HyperlinkUriAndAnchorConflictFailsLint()
    {
        DocxLintResult lint = Lint("docxpatch 1\n\nop set-hyperlink-target\ntarget M.L0001\nuri https://example.test/x\nanchor Mark\nend\n");

        Assert.False(lint.Success);
        Assert.Contains(lint.Diagnostics, static diagnostic => diagnostic.Code == "E4205");
    }

    [Fact]
    public static void ValidPatchLintsClean()
    {
        const string patchText = """
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            with Omega
            end

            op set-cell-shading
            target M.T0001.R01.C01
            fill 4472C4
            end
            """;
        DocxLintResult lint = Lint(patchText);

        Assert.True(lint.Success, string.Join("|", lint.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        Assert.Equal(2, lint.OperationCount);
        Assert.Equal(["replace-text", "set-cell-shading"], lint.Operations.Select(static operation => operation.Name).ToArray());
        Assert.Equal([1, 2], lint.Operations.Select(static operation => operation.Index).ToArray());
    }

    [Fact]
    public static void SyntaxErrorsSurfaceThroughLint()
    {
        DocxLintResult unknown = Lint("docxpatch 1\n\nop frobnicate\nend\n");

        Assert.False(unknown.Success);
        Assert.Equal(0, unknown.OperationCount);
        Assert.Contains(unknown.Diagnostics, static diagnostic => diagnostic.Code == "E2010");

        DocxLintResult badBoolean = Lint("docxpatch 1\n\nop set-cell\ntarget M.T0001.R01.C01\ntext X\nforce maybe\nend\n");

        Assert.False(badBoolean.Success);
        Assert.Contains(badBoolean.Diagnostics, static diagnostic => diagnostic.Code == "E2012");
    }

    [Fact]
    public static void ExclusiveGroupsAreMachineReadable()
    {
        Assert.True(DocxHelp.TryGetPatchOperation("set-content-control-choice", out DocxPatchOperationInfo choice));
        Assert.Contains(choice.ExclusiveAlternatives, group => group.Contains("value", StringComparer.Ordinal) && group.Contains("display-text", StringComparer.Ordinal));
        Assert.True(DocxHelp.TryGetPatchOperation("set-cell-shading", out DocxPatchOperationInfo shading));
        Assert.Contains(shading.ExclusiveAlternatives, group => group.Contains("fill", StringComparer.Ordinal) && group.Contains("clear", StringComparer.Ordinal));

        string topic = DocxHelp.RenderTopic("set-cell-shading");
        Assert.Contains("Exclusive fields (at most one):", topic, StringComparison.Ordinal);
        Assert.Contains("fill|clear", topic, StringComparison.Ordinal);
    }
}