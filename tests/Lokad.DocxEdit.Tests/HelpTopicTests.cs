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
}
