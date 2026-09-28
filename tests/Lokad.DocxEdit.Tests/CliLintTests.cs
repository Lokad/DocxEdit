using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// Serializes with CliTests and EditCaseTests: both redirect the process-wide Console.
[Collection("ConsoleCli")]
public static class CliLintTests
{
    private static string WritePatch(TempDirectory temp, string name, string patchText)
    {
        string path = Path.Combine(temp.Path, name);
        File.WriteAllText(path, patchText);
        return path;
    }

    [Fact]
    public static void CliLintTextPrintsOperations()
    {
        using TempDirectory temp = TempDirectory.Create();
        string patch = WritePatch(temp, "edits.docxpatch", "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Omega\nend\n");

        CliTests.CliResult result = CliTests.RunCli("lint", patch);

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("lint operations=1", result.Output, StringComparison.Ordinal);
        Assert.Contains("op 1 replace-text", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliLintMissingFieldExitsOne()
    {
        using TempDirectory temp = TempDirectory.Create();
        string patch = WritePatch(temp, "edits.docxpatch", "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nend\n");

        CliTests.CliResult result = CliTests.RunCli("lint", patch);

        Assert.Equal(1, result.ExitCode);
        Assert.Contains("E4202", result.Error, StringComparison.Ordinal);
        Assert.Contains("missing required field", result.Error, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliLintJsonEmitsResult()
    {
        using TempDirectory temp = TempDirectory.Create();
        string patch = WritePatch(temp, "edits.docxpatch", "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Omega\nend\n");

        CliTests.CliResult result = CliTests.RunCli("lint", patch, "--json");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("\"OperationCount\": 1", result.Output, StringComparison.Ordinal);
        Assert.Contains("\"Name\": \"replace-text\"", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliLintUnknownOperationExitsOne()
    {
        using TempDirectory temp = TempDirectory.Create();
        string patch = WritePatch(temp, "edits.docxpatch", "docxpatch 1\n\nop frobnicate\nend\n");

        CliTests.CliResult result = CliTests.RunCli("lint", patch);

        Assert.Equal(1, result.ExitCode);
        Assert.Contains("E2010", result.Error, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliLintMissingPositionalExitsTwoWithUsage()
    {
        CliTests.CliResult result = CliTests.RunCli("lint");

        Assert.Equal(2, result.ExitCode);
        Assert.Contains("Usage: docxedit lint", result.Error, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliLintRejectsForeignOption()
    {
        using TempDirectory temp = TempDirectory.Create();
        string patch = WritePatch(temp, "edits.docxpatch", "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Omega\nend\n");

        CliTests.CliResult result = CliTests.RunCli("lint", patch, "--id", "M.P0001");

        Assert.Equal(2, result.ExitCode);
        Assert.Contains("--id is not an option of the lint command.", result.Error, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliLintHelpTopicRenders()
    {
        CliTests.CliResult result = CliTests.RunCli("help", "lint");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("docxedit lint edits.docxpatch", result.Output, StringComparison.Ordinal);
    }
}