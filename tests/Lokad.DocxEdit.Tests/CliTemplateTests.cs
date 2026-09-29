using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// Serializes with CliTests and EditCaseTests: both redirect the process-wide Console.
[Collection("ConsoleCli")]
public static class CliTemplateTests
{
    private static string WriteFixture(TempDirectory temp, string name, MemoryStream docx)
    {
        string path = Path.Combine(temp.Path, name);
        using FileStream file = File.Create(path);
        docx.CopyTo(file);
        return path;
    }

    [Fact]
    public static void CliTemplateTextPrintsStarterBlock()
    {
        using TempDirectory temp = TempDirectory.Create();
        using MemoryStream docx = CreateDocx("Alpha Beta");
        string input = WriteFixture(temp, "input.docx", docx);

        CliTests.CliResult result = CliTests.RunCli("template", input, "--id", "M.P0001");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("docxpatch 1", result.Output, StringComparison.Ordinal);
        Assert.Contains("op replace-text", result.Output, StringComparison.Ordinal);
        Assert.Contains("target M.P0001", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliTemplateOutputPassesCheck()
    {
        using TempDirectory temp = TempDirectory.Create();
        using MemoryStream docx = CreateDocx("Alpha Beta");
        string input = WriteFixture(temp, "input.docx", docx);

        CliTests.CliResult template = CliTests.RunCli("template", input, "--id", "M.P0001");
        Assert.Equal(0, template.ExitCode);
        string patchPath = Path.Combine(temp.Path, "template.docxpatch");
        File.WriteAllText(patchPath, template.Output);

        CliTests.CliResult check = CliTests.RunCli("check", input, patchPath);

        Assert.Equal(0, check.ExitCode);
    }

    [Fact]
    public static void CliCellTemplateOutputPassesCheck()
    {
        using TempDirectory temp = TempDirectory.Create();
        using MemoryStream docx = CreateDocxWithSimpleTwoByTwoTable();
        string input = WriteFixture(temp, "input.docx", docx);

        CliTests.CliResult template = CliTests.RunCli("template", input, "--id", "M.T0001.R01.C01");
        Assert.Equal(0, template.ExitCode);
        string patchPath = Path.Combine(temp.Path, "template.docxpatch");
        File.WriteAllText(patchPath, template.Output);

        CliTests.CliResult check = CliTests.RunCli("check", input, patchPath);

        Assert.Equal(0, check.ExitCode);
    }

    [Fact]
    public static void CliImageTemplateOutputPassesCheck()
    {
        using TempDirectory temp = TempDirectory.Create();
        using MemoryStream docx = CreateDocxWithImage("png", "image/png", "old-png");
        string input = WriteFixture(temp, "input.docx", docx);

        CliTests.CliResult template = CliTests.RunCli("template", input, "--id", "M.I0001");
        Assert.Equal(0, template.ExitCode);
        string patchPath = Path.Combine(temp.Path, "template.docxpatch");
        File.WriteAllText(patchPath, template.Output);

        CliTests.CliResult check = CliTests.RunCli("check", input, patchPath);

        Assert.Equal(0, check.ExitCode);
    }

    [Fact]
    public static void CliTemplateJsonEmitsTemplateAndCapabilities()
    {
        using TempDirectory temp = TempDirectory.Create();
        using MemoryStream docx = CreateDocx("Alpha Beta");
        string input = WriteFixture(temp, "input.docx", docx);

        CliTests.CliResult result = CliTests.RunCli("template", input, "--id", "M.P0001", "--json");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("\"TargetId\": \"M.P0001\"", result.Output, StringComparison.Ordinal);
        Assert.Contains("op replace-text", result.Output, StringComparison.Ordinal);
        Assert.Contains("\"Operation\": \"replace-text\"", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliTemplateUnknownTargetExitsOne()
    {
        using TempDirectory temp = TempDirectory.Create();
        using MemoryStream docx = CreateDocx("Alpha");
        string input = WriteFixture(temp, "input.docx", docx);

        CliTests.CliResult result = CliTests.RunCli("template", input, "--id", "M.P0099");

        Assert.Equal(1, result.ExitCode);
        Assert.Contains("E1201", result.Error, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliTemplateMissingIdExitsTwoWithUsage()
    {
        using TempDirectory temp = TempDirectory.Create();
        using MemoryStream docx = CreateDocx("Alpha");
        string input = WriteFixture(temp, "input.docx", docx);

        CliTests.CliResult result = CliTests.RunCli("template", input);

        Assert.Equal(2, result.ExitCode);
        Assert.Contains("Usage: docxedit template", result.Error, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliTemplateRejectsForeignOption()
    {
        using TempDirectory temp = TempDirectory.Create();
        using MemoryStream docx = CreateDocx("Alpha");
        string input = WriteFixture(temp, "input.docx", docx);

        CliTests.CliResult result = CliTests.RunCli("template", input, "--id", "M.P0001", "--runs");

        Assert.Equal(2, result.ExitCode);
        Assert.Contains("--runs is not an option of the template command.", result.Error, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliTemplateHelpTopicRenders()
    {
        CliTests.CliResult result = CliTests.RunCli("help", "template");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("docxedit template input.docx --id M.P0001", result.Output, StringComparison.Ordinal);
        Assert.Contains("--track-changes", result.Output, StringComparison.Ordinal);
    }
}