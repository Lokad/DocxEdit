using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// Serializes with CliTests and EditCaseTests: both redirect the process-wide Console.
[Collection("ConsoleCli")]
public static class CliPolicyTests
{
    private const string PatchText = "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Omega\nend\n";

    private static string WriteDocx(TempDirectory temp, string name)
    {
        string path = Path.Combine(temp.Path, name);
        using MemoryStream docx = CreateDocx("Alpha Beta");
        using FileStream file = File.Create(path);
        docx.CopyTo(file);
        return path;
    }

    private static string WritePatch(TempDirectory temp, string name)
    {
        string path = Path.Combine(temp.Path, name);
        File.WriteAllText(path, PatchText);
        return path;
    }

    [Fact]
    public static void CliCheckJsonRecordsPolicy()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = WriteDocx(temp, "input.docx");
        string patch = WritePatch(temp, "edit.docxpatch");

        CliTests.CliResult result = CliTests.RunCli("check", input, patch, "--track-changes", "suggest", "--json");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("\"TrackChanges\": \"suggest\"", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliCheckTextHeaderRecordsPolicy()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = WriteDocx(temp, "input.docx");
        string patch = WritePatch(temp, "edit.docxpatch");

        CliTests.CliResult result = CliTests.RunCli("check", input, patch, "--track-changes", "require");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("track-changes=require", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliApplyReportRoundTripsPolicy()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = WriteDocx(temp, "input.docx");
        string patch = WritePatch(temp, "edit.docxpatch");
        string output = Path.Combine(temp.Path, "output.docx");
        string report = Path.Combine(temp.Path, "report.json");

        CliTests.CliResult apply = CliTests.RunCli("apply", input, patch, "--output", output, "--report", report, "--track-changes", "require");

        Assert.Equal(0, apply.ExitCode);
        Assert.Contains("track-changes=require", apply.Output, StringComparison.Ordinal);
        string reportJson = File.ReadAllText(report);
        Assert.Contains("\"TrackChanges\": \"require\"", reportJson, StringComparison.Ordinal);

        CliTests.CliResult changes = CliTests.RunCli("changes", output, "--operation-report", report);

        Assert.Equal(0, changes.ExitCode);
    }

    [Fact]
    public static void CliChangesAcceptsReportWithoutPolicy()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = WriteDocx(temp, "input.docx");
        string report = Path.Combine(temp.Path, "legacy.json");
        File.WriteAllText(report, "{\"Success\":true,\"Diagnostics\":[],\"Operations\":[],\"Author\":\"docxedit\",\"TimestampUtc\":\"2026-01-01T00:00:00Z\",\"ToolVersion\":\"1.0\"}");

        CliTests.CliResult changes = CliTests.RunCli("changes", input, "--operation-report", report);

        Assert.Equal(0, changes.ExitCode);
    }
}