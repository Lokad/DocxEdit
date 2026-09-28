using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// Serializes with CliTests and EditCaseTests: both redirect the process-wide Console.
[Collection("ConsoleCli")]
public static class CliCapabilitiesTests
{
    private static string WriteFixture(TempDirectory temp, string name, MemoryStream docx)
    {
        string path = Path.Combine(temp.Path, name);
        using FileStream file = File.Create(path);
        docx.CopyTo(file);
        return path;
    }

    [Fact]
    public static void CliCapabilitiesTextPrintsTargetAndOperations()
    {
        using TempDirectory temp = TempDirectory.Create();
        using MemoryStream docx = CreateDocx("Alpha Beta");
        string input = WriteFixture(temp, "input.docx", docx);

        CliTests.CliResult result = CliTests.RunCli("capabilities", input, "--id", "M.P0001");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("target M.P0001 paragraph main", result.Output, StringComparison.Ordinal);
        Assert.Contains("supported replace-text:", result.Output, StringComparison.Ordinal);
        Assert.Contains("supported set-style:", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliCapabilitiesJsonEmitsStructuredOperations()
    {
        using TempDirectory temp = TempDirectory.Create();
        using MemoryStream docx = CreateDocx("Alpha Beta");
        string input = WriteFixture(temp, "input.docx", docx);

        CliTests.CliResult result = CliTests.RunCli("capabilities", input, "--id", "M.P0001", "--json");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("\"TargetId\": \"M.P0001\"", result.Output, StringComparison.Ordinal);
        Assert.Contains("\"Kind\": \"paragraph\"", result.Output, StringComparison.Ordinal);
        Assert.Contains("\"Operation\": \"replace-text\"", result.Output, StringComparison.Ordinal);
        Assert.Contains("\"Support\": \"supported\"", result.Output, StringComparison.Ordinal);
        Assert.Contains("\"HelpTopic\": \"replace-text\"", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliCapabilitiesRequirePolicySurfacesBookmarkRefusal()
    {
        using TempDirectory temp = TempDirectory.Create();
        using MemoryStream docx = CreateDocx("Alpha Beta");
        string input = WriteFixture(temp, "input.docx", docx);

        CliTests.CliResult result = CliTests.RunCli("capabilities", input, "--id", "M.P0001", "--track-changes", "require");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("unsupported add-bookmark:", result.Output, StringComparison.Ordinal);
        Assert.Contains("E6001", result.Output, StringComparison.Ordinal);
        Assert.Contains("supported add-comment:", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliCapabilitiesContinuationCellNamesRootAlternative()
    {
        using TempDirectory temp = TempDirectory.Create();
        using MemoryStream docx = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
                          <w:p><w:r><w:t>North</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:vMerge/></w:tcPr>
                          <w:p><w:r><w:t>South</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc><w:p><w:r><w:t>Profit</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        string input = WriteFixture(temp, "input.docx", docx);

        CliTests.CliResult result = CliTests.RunCli("capabilities", input, "--id", "M.T0001.R02.C01");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("target M.T0001.R02.C01 cell main", result.Output, StringComparison.Ordinal);
        Assert.Contains("unsupported set-cell alternative=M.T0001.R01.C01:", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliCapabilitiesUnknownTargetExitsOne()
    {
        using TempDirectory temp = TempDirectory.Create();
        using MemoryStream docx = CreateDocx("Alpha");
        string input = WriteFixture(temp, "input.docx", docx);

        CliTests.CliResult result = CliTests.RunCli("capabilities", input, "--id", "M.P0099");

        Assert.Equal(1, result.ExitCode);
        Assert.Contains("E1201", result.Error, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliCapabilitiesMissingIdExitsTwoWithUsage()
    {
        using TempDirectory temp = TempDirectory.Create();
        using MemoryStream docx = CreateDocx("Alpha");
        string input = WriteFixture(temp, "input.docx", docx);

        CliTests.CliResult result = CliTests.RunCli("capabilities", input);

        Assert.Equal(2, result.ExitCode);
        Assert.Contains("Usage: docxedit capabilities", result.Error, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliCapabilitiesRejectsForeignOption()
    {
        using TempDirectory temp = TempDirectory.Create();
        using MemoryStream docx = CreateDocx("Alpha");
        string input = WriteFixture(temp, "input.docx", docx);

        CliTests.CliResult result = CliTests.RunCli("capabilities", input, "--id", "M.P0001", "--runs");

        Assert.Equal(2, result.ExitCode);
        Assert.Contains("--runs is not an option of the capabilities command.", result.Error, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliCapabilitiesRejectsInvalidTrackChangesMode()
    {
        using TempDirectory temp = TempDirectory.Create();
        using MemoryStream docx = CreateDocx("Alpha");
        string input = WriteFixture(temp, "input.docx", docx);

        CliTests.CliResult result = CliTests.RunCli("capabilities", input, "--id", "M.P0001", "--track-changes", "sometimes");

        Assert.Equal(2, result.ExitCode);
        Assert.Contains("Invalid value for --track-changes.", result.Error, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliCapabilitiesHelpTopicRenders()
    {
        CliTests.CliResult result = CliTests.RunCli("help", "capabilities");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("docxedit capabilities input.docx --id M.P0001", result.Output, StringComparison.Ordinal);
        Assert.Contains("--track-changes", result.Output, StringComparison.Ordinal);
    }
}