using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

[Collection("ConsoleCli")]
public static class CliCreateTests
{
    [Fact]
    public static void CreatesAndInspectsDocumentThroughFilesystemCli()
    {
        using TempDirectory temp = TempDirectory.Create();
        string path = Path.Combine(temp.Path, "created.docx");
        File.WriteAllText(path, "existing destination");
        CliTests.CliResult create = CliTests.RunCli("create", "--output", path, "--paper", "letter", "--orientation", "landscape");
        Assert.Equal(0, create.ExitCode);
        Assert.Contains("docxedit create: OK", create.Output);
        Assert.Equal(0, CliTests.RunCli("validate", path).ExitCode);
        CliTests.CliResult read = CliTests.RunCli("read", path);
        Assert.Equal(0, read.ExitCode);
        Assert.Contains("page-width-twips=15840 page-height-twips=12240", read.Output);
        Assert.Contains("margin-top-twips=1440", read.Output);
        Assert.Contains("M.P0001", read.Output);
        Assert.Single(Directory.GetFiles(temp.Path));
    }

    [Fact]
    public static void InvalidCreationPreservesExistingFile()
    {
        using TempDirectory temp = TempDirectory.Create();
        string path = Path.Combine(temp.Path, "created.docx");
        File.WriteAllText(path, "existing destination");
        CliTests.CliResult create = CliTests.RunCli("create", "--output", path, "--orientation", "invalid");
        Assert.Equal(2, create.ExitCode);
        Assert.Equal("existing destination", File.ReadAllText(path));
    }
}
