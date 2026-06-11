using System.Diagnostics;
using System.IO.Compression;
using System.Text;

namespace DocxEdit.Tests;

public static class CliTests
{
    [Fact]
    public static void CliReadReturnsZeroAndPrintsDocumentStructure()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = Path.Combine(temp.Path, "input.docx");
        CreateDocx(input);

        CliResult result = RunCli("read", input);

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("M.P0001", result.Output, StringComparison.Ordinal);
        Assert.Contains("Revenue increased", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliReturnsTwoForInvalidArguments()
    {
        CliResult result = RunCli("dump", "input.docx");

        Assert.Equal(2, result.ExitCode);
        Assert.Contains("Usage: docxedit dump", result.Error, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliCheckWritesJsonReport()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = Path.Combine(temp.Path, "input.docx");
        string patch = Path.Combine(temp.Path, "edits.docxpatch");
        string report = Path.Combine(temp.Path, "report.json");
        CreateDocx(input);
        File.WriteAllText(patch, "docxpatch 1" + Environment.NewLine);

        CliResult result = RunCli("check", input, patch, "--report", report);

        Assert.Equal(0, result.ExitCode);
        Assert.True(File.Exists(report));
        Assert.Contains("\"success\": true", File.ReadAllText(report), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public static void CliStylesPrintsTextOutput()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = Path.Combine(temp.Path, "input.docx");
        CreateDocx(input);

        CliResult result = RunCli("styles", input);

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("paragraph styleId=Normal", result.Output, StringComparison.Ordinal);
        Assert.Contains("table styleId=TableGrid", result.Output, StringComparison.Ordinal);
    }

    private static CliResult RunCli(params string[] args)
    {
        string repoRoot = FindRepoRoot();
        var startInfo = new ProcessStartInfo("dotnet")
        {
            WorkingDirectory = repoRoot,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        startInfo.ArgumentList.Add("run");
        startInfo.ArgumentList.Add("--project");
        startInfo.ArgumentList.Add(Path.Combine(repoRoot, "src", "DocxEdit.Cli", "DocxEdit.Cli.csproj"));
        startInfo.ArgumentList.Add("--");
        foreach (string arg in args)
        {
            startInfo.ArgumentList.Add(arg);
        }

        using Process process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Failed to start CLI process.");
        string output = process.StandardOutput.ReadToEnd();
        string error = process.StandardError.ReadToEnd();
        process.WaitForExit();
        return new CliResult(process.ExitCode, output, error);
    }

    private static string FindRepoRoot()
    {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "DocxEdit.slnx")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new InvalidOperationException("Could not find repository root.");
    }

    private static void CreateDocx(string path)
    {
        using FileStream file = File.Create(path);
        using var archive = new ZipArchive(file, ZipArchiveMode.Create);
        AddEntry(archive, "[Content_Types].xml", """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
              <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
            </Types>
            """);
        AddEntry(archive, "_rels/.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """);
        AddEntry(archive, "word/_rels/document.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
            </Relationships>
            """);
        AddEntry(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p><w:r><w:t>Revenue increased</w:t></w:r></w:p>
              </w:body>
            </w:document>
            """);
        AddEntry(archive, "word/styles.xml", """
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
                <w:name w:val="Normal"/>
              </w:style>
              <w:style w:type="table" w:styleId="TableGrid">
                <w:name w:val="Table Grid"/>
              </w:style>
            </w:styles>
            """);
    }

    private static void AddEntry(ZipArchive archive, string name, string text)
    {
        ZipArchiveEntry entry = archive.CreateEntry(name);
        using Stream stream = entry.Open();
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        stream.Write(bytes, 0, bytes.Length);
    }

    private sealed record CliResult(int ExitCode, string Output, string Error);

    private sealed class TempDirectory : IDisposable
    {
        private TempDirectory(string path)
        {
            Path = path;
        }

        public string Path { get; }

        public static TempDirectory Create()
        {
            string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "docxedit-tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(path);
            return new TempDirectory(path);
        }

        public void Dispose()
        {
            if (Directory.Exists(Path))
            {
                Directory.Delete(Path, recursive: true);
            }
        }
    }
}
