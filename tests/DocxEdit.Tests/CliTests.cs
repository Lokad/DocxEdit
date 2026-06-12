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
    public static void CliHelpIncludesRequiredExamples()
    {
        CliResult help = RunCli("--help");
        CliResult patchHelp = RunCli("help", "patch");

        Assert.Equal(0, help.ExitCode);
        Assert.Contains("docxedit read report.docx", help.Output, StringComparison.Ordinal);
        Assert.Contains("docxedit apply report.docx edits.docxpatch --output report.edited.docx", help.Output, StringComparison.Ordinal);
        Assert.Contains("tracked-change and comment markup", help.Output, StringComparison.Ordinal);
        Assert.Equal(0, patchHelp.ExitCode);
        Assert.Contains("find <<<", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("op set-cell", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("insert-after", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("expect-row-count", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("set-section-orientation", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("op replace-image", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("Unsupported fields are rejected", patchHelp.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliHasCommandSpecificHelpForAgentWorkflows()
    {
        CliResult dump = RunCli("help", "dump");
        CliResult changes = RunCli("help", "changes");
        CliResult check = RunCli("help", "check");
        CliResult apply = RunCli("help", "apply");

        Assert.Equal(0, dump.ExitCode);
        Assert.Contains("docxedit dump input.docx --id TARGET", dump.Output, StringComparison.Ordinal);
        Assert.Contains("markup=inserted-run", dump.Output, StringComparison.Ordinal);
        Assert.Contains("Runs array", dump.Output, StringComparison.Ordinal);
        Assert.Contains("not the same namespace", dump.Output, StringComparison.Ordinal);
        Assert.Equal(0, changes.ExitCode);
        Assert.Contains("GroupSummary", changes.Output, StringComparison.Ordinal);
        Assert.Contains("TargetSummary", changes.Output, StringComparison.Ordinal);
        Assert.Contains("CommentSummary", changes.Output, StringComparison.Ordinal);
        Assert.Contains("comment-anchor-target", changes.Output, StringComparison.Ordinal);
        Assert.Contains("PowerShell ConvertFrom-Json", changes.Output, StringComparison.Ordinal);
        Assert.Equal(0, check.ExitCode);
        Assert.Contains("--track-changes off|preserve|suggest|require", check.Output, StringComparison.Ordinal);
        Assert.Contains("operation line", check.Output, StringComparison.Ordinal);
        Assert.Equal(0, apply.ExitCode);
        Assert.Contains("--output path, -o path", apply.Output, StringComparison.Ordinal);
        Assert.Contains("require", apply.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliReadSummaryPrintsCompactCounts()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = Path.Combine(temp.Path, "input.docx");
        CreateDocx(input);

        CliResult result = RunCli("read", input, "--summary");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("paragraphs count=1", result.Output, StringComparison.Ordinal);
        Assert.Contains("images count=1", result.Output, StringComparison.Ordinal);
        Assert.Contains("story=\"main\" paragraphs=1", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("Revenue increased", result.Output, StringComparison.Ordinal);
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
    public static void CliCheckAcceptsEditOptions()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = Path.Combine(temp.Path, "input.docx");
        string patch = Path.Combine(temp.Path, "edits.docxpatch");
        CreateDocx(input);
        File.WriteAllText(patch, "docxpatch 1" + Environment.NewLine);

        CliResult result = RunCli(
            "check",
            input,
            patch,
            "--track-changes",
            "preserve",
            "--author",
            "Agent",
            "--timestamp-utc",
            "2026-01-02T03:04:05Z",
            "--verbose");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("docxedit check: OK", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliRejectsInvalidTrackChangesMode()
    {
        CliResult result = RunCli("check", "input.docx", "edits.docxpatch", "--track-changes", "maybe");

        Assert.Equal(2, result.ExitCode);
        Assert.Contains("Invalid value for --track-changes", result.Error, StringComparison.Ordinal);
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

    [Fact]
    public static void CliMediaExtractsReferencedImages()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = Path.Combine(temp.Path, "input.docx");
        string extract = Path.Combine(temp.Path, "media");
        CreateDocx(input);

        CliResult result = RunCli("media", input, "--extract", extract);

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("M.I0001", result.Output, StringComparison.Ordinal);
        string extracted = Path.Combine(extract, "M.I0001-image1.png");
        Assert.True(File.Exists(extracted));
        Assert.Equal("fake-png", File.ReadAllText(extracted));
    }

    [Fact]
    public static void CliApplyCanResolveImageAssetFiles()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = Path.Combine(temp.Path, "input.docx");
        string output = Path.Combine(temp.Path, "output.docx");
        string patch = Path.Combine(temp.Path, "edit.docxpatch");
        string asset = Path.Combine(temp.Path, "chart.png");
        CreateDocx(input);
        File.WriteAllText(asset, "new-png");
        File.WriteAllText(patch, $"""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset {asset}
            alt Inserted chart
            end
            """);

        CliResult result = RunCli("apply", input, patch, "-o", output);

        Assert.Equal(0, result.ExitCode);
        Assert.True(File.Exists(output));
        Assert.Equal("new-png", ReadEntry(output, "word/media/image2.png"));
    }

    [Fact]
    public static void CliChangesPrintsSafeTrackedMarkupSummary()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = Path.Combine(temp.Path, "changes.docx");
        CreateDocxWithTrackedChanges(input);

        CliResult result = RunCli("changes", input);

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("inserted-run count=1", result.Output, StringComparison.Ordinal);
        Assert.Contains("deleted-run count=1", result.Output, StringComparison.Ordinal);
        Assert.Contains("summary group=story key=\"main\" type=inserted-run count=1", result.Output, StringComparison.Ordinal);
        Assert.Contains("summary group=author key=\"Alice\" type=inserted-run count=1", result.Output, StringComparison.Ordinal);
        Assert.Contains("summary group=target key=\"M.P0001\" type=deleted-run count=1", result.Output, StringComparison.Ordinal);
        Assert.Contains("target-summary target=M.P0001 count=2", result.Output, StringComparison.Ordinal);
        Assert.Contains("target-status=targeted", result.Output, StringComparison.Ordinal);
        Assert.Contains("M.CH0001 inserted-run", result.Output, StringComparison.Ordinal);
        Assert.Contains("text-length=8", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("Inserted", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("Deleted", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliApplyPrintsOperationSummary()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = Path.Combine(temp.Path, "input.docx");
        string output = Path.Combine(temp.Path, "output.docx");
        string patch = Path.Combine(temp.Path, "edit.docxpatch");
        CreateTextOnlyDocx(input);
        File.WriteAllText(patch, """
            docxpatch 1

            op replace-text
            target M.P0001
            find Revenue
            with Margin
            end
            """);

        CliResult result = RunCli("apply", input, patch, "--output", output);

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("operation index=1 name=replace-text target=M.P0001 success=True", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliDumpJsonIncludesStructuredRuns()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = Path.Combine(temp.Path, "changes.docx");
        CreateDocxWithTrackedChanges(input);

        CliResult result = RunCli("dump", input, "--id", "M.P0001", "--runs", "--json");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("\"runs\"", result.Output, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("\"markupType\": \"inserted-run\"", result.Output, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("\"revisionId\": \"1\"", result.Output, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Deleted", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliReadSupportsOriginalTrackedChangeView()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = Path.Combine(temp.Path, "changes.docx");
        CreateDocxWithTrackedChanges(input);

        CliResult result = RunCli("read", input, "--view", "original");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("Deleted", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("Inserted", result.Output, StringComparison.Ordinal);
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
              <Default Extension="png" ContentType="image/png"/>
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
              <Relationship Id="rImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
            </Relationships>
            """);
        AddEntry(archive, "word/document.xml", """
            <w:document
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <w:body>
                <w:p>
                  <w:r><w:t>Revenue increased</w:t></w:r>
                  <w:r>
                    <w:drawing>
                      <wp:inline>
                        <a:graphic>
                          <a:graphicData>
                            <pic:pic>
                              <pic:blipFill>
                                <a:blip r:embed="rImage"/>
                              </pic:blipFill>
                            </pic:pic>
                          </a:graphicData>
                        </a:graphic>
                      </wp:inline>
                    </w:drawing>
                  </w:r>
                </w:p>
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
        AddEntry(archive, "word/media/image1.png", "fake-png");
    }

    private static void CreateDocxWithTrackedChanges(string path)
    {
        using FileStream file = File.Create(path);
        using var archive = new ZipArchive(file, ZipArchiveMode.Create);
        AddEntry(archive, "[Content_Types].xml", """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>
            """);
        AddEntry(archive, "_rels/.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """);
        AddEntry(archive, "word/_rels/document.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
            """);
        AddEntry(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:ins w:id="1" w:author="Alice" w:date="2026-06-01T12:00:00Z">
                    <w:r><w:t>Inserted</w:t></w:r>
                  </w:ins>
                  <w:del w:id="2" w:author="Bob" w:date="2026-06-02T12:00:00Z">
                    <w:r><w:delText>Deleted</w:delText></w:r>
                  </w:del>
                </w:p>
              </w:body>
            </w:document>
            """);
    }

    private static void CreateTextOnlyDocx(string path)
    {
        using FileStream file = File.Create(path);
        using var archive = new ZipArchive(file, ZipArchiveMode.Create);
        AddEntry(archive, "[Content_Types].xml", """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>
            """);
        AddEntry(archive, "_rels/.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """);
        AddEntry(archive, "word/_rels/document.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
            """);
        AddEntry(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p><w:r><w:t>Revenue increased</w:t></w:r></w:p>
              </w:body>
            </w:document>
            """);
    }

    private static void AddEntry(ZipArchive archive, string name, string text)
    {
        ZipArchiveEntry entry = archive.CreateEntry(name);
        using Stream stream = entry.Open();
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static string ReadEntry(string docxPath, string entryName)
    {
        using FileStream file = File.OpenRead(docxPath);
        using var archive = new ZipArchive(file, ZipArchiveMode.Read);
        ZipArchiveEntry entry = archive.GetEntry(entryName)
            ?? throw new InvalidDataException($"Missing {entryName}.");
        using Stream stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
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
