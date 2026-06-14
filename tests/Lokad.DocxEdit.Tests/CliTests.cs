using System.Diagnostics;
using System.IO.Compression;
using System.Text;

namespace Lokad.DocxEdit.Tests;

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
    public static void CliValidatePrintsStatus()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = Path.Combine(temp.Path, "input.docx");
        CreateDocx(input);

        CliResult text = RunCli("validate", input);
        CliResult package = RunCli("validate", input, "--profile", "package");
        CliResult json = RunCli("validate", input, "--json");

        Assert.Equal(0, text.ExitCode);
        Assert.Contains("docxedit validate: OK", text.Output, StringComparison.Ordinal);
        Assert.Contains("profile=structural", text.Output, StringComparison.Ordinal);
        Assert.Equal(0, package.ExitCode);
        Assert.Contains("profile=package", package.Output, StringComparison.Ordinal);
        Assert.Equal(0, json.ExitCode);
        Assert.Contains("\"Success\": true", json.Output, StringComparison.Ordinal);
        Assert.Contains("\"MainDocumentPartName\": \"/word/document.xml\"", json.Output, StringComparison.Ordinal);
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
        Assert.Equal(DocxHelp.RenderOverview(), help.Output);
        Assert.Contains("docxedit read report.docx", help.Output, StringComparison.Ordinal);
        Assert.Contains("docxedit context report.docx --id M.P0004", help.Output, StringComparison.Ordinal);
        Assert.Contains("docxedit validate report.docx", help.Output, StringComparison.Ordinal);
        Assert.Contains("docxedit apply report.docx edits.docxpatch --output report.edited.docx", help.Output, StringComparison.Ordinal);
        Assert.Contains("tracked-change and comment markup", help.Output, StringComparison.Ordinal);
        Assert.Equal(0, patchHelp.ExitCode);
        Assert.Equal(DocxHelp.RenderTopic("patch"), patchHelp.Output);
        Assert.Contains("find <<<", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("op set-cell", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("set-table-style", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("set-table-metadata", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("set-row-header", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("append-column", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("add-comment-reply", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("add-repeating-section-item", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("add-bookmark", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("refresh-field-result", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("E4314", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("E4315", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("E4316", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("insert-after", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("expect-row-count", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("set-section-orientation", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("op replace-image", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("Unsupported fields are rejected", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("Track-change support:", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("replace-text", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("operation | support class | support value | behavior", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("replace-text | text-run | tracked-simple", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("tracked-simple", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("set-cell", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("preserve-only", patchHelp.Output, StringComparison.Ordinal);
        Assert.Contains("unsupported", patchHelp.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliHasCommandSpecificHelpForAgentWorkflows()
    {
        CliResult dump = RunCli("help", "dump");
        CliResult context = RunCli("help", "context");
        CliResult changes = RunCli("help", "changes");
        CliResult validate = RunCli("help", "validate");
        CliResult check = RunCli("help", "check");
        CliResult apply = RunCli("help", "apply");

        Assert.Equal(0, dump.ExitCode);
        Assert.Contains("docxedit dump input.docx --id TARGET", dump.Output, StringComparison.Ordinal);
        Assert.Contains("markup=inserted-run", dump.Output, StringComparison.Ordinal);
        Assert.Contains("Runs array", dump.Output, StringComparison.Ordinal);
        Assert.Contains("not the same namespace", dump.Output, StringComparison.Ordinal);
        Assert.Contains("inspect Runs[]", dump.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("PowerShell", dump.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("ConvertFrom-Json", dump.Output, StringComparison.Ordinal);
        Assert.Equal(0, context.ExitCode);
        Assert.Contains("docxedit context input.docx --id TARGET", context.Output, StringComparison.Ordinal);
        Assert.Contains("--max-text is 0", context.Output, StringComparison.Ordinal);
        Assert.Contains("--radius N", context.Output, StringComparison.Ordinal);
        Assert.Equal(0, changes.ExitCode);
        Assert.Contains("GroupSummary", changes.Output, StringComparison.Ordinal);
        Assert.Contains("TargetSummary", changes.Output, StringComparison.Ordinal);
        Assert.Contains("CommentSummary", changes.Output, StringComparison.Ordinal);
        Assert.Contains("comment-anchor-target", changes.Output, StringComparison.Ordinal);
        Assert.Contains("target-source", changes.Output, StringComparison.Ordinal);
        Assert.Contains("paired-change-id", changes.Output, StringComparison.Ordinal);
        Assert.Contains("--operation-report path", changes.Output, StringComparison.Ordinal);
        Assert.Contains("Some JSON consumers", changes.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("PowerShell", changes.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("ConvertFrom-Json", changes.Output, StringComparison.Ordinal);
        Assert.Equal(0, validate.ExitCode);
        Assert.Contains("docxedit validate input.docx", validate.Output, StringComparison.Ordinal);
        Assert.Contains("--profile structural|package", validate.Output, StringComparison.Ordinal);
        Assert.Contains("--max-diagnostics N", validate.Output, StringComparison.Ordinal);
        Assert.Contains("WordprocessingML invariants", validate.Output, StringComparison.Ordinal);
        Assert.Contains("stable diagnostic codes", validate.Output, StringComparison.Ordinal);
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
    public static void CliContextPrintsStructureWithoutTextByDefault()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = Path.Combine(temp.Path, "input.docx");
        CreateDocx(input);

        CliResult result = RunCli("context", input, "--id", "M.P0001");

        Assert.Equal(0, result.ExitCode);
        Assert.Contains("target M.P0001 paragraph story=\"main\" text=\"\"", result.Output, StringComparison.Ordinal);
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
    public static void CliRejectsInvalidValidationProfile()
    {
        CliResult result = RunCli("validate", "input.docx", "--profile", "strict");

        Assert.Equal(2, result.ExitCode);
        Assert.Contains("Invalid value for --profile", result.Error, StringComparison.Ordinal);
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
    public static void CliApplyCanReplaceBookmarkText()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = Path.Combine(temp.Path, "input.docx");
        string output = Path.Combine(temp.Path, "output.docx");
        string patch = Path.Combine(temp.Path, "edit.docxpatch");
        CreateDocxWithBookmark(input);
        File.WriteAllText(patch, """
            docxpatch 1

            op replace-bookmark-text
            target M.B0001
            text Customer
            end
            """);

        CliResult check = RunCli("check", input, patch);
        CliResult apply = RunCli("apply", input, patch, "--output", output);

        Assert.Equal(0, check.ExitCode);
        Assert.Equal(0, apply.ExitCode);
        Assert.True(File.Exists(output));
        string xml = ReadEntry(output, "word/document.xml");
        Assert.Contains("w:bookmarkStart w:id=\"1\" w:name=\"ClientName\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:bookmarkEnd w:id=\"1\"", xml, StringComparison.Ordinal);
        Assert.Contains("Customer", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:t>Client</w:t>", xml, StringComparison.Ordinal);
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
        Assert.Contains("target-source=ancestor", result.Output, StringComparison.Ordinal);
        Assert.Contains("M.CH0001 inserted-run", result.Output, StringComparison.Ordinal);
        Assert.Contains("text-length=8", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("Inserted", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("Deleted", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliChangesIncludesCommentTextOnlyWhenRequested()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = Path.Combine(temp.Path, "comments.docx");
        CreateDocxWithComments(input);

        CliResult safe = RunCli("changes", input);
        CliResult withText = RunCli("changes", input, "--include-comment-text", "--max-comment-text", "7");

        Assert.Equal(0, safe.ExitCode);
        Assert.Contains("comment-summary comment-id=3", safe.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("Private comment text", safe.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("comment-text=\"", safe.Output, StringComparison.Ordinal);

        Assert.Equal(0, withText.ExitCode);
        Assert.Contains("comment-text-length=20", withText.Output, StringComparison.Ordinal);
        Assert.Contains("comment-text=\"Private\"", withText.Output, StringComparison.Ordinal);
        Assert.Contains("comment-text-truncated=true", withText.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("Private comment text", withText.Output, StringComparison.Ordinal);
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
    public static void CliChangesAnnotatesGeneratedRevisionsFromOperationReport()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = Path.Combine(temp.Path, "input.docx");
        string output = Path.Combine(temp.Path, "output.docx");
        string patch = Path.Combine(temp.Path, "edit.docxpatch");
        string report = Path.Combine(temp.Path, "apply-report.json");
        CreateTextOnlyDocx(input);
        File.WriteAllText(patch, """
            docxpatch 1

            op replace-text
            target M.P0001
            find increased
            with rose
            end
            """);

        CliResult apply = RunCli(
            "apply",
            input,
            patch,
            "--output",
            output,
            "--track-changes",
            "require",
            "--author",
            "Reviewer",
            "--timestamp-utc",
            "2026-01-01T00:00:00Z",
            "--report",
            report);
        CliResult changes = RunCli("changes", output, "--operation-report", report);

        Assert.Equal(0, apply.ExitCode);
        Assert.Equal(0, changes.ExitCode);
        Assert.Contains("operation-index=1", changes.Output, StringComparison.Ordinal);
        Assert.Contains("operation-name=replace-text", changes.Output, StringComparison.Ordinal);
        Assert.Contains("operation-target=M.P0001", changes.Output, StringComparison.Ordinal);
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
        startInfo.ArgumentList.Add(Path.Combine(repoRoot, "src", "Lokad.DocxEdit.Cli", "Lokad.DocxEdit.Cli.csproj"));
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
            if (File.Exists(Path.Combine(directory.FullName, "Lokad.DocxEdit.slnx")))
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

    private static void CreateDocxWithComments(string path)
    {
        using FileStream file = File.Create(path);
        using var archive = new ZipArchive(file, ZipArchiveMode.Create);
        AddEntry(archive, "[Content_Types].xml", """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
              <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
            </Types>
            """);
        AddEntry(archive, "_rels/.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """);
        AddEntry(archive, "word/_rels/document.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
            </Relationships>
            """);
        AddEntry(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:commentRangeStart w:id="3"/>
                  <w:r><w:t>Commented</w:t></w:r>
                  <w:commentRangeEnd w:id="3"/>
                  <w:r><w:commentReference w:id="3"/></w:r>
                </w:p>
              </w:body>
            </w:document>
            """);
        AddEntry(archive, "word/comments.xml", """
            <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:comment w:id="3" w:author="Reviewer" w:initials="RV" w:date="2026-06-07T12:00:00Z">
                <w:p><w:r><w:t>Private comment text</w:t></w:r></w:p>
              </w:comment>
            </w:comments>
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

    private static void CreateDocxWithBookmark(string path)
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
                  <w:r><w:t>Before </w:t></w:r>
                  <w:bookmarkStart w:id="1" w:name="ClientName"/>
                  <w:r><w:t>Client</w:t></w:r>
                  <w:bookmarkEnd w:id="1"/>
                  <w:r><w:t> After</w:t></w:r>
                </w:p>
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
