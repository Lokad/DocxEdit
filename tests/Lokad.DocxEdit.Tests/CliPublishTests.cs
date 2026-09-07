using System.IO.Compression;
using System.Text;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// PLAN B05: failed CLI operations must preserve existing files; outputs publish
// atomically (temp sibling + replace) only after successful editing; path
// collisions across document/patch/output/report/diagnostics fail upfront.
[Collection("ConsoleCli")]
public static class CliPublishTests
{
    [Fact]
    public static void FailedApplyPreservesExistingOutput()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = WriteDocx(temp.Path, "input.docx", "Alpha.");
        string patch = WriteText(temp.Path, "bad.docxpatch", "this is not a patch\n");
        string output = Path.Combine(temp.Path, "output.docx");
        File.WriteAllBytes(output, "SENTINEL-OLD-CONTENT"u8.ToArray());
        CliOutcome result = RunCli("apply", input, patch, "--output", output);
        Assert.Equal(1, result.ExitCode);
        Assert.Equal("SENTINEL-OLD-CONTENT", File.ReadAllText(output));
    }

    [Fact]
    public static void FailedGuardApplyPreservesExistingOutput()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = WriteDocx(temp.Path, "input.docx", "Alpha.");
        string patch = WriteText(temp.Path, "guard.docxpatch", "docxpatch 1\n\nop replace-text\ntarget M.P0001\nexpect-text Wrong guarded text\nfind Alpha\nwith Beta\nend\n");
        string output = Path.Combine(temp.Path, "output.docx");
        File.WriteAllBytes(output, "SENTINEL-OLD-CONTENT"u8.ToArray());
        CliOutcome result = RunCli("apply", input, patch, "--output", output);
        Assert.Equal(1, result.ExitCode);
        Assert.Equal("SENTINEL-OLD-CONTENT", File.ReadAllText(output));
    }

    [Fact]
    public static void SuccessfulApplyPublishesReadablePackage()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = WriteDocx(temp.Path, "input.docx", "Alpha.");
        string patch = WriteText(temp.Path, "good.docxpatch", "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n");
        string output = Path.Combine(temp.Path, "output.docx");
        CliOutcome result = RunCli("apply", input, patch, "--output", output);
        Assert.Equal(0, result.ExitCode);
        using var stream = File.OpenRead(output);
        DocxReadResult read = new DocxEditor().Read(stream);
        Assert.True(read.Success);
        Assert.Equal("Beta.", Assert.Single(read.Paragraphs).Text);
    }

    [Fact]
    public static void OutputInputCollisionIsRejected()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = WriteDocx(temp.Path, "input.docx", "Alpha.");
        string patch = WriteText(temp.Path, "good.docxpatch", "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n");
        byte[] before = File.ReadAllBytes(input);
        CliOutcome result = RunCli("apply", input, patch, "--output", input);
        Assert.Equal(2, result.ExitCode);
        Assert.Contains("--output", result.Error, StringComparison.Ordinal);
        Assert.Equal(before, File.ReadAllBytes(input));
    }

    [Fact]
    public static void OutputPatchCollisionIsRejected()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = WriteDocx(temp.Path, "input.docx", "Alpha.");
        string patch = WriteText(temp.Path, "good.docxpatch", "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n");
        string before = File.ReadAllText(patch);
        CliOutcome result = RunCli("apply", input, patch, "--output", patch);
        Assert.Equal(2, result.ExitCode);
        Assert.Contains("--output", result.Error, StringComparison.Ordinal);
        Assert.Equal(before, File.ReadAllText(patch));
    }

    [Fact]
    public static void ReportInputCollisionIsRejected()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = WriteDocx(temp.Path, "input.docx", "Alpha.");
        string patch = WriteText(temp.Path, "good.docxpatch", "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n");
        byte[] before = File.ReadAllBytes(input);
        CliOutcome result = RunCli("check", input, patch, "--report", input);
        Assert.Equal(2, result.ExitCode);
        Assert.Contains("--report", result.Error, StringComparison.Ordinal);
        Assert.Equal(before, File.ReadAllBytes(input));
    }

    [Fact]
    public static void DiagnosticsInputCollisionIsRejected()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = WriteDocx(temp.Path, "input.docx", "Alpha.");
        byte[] before = File.ReadAllBytes(input);
        CliOutcome result = RunCli("read", input, "--diagnostics", input);
        Assert.Equal(2, result.ExitCode);
        Assert.Contains("--diagnostics", result.Error, StringComparison.Ordinal);
        Assert.Equal(before, File.ReadAllBytes(input));
    }

    private static string WriteDocx(string directory, string fileName, string paragraphText)
    {
        string path = Path.Combine(directory, fileName);
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "</Types>");
            AddEntry(archive, "_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rDocument\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + "</Relationships>");
            AddEntry(archive, "word/_rels/document.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\" />");
            AddEntry(archive, "word/document.xml", "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>"
                + "<w:p><w:r><w:t>" + paragraphText + "</w:t></w:r></w:p>"
                + "</w:body></w:document>");
        }
        File.WriteAllBytes(path, stream.ToArray());
        return path;
    }

    private static string WriteText(string directory, string fileName, string text)
    {
        string path = Path.Combine(directory, fileName);
        File.WriteAllText(path, text);
        return path;
    }

    private static CliOutcome RunCli(params string[] args)
    {
        TextWriter savedOut = Console.Out;
        TextWriter savedError = Console.Error;
        var output = new StringWriter();
        var error = new StringWriter();
        try
        {
            Console.SetOut(output);
            Console.SetError(error);
            int exitCode = ProgramMain.Run(args);
            return new CliOutcome(exitCode, output.ToString(), error.ToString());
        }
        finally
        {
            Console.SetOut(savedOut);
            Console.SetError(savedError);
        }
    }



    private sealed record CliOutcome(int ExitCode, string Output, string Error);
}
