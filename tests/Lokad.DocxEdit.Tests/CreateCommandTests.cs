using System.Text;
using System.Text.Json;
using static Lokad.DocxEdit.Tests.HostedCommandTests;

namespace Lokad.DocxEdit.Tests;

public static class CreateCommandTests
{
    [Theory]
    [InlineData("a4", "portrait", 11906, 16838)]
    [InlineData("a4", "landscape", 16838, 11906)]
    [InlineData("letter", "portrait", 12240, 15840)]
    [InlineData("letter", "landscape", 15840, 12240)]
    public static async Task HostedCreatePublishesDocumentAndCanonicalReports(string paper, string orientation, int width, int height)
    {
        var host = new MemoryHost();
        host.Files["/out.docx"] = "old"u8.ToArray();
        int exit = await DocxCommand.RunAsync(
            ["create", "-o", "/out.docx", "--paper", paper, "--orientation", orientation, "--report", "/report.json", "--diagnostics", "/diagnostics.json", "--json", "--compact"],
            host, CancellationToken.None);
        Assert.Equal(0, exit);
        Assert.Empty(host.Inputs);
        Assert.Empty(host.Error.ToString());
        using var document = new MemoryStream(host.Files["/out.docx"]);
        DocxSectionInfo section = Assert.Single(new DocxEditor().Read(document).Sections);
        Assert.Equal(width, section.PageWidthTwips);
        Assert.Equal(height, section.PageHeightTwips);
        using JsonDocument report = JsonDocument.Parse(host.Files["/report.json"]);
        Assert.Equal(paper, report.RootElement.GetProperty("PaperSize").GetString());
        Assert.Equal(orientation, report.RootElement.GetProperty("Orientation").GetString());
        Assert.True(report.RootElement.GetProperty("Success").GetBoolean());
        Assert.Equal(report.RootElement.GetRawText(), host.Output.ToString().Trim());
        Assert.Equal("[]", Encoding.UTF8.GetString(host.Files["/diagnostics.json"]));
    }

    [Fact]
    public static async Task BinaryStdoutHasOnlyDocumentBytesAndStatusGoesToStderr()
    {
        var host = new MemoryHost();
        int exit = await DocxCommand.RunAsync(["create", "--output", "-"], host, CancellationToken.None);
        Assert.Equal(0, exit);
        Assert.Empty(host.Output.ToString());
        Assert.Contains("docxedit create: OK (paper=a4 orientation=portrait)", host.Error.ToString());
        host.BinaryOutput.Position = 0;
        Assert.True(new DocxEditor().Validate(host.BinaryOutput).Success);
    }

    [Theory]
    [InlineData("create")]
    [InlineData("create", "unexpected", "-o", "/out.docx")]
    [InlineData("create", "--output", "")]
    [InlineData("create", "-o", "/out.docx", "--paper", "legal")]
    [InlineData("create", "-o", "/out.docx", "--orientation", "sideways")]
    [InlineData("create", "-o", "/out.docx", "--paper")]
    [InlineData("create", "-o", "/out.docx", "--orientation")]
    [InlineData("create", "-o", "/out.docx", "--track-changes", "off")]
    [InlineData("read", "/in.docx", "--paper", "a4")]
    [InlineData("create", "-o", "-", "--json")]
    [InlineData("create", "-o", "/out.docx", "--report", "-")]
    [InlineData("create", "-o", "/out.docx", "--diagnostics", "-")]
    [InlineData("create", "-o", "/out.docx", "--report", "/./out.docx")]
    [InlineData("create", "-o", "/out.docx", "--diagnostics", "/./out.docx")]
    [InlineData("create", "-o", "/out.docx", "--report", "/report.json", "--diagnostics", "/./report.json")]
    public static async Task InvalidUsageHasNoIo(params string[] args)
    {
        var host = new MemoryHost();
        Assert.Equal(2, await DocxCommand.RunAsync(args, host, CancellationToken.None));
        Assert.NotEmpty(host.Error.ToString());
        Assert.Empty(host.Inputs);
        Assert.Empty(host.Published);
        Assert.Empty(host.BinaryOutput.ToArray());
    }

    [Theory]
    [InlineData("/out.docx")]
    [InlineData("-")]
    public static async Task QuotaFailurePreservesDestinationAndStdout(string output)
    {
        var host = new MemoryHost();
        host.Files["/out.docx"] = "old"u8.ToArray();
        int exit = await DocxCommand.RunAsync(["create", "-o", output], host,
            new DocxCommandOptions { Quotas = new(1, 100000, 100000) }, CancellationToken.None);
        Assert.Equal(1, exit);
        Assert.Equal("old"u8.ToArray(), host.Files["/out.docx"]);
        Assert.Empty(host.Published);
        Assert.Empty(host.BinaryOutput.ToArray());
        Assert.Contains("E0001", host.Error.ToString());
    }

    [Fact]
    public static async Task CanceledOrFailedPublicationPreservesDestination()
    {
        var host = new MemoryHost();
        host.Files["/out.docx"] = "old"u8.ToArray();
        using var canceled = new CancellationTokenSource();
        host.OnPublish = canceled.Cancel;
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => DocxCommand.RunAsync(["create", "-o", "/out.docx"], host, canceled.Token));
        Assert.Equal("old"u8.ToArray(), host.Files["/out.docx"]);
        Assert.Empty(host.Published);
        host.OnPublish = () => throw new IOException("Simulated publication failure.");
        Assert.Equal(4, await DocxCommand.RunAsync(["create", "-o", "/out.docx"], host, CancellationToken.None));
        Assert.Equal("old"u8.ToArray(), host.Files["/out.docx"]);
        Assert.Empty(host.Published);
    }

    [Fact]
    public static void CreationIsDiscoverableThroughSharedHelp()
    {
        Assert.True(DocxHelp.TryGetCommand("create", out DocxCommandInfo command));
        Assert.Equal(0, command.MaxPositionals);
        string help = DocxHelp.RenderTopic("create");
        Assert.Contains("--paper a4|letter", help);
        Assert.Contains("--orientation portrait|landscape", help);
        Assert.Contains("docxedit create", DocxHelp.RenderOverview());
        Assert.Contains("PageWidthTwips", DocxHelp.RenderTopic("read"));
    }
}
