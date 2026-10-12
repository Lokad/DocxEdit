using System.Text;
using System.Text.Json;
using static Lokad.DocxEdit.Tests.DocxTestFixtures;
using static Lokad.DocxEdit.Tests.HostedCommandTests;

namespace Lokad.DocxEdit.Tests;

public static class HostedAssetTests
{
    [Theory]
    [InlineData("png")]
    [InlineData("jpeg")]
    [InlineData("svg")]
    [InlineData("webp")]
    public static async Task ReusedSnapshotPreservesEveryFormatAndIndependentExports(string extension)
    {
        byte[] bytes = extension switch
        {
            "png" => CreatePngBytes(4, 2), "jpeg" => CreateJpegBytes(4, 2),
            "svg" => SvgImageTests.Diagram(), _ => WebpImageTests.Fixture()
        };
        var host = CreateHost();
        string asset = "/asset." + extension;
        host.Files[asset] = bytes;
        host.Files["/edit.patch"] = Encoding.UTF8.GetBytes($"docxpatch 1\nop insert-image-after\ntarget M.P0001\nas first\nasset {asset}\nend\nop insert-image-after\ntarget M.P0001\nasset {asset}\nend\nop replace-image\ntarget @first\nasset {asset}\nend\n");
        Assert.Equal(0, await DocxCommand.RunAsync(["apply", "/input.docx", "/edit.patch", "-o", "/out.docx"], host, CancellationToken.None));
        Assert.Equal(asset, Assert.Single(host.AssetRequests));
        using var document = new MemoryStream(host.Files["/out.docx"]);
        var media = new DocxEditor().ExtractMedia(document);
        Assert.True(media.Success);
        Assert.Equal(2, media.Files.Count);
        Assert.All(media.Files, file => Assert.Equal(bytes, file.Content));
        media.Files[0].Content[0] ^= 0xFF;
        Assert.Equal(bytes, media.Files[1].Content);
        Assert.Equal(bytes, host.Files[asset]);
    }

    [Fact]
    public static void GeneralProvidersMayChangeBytesBetweenOperations()
    {
        using MemoryStream input = CreateDocxWithBody("<w:p><w:r><w:t>Start</w:t></w:r></w:p>");
        using var output = new MemoryStream();
        var provider = new AdvancingProvider();
        var result = new DocxEditor().Apply(input, new StringReader("docxpatch 1\nop insert-image-after\ntarget M.P0001\nas image\nasset changing\nend\nop replace-image\ntarget @image\nasset changing\nend\n"), output, new DocxEditOptions { AssetProvider = provider });
        Assert.True(result.Success);
        Assert.Equal(2, provider.Opens);
        output.Position = 0;
        DocxMediaFile file = Assert.Single(new DocxEditor().ExtractMedia(output).Files);
        Assert.Equal("image/png", file.ContentType);
        Assert.Equal(CreatePngBytes(4, 2), file.Content);
    }

    private sealed class AdvancingProvider : IDocxAssetProvider
    {
        public int Opens { get; private set; }
        public bool TryOpen(string reference, out Stream stream, out string? contentTypeHint, out string? fileNameHint)
        {
            Opens++;
            stream = new MemoryStream(Opens == 1 ? SvgImageTests.Diagram() : CreatePngBytes(4, 2), writable: false);
            contentTypeHint = fileNameHint = null;
            return true;
        }
    }

    private const string AssetName = "/assets/chart image.png";
    private const string ImagePatch = "docxpatch 1\nop insert-image-after\ntarget M.P0001\nas picture\nasset \"/assets/chart image.png\"\nend\nop replace-image\ntarget @picture\nasset <<<\n/assets/chart image.png\n>>>\nend\n";

    [Fact]
    public static async Task DelayedAssetsResolveDecodedReferencesOnceAndExportThroughHost()
    {
        var host = ImageHost();
        int exit = await DocxCommand.RunAsync(["apply", "/input.docx", "/edit.patch", "-o", "/out.docx"], host, CancellationToken.None);
        Assert.True(exit == 0, host.Error.ToString());
        Assert.Equal(AssetName, Assert.Single(host.AssetRequests));
        Assert.All(host.Inputs, input => Assert.True(input.Disposed));
        using var output = new MemoryStream(host.Files["/out.docx"]);
        var image = Assert.Single(new DocxEditor().ExtractMedia(output).Files);
        Assert.Equal(host.Files[AssetName], image.Content);
        Assert.Equal("image/png", image.ContentType);

        exit = await DocxCommand.RunAsync(["media", "/out.docx", "--extract", "/exports", "--id", "M.I0001", "--json"], host, CancellationToken.None);
        Assert.True(exit == 0, host.Error.ToString());
        Assert.Equal(image.Content, host.Files["/exports/" + image.FileName]);
    }

    [Theory]
    [InlineData("check")]
    [InlineData("apply")]
    public static async Task MissingAssetsKeepExistingDiagnosticAndPublishNothing(string command)
    {
        var host = ImageHost();
        host.Files.Remove(AssetName);
        int exit = await DocxCommand.RunAsync(Arguments(command), host, CancellationToken.None);
        Assert.Equal(1, exit);
        Assert.Contains("E5202", host.Output.ToString());
        Assert.Empty(host.Published);
        Assert.Single(host.AssetRequests);
    }

    [Theory]
    [InlineData("check")]
    [InlineData("apply")]
    public static async Task OversizeAssetIsBoundedAndRetainsEngineDiagnostic(string command)
    {
        var host = ImageHost();
        host.Files[AssetName] = new byte[9000];
        var options = new DocxCommandOptions { Quotas = new DocxPackageLimits(100, 32000, 8000) };
        int exit = await DocxCommand.RunAsync(Arguments(command), host, options, CancellationToken.None);
        Assert.Equal(1, exit);
        Assert.Contains("E5207", host.Output.ToString());
        Assert.Empty(host.Published);
        Assert.All(host.Inputs, input => Assert.True(input.Disposed));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public static async Task CancellationDuringAssetOpenOrReadPropagates(bool duringRead)
    {
        var host = ImageHost();
        using var canceled = new CancellationTokenSource();
        if (duringRead) host.OnAssetRead = canceled.Cancel;
        else host.OnAssetOpen = canceled.Cancel;
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => DocxCommand.RunAsync(Arguments("apply"), host, canceled.Token));
        Assert.Empty(host.Published);
        Assert.Empty(host.Output.ToString());
        Assert.All(host.Inputs, input => Assert.True(input.Disposed));
    }

    [Fact]
    public static async Task UnreachedAssetFailureDoesNotReplaceEarlierGuardFailure()
    {
        var host = ImageHost();
        host.AssetReadError = new IOException("synthetic asset read failure");
        string guarded = "docxpatch 1\nop replace-text\ntarget M.P0001\nexpect-text stale\nfind Alpha\nwith Beta\nend\n" + ImagePatch["docxpatch 1\n".Length..];
        host.Files["/edit.patch"] = Encoding.UTF8.GetBytes(guarded);
        int exit = await DocxCommand.RunAsync(Arguments("apply"), host, CancellationToken.None);
        Assert.Equal(1, exit);
        Assert.Contains("E3201", host.Output.ToString());
        Assert.DoesNotContain("synthetic asset read failure", host.Output.ToString());
        Assert.All(host.Inputs, input => Assert.True(input.Disposed));
    }

    [Fact]
    public static async Task ParseFailureDoesNotOpenAssets()
    {
        var host = ImageHost();
        host.Files["/edit.patch"] = Encoding.UTF8.GetBytes("not a patch\n" + ImagePatch);
        Assert.Equal(1, await DocxCommand.RunAsync(Arguments("apply"), host, CancellationToken.None));
        Assert.Empty(host.AssetRequests);
        Assert.Empty(host.Published);
    }

    [Fact]
    public static async Task CheckAndApplyUseSameAssetAndOperationSemantics()
    {
        var check = ImageHost();
        var apply = ImageHost();
        Assert.Equal(0, await DocxCommand.RunAsync(Arguments("check"), check, CancellationToken.None));
        Assert.Equal(0, await DocxCommand.RunAsync(Arguments("apply"), apply, CancellationToken.None));
        using var checkResult = JsonDocument.Parse(check.Output.ToString());
        using var applyResult = JsonDocument.Parse(apply.Output.ToString());
        Assert.Equal(checkResult.RootElement.GetProperty("Operations").GetRawText(), applyResult.RootElement.GetProperty("Operations").GetRawText());
        Assert.Empty(check.Published);
        Assert.Equal("/out.docx", Assert.Single(apply.Published));
    }

    private static MemoryHost ImageHost()
    {
        var host = CreateHost();
        host.Files["/edit.patch"] = Encoding.UTF8.GetBytes(ImagePatch);
        host.Files[AssetName] = CreatePngBytes(4, 3);
        return host;
    }

    private static string[] Arguments(string command) => command == "apply"
        ? [command, "/input.docx", "/edit.patch", "--json", "-o", "/out.docx"]
        : [command, "/input.docx", "/edit.patch", "--json"];
}
