using System.Text;
using System.Text.Json;
using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

public static class HostedCommandTests
{
    private const string Patch = "docxpatch 1\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n";

    [Theory]
    [InlineData("read", "/input.docx")]
    [InlineData("outline", "/input.docx")]
    [InlineData("find", "/input.docx", "Alpha")]
    [InlineData("dump", "/input.docx", "--id", "M.P0001")]
    [InlineData("context", "/input.docx", "--id", "M.P0001")]
    [InlineData("capabilities", "/input.docx", "--id", "M.P0001")]
    [InlineData("template", "/input.docx", "--id", "M.P0001")]
    [InlineData("styles", "/input.docx")]
    [InlineData("media", "/input.docx")]
    [InlineData("validate", "/input.docx")]
    [InlineData("changes", "/input.docx")]
    [InlineData("lint", "/edit.patch")]
    [InlineData("check", "/input.docx", "/edit.patch")]
    [InlineData("catalog")]
    [InlineData("version")]
    [InlineData("help", "apply")]
    public static async Task CommandsRunWithoutFilesystemOrConsole(params string[] arguments)
    {
        var host = CreateHost();
        if (arguments[0] != "help") arguments = [.. arguments, "--json"];
        int exit = await DocxCommand.RunAsync(arguments, host, CancellationToken.None);
        Assert.True(exit == 0, host.Error.ToString());
        Assert.NotEmpty(host.Output.ToString());
        Assert.All(host.Inputs, stream => Assert.True(stream.Disposed));
    }

    [Fact]
    public static async Task ApplyPublishesReadableDocumentAndCanonicalJsonReport()
    {
        var host = CreateHost();
        int exit = await DocxCommand.RunAsync(["apply", "/input.docx", "/edit.patch", "-o", "/out.docx", "--report", "/report.json", "--json"], host, CancellationToken.None);
        Assert.Equal(0, exit);
        using var output = new MemoryStream(host.Files["/out.docx"]);
        Assert.Equal("Beta", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        using JsonDocument report = JsonDocument.Parse(host.Files["/report.json"]);
        Assert.Equal("off", report.RootElement.GetProperty("TrackChanges").GetString());
        Assert.True(report.RootElement.GetProperty("Success").GetBoolean());
        Assert.Equal(JsonDocument.Parse(host.Output.ToString()).RootElement.GetRawText(), report.RootElement.GetRawText());
        Assert.Equal("Beta", ReadText(host.Files["/out.docx"]));
        Assert.Equal("Alpha", ReadText(host.Files["/input.docx"]));
    }

    [Fact]
    public static async Task FailedApplyKeepsDestinationAndBinaryStdoutClean()
    {
        var host = CreateHost();
        host.Files["/edit.patch"] = Encoding.UTF8.GetBytes(Patch.Replace("find Alpha", "expect-text stale\nfind Alpha"));
        host.Files["/out.docx"] = "old destination"u8.ToArray();
        int exit = await DocxCommand.RunAsync(["apply", "/input.docx", "/edit.patch", "-o", "/out.docx"], host, CancellationToken.None);
        Assert.Equal(1, exit);
        Assert.Equal("old destination", Encoding.UTF8.GetString(host.Files["/out.docx"]));
        Assert.Empty(host.Published);
        exit = await DocxCommand.RunAsync(["apply", "/input.docx", "/edit.patch", "-o", "-"], host, CancellationToken.None);
        Assert.Equal(1, exit);
        Assert.Empty(host.BinaryOutput.ToArray());
        Assert.Contains("FAILED", host.Error.ToString());
    }

    [Fact]
    public static async Task BinaryStdoutAndTextStdinAreHostMediated()
    {
        var host = CreateHost();
        host.Files["-"] = host.Files["/edit.patch"];
        int exit = await DocxCommand.RunAsync(["apply", "/input.docx", "-", "-o", "-"], host, CancellationToken.None);
        Assert.Equal(0, exit);
        Assert.Equal("Beta", ReadText(host.BinaryOutput.ToArray()));
        Assert.Empty(host.Output.ToString());
        Assert.Contains("docxedit apply: OK", host.Error.ToString());
        host.Files["-"] = host.BinaryOutput.ToArray();
        Assert.Equal(0, await DocxCommand.RunAsync(["read", "-", "--summary"], host, CancellationToken.None));
        Assert.Contains("paragraphs count=1", host.Output.ToString());
    }

    [Fact]
    public static async Task PathCollisionsUseHostRulesBeforeAnyIo()
    {
        var host = CreateHost();
        int exit = await DocxCommand.RunAsync(["apply", "/input.docx", "/edit.patch", "-o", "/./input.docx"], host, CancellationToken.None);
        Assert.Equal(2, exit);
        Assert.Empty(host.Inputs);
        Assert.Empty(host.Published);
    }

    [Fact]
    public static async Task CancellationDuringAsyncReadPropagatesAndDisposesInput()
    {
        var host = CreateHost();
        using var canceled = new CancellationTokenSource();
        host.OnRead = canceled.Cancel;
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => DocxCommand.RunAsync(["read", "/input.docx"], host, canceled.Token));
        Assert.True(Assert.Single(host.Inputs).Disposed);
        Assert.Empty(host.Output.ToString());
        Assert.Empty(host.Published);
    }

    [Fact]
    public static async Task CancellationAtPublicationKeepsOldDestination()
    {
        var host = CreateHost();
        host.Files["/out.docx"] = "old"u8.ToArray();
        using var canceled = new CancellationTokenSource();
        host.OnPublish = canceled.Cancel;
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => DocxCommand.RunAsync(["apply", "/input.docx", "/edit.patch", "-o", "/out.docx"], host, canceled.Token));
        Assert.Equal("old", Encoding.UTF8.GetString(host.Files["/out.docx"]));
        Assert.Empty(host.Published);
    }

    [Fact]
    public static async Task ConcurrentHostsKeepOutputSeparate()
    {
        var first = CreateHost();
        var second = CreateHost();
        second.Files["/input.docx"] = MakeDocument("Other");
        int[] exits = await Task.WhenAll(DocxCommand.RunAsync(["read", "/input.docx"], first, CancellationToken.None), DocxCommand.RunAsync(["read", "/input.docx"], second, CancellationToken.None));
        Assert.All(exits, exit => Assert.Equal(0, exit));
        Assert.Contains("Alpha", first.Output.ToString());
        Assert.DoesNotContain("Other", first.Output.ToString());
        Assert.Contains("Other", second.Output.ToString());
        Assert.DoesNotContain("Alpha", second.Output.ToString());
    }

    [Fact]
    public static async Task HostPackageQuotaAppliesBeforeUnboundedInputBuffering()
    {
        var host = CreateHost();
        var options = new DocxCommandOptions { Quotas = new DocxPackageLimits(100, 20, 20) };
        int exit = await DocxCommand.RunAsync(["read", "/input.docx"], host, options, CancellationToken.None);
        Assert.Equal(4, exit);
        Assert.Contains("quota", host.Error.ToString());
        Assert.True(Assert.Single(host.Inputs).Disposed);
    }

    private static string ReadText(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        return Assert.Single(new DocxEditor().Read(stream).Paragraphs).Text;
    }

    internal static byte[] MakeDocument(string text)
    {
        using var stream = CreateDocxWithBody($"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>");
        return stream.ToArray();
    }

    internal static MemoryHost CreateHost()
    {
        var host = new MemoryHost();
        host.Files["/input.docx"] = MakeDocument("Alpha");
        host.Files["/edit.patch"] = Encoding.UTF8.GetBytes(Patch);
        return host;
    }

    internal sealed class MemoryHost : IDocxCommandHost, IDocxAssetProvider
    {
        public Dictionary<string, byte[]> Files { get; } = new(StringComparer.Ordinal);
        public List<string> Published { get; } = [];
        public List<AsyncInput> Inputs { get; } = [];
        public StringWriter Output { get; } = new();
        public StringWriter Error { get; } = new();
        public MemoryStream BinaryOutput { get; } = new();
        public Action? OnRead { get; set; }
        public Action? OnPublish { get; set; }
        public TextWriter StandardOutput => Output;
        public TextWriter StandardError => Error;
        public bool PathsEqual(string first, string second) => first.Replace("/./", "/") == second.Replace("/./", "/");
        public string CombinePath(string directory, string fileName) => directory.TrimEnd('/') + "/" + fileName;
        public IDocxAssetProvider GetAssetProvider(string patchPath) => this;

        public async ValueTask<Stream> OpenReadAsync(string path, CancellationToken cancellationToken)
        {
            await Task.Yield();
            cancellationToken.ThrowIfCancellationRequested();
            var input = new AsyncInput(Files[path], OnRead);
            Inputs.Add(input);
            return input;
        }

        public async ValueTask<TextReader> OpenTextAsync(string path, CancellationToken cancellationToken) =>
            new StreamReader(await OpenReadAsync(path, cancellationToken));

        public async ValueTask PublishFileAsync(string path, Stream contents, CancellationToken cancellationToken)
        {
            using var staged = new MemoryStream();
            await contents.CopyToAsync(staged, cancellationToken);
            OnPublish?.Invoke();
            cancellationToken.ThrowIfCancellationRequested();
            Files[path] = staged.ToArray();
            Published.Add(path);
        }

        public async ValueTask WriteStandardOutputAsync(Stream contents, CancellationToken cancellationToken) =>
            await contents.CopyToAsync(BinaryOutput, cancellationToken);

        public bool TryOpen(string reference, out Stream stream, out string? contentTypeHint, out string? fileNameHint)
        {
            contentTypeHint = null;
            fileNameHint = reference;
            if (Files.TryGetValue(reference, out byte[]? bytes))
            {
                stream = new MemoryStream(bytes, writable: false);
                return true;
            }
            stream = Stream.Null;
            return false;
        }
    }

    internal sealed class AsyncInput(byte[] bytes, Action? onRead) : MemoryStream(bytes, writable: false)
    {
        private readonly byte[] _bytes = bytes;
        public bool Disposed { get; private set; }
        public override int Read(byte[] buffer, int offset, int count) => throw new InvalidOperationException("Synchronous host read.");
        public override int Read(Span<byte> buffer) => throw new InvalidOperationException("Synchronous host read.");
        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            await Task.Yield();
            onRead?.Invoke();
            cancellationToken.ThrowIfCancellationRequested();
            int length = (int)Math.Min(buffer.Length, Length - Position);
            _bytes.AsMemory((int)Position, length).CopyTo(buffer);
            Position += length;
            return length;
        }
        protected override void Dispose(bool disposing)
        {
            Disposed = true;
            base.Dispose(disposing);
        }
    }
}
