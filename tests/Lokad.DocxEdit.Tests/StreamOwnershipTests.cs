using System.IO.Compression;
using System.Text;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// PLAN B09: stream ownership follows one matrix on every exit.
// Input is disposed iff LeaveInputOpen is false (parse/load/edit/save/cancel
// failures and success); output iff LeaveOutputOpen is false; the patch reader
// stays caller-owned; provider-opened asset streams are engine-owned.
public static class StreamOwnershipTests
{
    [Fact]
    public static void SeekableLoadFailureDisposesInputWhenNotLeftOpen()
    {
        var input = new TrackedStream("not a zip"u8.ToArray(), writable: false);
        var options = new DocxReadOptions { LeaveInputOpen = false };
        DocxReadResult result = new DocxEditor().Read(input, options);
        Assert.False(result.Success);
        Assert.True(input.Disposed);
    }

    [Fact]
    public static void NonSeekableOverQuotaDisposesInputWhenNotLeftOpen()
    {
        var inner = new MemoryStream(new byte[1024], writable: false);
        var input = new NonSeekableTrackedStream(inner);
        var options = new DocxReadOptions { LeaveInputOpen = false, Quotas = DocxPackageLimits.Default with { MaxUncompressedBytes = 10 } };
        DocxReadResult result = new DocxEditor().Read(input, options);
        Assert.False(result.Success);
        Assert.True(input.Disposed);
    }

    [Fact]
    public static void ParseFailureDisposesInputAndOutputWhenNotLeftOpen()
    {
        using MemoryStream source = CreateDocx("Alpha.");
        var input = new TrackedStream(source.ToArray(), writable: false);
        var output = new TrackedStream(null, writable: true);
        var options = new DocxEditOptions { LeaveInputOpen = false, LeaveOutputOpen = false };
        using var patch = new StringReader("this is not a patch\n");
        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output, options);
        Assert.False(apply.Success);
        Assert.True(input.Disposed);
        Assert.True(output.Disposed);
    }

    [Fact]
    public static void EditFailureDisposesInputAndOutputWhenNotLeftOpen()
    {
        using MemoryStream source = CreateDocx("Alpha.");
        var input = new TrackedStream(source.ToArray(), writable: false);
        var output = new TrackedStream(null, writable: true);
        var options = new DocxEditOptions { LeaveInputOpen = false, LeaveOutputOpen = false };
        using var patch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nexpect-text Wrong guarded text\nfind Alpha\nwith Beta\nend\n");
        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output, options);
        Assert.False(apply.Success);
        Assert.True(input.Disposed);
        Assert.True(output.Disposed);
    }

    [Fact]
    public static void CancellationDisposesOwnedStreams()
    {
        using MemoryStream source = CreateDocx("Alpha.");
        var input = new TrackedStream(source.ToArray(), writable: false);
        var output = new TrackedStream(null, writable: true);
        var options = new DocxEditOptions { LeaveInputOpen = false, LeaveOutputOpen = false };
        using var patch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n");
        using var canceled = new CancellationTokenSource();
        canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() => new DocxEditor().Apply(input, patch, output, options, canceled.Token));
        Assert.True(input.Disposed);
        Assert.True(output.Disposed);
    }

    [Fact]
    public static void LeftOpenStreamsStayOpenOnFailure()
    {
        using MemoryStream source = CreateDocx("Alpha.");
        var input = new TrackedStream(source.ToArray(), writable: false);
        var output = new TrackedStream(null, writable: true);
        var options = new DocxEditOptions { LeaveInputOpen = true, LeaveOutputOpen = true };
        using var patch = new StringReader("this is not a patch\n");
        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output, options);
        Assert.False(apply.Success);
        Assert.False(input.Disposed);
        Assert.False(output.Disposed);
    }

    [Fact]
    public static void SuccessfulApplyDisposesOutputWhenNotLeftOpen()
    {
        using MemoryStream source = CreateDocx("Alpha.");
        var input = new TrackedStream(source.ToArray(), writable: false);
        var output = new TrackedStream(null, writable: true);
        var options = new DocxEditOptions { LeaveInputOpen = true, LeaveOutputOpen = false };
        using var patch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n");
        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output, options);
        Assert.True(apply.Success);
        Assert.False(input.Disposed);
        Assert.True(output.Disposed);
    }

    [Fact]
    public static void FailedAssetStreamIsDisposedByEngine()
    {
        using MemoryStream input = CreateMinimalDocx();
        var provider = new OversizedAssetProvider();
        var options = new DocxEditOptions { Quotas = DocxPackageLimits.Default with { MaxSinglePartBytes = 300 }, AssetProvider = provider };
        using var patch = new StringReader("docxpatch 1\n\nop insert-image-after\ntarget M.P0001\nasset photo\nalt Caption\nend\n");
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output, options);
        Assert.False(apply.Success);
        Assert.Contains(apply.Diagnostics, static d => d.Code == "E5207");
        Assert.True(provider.OpenedStream is not null && provider.OpenedStream.Disposed);
    }

    private static MemoryStream CreateDocx(string paragraphText)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", RootRelsXml());
            AddEntry(archive, "word/_rels/document.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\" />");
            AddEntry(archive, "word/document.xml", "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>"
                + "<w:p><w:r><w:t>" + paragraphText + "</w:t></w:r></w:p>"
                + "</w:body></w:document>");
        }
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateMinimalDocx()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
            AddEntry(archive, "_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rD\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
            AddEntry(archive, "word/_rels/document.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>");
            AddEntry(archive, "word/document.xml", "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>Alpha.</w:t></w:r></w:p></w:body></w:document>");
        }
        stream.Position = 0;
        return stream;
    }

    private static string ContentTypesXml()
    {
        return "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
            + "</Types>";
    }

    private static string RootRelsXml()
    {
        return "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rDocument\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
            + "</Relationships>";
    }


    private sealed class TrackedStream : MemoryStream
    {
        public bool Disposed { get; private set; }
        public TrackedStream(byte[]? bytes, bool writable)
            : base(GetBuffer(bytes, writable), writable)
        {
        }

        private static byte[] GetBuffer(byte[]? bytes, bool writable)
        {
            if (bytes is not null)
            {
                return bytes;
            }

            // Expandable empty buffer for writable outputs; fixed callers pass
            // their bytes with matching writability instead.
            return new byte[65536];
        }
        protected override void Dispose(bool disposing)
        {
            Disposed = true;
            base.Dispose(disposing);
        }
    }

    private sealed class NonSeekableTrackedStream : MemoryStream
    {
        public bool Disposed { get; private set; }
        public NonSeekableTrackedStream(MemoryStream inner)
            : base(inner.ToArray(), writable: false)
        {
        }
        public override bool CanSeek => false;
        protected override void Dispose(bool disposing)
        {
            Disposed = true;
            base.Dispose(disposing);
        }
    }

    private sealed class OversizedAssetProvider : IDocxAssetProvider
    {
        public TrackedStream? OpenedStream { get; private set; }
        public bool TryOpen(string requestedReference, out Stream stream, out string? contentTypeHint, out string? fileNameHint)
        {
            OpenedStream = new TrackedStream(new byte[1024], writable: false);
            stream = OpenedStream;
            contentTypeHint = null;
            fileNameHint = "photo.png";
            return true;
        }
    }
}
