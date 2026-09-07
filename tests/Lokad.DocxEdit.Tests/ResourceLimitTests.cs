using System.IO.Compression;
using System.Text;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// PLAN B08: quotas bind reads before allocation/parsing at ZIP/XML, patch,
// asset, and edited-package boundaries; cancellation aborts bounded reads.
public static class ResourceLimitTests
{
    [Fact]
    public static void OversizedContentTypesRejected()
    {
        using MemoryStream input = CreateDocxWithLargeContentTypes(600);
        var options = new DocxReadOptions { Quotas = DocxPackageLimits.Default with { MaxSinglePartBytes = 100 } };
        DocxReadResult result = new DocxEditor().Read(input, options);
        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static d => d.Code == "E0001");
    }

    [Fact]
    public static void OversizedPatchRejectedWithoutFullRead()
    {
        string patchText = "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith " + new string('B', 2000) + "\nend\n";
        var options = new DocxEditOptions { MaxPatchChars = 100 };
        using MemoryStream input = CreateDocx("Alpha.");
        DocxCheckResult check = new DocxEditor().Check(input, new StringReader(patchText), options);
        Assert.False(check.Success);
        Assert.Contains(check.Diagnostics, static d => d.Code == "E2014");
    }

    [Fact]
    public static void OversizedAssetRejectedWithoutFullCopy()
    {
        using MemoryStream input = CreateMinimalDocx();
        var provider = new SizedAssetProvider(8 * 1024);
        var options = new DocxEditOptions { Quotas = DocxPackageLimits.Default with { MaxSinglePartBytes = 500 }, AssetProvider = provider };
        using var patch = new StringReader("docxpatch 1\n\nop insert-image-after\ntarget M.P0001\nasset photo\nalt Caption\nend\n");
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output, options);
        Assert.False(apply.Success);
        Assert.Contains(apply.Diagnostics, static d => d.Code == "E5207");
    }

    [Fact]
    public static void EditedPackageGrowthRejected()
    {
        using MemoryStream checkInput = CreateMinimalDocx();
        var provider = new SizedAssetProvider(5 * 1024);
        var options = new DocxEditOptions { Quotas = DocxPackageLimits.Default with { MaxUncompressedBytes = 2000 }, AssetProvider = provider };
        using var checkPatch = new StringReader("docxpatch 1\n\nop insert-image-after\ntarget M.P0001\nasset photo\nalt Caption\nend\n");
        DocxCheckResult check = new DocxEditor().Check(checkInput, checkPatch, options);
        Assert.False(check.Success);
        Assert.Contains(check.Diagnostics, static d => d.Code == "E0001");
        using MemoryStream applyInput = CreateMinimalDocx();
        using var applyPatch = new StringReader("docxpatch 1\n\nop insert-image-after\ntarget M.P0001\nasset photo\nalt Caption\nend\n");
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, applyPatch, output, options);
        Assert.False(apply.Success);
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public static void CancelledPatchReadAborts()
    {
        using var cts = new CancellationTokenSource();
        using MemoryStream input = CreateDocx("Alpha.");
        var reader = new CancelAfterReadsReader(cts, 3);
        Assert.Throws<OperationCanceledException>(() => new DocxEditor().Check(input, reader, new DocxEditOptions(), cts.Token));
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

    private static MemoryStream CreateDocxWithLargeContentTypes(int totalBytes)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            var padding = new string('x', Math.Max(0, totalBytes - 300));
            AddEntry(archive, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\" print=\"" + padding + "\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
            AddEntry(archive, "_rels/.rels", RootRelsXml());
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


    private sealed class SizedAssetProvider : IDocxAssetProvider
    {
        private readonly int size;
        public SizedAssetProvider(int size)
        {
            this.size = size;
        }
        public bool TryOpen(string requestedReference, out Stream stream, out string? contentTypeHint, out string? fileNameHint)
        {
            // Structurally valid PNG of the requested size: signature, IHDR,
            // IDAT carrying the padding, IEND terminating the file exactly.
            // CRCs stay zero; validation is structural, not a decoder.
            int dataLength = Math.Max(0, size - 57);
            byte[] bytes = new byte[57 + dataLength];
            byte[] signature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
            Array.Copy(signature, bytes, signature.Length);
            WriteBigEndian(bytes, 8, 13);
            byte[] ihdrTag = [0x49, 0x48, 0x44, 0x52];
            Array.Copy(ihdrTag, 0, bytes, 12, ihdrTag.Length);
            WriteBigEndian(bytes, 16, 100);
            WriteBigEndian(bytes, 20, 50);
            bytes[24] = 0x08;
            bytes[25] = 0x02;
            WriteBigEndian(bytes, 33, dataLength);
            byte[] idatTag = [0x49, 0x44, 0x41, 0x54];
            Array.Copy(idatTag, 0, bytes, 37, idatTag.Length);
            int iend = 37 + 4 + dataLength + 4;
            WriteBigEndian(bytes, iend, 0);
            byte[] iendTag = [0x49, 0x45, 0x4E, 0x44];
            Array.Copy(iendTag, 0, bytes, iend + 4, iendTag.Length);
            bytes[iend + 8] = 0xAE;
            bytes[iend + 9] = 0x42;
            bytes[iend + 10] = 0x60;
            bytes[iend + 11] = 0x82;
            stream = new MemoryStream(bytes, writable: false);
            contentTypeHint = null;
            fileNameHint = "photo.png";
            return true;
        }
    }

    private static void WriteBigEndian(byte[] bytes, int offset, int value)
    {
        bytes[offset] = (byte)((value >> 24) & 0xFF);
        bytes[offset + 1] = (byte)((value >> 16) & 0xFF);
        bytes[offset + 2] = (byte)((value >> 8) & 0xFF);
        bytes[offset + 3] = (byte)(value & 0xFF);
    }

    private sealed class CancelAfterReadsReader : TextReader
    {
        private readonly CancellationTokenSource source;
        private readonly int cancelAfter;
        private int reads;
        private readonly string chunk = new string('x', 100);
        public CancelAfterReadsReader(CancellationTokenSource source, int cancelAfter)
        {
            this.source = source;
            this.cancelAfter = cancelAfter;
        }
        public override int Read(char[] buffer, int index, int count)
        {
            reads++;
            if (reads == cancelAfter)
            {
                source.Cancel();
            }
            if (reads > 500)
            {
                return 0;
            }
            int take = Math.Min(count, chunk.Length);
            chunk.CopyTo(0, buffer, index, take);
            return take;
        }
    }
}
