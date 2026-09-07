using System.IO.Compression;
using System.Text;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// PLAN B10: image bytes are validated by signature and structure; provider
// MIME/name hints are only hints. Magic wins; disagreements and truncated
// structures fail before publication in check and apply alike.
public static class ImageValidationTests
{
    [Fact]
    public static void TextFileNamedPngIsRejected()
    {
        using MemoryStream input = CreateDocx("Alpha.");
        var provider = new TestAssetProvider("note.png", "just some text, not an image"u8.ToArray(), null, "note.png");
        var options = new DocxEditOptions { AssetProvider = provider };
        using var patch = new StringReader("docxpatch 1\n\nop insert-image-after\ntarget M.P0001\nasset note\nend\n");
        DocxCheckResult check = new DocxEditor().Check(input, patch, options);
        Assert.False(check.Success);
        Assert.Contains(check.Diagnostics, static d => d.Code == "E5203");
        using MemoryStream applyInput = CreateDocx("Alpha.");
        using var applyPatch = new StringReader("docxpatch 1\n\nop insert-image-after\ntarget M.P0001\nasset note\nend\n");
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, applyPatch, output, options);
        Assert.False(apply.Success);
        Assert.Contains(apply.Diagnostics, static d => d.Code == "E5203");
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public static void TruncatedPngIsRejected()
    {
        byte[] full = ValidPng(10, 20);
        byte[] truncated = full[..^12];
        using MemoryStream input = CreateDocx("Alpha.");
        var provider = new TestAssetProvider("cut.png", truncated, null, "cut.png");
        var options = new DocxEditOptions { AssetProvider = provider };
        using var patch = new StringReader("docxpatch 1\n\nop insert-image-after\ntarget M.P0001\nasset cut\nend\n");
        DocxCheckResult check = new DocxEditor().Check(input, patch, options);
        Assert.False(check.Success);
        Assert.Contains(check.Diagnostics, static d => d.Code == "E5203");
    }

    [Fact]
    public static void TruncatedJpegIsRejected()
    {
        byte[] full = ValidJpeg(30, 40);
        byte[] truncated = full[..^2];
        using MemoryStream input = CreateDocx("Alpha.");
        var provider = new TestAssetProvider("cut.jpg", truncated, null, "cut.jpg");
        var options = new DocxEditOptions { AssetProvider = provider };
        using var patch = new StringReader("docxpatch 1\n\nop insert-image-after\ntarget M.P0001\nasset cut\nend\n");
        DocxCheckResult check = new DocxEditor().Check(input, patch, options);
        Assert.False(check.Success);
        Assert.Contains(check.Diagnostics, static d => d.Code == "E5203");
    }

    [Fact]
    public static void ContradictoryHintIsRejected()
    {
        byte[] jpeg = ValidJpeg(30, 40);
        using MemoryStream input = CreateDocx("Alpha.");
        var provider = new TestAssetProvider("photo", jpeg, "image/png", "photo");
        var options = new DocxEditOptions { AssetProvider = provider };
        using var patch = new StringReader("docxpatch 1\n\nop insert-image-after\ntarget M.P0001\nasset photo\nend\n");
        DocxCheckResult check = new DocxEditor().Check(input, patch, options);
        Assert.False(check.Success);
        Assert.Contains(check.Diagnostics, static d => d.Code == "E5203");
    }

    [Fact]
    public static void AgreeingHintWithOddNameSucceeds()
    {
        byte[] png = ValidPng(10, 20);
        using MemoryStream input = CreateDocx("Alpha.");
        var provider = new TestAssetProvider("oddname.jpg", png, "image/png", "oddname.jpg");
        var options = new DocxEditOptions { AssetProvider = provider };
        using var patch = new StringReader("docxpatch 1\n\nop insert-image-after\ntarget M.P0001\nasset odd\nend\n");
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output, options);
        Assert.True(apply.Success);
        output.Position = 0;
        Assert.Equal(png, ReadEntryBytes(output, "word/media/image1.png"));
    }

    [Fact]
    public static void ValidPngInsertRoundTripsBytes()
    {
        byte[] png = ValidPng(200, 100);
        using MemoryStream input = CreateDocx("Alpha.");
        var provider = new TestAssetProvider("chart.png", png, null, "chart.png");
        var options = new DocxEditOptions { AssetProvider = provider };
        using var patch = new StringReader("docxpatch 1\n\nop insert-image-after\ntarget M.P0001\nasset chart\nend\n");
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output, options);
        Assert.True(apply.Success);
        output.Position = 0;
        Assert.Equal(png, ReadEntryBytes(output, "word/media/image1.png"));
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.True(read.Success);
        DocxImageInfo image = Assert.Single(read.Images);
        Assert.Equal("image/png", image.ContentType);
    }

    [Fact]
    public static void ValidJpegReplaceRoundTripsBytes()
    {
        byte[] jpeg = ValidJpeg(120, 60);
        using MemoryStream input = CreateDocxWithImage(jpeg, "image/jpeg");
        var provider = new TestAssetProvider("photo.jpeg", jpeg, null, "photo.jpeg");
        var options = new DocxEditOptions { AssetProvider = provider };
        using var patch = new StringReader("docxpatch 1\n\nop replace-image\ntarget M.I0001\nasset photo\nend\n");
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output, options);
        Assert.True(apply.Success);
        output.Position = 0;
        Assert.Equal(jpeg, ReadEntryBytes(output, "word/media/image1.jpeg"));
    }

    internal static byte[] ValidPng(int width, int height)
    {
        using var stream = new MemoryStream();
        stream.Write(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A });
        WriteChunk(stream, "IHDR", IHDRData(width, height));
        WriteChunk(stream, "IDAT", []);
        WriteChunk(stream, "IEND", []);
        return stream.ToArray();
    }

    internal static byte[] ValidJpeg(int width, int height)
    {
        var bytes = new List<byte> { 0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00 };
        bytes.AddRange([0xFF, 0xC0, 0x00, 0x0B, 0x08]);
        bytes.Add((byte)((height >> 8) & 0xFF));
        bytes.Add((byte)(height & 0xFF));
        bytes.Add((byte)((width >> 8) & 0xFF));
        bytes.Add((byte)(width & 0xFF));
        bytes.AddRange([0x01, 0x01, 0x11, 0x00]);
        bytes.AddRange([0xFF, 0xDA, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3F, 0x00, 0x01, 0xD2, 0xCF, 0x20, 0xFF, 0xD9]);
        return bytes.ToArray();
    }

    private static byte[] IHDRData(int width, int height)
    {
        return
        [
            (byte)((width >> 24) & 0xFF), (byte)((width >> 16) & 0xFF), (byte)((width >> 8) & 0xFF), (byte)(width & 0xFF),
            (byte)((height >> 24) & 0xFF), (byte)((height >> 16) & 0xFF), (byte)((height >> 8) & 0xFF), (byte)(height & 0xFF),
            0x08, 0x02, 0x00, 0x00, 0x00
        ];
    }

    private static void WriteChunk(MemoryStream stream, string type, byte[] data)
    {
        stream.Write([(byte)((data.Length >> 24) & 0xFF), (byte)((data.Length >> 16) & 0xFF), (byte)((data.Length >> 8) & 0xFF), (byte)(data.Length & 0xFF)]);
        byte[] typeBytes = Encoding.ASCII.GetBytes(type);
        stream.Write(typeBytes);
        stream.Write(data);
        stream.Write([0, 0, 0, 0]);
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

    private static MemoryStream CreateDocxWithImage(byte[] imageBytes, string contentType)
    {
        string extension = contentType == "image/jpeg" ? "jpeg" : "png";
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Default Extension=\"" + extension + "\" ContentType=\"" + contentType + "\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "</Types>");
            AddEntry(archive, "_rels/.rels", RootRelsXml());
            AddEntry(archive, "word/_rels/document.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rImage\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image1." + extension + "\"/>"
                + "</Relationships>");
            AddEntry(archive, "word/document.xml", "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
                + "<w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"914400\" cy=\"457200\"/><wp:docPr id=\"1\" name=\"Picture 1\"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed=\"rImage\"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>");
            ZipArchiveEntry media = archive.CreateEntry("word/media/image1." + extension);
            using Stream mediaStream = media.Open();
            mediaStream.Write(imageBytes);
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


    private sealed class TestAssetProvider : IDocxAssetProvider
    {
        private readonly string reference;
        private readonly byte[] bytes;
        private readonly string? contentTypeHint;
        private readonly string fileNameHint;
        public TestAssetProvider(string reference, byte[] bytes, string? contentTypeHint, string fileNameHint)
        {
            this.reference = reference;
            this.bytes = bytes;
            this.contentTypeHint = contentTypeHint;
            this.fileNameHint = fileNameHint;
        }
        public bool TryOpen(string requestedReference, out Stream stream, out string? contentTypeHint, out string? fileNameHint)
        {
            stream = new MemoryStream(bytes, writable: false);
            contentTypeHint = this.contentTypeHint;
            fileNameHint = this.fileNameHint;
            return true;
        }
    }
}
