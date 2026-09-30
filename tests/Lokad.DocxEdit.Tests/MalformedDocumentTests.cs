using System.IO.Compression;
using System.Text;
using System.Text.Json;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// PLAN B12: malformed document data becomes failed results with diagnostics at
// its boundary (never an unhandled throw / exit 4); programmer errors still throw.
[Collection("ConsoleCli")]
public static class MalformedDocumentTests
{
    [Fact]
    public static void DuplicateStyleIdsFailReadWithDiagnostic()
    {
        using MemoryStream input = CreateDocxWithDuplicateStyles();
        DocxReadResult result = new DocxEditor().Read(input);
        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static d => d.Code == "E0001");
        Assert.Contains(result.Diagnostics, static d => d.Message.Contains("Dup", StringComparison.Ordinal));
    }

    [Fact]
    public static void DuplicateStyleIdsKeepCliJsonPredictable()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = Path.Combine(temp.Path, "input.docx");
        using (MemoryStream source = CreateDocxWithDuplicateStyles())
        using (FileStream file = File.Create(input))
        {
            source.CopyTo(file);
        }

        CliOutcome result = RunCli("read", input, "--json");

        Assert.Equal(1, result.ExitCode);
        using JsonDocument json = JsonDocument.Parse(result.Output);
        Assert.False(json.RootElement.GetProperty("Success").GetBoolean());
        Assert.Contains("E0001", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void HugeCropPercentFailsValidateWithoutThrow()
    {
        using MemoryStream input = CreateDocxWithHugeCrop();
        DocxValidateResult result = new DocxEditor().Validate(input);
        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static d => d.Code == "E9110");
    }

    [Fact]
    public static void HugePatchDimensionFailsWithoutThrow()
    {
        using MemoryStream input = CreateDocx("Alpha.");
        using var patch = new StringReader("docxpatch 1\n\nop insert-image-after\ntarget M.P0001\nasset chart.png\nwidth 1e999in\nend\n");
        var assets = new TestAssetProvider("chart.png", ValidPng(4, 3), null, "chart.png");
        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { AssetProvider = assets });
        Assert.False(result.Success);
        Assert.DoesNotContain(result.Diagnostics, static d => d.Code == "E0001");
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
            AddEntry(archive, "word/styles.xml", "<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" />");
        }
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithDuplicateStyles()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesWithStylesXml());
            AddEntry(archive, "_rels/.rels", RootRelsXml());
            AddEntry(archive, "word/_rels/document.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rStyles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>");
            AddEntry(archive, "word/document.xml", "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>"
                + "<w:p><w:r><w:t>Alpha.</w:t></w:r></w:p>"
                + "</w:body></w:document>");
            AddEntry(archive, "word/styles.xml", "<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                + "<w:style w:type=\"paragraph\" w:styleId=\"Dup\"><w:name w:val=\"First\"/></w:style>"
                + "<w:style w:type=\"paragraph\" w:styleId=\"Dup\"><w:name w:val=\"Second\"/></w:style>"
                + "</w:styles>");
        }
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithHugeCrop()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", RootRelsXml());
            AddEntry(archive, "word/_rels/document.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rImage\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image1.png\"/></Relationships>");
            AddEntry(archive, "word/document.xml", "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
                + "<w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"914400\" cy=\"457200\"/><wp:docPr id=\"1\" name=\"Picture 1\"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed=\"rImage\"/><a:srcRect l=\"99999999999999999999999%\"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>");
            AddEntry(archive, "word/media/image1.png", "fakepng");
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

    private static string ContentTypesWithStylesXml()
    {
        return "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
            + "<Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>"
            + "</Types>";
    }

    private static string RootRelsXml()
    {
        return "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rDocument\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
            + "</Relationships>";
    }

    private static CliOutcome RunCli(params string[] args)
    {
        var captured = DocxTestFixtures.CaptureCliOutput(null, args);
        return new CliOutcome(captured.ExitCode, captured.Output, captured.Error);
    }


    private static byte[] ValidPng(int width, int height)
    {
        using var stream = new MemoryStream();
        stream.Write(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A });
        WriteChunk(stream, "IHDR", IHDRData(width, height));
        WriteChunk(stream, "IDAT", []);
        WriteChunk(stream, "IEND", []);
        return stream.ToArray();
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


    private sealed record CliOutcome(int ExitCode, string Output, string Error);

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
