using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using static Lokad.DocxEdit.Tests.DocxTestFixtures;
using static Lokad.DocxEdit.Tests.ImageFormatTestTools;

namespace Lokad.DocxEdit.Tests;

public static class WebpImageTests
{
    // Synthetic 4x2 color grids encoded with Pillow/libwebp. Embedded constants
    // keep tests independent of encoders, decoders, downloads, and Word.
    internal static byte[] Fixture(string variant = "Lossy") => Convert.FromBase64String(variant switch
    {
        "Lossy" => "UklGRlAAAABXRUJQVlA4IEQAAADwAQCdASoEAAIAAUAmJagCdLoAAwkG/ewA/ub4N/Tq6iNS//Q5a7RPsD/0OWu0Tuz/WXHA+//yb/KKUd7F9//k37dAAA==",
        "Lossless" => "UklGRjAAAABXRUJQVlA4TCMAAAAvA0AAAD8weAKBZLDwBxva+Q8NAMMYA4EAGVvQ4kGLiP4HHgA=",
        "TransparentLossless" => "UklGRjQAAABXRUJQVlA4TCgAAAAvA0AAEDcweAKBZLDwBxtagUAyYqgjhkQAk21DIEDGFrR48CGi/4EH",
        "TransparentLossy" => "UklGRnYAAABXRUJQVlA4WAoAAAAQAAAAAwAAAQAAQUxQSAkAAAAAAFCg8ABQoPAAVlA4IEYAAAAwAgCdASoEAAIAAUAmJagCdLoB+AADIQb97AD+3iGdnt9DPciQH2ekoP/PZy1h/YrhdOf/g1zRfn/8m63pOAl8//k38gAA",
        _ => throw new ArgumentException("Unknown synthetic variant.", nameof(variant))
    });

    [Theory]
    [InlineData("Lossy")]
    [InlineData("Lossless")]
    [InlineData("TransparentLossless")]
    [InlineData("TransparentLossy")]
    public static void NativeStillImageInsertionAndEditsPreserveEncodedBytes(string variant)
    {
        byte[] bytes = Fixture(variant);
        using var input = NewDocument();
        using var inserted = Insert(input, bytes, "webp");
        AssertImage(inserted, bytes, "image/webp", ".webp", 38100, 19050);
        using var edited = Apply(inserted, "op set-image-size\ntarget M.I0001\nheight 1in\nexpect-content-type image/webp\nend\nop set-image-metadata\ntarget M.I0001\nalt Transparent grid\ntitle Grid title\nname Grid name\nend\nop set-image-crop\ntarget M.I0001\nleft-percent 10\nright-percent 5\nend\n");
        AssertImage(edited, bytes, "image/webp", ".webp", 1828800, 914400);
        edited.Position = 0;
        var read = new DocxEditor().Read(edited);
        DocxImageInfo info = Assert.Single(read.Images);
        Assert.Equal("Transparent grid", info.Description);
        Assert.Equal("Grid title", info.Title);
        Assert.Equal("Grid name", info.Name);
        Assert.Equal(10, info.CropLeftPercent);
        Assert.Equal(5, info.CropRightPercent);
        Assert.Contains(read.Paragraphs, p => p.Text == "Unrelated content.");
        XDocument xml = XDocument.Parse(ReadDocumentXml(edited));
        string relationship = (string)xml.Descendants(A + "blip").Single().Attribute(R + "embed")!;
        var rels = XDocument.Parse(ReadEntry(edited, "word/_rels/document.xml.rels"));
        Assert.Equal("media/image1.webp", (string?)rels.Descendants(Rel + "Relationship").Single(r => (string?)r.Attribute("Id") == relationship).Attribute("Target"));
        CommentLifecycleTests.AssertValid(edited);
    }

    [Theory]
    [InlineData("png")]
    [InlineData("jpeg")]
    [InlineData("svg")]
    [InlineData("webp")]
    public static void FormatInterchangePreservesLayoutAndTargetsDisplayedBytes(string extension)
    {
        byte[] other = extension switch
        {
            "png" => CreatePngBytes(40, 20),
            "jpeg" => CreateJpegBytes(40, 20),
            "svg" => SvgImageTests.Diagram(),
            _ => Fixture("TransparentLossless")
        };
        string otherType = extension switch { "jpeg" => "image/jpeg", "svg" => "image/svg+xml", _ => "image/" + extension };
        using var input = NewDocument();
        using var initial = Insert(input, other, extension, "width 2in\n");
        using var webp = Replace(initial, Fixture(), "webp");
        AssertImage(webp, Fixture(), "image/webp", ".webp", 1828800, 914400);
        Assert.Empty(XDocument.Parse(ReadDocumentXml(webp)).Descendants(Asvg + "svgBlip"));
        using var restored = Replace(webp, other, extension);
        AssertImage(restored, other, otherType, "." + extension, 1828800, 914400);
        CommentLifecycleTests.AssertValid(restored);
    }

    [Theory]
    [InlineData("M.P0001", "M.I0001")]
    [InlineData("H001.P0001", "H001.I0001")]
    [InlineData("F001.P0001", "F001.I0001")]
    public static void EveryStorySupportsNativeWebpReplacementAndDeletion(string paragraph, string image)
    {
        using var input = CreateDocxWithHeaderFooter("Header", "Footer");
        using var inserted = Insert(input, Fixture(), "webp", "width 2in\n", paragraph);
        using var replaced = Replace(inserted, Fixture("TransparentLossy"), "webp", image);
        replaced.Position = 0;
        var file = Assert.Single(new DocxEditor().ExtractMedia(replaced, new DocxMediaOptions { IncludeHeadersFooters = true }).Files);
        Assert.Equal(image, file.ImageId.ToWireValue());
        Assert.Equal("image/webp", file.ContentType);
        Assert.Equal(Fixture("TransparentLossy"), file.Content);
        using var deleted = Apply(replaced, $"op delete-image\ntarget {image}\nend\n");
        deleted.Position = 0;
        Assert.Empty(new DocxEditor().Read(deleted, new DocxReadOptions { IncludeHeadersFooters = true }).Images);
        deleted.Position = 0;
        Assert.True(new DocxEditor().Validate(deleted).Success);
        using var archive = new ZipArchive(deleted, ZipArchiveMode.Read, leaveOpen: true);
        Assert.DoesNotContain(archive.Entries, entry => entry.FullName.StartsWith("word/media/", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public static void SharedNativeAssetSurvivesReplacementAndDeletion(bool replace, bool header)
    {
        using var input = NewDocument();
        using var shared = Insert(input, Fixture(), "webp");
        if (header) AddHeaderUse(shared);
        else EditXml(shared, "word/document.xml", doc =>
        {
            var copy = new XElement(doc.Descendants(W + "p").Single(p => p.Descendants(W + "drawing").Any()));
            copy.Descendants(Wp + "docPr").Single().SetAttributeValue("id", "2");
            doc.Descendants(W + "sectPr").Single().AddBeforeSelf(copy);
        });
        using var changed = replace ? Replace(shared, Fixture("TransparentLossless"), "webp") : Apply(shared, "op delete-image\ntarget M.I0001\nend\n");
        changed.Position = 0;
        var media = new DocxEditor().ExtractMedia(changed, new DocxMediaOptions { IncludeHeadersFooters = true });
        Assert.Equal(replace ? 2 : 1, media.Files.Count);
        Assert.Equal(Fixture(), media.Files[^1].Content);
        if (replace) Assert.Equal(Fixture("TransparentLossless"), media.Files[0].Content);
        CommentLifecycleTests.AssertValid(changed);
    }

    [Fact]
    public static void WordSvgPreviewIsReplacedByNativeWebp()
    {
        using var word = SvgImageTests.WordSvg();
        using var changed = Replace(word, Fixture(), "webp");
        AssertImage(changed, Fixture(), "image/webp", ".webp");
        Assert.Empty(XDocument.Parse(ReadDocumentXml(changed)).Descendants(Asvg + "svgBlip"));
        using var zip = new ZipArchive(changed, ZipArchiveMode.Read, leaveOpen: true);
        Assert.Null(zip.GetEntry("word/media/preview.png"));
        CommentLifecycleTests.AssertValid(changed);
    }

    [Fact]
    public static void MetadataAndUnknownChunksRoundTripWithoutReencoding()
    {
        byte[] extended = Fixture("TransparentLossy");
        extended[20] |= 0x0C;
        byte[] bytes = AppendChunks(extended, ("EXIF", "synthetic exif"u8.ToArray()), ("XMP ", "synthetic xmp"u8.ToArray()), ("TEST", new byte[] { 1, 2, 3 }));
        using var input = NewDocument();
        using var inserted = Insert(input, bytes, "webp");
        using var edited = Apply(inserted, "op set-image-size\ntarget M.I0001\nwidth 2in\nend\n");
        AssertImage(edited, bytes, "image/webp", ".webp");
    }

    [Fact]
    public static void CompressedAlphaRoundTripsAndAnimatedInputIsDiagnosed()
    {
        byte[] compressed = Convert.FromBase64String("UklGRpQAAABXRUJQVlA4WAoAAAAQAAAADwAABwAAQUxQSAwAAAABHzBQERGOkkf0PwBWUDggYgAAAHACAJ0BKhAACAABQCYlqAJ0ugEmAGGAHUKi12AA/vuYYXwhMKNoR5mb7+z+WuyHiib0wgS+JwJlAf9MIE0BvUmrTo+yuO0fjDU4P/6tuY3/VtzG9lvGf/I0iqGiK/wPxmAA");
        using var input = NewDocument();
        using var inserted = Insert(input, compressed, "webp");
        AssertImage(inserted, compressed, "image/webp", ".webp", 152400, 76200);
        byte[] animated = AnimatedFixture();
        AssertRejected(animated, "webp", "E5203");
        input.Position = 0;
        var check = new DocxEditor().Check(input, new StringReader("docxpatch 1\nop insert-image-after\ntarget M.P0001\nasset animated.webp\nend\n"),
            new DocxEditOptions { AssetProvider = new MemoryAssetProvider("animated.webp", animated, "image/webp", null) });
        Assert.Contains(check.Diagnostics, d => d.Message.Contains("animated WebP", StringComparison.Ordinal));
    }

    [Fact]
    public static void ExistingAnimatedMediaCanBeInspectedAndExtractedWithoutConversion()
    {
        using var input = NewDocument();
        using var document = Insert(input, Fixture(), "webp");
        byte[] animated = AnimatedFixture();
        SetEntry(document, "word/media/image1.webp", animated);
        using var edited = Apply(document, "op set-image-alt\ntarget M.I0001\nalt Existing animation\nend\n");
        AssertImage(edited, animated, "image/webp", ".webp");
    }

    private static byte[] AnimatedFixture() => Convert.FromBase64String("UklGRoQAAABXRUJQVlA4WAoAAAASAAAAAwAAAQAAQU5JTQYAAAAAAAAAAABBTk1GKAAAAAAAAAAAAAMAAAEAAGQAAAJWUDhMDwAAAC8DQAAQBxD9jyIHIqL/AQBBTk1GKAAAAAAAAAAAAAMAAAEAAGQAAAJWUDhMDwAAAC8DQAAQBxDR/+IFIqL/AQA=");

    [Theory]
    [InlineData("svg")]
    [InlineData("webp")]
    public static void MissingStoredExtentsUseSourceAspectRatio(string extension)
    {
        byte[] bytes = extension == "svg" ? SvgImageTests.Diagram() : Fixture();
        using var input = NewDocument();
        using var inserted = Insert(input, bytes, extension);
        EditXml(inserted, "word/document.xml", doc => doc.Descendants(Wp + "extent").Remove());
        using var resized = Apply(inserted, "op set-image-size\ntarget M.I0001\nwidth 2in\nend\n");
        AssertImage(resized, bytes, extension == "svg" ? "image/svg+xml" : "image/webp", "." + extension, 1828800, 914400);
        CommentLifecycleTests.AssertValid(resized);
    }

    public static IEnumerable<object[]> MalformedAssets()
    {
        yield return ["signature only", "RIFF\0\0\0\0WEBP"u8.ToArray()];
        yield return ["truncated payload", Fixture()[..^2]];
        var size = Fixture(); size[4] = 0; yield return ["wrong RIFF size", size];
        var chunk = Fixture(); BinaryPrimitives.WriteUInt32LittleEndian(chunk.AsSpan(16, 4), uint.MaxValue); yield return ["oversized chunk", chunk];
        var padding = Fixture("Lossless"); padding[^1] = 1; yield return ["nonzero padding", padding];
        var version = Fixture("Lossless"); version[24] |= 0x20; yield return ["unsupported lossless version", version];
        var frame = Fixture(); frame[23] = 0; yield return ["bad VP8 frame signature", frame];
        var dimensions = Fixture(); dimensions[26] = dimensions[27] = 0; yield return ["zero VP8 width", dimensions];
        var canvas = Fixture("TransparentLossy"); canvas[24] = 4; yield return ["canvas mismatch", canvas];
        var animation = Fixture("TransparentLossy"); animation[20] |= 2; yield return ["animation flag", animation];
        yield return ["animation chunk", AppendChunks(Fixture(), ("ANIM", new byte[6]))];
        yield return ["duplicate bitstream", AppendChunks(Fixture("Lossless"), ("VP8L", Fixture("Lossless")[20..^1]))];
        var alpha = Fixture("TransparentLossy"); alpha[24] = alpha[62] = 2; yield return ["raw alpha mismatch", alpha];
        var missingImage = Fixture()[..12]; BinaryPrimitives.WriteUInt32LittleEndian(missingImage.AsSpan(4, 4), 4); yield return ["missing image", missingImage];
    }

    [Theory]
    [MemberData(nameof(MalformedAssets))]
    public static void MalformedOrAnimatedWebpFailsCheckAndApply(string scenario, byte[] bytes)
    {
        Assert.NotEmpty(scenario);
        AssertRejected(bytes, "webp", "E5203");
    }

    [Fact]
    public static void WebpQuotaAndMismatchedHintsFailBeforePublication()
    {
        AssertRejected(AppendChunks(Fixture(), ("TEST", new byte[9000])), "webp", "E5207", options: new DocxEditOptions { Quotas = new DocxPackageLimits(100, 100000, 8000) });
        AssertRejected(Fixture(), "png", "E5203");
        AssertRejected(CreatePngBytes(4, 2), "webp", "E5203");
    }

    [Fact]
    public static async Task HostedCommandsUseAndExtractOriginalWebp() => await AssertHosted(Fixture("TransparentLossless"), "webp", "image/webp");

    [Fact]
    public static async Task CanceledWebpAssetDoesNotPublish()
    {
        var host = HostedCommandTests.CreateHost();
        host.Files["/asset.webp"] = Fixture();
        host.Files["/edit.patch"] = "docxpatch 1\nop insert-image-after\ntarget M.P0001\nasset /asset.webp\nend\n"u8.ToArray();
        using var canceled = new CancellationTokenSource();
        host.OnAssetRead = canceled.Cancel;
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => DocxCommand.RunAsync(["apply", "/input.docx", "/edit.patch", "-o", "/out.docx"], host, canceled.Token));
        Assert.Empty(host.Published);
    }

    private static byte[] AppendChunks(byte[] original, params (string Kind, byte[] Payload)[] chunks)
    {
        using var stream = new MemoryStream();
        stream.Write(original);
        foreach (var (kind, payload) in chunks)
        {
            stream.Write(Encoding.ASCII.GetBytes(kind));
            var length = new byte[4];
            BinaryPrimitives.WriteUInt32LittleEndian(length, (uint)payload.Length);
            stream.Write(length);
            stream.Write(payload);
            if ((payload.Length & 1) != 0) stream.WriteByte(0);
        }
        byte[] bytes = stream.ToArray();
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(4, 4), (uint)(bytes.Length - 8));
        return bytes;
    }
}
