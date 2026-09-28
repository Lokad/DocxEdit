using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml.Linq;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

public static class PatchImageTests
{

    [Fact]
    public static void ExtractMediaReturnsInventoriedBytes()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");

        DocxMediaExtractResult result = new DocxEditor().ExtractMedia(input);

        Assert.True(result.Success);
        DocxMediaFile file = Assert.Single(result.Files);
        Assert.Equal("M.I0001", file.ImageId.ToWireValue());
        Assert.Equal("/word/media/image1.png", file.PartName);
        Assert.Equal("M.I0001-image1.png", file.FileName);
        Assert.Equal("old-png", Encoding.UTF8.GetString(file.Content));
    }

    [Fact]
    public static void ApplyReplaceImageUsesAssetProviderForPng()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", CreatePngBytes(4, 3), null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-image
            target M.I0001
            asset chart.png
            alt Updated chart
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal(CreatePngBytes(4, 3), ReadEntryBytes(output, "word/media/image1.png"));
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("descr=\"Updated chart\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("descr=\"Old chart\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyReplaceImagePreservesAnchoredLayoutAndCropMetadata()
    {
        using MemoryStream input = CreateDocxWithAnchoredImage();
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", CreatePngBytes(4, 3), null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-image
            target M.I0001
            asset chart.png
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal(CreatePngBytes(4, 3), ReadEntryBytes(output, "word/media/image1.png"));
        output.Position = 0;
        DocxImageInfo image = Assert.Single(new DocxEditor().Read(output).Images);
        Assert.Equal("anchor", image.LayoutKind);
        Assert.Equal(914400, image.WidthEmu);
        Assert.Equal(457200, image.HeightEmu);
        Assert.Equal("wrapSquare", image.WrapMode);
        Assert.Equal(10L, image.WrapDistanceTopEmu);
        Assert.Equal(20L, image.WrapDistanceBottomEmu);
        Assert.Equal(30L, image.WrapDistanceLeftEmu);
        Assert.Equal(40L, image.WrapDistanceRightEmu);
        Assert.Equal("column", image.HorizontalPositionRelativeFrom);
        Assert.Equal(123L, image.HorizontalPositionOffsetEmu);
        Assert.Equal("paragraph", image.VerticalPositionRelativeFrom);
        Assert.Equal("top", image.VerticalPositionAlign);
        Assert.Equal(1m, image.CropLeftPercent);
        Assert.Equal(2m, image.CropTopPercent);
        Assert.Equal(3m, image.CropRightPercent);
        Assert.Equal(4m, image.CropBottomPercent);
    }

    [Fact]
    public static void CheckReplaceImageRejectsFailedContentTypeGuard()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        var assets = new MemoryAssetProvider("chart.png", "new-png", null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-image
            target M.I0001
            expect-content-type image/jpeg
            asset chart.png
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { AssetProvider = assets });

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E3201" && diagnostic.TargetId == "M.I0001");
    }

    [Fact]
    public static void ApplyReplaceImageUsesAssetProviderForJpeg()
    {
        using MemoryStream input = CreateDocxWithImage("jpeg", "image/jpeg", "old-jpeg");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("photo.jpeg", CreateJpegBytes(4, 3), null, "photo.jpeg");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-image
            target M.I0001
            asset photo.jpeg
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal(CreateJpegBytes(4, 3), ReadEntryBytes(output, "word/media/image1.jpeg"));
    }

    [Fact]
    public static void CheckReplaceImageFailsWithoutAssetProvider()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-image
            target M.I0001
            asset chart.png
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E5201");
    }

    [Fact]
    public static void CheckReplaceImageRejectsUnsupportedAssetType()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        var assets = new MemoryAssetProvider("asset.bin", "not-image", null, "asset.bin");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-image
            target M.I0001
            asset asset.bin
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { AssetProvider = assets });

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E5203");
    }

    [Fact]
    public static void ApplySetImageAltUpdatesDrawingProperties()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-alt
            target M.I0001
            alt Updated chart
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("descr=\"Updated chart\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("descr=\"Old chart\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetImageMetadataUpdatesDrawingProperties()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-metadata
            target M.I0001
            alt Updated chart
            title Revenue chart
            name Revenue picture
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxImageInfo image = Assert.Single(new DocxEditor().Read(output).Images);
        Assert.Equal("Updated chart", image.Description);
        Assert.Equal("Revenue chart", image.Title);
        Assert.Equal("Revenue picture", image.Name);
    }

    [Fact]
    public static void ApplySetImageSizePreservesAspectWhenOnlyWidthIsGiven()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-size
            target M.I0001
            width 2in
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxImageInfo image = Assert.Single(new DocxEditor().Read(output).Images);
        Assert.Equal(1_828_800, image.WidthEmu);
        Assert.Equal(914_400, image.HeightEmu);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<wp:extent cx=\"1828800\" cy=\"914400\"", xml, StringComparison.Ordinal);
        Assert.Contains("<a:ext cx=\"1828800\" cy=\"914400\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetImageWrapUpdatesAnchorModeAndDistances()
    {
        using MemoryStream input = CreateDocxWithAnchoredImage();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-wrap
            target M.I0001
            mode top-bottom
            dist-top 1pt
            dist-right 2pt
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxImageInfo image = Assert.Single(new DocxEditor().Read(output).Images);
        Assert.Equal("wrapTopAndBottom", image.WrapMode);
        Assert.Equal(12_700, image.WrapDistanceTopEmu);
        Assert.Equal(20, image.WrapDistanceBottomEmu);
        Assert.Equal(30, image.WrapDistanceLeftEmu);
        Assert.Equal(25_400, image.WrapDistanceRightEmu);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<wp:wrapTopAndBottom", xml, StringComparison.Ordinal);
        Assert.Contains("distT=\"12700\"", xml, StringComparison.Ordinal);
        Assert.Contains("distR=\"25400\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("wrapSquare", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckSetImageWrapRejectsInlineImage()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-wrap
            target M.I0001
            mode square
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E5205");
    }

    [Fact]
    public static void ApplySetImagePositionUpdatesAnchorPosition()
    {
        using MemoryStream input = CreateDocxWithAnchoredImage();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-position
            target M.I0001
            horizontal-relative page
            horizontal-offset -0.25in
            vertical-relative paragraph
            vertical-align bottom
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxImageInfo image = Assert.Single(new DocxEditor().Read(output).Images);
        Assert.Equal("page", image.HorizontalPositionRelativeFrom);
        Assert.Equal(-228_600, image.HorizontalPositionOffsetEmu);
        Assert.Null(image.HorizontalPositionAlign);
        Assert.Equal("paragraph", image.VerticalPositionRelativeFrom);
        Assert.Null(image.VerticalPositionOffsetEmu);
        Assert.Equal("bottom", image.VerticalPositionAlign);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<wp:positionH relativeFrom=\"page\"><wp:posOffset>-228600</wp:posOffset></wp:positionH>", xml, StringComparison.Ordinal);
        Assert.Contains("<wp:positionV relativeFrom=\"paragraph\"><wp:align>bottom</wp:align></wp:positionV>", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckSetImagePositionRejectsInlineImage()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-position
            target M.I0001
            horizontal-offset 1in
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E5205");
    }

    [Fact]
    public static void ApplySetImageCropUpdatesSourceRectangle()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-crop
            target M.I0001
            left-percent 12.5
            top-percent 5
            right-percent 2.25
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxImageInfo image = Assert.Single(new DocxEditor().Read(output).Images);
        Assert.Equal(12.5m, image.CropLeftPercent);
        Assert.Equal(5m, image.CropTopPercent);
        Assert.Equal(2.25m, image.CropRightPercent);
        Assert.Null(image.CropBottomPercent);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<a:srcRect l=\"12500\" t=\"5000\" r=\"2250\" />", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckSetImageCropRejectsEmptyHorizontalCrop()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-crop
            target M.I0001
            left-percent 60
            right-percent 40
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E5208");
    }

    [Fact]
    public static void ApplyDeleteImageRemovesDrawingAndPreservesParagraph()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-image
            target M.I0001
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Empty(read.Images);
        Assert.Single(read.Paragraphs);
        output.Position = 0;
        Assert.DoesNotContain("<w:drawing>", ReadDocumentXml(output), StringComparison.Ordinal);
        output.Position = 0;
        Assert.False(EntryExists(output, "word/media/image1.png"));
        output.Position = 0;
        string relationships = ReadEntry(output, "word/_rels/document.xml.rels");
        Assert.DoesNotContain("rImage", relationships, StringComparison.Ordinal);
        Assert.DoesNotContain("media/image1.png", relationships, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyImageOperationsCanTargetHeaderAndFooterImages()
    {
        using MemoryStream input = CreateDocxWithHeaderFooterImages();
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("header.png", CreatePngBytes(4, 3), null, "header.png");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-image
            target H001.I0001
            asset header.png
            alt Updated header image
            end

            op delete-image
            target F001.I0001
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal(CreatePngBytes(4, 3), ReadEntryBytes(output, "word/media/header.png"));
        output.Position = 0;
        Assert.False(EntryExists(output, "word/media/footer.png"));
        output.Position = 0;
        Assert.Contains("descr=\"Updated header image\"", ReadEntry(output, "word/header1.xml"), StringComparison.Ordinal);
        output.Position = 0;
        Assert.DoesNotContain("rFooterImage", ReadEntry(output, "word/_rels/footer1.xml.rels"), StringComparison.Ordinal);
        output.Position = 0;
        DocxImageInfo image = Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { IncludeHeadersFooters = true }).Images);
        Assert.Equal("H001.I0001", image.Id.ToWireValue());
        Assert.Equal("/word/media/header.png", image.PartName);
    }

    [Fact]
    public static void ApplyInsertImageAfterAddsInlinePngImage()
    {
        using MemoryStream input = CreateDocx("Intro");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", CreatePngBytes(4, 3), null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset chart.png
            width 1in
            alt Inserted chart
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        DocxImageInfo image = Assert.Single(read.Images);
        Assert.Equal("/word/media/image1.png", image.PartName);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(2, read.Paragraphs.Count);
        output.Position = 0;
        Assert.Equal(CreatePngBytes(4, 3), ReadEntryBytes(output, "word/media/image1.png"));
        output.Position = 0;
        Assert.Contains("ContentType=\"image/png\"", ReadEntry(output, "[Content_Types].xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckRejectsImageAssetWithMismatchedExtensionAndContent()
    {
        byte[] jpegBytes = [0xFF, 0xD8, 0xFF, 0x00];
        var assets = new MemoryAssetProvider("chart.png", jpegBytes, null, "chart.png");
        using MemoryStream input = CreateDocx("Intro");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset chart.png
            width 1in
            height 1in
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { AssetProvider = assets });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal("E5203", diagnostic.Code);
        Assert.Contains("chart.png", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyInsertImageRejectsEmptyImageAsset()
    {
        using MemoryStream input = CreateDocx("Intro");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", [], null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset chart.png
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E5203");
    }

    [Fact]
    public static void ApplyInsertImageAfterPreservesPngAspectRatioWhenOnlyWidthIsSupplied()
    {
        using MemoryStream input = CreateDocx("Intro");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", CreatePngBytes(200, 100), null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset chart.png
            width 2in
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        (long cx, long cy) = ReadFirstInlineImageExtent(output);
        Assert.Equal(1_828_800, cx);
        Assert.Equal(914_400, cy);
    }

    [Fact]
    public static void ApplyInsertImageAfterInfersPngDimensionsWhenNoneAreSupplied()
    {
        using MemoryStream input = CreateDocx("Intro");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", CreatePngBytes(200, 100), null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset chart.png
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        (long cx, long cy) = ReadFirstInlineImageExtent(output);
        Assert.Equal(1_905_000, cx);
        Assert.Equal(952_500, cy);
    }

    [Fact]
    public static void ApplyInsertImageAfterPreservesJpegAspectRatioWhenOnlyHeightIsSupplied()
    {
        using MemoryStream input = CreateDocx("Intro");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("photo.jpeg", CreateJpegBytes(120, 60), null, "photo.jpeg");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset photo.jpeg
            height 1in
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        (long cx, long cy) = ReadFirstInlineImageExtent(output);
        Assert.Equal(1_828_800, cx);
        Assert.Equal(914_400, cy);
    }

    [Fact]
    public static void ApplyInsertImageAfterAddsInlineJpegImage()
    {
        using MemoryStream input = CreateDocx("Intro");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("photo.jpeg", CreateJpegBytes(4, 3), null, "photo.jpeg");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset photo.jpeg
            width 72pt
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        DocxImageInfo image = Assert.Single(new DocxEditor().Read(output).Images);
        Assert.Equal("/word/media/image1.jpeg", image.PartName);
        Assert.Equal("image/jpeg", image.ContentType);
        output.Position = 0;
        Assert.Equal(CreateJpegBytes(4, 3), ReadEntryBytes(output, "word/media/image1.jpeg"));
    }

    [Fact]
    public static void ApplyInsertImageAfterAllocatesUniqueDocPrId()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", CreatePngBytes(4, 3), null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset chart.png
            width 1in
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<wp:docPr id=\"1\"", xml, StringComparison.Ordinal);
        Assert.Contains("<wp:docPr id=\"2\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyInsertImageAfterTwiceRefreshesContentTypesAndRelationships()
    {
        using MemoryStream input = CreateDocx("Intro");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider(
            ("chart.png", CreatePngBytes(4, 3), null, "chart.png"),
            ("logo.png", CreatePngBytes(5, 6), null, "logo.png"));
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset chart.png
            width 1in
            end

            op insert-image-after
            target M.P0001
            asset logo.png
            width 72pt
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal(2, read.Images.Count);
        Assert.Contains(read.Images, image => image.PartName == "/word/media/image1.png" && image.ContentType == "image/png");
        Assert.Contains(read.Images, image => image.PartName == "/word/media/image2.png" && image.ContentType == "image/png");

        output.Position = 0;
        Assert.Equal(CreatePngBytes(4, 3), ReadEntryBytes(output, "word/media/image1.png"));
        output.Position = 0;
        Assert.Equal(CreatePngBytes(5, 6), ReadEntryBytes(output, "word/media/image2.png"));

        output.Position = 0;
        string contentTypes = ReadEntry(output, "[Content_Types].xml");
        Assert.Contains("PartName=\"/word/media/image1.png\"", contentTypes, StringComparison.Ordinal);
        Assert.Contains("ContentType=\"image/png\"", contentTypes, StringComparison.Ordinal);
        Assert.Contains("PartName=\"/word/media/image2.png\"", contentTypes, StringComparison.Ordinal);

        output.Position = 0;
        string relationships = ReadEntry(output, "word/_rels/document.xml.rels");
        Assert.Contains("Target=\"media/image1.png\"", relationships, StringComparison.Ordinal);
        Assert.Contains("Target=\"media/image2.png\"", relationships, StringComparison.Ordinal);
        Assert.Equal(2, CountOccurrences(relationships, "relationships/image"));
    }

    public static (long Cx, long Cy) ReadFirstInlineImageExtent(Stream docx)
    {
        string xml = ReadDocumentXml(docx);
        XElement extent = XDocument.Parse(xml)
            .Descendants(XName.Get("extent", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"))
            .First();
        XAttribute cx = extent.Attribute("cx") ?? throw new InvalidDataException("Test fixture image extent lacks cx.");
        XAttribute cy = extent.Attribute("cy") ?? throw new InvalidDataException("Test fixture image extent lacks cy.");
        return ((long)cx, (long)cy);
    }

    [Fact]
    public static void ApplySetImageAltWithMatchingExpectAltSucceeds()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-alt
            target M.I0001
            expect-alt Old chart
            alt Updated chart
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Contains("descr=" + (char)34 + "Updated chart" + (char)34, ReadDocumentXml(output), StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckSetImageAltWithMismatchedExpectAltFails()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-alt
            target M.I0001
            expect-alt Stale chart
            alt Updated chart
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, static d => d.Code == "E3201");
        Assert.Equal(5, diagnostic.Line);
        Assert.Equal(1, diagnostic.Column);
    }

    [Fact]
    public static void ApplySetImageMetadataWithMatchingExpectNameSucceeds()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-metadata
            target M.I0001
            expect-name Picture 1
            name Revenue picture
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Equal("Revenue picture", Assert.Single(new DocxEditor().Read(output).Images).Name);
    }

    [Fact]
    public static void CheckSetImageMetadataWithMissingExpectTitleFails()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-metadata
            target M.I0001
            expect-title Some title
            title New title
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E3201");
    }

}
