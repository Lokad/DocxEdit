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

    private static MemoryStream CreateDocxWithTwoImages()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            void Add(string name, string content)
            {
                ZipArchiveEntry entry = archive.CreateEntry(name);
                using var writer = new StreamWriter(entry.Open());
                writer.Write(content);
            }

            Add("[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """);
            Add("_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            Add("word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
                  <Relationship Id="rImage2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/>
                </Relationships>
                """);
            Add("word/document.xml", """
                <w:document
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:drawing>
                          <wp:inline>
                            <wp:extent cx="914400" cy="457200"/>
                            <wp:docPr id="1" name="Picture 1" descr="First chart"/>
                            <a:graphic>
                              <a:graphicData>
                                <pic:pic>
                                  <pic:blipFill>
                                    <a:blip r:embed="rImage1"/>
                                  </pic:blipFill>
                                </pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:p>
                      <w:r>
                        <w:drawing>
                          <wp:inline>
                            <wp:extent cx="914400" cy="457200"/>
                            <wp:docPr id="2" name="Picture 2" descr="Second chart"/>
                            <a:graphic>
                              <a:graphicData>
                                <pic:pic>
                                  <pic:blipFill>
                                    <a:blip r:embed="rImage2"/>
                                  </pic:blipFill>
                                </pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                    </w:p>
                  </w:body>
                </w:document>
                """);
            Add("word/media/image1.png", "old-png-1");
            Add("word/media/image2.png", "old-png-2");
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithSharedHeaderMedia()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            void Add(string name, string content)
            {
                ZipArchiveEntry entry = archive.CreateEntry(name);
                using var writer = new StreamWriter(entry.Open());
                writer.Write(content);
            }

            Add("[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                </Types>
                """);
            Add("_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            Add("word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/shared.png"/>
                  <Relationship Id="rHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                </Relationships>
                """);
            Add("word/_rels/header1.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/shared.png"/>
                  <Relationship Id="rUnique" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/unique.png"/>
                </Relationships>
                """);
            Add("word/document.xml", """
                <w:document
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:drawing>
                          <wp:inline>
                            <wp:extent cx="914400" cy="457200"/>
                            <wp:docPr id="1" name="Picture 1" descr="MainShared"/>
                            <a:graphic>
                              <a:graphicData>
                                <pic:pic>
                                  <pic:blipFill>
                                    <a:blip r:embed="rShared"/>
                                  </pic:blipFill>
                                </pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                    </w:p>
                  </w:body>
                </w:document>
                """);
            Add("word/header1.xml", """
                <w:hdr
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <w:p>
                    <w:r>
                      <w:drawing>
                        <wp:inline>
                          <wp:extent cx="914400" cy="457200"/>
                          <wp:docPr id="2" name="Picture 2" descr="HeaderShared"/>
                          <a:graphic>
                            <a:graphicData>
                              <pic:pic>
                                <pic:blipFill>
                                  <a:blip r:embed="rShared"/>
                                </pic:blipFill>
                              </pic:pic>
                            </a:graphicData>
                          </a:graphic>
                        </wp:inline>
                      </w:drawing>
                    </w:r>
                  </w:p>
                  <w:p>
                    <w:r>
                      <w:drawing>
                        <wp:inline>
                          <wp:extent cx="914400" cy="457200"/>
                          <wp:docPr id="3" name="Picture 3" descr="HeaderUnique"/>
                          <a:graphic>
                            <a:graphicData>
                              <pic:pic>
                                <pic:blipFill>
                                  <a:blip r:embed="rUnique"/>
                                </pic:blipFill>
                              </pic:pic>
                            </a:graphicData>
                          </a:graphic>
                        </wp:inline>
                      </w:drawing>
                    </w:r>
                  </w:p>
                </w:hdr>
                """);
            Add("word/media/shared.png", "shared-bytes");
            Add("word/media/unique.png", "unique-bytes");
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithRepeatedMedia()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            void Add(string name, string content)
            {
                ZipArchiveEntry entry = archive.CreateEntry(name);
                using var writer = new StreamWriter(entry.Open());
                writer.Write(content);
            }

            Add("[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """);
            Add("_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            Add("word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/shared.png"/>
                </Relationships>
                """);
            Add("word/document.xml", """
                <w:document
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:drawing>
                          <wp:inline>
                            <wp:extent cx="914400" cy="457200"/>
                            <wp:docPr id="1" name="Picture 1" descr="First shared"/>
                            <a:graphic>
                              <a:graphicData>
                                <pic:pic>
                                  <pic:blipFill>
                                    <a:blip r:embed="rShared"/>
                                  </pic:blipFill>
                                </pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:p>
                      <w:r>
                        <w:drawing>
                          <wp:inline>
                            <wp:extent cx="914400" cy="457200"/>
                            <wp:docPr id="2" name="Picture 2" descr="Second shared"/>
                            <a:graphic>
                              <a:graphicData>
                                <pic:pic>
                                  <pic:blipFill>
                                    <a:blip r:embed="rShared"/>
                                  </pic:blipFill>
                                </pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                    </w:p>
                  </w:body>
                </w:document>
                """);
            Add("word/media/shared.png", "shared-bytes");
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithDeletedParagraphImage()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            void Add(string name, string content)
            {
                ZipArchiveEntry entry = archive.CreateEntry(name);
                using var writer = new StreamWriter(entry.Open());
                writer.Write(content);
            }

            Add("[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """);
            Add("_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            Add("word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/a.png"/>
                  <Relationship Id="rB" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/b.png"/>
                  <Relationship Id="rC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/c.png"/>
                </Relationships>
                """);
            Add("word/document.xml", """
                <w:document
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:drawing>
                          <wp:inline>
                            <wp:extent cx="914400" cy="457200"/>
                            <wp:docPr id="1" name="Picture 1" descr="VisibleA"/>
                            <a:graphic>
                              <a:graphicData>
                                <pic:pic>
                                  <pic:blipFill>
                                    <a:blip r:embed="rA"/>
                                  </pic:blipFill>
                                </pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:del w:id="9" w:author="Example" w:date="2026-01-01T00:00:00Z">
                      <w:p>
                        <w:r>
                          <w:drawing>
                            <wp:inline>
                              <wp:extent cx="914400" cy="457200"/>
                              <wp:docPr id="2" name="Picture 2" descr="HiddenB"/>
                              <a:graphic>
                                <a:graphicData>
                                  <pic:pic>
                                    <pic:blipFill>
                                      <a:blip r:embed="rB"/>
                                    </pic:blipFill>
                                  </pic:pic>
                                </a:graphicData>
                              </a:graphic>
                            </wp:inline>
                          </w:drawing>
                        </w:r>
                      </w:p>
                    </w:del>
                    <w:p>
                      <w:r>
                        <w:drawing>
                          <wp:inline>
                            <wp:extent cx="914400" cy="457200"/>
                            <wp:docPr id="3" name="Picture 3" descr="VisibleC"/>
                            <a:graphic>
                              <a:graphicData>
                                <pic:pic>
                                  <pic:blipFill>
                                    <a:blip r:embed="rC"/>
                                  </pic:blipFill>
                                </pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                    </w:p>
                  </w:body>
                </w:document>
                """);
            Add("word/media/a.png", "bytes-a");
            Add("word/media/b.png", "bytes-b");
            Add("word/media/c.png", "bytes-c");
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithSdtParagraphImage()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            void Add(string name, string content)
            {
                ZipArchiveEntry entry = archive.CreateEntry(name);
                using var writer = new StreamWriter(entry.Open());
                writer.Write(content);
            }

            Add("[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """);
            Add("_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            Add("word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/a.png"/>
                  <Relationship Id="rS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/s.png"/>
                  <Relationship Id="rC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/c.png"/>
                </Relationships>
                """);
            Add("word/document.xml", """
                <w:document
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:drawing>
                          <wp:inline>
                            <wp:extent cx="914400" cy="457200"/>
                            <wp:docPr id="1" name="Picture 1" descr="VisibleA"/>
                            <a:graphic>
                              <a:graphicData>
                                <pic:pic>
                                  <pic:blipFill>
                                    <a:blip r:embed="rA"/>
                                  </pic:blipFill>
                                </pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:sdt>
                      <w:sdtPr><w:alias w:val="Box"/></w:sdtPr>
                      <w:sdtContent>
                        <w:p>
                          <w:r>
                            <w:drawing>
                              <wp:inline>
                                <wp:extent cx="914400" cy="457200"/>
                                <wp:docPr id="9" name="Picture 9" descr="SdtInner"/>
                                <a:graphic>
                                  <a:graphicData>
                                    <pic:pic>
                                      <pic:blipFill>
                                        <a:blip r:embed="rS"/>
                                      </pic:blipFill>
                                    </pic:pic>
                                  </a:graphicData>
                                </a:graphic>
                              </wp:inline>
                            </w:drawing>
                          </w:r>
                        </w:p>
                      </w:sdtContent>
                    </w:sdt>
                    <w:p>
                      <w:r>
                        <w:drawing>
                          <wp:inline>
                            <wp:extent cx="914400" cy="457200"/>
                            <wp:docPr id="3" name="Picture 3" descr="VisibleC"/>
                            <a:graphic>
                              <a:graphicData>
                                <pic:pic>
                                  <pic:blipFill>
                                    <a:blip r:embed="rC"/>
                                  </pic:blipFill>
                                </pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                    </w:p>
                  </w:body>
                </w:document>
                """);
            Add("word/media/a.png", "bytes-a");
            Add("word/media/s.png", "bytes-s");
            Add("word/media/c.png", "bytes-c");
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithDeletedRowImage()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            void Add(string name, string content)
            {
                ZipArchiveEntry entry = archive.CreateEntry(name);
                using var writer = new StreamWriter(entry.Open());
                writer.Write(content);
            }

            Add("[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """);
            Add("_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            Add("word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/a.png"/>
                  <Relationship Id="rB" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/b.png"/>
                  <Relationship Id="rC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/c.png"/>
                </Relationships>
                """);
            Add("word/document.xml", """
                <w:document
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
                      <w:tr><w:tc><w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="Picture 1" descr="VisibleA"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rA"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:tc></w:tr>
                      <w:tr><w:trPr><w:del w:id="10" w:author="Example" w:date="2026-01-01T00:00:00Z"/></w:trPr><w:tc><w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="2" name="Picture 2" descr="HiddenB"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rB"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:tc></w:tr>
                      <w:tr><w:tc><w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="3" name="Picture 3" descr="VisibleC"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rC"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                  </w:body>
                </w:document>
                """);
            Add("word/media/a.png", "bytes-a");
            Add("word/media/b.png", "bytes-b");
            Add("word/media/c.png", "bytes-c");
        }

        stream.Position = 0;
        return stream;
    }

    [Fact]
    public static void DeleteDoesNotRetargetLaterExplicitImage()
    {
        using MemoryStream input = CreateDocxWithTwoImages();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-image
            target M.I0001
            end

            op set-image-alt
            target M.I0002
            expect-alt Second chart
            alt Updated chart
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        output.Position = 0;
        Assert.Equal("Updated chart", Assert.Single(new DocxEditor().Read(output).Images).Description);
    }

    [Fact]
    public static void DeletedImageTargetFailsInsteadOfEditingNeighbour()
    {
        using MemoryStream input = CreateDocxWithTwoImages();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-image
            target M.I0001
            end

            op set-image-alt
            target M.I0001
            alt Updated chart
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);
        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }


    [Fact]
    public static void SharedHeaderPlacementIdentityAddressesDiscoveredDrawing()
    {
        using MemoryStream probe = CreateDocxWithSharedHeaderMedia();
        DocxReadResult before = new DocxEditor().Read(probe, new DocxReadOptions { IncludeHeadersFooters = true });
        Assert.True(before.Success);
        Assert.Equal(3, before.Images.Count);
        Assert.Equal("M.I0001", before.Images[0].Id.ToWireValue());
        Assert.Equal("MainShared", before.Images[0].Description);
        Assert.Equal("H001.I0001", before.Images[1].Id.ToWireValue());
        Assert.Equal("HeaderShared", before.Images[1].Description);
        Assert.Equal("H001.I0002", before.Images[2].Id.ToWireValue());
        Assert.Equal("HeaderUnique", before.Images[2].Description);

        using MemoryStream input = CreateDocxWithSharedHeaderMedia();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-alt
            target H001.I0002
            alt CHANGED
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));

        output.Position = 0;
        string header = ReadEntry(output, "word/header1.xml");
        Assert.Contains("descr=\"HeaderShared\"", header, StringComparison.Ordinal);
        Assert.Contains("descr=\"CHANGED\"", header, StringComparison.Ordinal);
        Assert.DoesNotContain("descr=\"HeaderUnique\"", header, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains("descr=\"MainShared\"", ReadDocumentXml(output), StringComparison.Ordinal);

        output.Position = 0;
        DocxReadResult after = new DocxEditor().Read(output, new DocxReadOptions { IncludeHeadersFooters = true });
        Assert.Equal("MainShared", after.Images[0].Description);
        Assert.Equal("HeaderShared", after.Images[1].Description);
        Assert.Equal("CHANGED", after.Images[2].Description);
    }

    [Fact]
    public static void RepeatedMediaPlacementsHaveDistinctPlacementIds()
    {
        using MemoryStream probe = CreateDocxWithRepeatedMedia();
        DocxReadResult before = new DocxEditor().Read(probe);
        Assert.True(before.Success);
        Assert.Equal(2, before.Images.Count);
        Assert.Equal("M.I0001", before.Images[0].Id.ToWireValue());
        Assert.Equal("First shared", before.Images[0].Description);
        Assert.Equal("M.I0002", before.Images[1].Id.ToWireValue());
        Assert.Equal("Second shared", before.Images[1].Description);

        using MemoryStream input = CreateDocxWithRepeatedMedia();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-alt
            target M.I0001
            alt CHANGED
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));

        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("descr=\"CHANGED\"", xml, StringComparison.Ordinal);
        Assert.Contains("descr=\"Second shared\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("descr=\"First shared\"", xml, StringComparison.Ordinal);

        output.Position = 0;
        DocxReadResult after = new DocxEditor().Read(output);
        Assert.Equal("CHANGED", after.Images[0].Description);
        Assert.Equal("Second shared", after.Images[1].Description);
    }

    [Fact]
    public static void RevisionHiddenDrawingOwnsNoPlacementId()
    {
        using MemoryStream probe = CreateDocxWithDeletedParagraphImage();
        DocxReadResult before = new DocxEditor().Read(probe);
        Assert.True(before.Success);
        Assert.Equal(2, before.Images.Count);
        Assert.Equal("M.I0001", before.Images[0].Id.ToWireValue());
        Assert.Equal("VisibleA", before.Images[0].Description);
        Assert.Equal("M.I0002", before.Images[1].Id.ToWireValue());
        Assert.Equal("VisibleC", before.Images[1].Description);

        using MemoryStream input = CreateDocxWithDeletedParagraphImage();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-alt
            target M.I0002
            alt CHANGED
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));

        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("descr=\"VisibleA\"", xml, StringComparison.Ordinal);
        Assert.Contains("descr=\"HiddenB\"", xml, StringComparison.Ordinal);
        Assert.Contains("descr=\"CHANGED\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("descr=\"VisibleC\"", xml, StringComparison.Ordinal);

        output.Position = 0;
        DocxReadResult after = new DocxEditor().Read(output);
        Assert.Equal("VisibleA", after.Images[0].Description);
        Assert.Equal("CHANGED", after.Images[1].Description);
    }

    [Fact]
    public static void PlacementCapabilitiesAgreeWithDiscovery()
    {
        using (MemoryStream input = CreateDocxWithRepeatedMedia())
        {
            AssertCapabilitiesDescribeImage(input, "M.I0001");
            AssertCapabilitiesDescribeImage(input, "M.I0002");
        }

        using (MemoryStream input = CreateDocxWithDeletedParagraphImage())
        {
            AssertCapabilitiesDescribeImage(input, "M.I0001");
            AssertCapabilitiesDescribeImage(input, "M.I0002");
        }
    }

    private static void AssertCapabilitiesDescribeImage(MemoryStream input, string targetId)
    {
        input.Position = 0;
        DocxCapabilitiesResult result = new DocxEditor().GetCapabilities(input, targetId, new DocxCapabilitiesOptions { TrackChanges = TrackChangesMode.Off });
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        Assert.Equal(targetId, result.TargetId);
        Assert.NotNull(result.Capabilities);
        Assert.Equal("image", result.Capabilities!.Kind);
        Assert.Equal("supported", result.Capabilities.Operations.First(static operation => operation.Operation == "set-image-alt").Support);

        input.Position = 0;
        using var patch = new StringReader("docxpatch 1\n\nop set-image-alt\ntarget " + targetId + "\nalt Updated\nend\n");
        DocxCheckResult check = new DocxEditor().Check(input, patch);
        Assert.True(check.Success, string.Join("|", check.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void BlockSdtDrawingOwnsNoPlacementId()
    {
        using MemoryStream probe = CreateDocxWithSdtParagraphImage();
        DocxReadResult before = new DocxEditor().Read(probe);
        Assert.True(before.Success);
        Assert.Equal(2, before.Images.Count);
        Assert.Equal("M.I0001", before.Images[0].Id.ToWireValue());
        Assert.Equal("VisibleA", before.Images[0].Description);
        Assert.Equal("M.I0002", before.Images[1].Id.ToWireValue());
        Assert.Equal("VisibleC", before.Images[1].Description);

        using MemoryStream input = CreateDocxWithSdtParagraphImage();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-alt
            target M.I0002
            alt CHANGED
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));

        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("descr=\"VisibleA\"", xml, StringComparison.Ordinal);
        Assert.Contains("descr=\"SdtInner\"", xml, StringComparison.Ordinal);
        Assert.Contains("descr=\"CHANGED\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("descr=\"VisibleC\"", xml, StringComparison.Ordinal);

        output.Position = 0;
        DocxReadResult after = new DocxEditor().Read(output);
        Assert.Equal("VisibleA", after.Images[0].Description);
        Assert.Equal("CHANGED", after.Images[1].Description);
    }

    [Fact]
    public static void DeletedRowDrawingOwnsNoPlacementId()
    {
        using MemoryStream probe = CreateDocxWithDeletedRowImage();
        DocxReadResult before = new DocxEditor().Read(probe);
        Assert.True(before.Success);
        Assert.Equal(2, before.Images.Count);
        Assert.Equal("M.I0001", before.Images[0].Id.ToWireValue());
        Assert.Equal("VisibleA", before.Images[0].Description);
        Assert.Equal("M.I0002", before.Images[1].Id.ToWireValue());
        Assert.Equal("VisibleC", before.Images[1].Description);

        using MemoryStream input = CreateDocxWithDeletedRowImage();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-alt
            target M.I0002
            alt CHANGED
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));

        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("descr=\"VisibleA\"", xml, StringComparison.Ordinal);
        Assert.Contains("descr=\"HiddenB\"", xml, StringComparison.Ordinal);
        Assert.Contains("descr=\"CHANGED\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("descr=\"VisibleC\"", xml, StringComparison.Ordinal);

        output.Position = 0;
        DocxReadResult after = new DocxEditor().Read(output);
        Assert.Equal("VisibleA", after.Images[0].Description);
        Assert.Equal("CHANGED", after.Images[1].Description);
    }

}
