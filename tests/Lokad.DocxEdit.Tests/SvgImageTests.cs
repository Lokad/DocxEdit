using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using static Lokad.DocxEdit.Tests.DocxTestFixtures;
using static Lokad.DocxEdit.Tests.ImageFormatTestTools;

namespace Lokad.DocxEdit.Tests;

public static class SvgImageTests
{
    internal static byte[] Diagram(string dimensions = "viewBox=\"0 0 200 100\"") =>
        Encoding.UTF8.GetBytes($"<svg xmlns=\"http://www.w3.org/2000/svg\" {dimensions}><rect width=\"200\" height=\"100\" fill=\"blue\"/><path d=\"M0 0L200 100\" stroke=\"white\"/></svg>");

    [Theory]
    [InlineData("viewBox=\"0 0 200 100\"", "", 1905000, 952500)]
    [InlineData("width=\"2in\" height=\"1in\"", "", 1828800, 914400)]
    [InlineData("width=\"100%\" viewBox=\"0,0,200,100\"", "width 2in\n", 1828800, 914400)]
    [InlineData("width=\"200.5px\" viewBox=\"0 0 200 100\"", "height 1in\n", 1828800, 914400)]
    [InlineData("", "width 2in\nheight 1in\n", 1828800, 914400)]
    public static void OriginalVectorBytesAndDimensionsRoundTrip(string dimensions, string fields, long width, long height)
    {
        byte[] bytes = Diagram(dimensions);
        using var input = NewDocument();
        using var output = Insert(input, bytes, "svg", fields);
        AssertImage(output, bytes, "image/svg+xml", ".svg", width, height);
        CommentLifecycleTests.AssertValid(output);
        XDocument xml = XDocument.Parse(ReadDocumentXml(output));
        Assert.Single(xml.Descendants(Asvg + "svgBlip"));
        output.Position = 0;
        Assert.Contains(new DocxEditor().Read(output).Paragraphs, p => p.Text == "Unrelated content.");
    }

    [Fact]
    public static void VectorEditsAndRasterInterchangePreserveLayoutAndBytes()
    {
        using var input = NewDocument();
        using var vector = Insert(input, Diagram(), "svg", "width 2in\n");
        using var edited = Apply(vector, "op set-image-size\ntarget M.I0001\nwidth 4in\nexpect-content-type image/svg+xml\nend\nop set-image-metadata\ntarget M.I0001\nalt New diagram\ntitle Diagram title\nname Diagram name\nend\nop set-image-crop\ntarget M.I0001\nleft-percent 10\nend\n");
        AssertImage(edited, Diagram(), "image/svg+xml", ".svg", 3657600, 1828800);
        edited.Position = 0;
        DocxImageInfo info = Assert.Single(new DocxEditor().Read(edited).Images);
        Assert.Equal("New diagram", info.Description);
        Assert.Equal("Diagram title", info.Title);
        Assert.Equal(10, info.CropLeftPercent);
        byte[] png = CreatePngBytes(20, 10);
        using var raster = Replace(edited, png, "png");
        AssertImage(raster, png, "image/png", ".png", 3657600, 1828800);
        Assert.Empty(XDocument.Parse(ReadDocumentXml(raster)).Descendants(Asvg + "svgBlip"));
        using var restored = Replace(raster, Diagram("width=\"600\" height=\"300\""), "svg");
        using var changed = Replace(restored, Diagram(), "svg");
        AssertImage(changed, Diagram(), "image/svg+xml", ".svg", 3657600, 1828800);
        using var deleted = Apply(changed, "op delete-image\ntarget M.I0001\nend\n");
        deleted.Position = 0;
        Assert.Empty(new DocxEditor().Read(deleted).Images);
        using var archive = new ZipArchive(deleted, ZipArchiveMode.Read, leaveOpen: true);
        Assert.DoesNotContain(archive.Entries, entry => entry.FullName.StartsWith("word/media/", StringComparison.Ordinal));
        CommentLifecycleTests.AssertValid(restored);
    }

    [Theory]
    [InlineData("M.P0001", "M.I0001")]
    [InlineData("H001.P0001", "H001.I0001")]
    [InlineData("F001.P0001", "F001.I0001")]
    public static void EveryStorySupportsVectorInsertReplaceAndDelete(string paragraph, string image)
    {
        using var input = CreateDocxWithHeaderFooter("Header", "Footer");
        using var inserted = Insert(input, Diagram(), "svg", target: paragraph);
        inserted.Position = 0;
        Assert.Equal(image, Assert.Single(new DocxEditor().Read(inserted, new DocxReadOptions { IncludeHeadersFooters = true }).Images).Id.ToWireValue());
        using var replaced = Replace(inserted, CreateJpegBytes(20, 10), "jpeg", image);
        replaced.Position = 0;
        Assert.Equal("image/jpeg", Assert.Single(new DocxEditor().ExtractMedia(replaced, new DocxMediaOptions { IncludeHeadersFooters = true }).Files).ContentType);
        using var deleted = Apply(replaced, $"op delete-image\ntarget {image}\nend\n");
        deleted.Position = 0;
        Assert.Empty(new DocxEditor().Read(deleted, new DocxReadOptions { IncludeHeadersFooters = true }).Images);
        deleted.Position = 0;
        Assert.True(new DocxEditor().Validate(deleted).Success);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public static void WordVectorAndPreviewAreOneImageAndSharedUsesSurvive(bool replace)
    {
        using var word = WordSvg(shared: true);
        word.Position = 0;
        DocxMediaExtractResult before = new DocxEditor().ExtractMedia(word);
        Assert.Equal(2, before.Files.Count);
        Assert.All(before.Files, file => Assert.Equal(Diagram(), file.Content));
        using var changed = replace ? Replace(word, CreatePngBytes(20, 10), "png") : Apply(word, "op delete-image\ntarget M.I0001\nend\n");
        changed.Position = 0;
        var files = new DocxEditor().ExtractMedia(changed).Files;
        Assert.Equal(replace ? 2 : 1, files.Count);
        Assert.Equal(Diagram(), files[^1].Content);
        Assert.Equal(CreatePngBytes(2, 1), ReadEntryBytes(changed, "word/media/preview.png"));
        CommentLifecycleTests.AssertValid(changed);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public static void WordPreviewIsRemovedWithExclusiveVector(bool replace)
    {
        using var word = WordSvg();
        using var changed = replace ? Replace(word, Diagram("width=\"100\" height=\"50\""), "svg") : Apply(word, "op delete-image\ntarget M.I0001\nend\n");
        using var archive = new ZipArchive(changed, ZipArchiveMode.Read, leaveOpen: true);
        Assert.Null(archive.GetEntry("word/media/preview.png"));
        if (replace) AssertImage(changed, Diagram("width=\"100\" height=\"50\""), "image/svg+xml", ".svg", 1905000, 952500);
    }

    [Theory]
    [InlineData("<svg/>", "E5203")]
    [InlineData("<svg xmlns=\"http://www.w3.org/2000/svg\"><path></svg>", "E5203")]
    [InlineData("<!DOCTYPE svg [<!ENTITY x 'bad'>]><svg xmlns=\"http://www.w3.org/2000/svg\">&x;</svg>", "E5203")]
    [InlineData("<svg xmlns=\"http://www.w3.org/2000/svg\"><script/></svg>", "E5203")]
    [InlineData("<svg xmlns=\"http://www.w3.org/2000/svg\"><image href=\"remote.png\"/></svg>", "E5203")]
    [InlineData("<svg xmlns=\"http://www.w3.org/2000/svg\"/>", "E5206")]
    [InlineData("<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 -1 3\"/>", "E5206")]
    public static void InvalidOrUnsizedSvgFailsCheckAndApply(string svg, string code) =>
        AssertRejected(Encoding.UTF8.GetBytes(svg), "svg", code);

    [Fact]
    public static async Task HostedCommandsUseAndExtractSvgWithoutConversion() => await AssertHosted(Diagram(), "svg", "image/svg+xml");

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public static void VectorAndPreviewSharedAcrossStoriesSurvive(bool replace)
    {
        using var word = WordSvg();
        AddHeaderUse(word);
        using var changed = replace ? Replace(word, Diagram("width=\"20\" height=\"10\""), "svg") : Apply(word, "op delete-image\ntarget M.I0001\nend\n");
        changed.Position = 0;
        var media = new DocxEditor().ExtractMedia(changed, new DocxMediaOptions { IncludeHeadersFooters = true });
        Assert.Equal(Diagram(), Assert.Single(media.Files, file => file.ImageId.ToWireValue() == "H001.I0001").Content);
        Assert.Equal(CreatePngBytes(2, 1), ReadEntryBytes(changed, "word/media/preview.png"));
        CommentLifecycleTests.AssertValid(changed);
    }

    [Fact]
    public static void SvgQuotaAndDimensionOverflowFailBeforePublication()
    {
        byte[] oversized = Encoding.UTF8.GetBytes("<svg xmlns=\"http://www.w3.org/2000/svg\"><!--" + new string('x', 9000) + "--></svg>");
        AssertRejected(oversized, "svg", "E5207", options: new DocxEditOptions { Quotas = new DocxPackageLimits(100, 100000, 8000) });
        AssertRejected(Diagram("width=\"1e30\" height=\"1e30\""), "svg", "E5206");
        AssertRejected(Diagram(), "svg", "E5206", "width 0px\nheight 1in\n");
        AssertRejected(Diagram(), "png", "E5203");
    }

    [Fact]
    public static async Task CanceledSvgHostDoesNotPublish()
    {
        var host = HostedCommandTests.CreateHost();
        host.Files["/asset.svg"] = Diagram();
        host.Files["/edit.patch"] = "docxpatch 1\nop insert-image-after\ntarget M.P0001\nasset /asset.svg\nend\n"u8.ToArray();
        using var canceled = new CancellationTokenSource();
        host.OnAssetRead = canceled.Cancel;
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => DocxCommand.RunAsync(["apply", "/input.docx", "/edit.patch", "-o", "/out.docx"], host, canceled.Token));
        Assert.Empty(host.Published);
    }

    [Fact]
    public static void ValidatorChecksTheVectorRelationship()
    {
        using var word = WordSvg();
        EditXml(word, "word/document.xml", doc => doc.Descendants(Asvg + "svgBlip").Single().SetAttributeValue(R + "embed", "missing"));
        word.Position = 0;
        Assert.Contains(new DocxEditor().Validate(word).Diagnostics, d => d.Code == "E9105");
    }

    internal static MemoryStream WordSvg(bool shared = false)
    {
        using var input = NewDocument();
        var word = Insert(input, Diagram(), "svg");
        EditXml(word, "word/document.xml", doc =>
        {
            doc.Descendants(A + "blip").Single().SetAttributeValue(R + "embed", "rPreview");
            if (!shared) return;
            var copy = new XElement(doc.Descendants(W + "p").Single(p => p.Descendants(W + "drawing").Any()));
            copy.Descendants(Wp + "docPr").Single().SetAttributeValue("id", "2");
            doc.Descendants(W + "sectPr").Single().AddBeforeSelf(copy);
        });
        EditXml(word, "word/_rels/document.xml.rels", doc => doc.Root!.Add(new XElement(Rel + "Relationship", new XAttribute("Id", "rPreview"), new XAttribute("Type", R.NamespaceName + "/image"), new XAttribute("Target", "media/preview.png"))));
        EditXml(word, "[Content_Types].xml", doc => doc.Root!.Add(new XElement(Ct + "Default", new XAttribute("Extension", "png"), new XAttribute("ContentType", "image/png"))));
        SetEntry(word, "word/media/preview.png", CreatePngBytes(2, 1));
        word.Position = 0;
        return word;
    }
}

internal static class ImageFormatTestTools
{
    internal static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    internal static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    internal static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    internal static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    internal static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
    internal static readonly XNamespace Ct = "http://schemas.openxmlformats.org/package/2006/content-types";
    internal static readonly XNamespace Asvg = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";

    internal static MemoryStream NewDocument()
    {
        using var empty = new MemoryStream();
        Assert.True(new DocxEditor().Create(empty).Success);
        return Apply(empty, "op replace-paragraph\ntarget M.P0001\ntext Unrelated content.\nend\n");
    }

    internal static MemoryStream Apply(MemoryStream input, string operations, DocxEditOptions? options = null)
    {
        string patch = "docxpatch 1\n" + operations;
        var editor = new DocxEditor();
        input.Position = 0;
        var check = editor.Check(input, new StringReader(patch), options ?? new());
        Assert.True(check.Success, string.Join("; ", check.Diagnostics));
        input.Position = 0;
        var output = new MemoryStream();
        var applied = editor.Apply(input, new StringReader(patch), output, options ?? new());
        Assert.True(applied.Success, string.Join("; ", applied.Diagnostics));
        output.Position = 0;
        return output;
    }

    internal static MemoryStream Insert(MemoryStream input, byte[] bytes, string extension, string fields = "", string target = "M.P0001") =>
        Apply(input, $"op insert-image-after\ntarget {target}\nasset asset.{extension}\n{fields}end\n",
            new() { AssetProvider = new MemoryAssetProvider($"asset.{extension}", bytes, null, null) });

    internal static MemoryStream Replace(MemoryStream input, byte[] bytes, string extension, string target = "M.I0001") =>
        Apply(input, $"op replace-image\ntarget {target}\nasset asset.{extension}\nend\n",
            new() { AssetProvider = new MemoryAssetProvider($"asset.{extension}", bytes, null, null) });

    internal static void AssertImage(MemoryStream document, byte[] bytes, string type, string extension, long? width = null, long? height = null)
    {
        document.Position = 0;
        var read = new DocxEditor().Read(document);
        Assert.True(read.Success, string.Join("; ", read.Diagnostics));
        DocxImageInfo image = Assert.Single(read.Images);
        Assert.Equal(type, image.ContentType);
        if (width is not null) Assert.Equal(width, image.WidthEmu);
        if (height is not null) Assert.Equal(height, image.HeightEmu);
        document.Position = 0;
        var extracted = new DocxEditor().ExtractMedia(document);
        Assert.True(extracted.Success);
        DocxMediaFile file = Assert.Single(extracted.Files);
        Assert.Equal(bytes, file.Content);
        Assert.Equal(type, file.ContentType);
        Assert.EndsWith(extension, file.FileName, StringComparison.Ordinal);
    }

    internal static void AssertRejected(byte[] bytes, string extension, string code, string fields = "", DocxEditOptions? options = null)
    {
        using var input = NewDocument();
        var editor = new DocxEditor();
        string patch = $"docxpatch 1\nop insert-image-after\ntarget M.P0001\nasset asset.{extension}\n{fields}end\n";
        options = new DocxEditOptions
        {
            Quotas = options?.Quotas ?? new DocxEditOptions().Quotas,
            AssetProvider = new MemoryAssetProvider($"asset.{extension}", bytes, null, null)
        };
        input.Position = 0;
        var check = editor.Check(input, new StringReader(patch), options);
        input.Position = 0;
        using var output = new MemoryStream();
        var apply = editor.Apply(input, new StringReader(patch), output, options);
        Assert.False(check.Success);
        Assert.False(apply.Success);
        Assert.Contains(check.Diagnostics, d => d.Code == code);
        Assert.Equal(check.Diagnostics.Select(d => d.Code), apply.Diagnostics.Select(d => d.Code));
        Assert.Equal(0, output.Length);
    }

    internal static async Task AssertHosted(byte[] bytes, string extension, string type)
    {
        var host = HostedCommandTests.CreateHost();
        host.Files[$"/asset.{extension}"] = bytes;
        host.Files["/edit.patch"] = Encoding.UTF8.GetBytes($"docxpatch 1\nop insert-image-after\ntarget M.P0001\nas image\nasset /asset.{extension}\nend\nop set-image-alt\ntarget @image\nalt Updated\nexpect-content-type {type}\nend\n");
        Assert.Equal(0, await DocxCommand.RunAsync(["check", "/input.docx", "/edit.patch"], host, CancellationToken.None));
        Assert.Equal(0, await DocxCommand.RunAsync(["apply", "/input.docx", "/edit.patch", "-o", "/out.docx"], host, CancellationToken.None));
        Assert.Equal(0, await DocxCommand.RunAsync(["media", "/out.docx", "--extract", "/media"], host, CancellationToken.None));
        Assert.Equal(bytes, Assert.Single(host.Files, f => f.Key.StartsWith("/media/", StringComparison.Ordinal)).Value);
    }

    internal static void EditXml(MemoryStream package, string name, Action<XDocument> edit)
    {
        package.Position = 0;
        var doc = XDocument.Parse(ReadEntry(package, name));
        edit(doc);
        SetEntry(package, name, Encoding.UTF8.GetBytes(doc.ToString()));
    }

    internal static void SetEntry(MemoryStream package, string name, byte[] bytes)
    {
        package.Position = 0;
        using (var archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true))
        {
            archive.GetEntry(name)?.Delete();
            using var stream = archive.CreateEntry(name).Open();
            stream.Write(bytes);
        }
        package.Position = 0;
    }

    internal static void AddHeaderUse(MemoryStream package)
    {
        var main = XDocument.Parse(ReadDocumentXml(package));
        var paragraph = new XElement(main.Descendants(W + "p").Single(p => p.Descendants(W + "drawing").Any()));
        SetEntry(package, "word/header1.xml", Encoding.UTF8.GetBytes(new XElement(W + "hdr", paragraph).ToString()));
        var rels = XDocument.Parse(ReadEntry(package, "word/_rels/document.xml.rels"));
        var imageRels = rels.Root!.Elements().Where(r => ((string?)r.Attribute("Type"))?.EndsWith("/image", StringComparison.Ordinal) == true);
        SetEntry(package, "word/_rels/header1.xml.rels", Encoding.UTF8.GetBytes(new XElement(Rel + "Relationships", imageRels.Select(r => new XElement(r))).ToString()));
        EditXml(package, "[Content_Types].xml", doc => doc.Root!.Add(new XElement(Ct + "Override", new XAttribute("PartName", "/word/header1.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"))));
        EditXml(package, "word/_rels/document.xml.rels", doc => doc.Root!.Add(new XElement(Rel + "Relationship", new XAttribute("Id", "rHeader"), new XAttribute("Type", R.NamespaceName + "/header"), new XAttribute("Target", "header1.xml"))));
        EditXml(package, "word/document.xml", doc => doc.Descendants(W + "sectPr").Single().AddFirst(new XElement(W + "headerReference", new XAttribute(W + "type", "default"), new XAttribute(R + "id", "rHeader"))));
    }
}
