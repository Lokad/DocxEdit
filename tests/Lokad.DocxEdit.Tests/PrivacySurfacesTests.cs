using System.IO.Compression;
using System.Text;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// PLAN B02: low-text surfaces drop body text in text and structured output,
// while IDs, counts, and structural metadata are retained.
[Collection("ConsoleCli")]
public static class PrivacySurfacesTests
{
    private const string ParaMarker = "PXPARAalpha09";
    private const string CellMarker = "PXCELLbeta09";
    private const string FieldResultMarker = "PXFIELDgamma09";
    private const string FieldCodeMarker = "PXCODEfield09";
    private const string ImageDescrMarker = "PXIMGdelta09";
    private const string BookmarkMarker = "PXBMname09";
    private const string HyperlinkMarker = "pxuri-marker09";
    private const string CommentMarker = "PXCOMMENTepsilon09";

    [Fact]
    public static void ReadMaxTextZeroDropsBodyTextButRetainsMetadata()
    {
        using MemoryStream input = CreatePrivacyDocx();
        DocxReadResult result = new DocxEditor().Read(input, new DocxReadOptions { MaxText = 0 });
        Assert.True(result.Success);
        foreach (DocxParagraphInfo paragraph in result.Paragraphs)
        {
            Assert.Equal(string.Empty, paragraph.Text);
            foreach (DocxRunInfo run in paragraph.Runs)
            {
                Assert.Equal(string.Empty, run.Text);
            }
        }
        foreach (DocxTableInfo table in result.Tables)
        {
            foreach (DocxTableCellInfo cell in table.Cells)
            {
                Assert.Equal(string.Empty, cell.Text);
            }
        }
        foreach (DocxFieldInfo field in result.Fields)
        {
            Assert.Equal(string.Empty, field.CachedResultText);
        }
        foreach (DocxImageInfo image in result.Images)
        {
            Assert.True(image.Description is null || image.Description.Length == 0);
        }
        Assert.DoesNotContain(ParaMarker, result.Text, StringComparison.Ordinal);
        Assert.DoesNotContain(CellMarker, result.Text, StringComparison.Ordinal);
        Assert.DoesNotContain(FieldResultMarker, result.Text, StringComparison.Ordinal);
        Assert.DoesNotContain(ImageDescrMarker, result.Text, StringComparison.Ordinal);
        Assert.Contains(result.Bookmarks, static b => b.Name.Contains(BookmarkMarker, StringComparison.Ordinal));
        Assert.Contains(result.Fields, static f => f.Code.Contains(FieldCodeMarker, StringComparison.Ordinal));
        Assert.Contains(result.Hyperlinks, static h => (h.Uri ?? string.Empty).Contains(HyperlinkMarker, StringComparison.Ordinal));
    }

    [Fact]
    public static void ReadSummaryPresetDropsBodyText()
    {
        using MemoryStream input = CreatePrivacyDocx();
        DocxReadResult result = new DocxEditor().Read(input, DocxPrivacyPresets.ReadSummary);
        Assert.True(result.Success);
        Assert.All(result.Paragraphs, static p => Assert.Equal(string.Empty, p.Text));
        Assert.DoesNotContain(ParaMarker, result.Text, StringComparison.Ordinal);
        Assert.DoesNotContain(CellMarker, result.Text, StringComparison.Ordinal);
        string summary = DocxPrivacyPresets.RenderReadSummary(result);
        Assert.DoesNotContain(ParaMarker, summary, StringComparison.Ordinal);
        Assert.Contains("paragraphs count=", summary, StringComparison.Ordinal);
    }

    [Fact]
    public static void ContextDefaultDropsTextAndOptInRestoresIt()
    {
        using MemoryStream input = CreatePrivacyDocx();
        DocxContextResult metadata = new DocxEditor().Context(input, "M.P0001", DocxPrivacyPresets.ContextMetadataOnly);
        Assert.True(metadata.Success);
        Assert.All(metadata.Items, static item => Assert.Equal(string.Empty, item.Text));
        Assert.DoesNotContain(ParaMarker, metadata.Text, StringComparison.Ordinal);
        input.Position = 0;
        DocxContextResult withText = new DocxEditor().Context(input, "M.P0001", new DocxContextOptions { Radius = 0, MaxText = 4000 });
        Assert.True(withText.Success);
        Assert.Contains(ParaMarker, withText.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void RenderContextMetadataStripsTextFromBearingResult()
    {
        using MemoryStream input = CreatePrivacyDocx();
        DocxContextResult bearing = new DocxEditor().Context(input, "M.P0001", new DocxContextOptions { Radius = 0, MaxText = 4000 });
        Assert.True(bearing.Success);
        Assert.Contains(ParaMarker, bearing.Text, StringComparison.Ordinal);
        string redacted = DocxPrivacyPresets.RenderContextMetadata(bearing);
        Assert.DoesNotContain(ParaMarker, redacted, StringComparison.Ordinal);
        Assert.Contains("M.P0001", redacted, StringComparison.Ordinal);
    }

    [Fact]
    public static void RenderChangesMarkupStripsCommentText()
    {
        using MemoryStream input = CreateCommentDocx();
        DocxChangesResult bearing = new DocxEditor().Changes(input, new DocxChangesOptions { IncludeCommentText = true, MaxCommentText = 4000 });
        Assert.True(bearing.Success);
        Assert.Contains(bearing.Changes, static c => c.CommentTextSnippet != null && c.CommentTextSnippet.Contains(CommentMarker, StringComparison.Ordinal));
        string redacted = DocxPrivacyPresets.RenderChangesMarkup(bearing);
        Assert.DoesNotContain(CommentMarker, redacted, StringComparison.Ordinal);
        DocxChangesResult markupOnly = new DocxEditor().Changes(input, DocxPrivacyPresets.ChangesMarkupOnly);
        Assert.True(markupOnly.Success);
        Assert.DoesNotContain(markupOnly.Changes, static c => c.CommentTextSnippet != null && c.CommentTextSnippet.Contains(CommentMarker, StringComparison.Ordinal));
    }

    [Fact]
    public static void CliSummaryTextOmitsBodyText()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = WritePrivacyDocx(temp.Path);
        CliOutcome result = RunCli("read", input, "--summary");
        Assert.Equal(0, result.ExitCode);
        Assert.DoesNotContain(ParaMarker, result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain(CellMarker, result.Output, StringComparison.Ordinal);
        Assert.Contains("paragraphs count=", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliSummaryJsonOmitsPayloadText()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = WritePrivacyDocx(temp.Path);
        CliOutcome result = RunCli("read", input, "--summary", "--json");
        Assert.Equal(0, result.ExitCode);
        Assert.DoesNotContain(ParaMarker, result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain(CellMarker, result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain(FieldResultMarker, result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain(ImageDescrMarker, result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public static void CliMaxTextZeroJsonOmitsBodyText()
    {
        using TempDirectory temp = TempDirectory.Create();
        string input = WritePrivacyDocx(temp.Path);
        CliOutcome result = RunCli("read", input, "--max-text", "0", "--json");
        Assert.Equal(0, result.ExitCode);
        Assert.DoesNotContain(ParaMarker, result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain(CellMarker, result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain(FieldResultMarker, result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain(ImageDescrMarker, result.Output, StringComparison.Ordinal);
    }

    private static MemoryStream CreatePrivacyDocx()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", RootRelsXml());
            AddEntry(archive, "word/_rels/document.xml.rels", DocumentRelsXml());
            AddEntry(archive, "word/document.xml", DocumentXml());
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
            + "<Default Extension=\"png\" ContentType=\"image/png\"/>"
            + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
            + "</Types>";
    }

    private static string RootRelsXml()
    {
        return "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rDocument\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
            + "</Relationships>";
    }

    private static string DocumentRelsXml()
    {
        return "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rHyperlink\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://pxuri-marker09.example/test\" TargetMode=\"External\"/>"
            + "<Relationship Id=\"rImage\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image1.png\"/>"
            + "</Relationships>";
    }

    private static string DocumentXml()
    {
        string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        string r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        string wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        string a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        string pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        return "<w:document xmlns:w=\"" + w + "\" xmlns:r=\"" + r + "\" xmlns:wp=\"" + wp + "\" xmlns:a=\"" + a + "\" xmlns:pic=\"" + pic + "\">"
            + "<w:body>"
            + "<w:p><w:bookmarkStart w:id=\"11\" w:name=\"PXBMname09\"/><w:r><w:t>" + ParaMarker + "</w:t></w:r><w:bookmarkEnd w:id=\"11\"/></w:p>"
            + "<w:p><w:fldSimple w:instr=\" REF PXCODEfield09 \\h \"><w:r><w:t>" + FieldResultMarker + "</w:t></w:r></w:fldSimple></w:p>"
            + "<w:p><w:hyperlink r:id=\"rHyperlink\"><w:r><w:t>Link</w:t></w:r></w:hyperlink></w:p>"
            + "<w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"914400\" cy=\"457200\"/><wp:docPr id=\"1\" name=\"Picture 1\" descr=\"" + ImageDescrMarker + "\"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed=\"rImage\"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
            + "<w:tbl><w:tblPr><w:tblStyle w:val=\"TableGrid\"/></w:tblPr><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>" + CellMarker + "</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
            + "</w:body>"
            + "</w:document>";
    }

    private static MemoryStream CreateCommentDocx()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", CommentContentTypesXml());
            AddEntry(archive, "_rels/.rels", RootRelsXml());
            AddEntry(archive, "word/_rels/document.xml.rels", CommentDocumentRelsXml());
            AddEntry(archive, "word/document.xml", CommentDocumentXml());
            AddEntry(archive, "word/comments.xml", CommentCommentsXml());
        }
        stream.Position = 0;
        return stream;
    }

    private static string CommentContentTypesXml()
    {
        return "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
            + "<Override PartName=\"/word/comments.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml\"/>"
            + "</Types>";
    }

    private static string CommentDocumentRelsXml()
    {
        return "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rComments\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"comments.xml\"/>"
            + "</Relationships>";
    }

    private static string CommentDocumentXml()
    {
        string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        return "<w:document xmlns:w=\"" + w + "\">"
            + "<w:body>"
            + "<w:p><w:commentRangeStart w:id=\"7\"/><w:r><w:t>Commented</w:t></w:r><w:commentRangeEnd w:id=\"7\"/><w:r><w:commentReference w:id=\"7\"/></w:r></w:p>"
            + "</w:body>"
            + "</w:document>";
    }

    private static string CommentCommentsXml()
    {
        string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        return "<w:comments xmlns:w=\"" + w + "\">"
            + "<w:comment w:id=\"7\" w:author=\"Reviewer\"><w:p><w:r><w:t>" + CommentMarker + "</w:t></w:r></w:p></w:comment>"
            + "</w:comments>";
    }

    private static string WritePrivacyDocx(string directory)
    {
        string path = Path.Combine(directory, "privacy.docx");
        using MemoryStream source = CreatePrivacyDocx();
        using FileStream file = File.Create(path);
        source.CopyTo(file);
        return path;
    }

    private static CliOutcome RunCli(params string[] args)
    {
        var captured = DocxTestFixtures.CaptureCliOutput(null, args);
        return new CliOutcome(captured.ExitCode, captured.Output, captured.Error);
    }



    private sealed record CliOutcome(int ExitCode, string Output, string Error);
}
