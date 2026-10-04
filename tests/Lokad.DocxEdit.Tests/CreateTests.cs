using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;

namespace Lokad.DocxEdit.Tests;

public static class CreateTests
{
    [Theory]
    [InlineData(DocxPaperSize.A4, DocxOrientation.Portrait, 11906, 16838)]
    [InlineData(DocxPaperSize.A4, DocxOrientation.Landscape, 16838, 11906)]
    [InlineData(DocxPaperSize.Letter, DocxOrientation.Portrait, 12240, 15840)]
    [InlineData(DocxPaperSize.Letter, DocxOrientation.Landscape, 15840, 12240)]
    public static void CreatesValidEditablePackage(DocxPaperSize paper, DocxOrientation orientation, int width, int height)
    {
        var editor = new DocxEditor();
        using var output = new MemoryStream();
        DocxCreateResult result = editor.Create(output, new DocxCreateOptions { PaperSize = paper, Orientation = orientation });
        Assert.True(result.Success, string.Join("; ", result.Diagnostics));
        Assert.Empty(result.Diagnostics);
        Assert.Equal(paper, result.PaperSize);
        Assert.Equal(orientation, result.Orientation);
        Assert.True(output.CanWrite);

        output.Position = 0;
        DocxValidateResult validation = editor.Validate(output);
        Assert.True(validation.Success);
        Assert.Empty(validation.Diagnostics);
        output.Position = 0;
        DocxReadResult read = editor.Read(output);
        Assert.True(read.Success);
        DocxParagraphInfo paragraph = Assert.Single(read.Paragraphs);
        Assert.Equal("M.P0001", paragraph.Id.ToWireValue());
        Assert.Empty(paragraph.Text);
        Assert.Empty(read.Tables);
        DocxSectionInfo section = Assert.Single(read.Sections);
        Assert.Equal("M.S0001", section.Id.ToWireValue());
        Assert.Equal(orientation, section.Orientation);
        Assert.Equal(width, section.PageWidthTwips);
        Assert.Equal(height, section.PageHeightTwips);
        Assert.Equal(1440, section.MarginTopTwips);
        Assert.Equal(1440, section.MarginBottomTwips);
        Assert.Equal(1440, section.MarginLeftTwips);
        Assert.Equal(1440, section.MarginRightTwips);
        Assert.Equal(720, section.HeaderDistanceTwips);
        Assert.Equal(720, section.FooterDistanceTwips);
        Assert.Equal(0, section.GutterTwips);
        Assert.Equal(1, section.Columns);

        output.Position = 0;
        using (var zip = new ZipArchive(output, ZipArchiveMode.Read, leaveOpen: true))
        {
            using Stream document = zip.GetEntry("word/document.xml")!.Open();
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XElement body = XDocument.Load(document).Root!.Element(w + "body")!;
            Assert.Equal(new[] { w + "p", w + "sectPr" }, body.Elements().Select(e => e.Name));
            XElement pageSize = body.Element(w + "sectPr")!.Element(w + "pgSz")!;
            Assert.Equal(width, (int)pageSize.Attribute(w + "w")!);
            Assert.Equal(height, (int)pageSize.Attribute(w + "h")!);
            Assert.Equal(orientation.ToWireValue(), (string?)pageSize.Attribute(w + "orient"));
        }

        output.Position = 0;
        using var patch = new StringReader("""
            docxpatch 1
            op replace-paragraph
            target M.P0001
            text First paragraph
            end
            op insert-after
            target M.P0001
            text First heading
            style Heading1
            end
            """);
        using var edited = new MemoryStream();
        DocxApplyResult applied = editor.Apply(output, patch, edited);
        Assert.True(applied.Success, string.Join("; ", applied.Diagnostics));
        edited.Position = 0;
        DocxReadResult after = editor.Read(edited);
        Assert.Equal(new[] { "First paragraph", "First heading" }, after.Paragraphs.Select(p => p.Text));
        Assert.Equal(1, after.Paragraphs[1].HeadingLevel);
        Assert.Equal(section, Assert.Single(after.Sections));
    }

    [Fact]
    public static void DefaultsAndStylesAreExplicitAndJsonUsesWireValues()
    {
        using var output = new MemoryStream();
        var editor = new DocxEditor();
        DocxCreateResult result = editor.Create(output);
        Assert.True(result.Success);
        Assert.Equal(DocxPaperSize.A4, result.PaperSize);
        Assert.Equal(DocxOrientation.Portrait, result.Orientation);
        output.Position = 0;
        DocxStylesResult styles = editor.Styles(output);
        Assert.True(styles.Success);
        Assert.Equal(10, styles.Styles.Count);
        Assert.Contains(styles.Styles, s => s.StyleId == "Normal");
        for (int level = 1; level <= 9; level++)
        {
            Assert.Contains(styles.Styles, s => s.StyleId == $"Heading{level}");
        }
        JsonSerializerOptions jsonOptions = DocxJson.CreateOptions(false);
        string json = JsonSerializer.Serialize(result, jsonOptions);
        Assert.Contains("\"PaperSize\":\"a4\"", json);
        Assert.Contains("\"Orientation\":\"portrait\"", json);
        Assert.Equal(result.PaperSize, JsonSerializer.Deserialize<DocxCreateResult>(json, jsonOptions)!.PaperSize);
        Assert.Equal(DocxPaperSize.Letter, JsonSerializer.Deserialize<DocxPaperSize>("\"letter\"", jsonOptions));
        Assert.Throws<JsonException>(() => JsonSerializer.Deserialize<DocxPaperSize>("\"legal\"", jsonOptions));
    }

    [Theory]
    [InlineData(0, 100000, 100000)]
    [InlineData(100, 10, 100000)]
    [InlineData(100, 100000, 10)]
    public static void QuotaFailureDoesNotTouchOutput(int entries, long total, long part)
    {
        using var output = new MemoryStream();
        output.Write("existing"u8);
        long position = output.Position;
        DocxCreateResult result = new DocxEditor().Create(output, new DocxCreateOptions { Quotas = new(entries, total, part) });
        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, d => d.Severity == DocxSeverity.Error && d.Code == "E0001");
        Assert.Equal(position, output.Position);
        Assert.Equal("existing"u8.ToArray(), output.ToArray());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public static void OwnershipHoldsOnSuccessFailureAndCancellation(bool leaveOpen)
    {
        foreach (string outcome in new[] { "success", "quota", "cancel", "io", "enum" })
        {
            using var output = new OutputStream { FailWrites = outcome == "io" };
            var options = new DocxCreateOptions
            {
                LeaveOutputOpen = leaveOpen,
                PaperSize = outcome == "enum" ? (DocxPaperSize)99 : DocxPaperSize.A4,
                Quotas = outcome == "quota" ? new(0, 0, 0) : DocxPackageLimits.Default
            };
            var editor = new DocxEditor();
            if (outcome == "cancel")
            {
                Assert.Throws<OperationCanceledException>(() => editor.Create(output, options, new CancellationToken(true)));
                Assert.Empty(output.ToArray());
            }
            else if (outcome == "io") Assert.Throws<IOException>(() => editor.Create(output, options));
            else if (outcome == "enum") Assert.Throws<ArgumentOutOfRangeException>(() => editor.Create(output, options));
            else Assert.Equal(outcome == "success", editor.Create(output, options).Success);
            Assert.Equal(!leaveOpen, output.Disposed);
        }
    }

    [Fact]
    public static void SupportsNonSeekableOutputAndRejectsInvalidArguments()
    {
        var editor = new DocxEditor();
        using var output = new OutputStream { NonSeekable = true };
        Assert.True(editor.Create(output).Success);
        using var input = new MemoryStream(output.ToArray());
        Assert.True(editor.Validate(input).Success);
        Assert.Throws<ArgumentNullException>(() => editor.Create(null!));
        Assert.Throws<ArgumentNullException>(() => editor.Create(output, null!));
        using var readOnly = new MemoryStream([], writable: false);
        Assert.Throws<ArgumentException>(() => editor.Create(readOnly));
        Assert.Throws<ArgumentOutOfRangeException>(() => editor.Create(output, new DocxCreateOptions { Orientation = (DocxOrientation)99 }));
    }

    private sealed class OutputStream : MemoryStream
    {
        public bool NonSeekable { get; init; }
        public bool FailWrites { get; init; }
        public bool Disposed { get; private set; }
        public override bool CanSeek => !NonSeekable && base.CanSeek;
        public override long Position
        {
            get => NonSeekable ? throw new NotSupportedException() : base.Position;
            set { if (NonSeekable) throw new NotSupportedException(); base.Position = value; }
        }
        public override void Write(byte[] buffer, int offset, int count)
        {
            if (FailWrites) throw new IOException("Simulated output failure.");
            base.Write(buffer, offset, count);
        }
        public override void Write(ReadOnlySpan<byte> buffer)
        {
            if (FailWrites) throw new IOException("Simulated output failure.");
            base.Write(buffer);
        }
        protected override void Dispose(bool disposing) { Disposed = true; base.Dispose(disposing); }
    }
}
