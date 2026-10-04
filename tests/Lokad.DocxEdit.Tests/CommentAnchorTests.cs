using System.Xml.Linq;
using static Lokad.DocxEdit.Tests.DocxTestFixtures;
using static Lokad.DocxEdit.Tests.CommentLifecycleTests;

namespace Lokad.DocxEdit.Tests;

public static class CommentAnchorTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string Cell = "M.T0001.R01.C01";
    private const string CellBody = """
        <w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid><w:tr><w:tc>
        <w:p><w:r><w:t>First estimate.</w:t></w:r></w:p>
        <w:p><w:r><w:t>Second estimate.</w:t></w:r></w:p>
        </w:tc></w:tr></w:tbl>
        """;

    public static TheoryData<string> ProtectedNeighbors => new()
    {
        "<w:bookmarkStart w:id=\"8\" w:name=\"note\"/><w:r><w:t>Inside</w:t></w:r><w:bookmarkEnd w:id=\"8\"/>",
        "<w:hyperlink w:anchor=\"note\"><w:r><w:t>Inside</w:t></w:r></w:hyperlink>",
        "<w:fldSimple w:instr=\"DATE\"><w:r><w:t>Inside</w:t></w:r></w:fldSimple>",
        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText>DATE</w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>Inside</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r>",
        "<w:ins w:id=\"8\" w:author=\"Reviewer\" w:date=\"2026-10-04T12:00:00Z\"><w:r><w:t>Inside</w:t></w:r></w:ins>",
        "<w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t>Inside</w:t></w:r></w:sdtContent></w:sdt>",
        "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>x</m:t></m:r></m:oMath>",
    };

    [Theory]
    [MemberData(nameof(ProtectedNeighbors))]
    public static void CommentBesideProtectedContentKeepsExactBoundaries(string protectedXml)
    {
        using MemoryStream input = CreateDocxWithBody("<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Before</w:t></w:r>" + protectedXml + "<w:r><w:t>After</w:t></w:r></w:p>");
        using MemoryStream output = Apply(input, """
            docxpatch 1
            op add-comment
            target M.P0001
            anchor-text Before
            text First note.
            end
            op add-comment
            target M.P0001
            anchor-text After
            text Second note.
            end
            """);
        XDocument xml = XDocument.Parse(ReadDocumentXml(output));
        XElement paragraph = xml.Descendants(W + "p").First();
        Assert.Equal("Before", AnchoredText(paragraph, "0"));
        Assert.Equal("After", AnchoredText(paragraph, "1"));
        XElement firstEnd = paragraph.Elements(W + "commentRangeEnd").First();
        XElement secondStart = paragraph.Elements(W + "commentRangeStart").Last();
        var between = firstEnd.ElementsAfterSelf().TakeWhile(e => e != secondStart)
            .Where(e => !e.Descendants(W + "commentReference").Any()).ToArray();
        XElement[] original = XElement.Parse($"<w:p xmlns:w=\"{W}\">{protectedXml}</w:p>").Elements().ToArray();
        Assert.Equal(original.Length, between.Length);
        for (int i = 0; i < original.Length; i++) Assert.True(XNode.DeepEquals(original[i], between[i]));
        Assert.Single(paragraph.Descendants(W + "b"));
        output.Position = 0;
        Assert.True(new DocxEditor().Validate(output).Success);
    }

    [Theory]
    [MemberData(nameof(ProtectedNeighbors))]
    public static void CommentCrossingProtectedContentRefuses(string protectedXml)
    {
        using MemoryStream input = CreateDocxWithBody("<w:p><w:r><w:t>Before</w:t></w:r>" + protectedXml + "<w:r><w:t>After</w:t></w:r></w:p>");
        string anchor = protectedXml.Contains("m:oMath", StringComparison.Ordinal) ? "BeforeAfter" : "BeforeInsideAfter";
        var result = new DocxEditor().Check(input, new StringReader($"docxpatch 1\nop add-comment\ntarget M.P0001\nanchor-text {anchor}\ntext Note\nend\n"));
        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, d => d.Code is "E4305" or "E4317");
    }

    [Fact]
    public static void SelectedCommentRefusesStoryRangeEnteringFromEarlierParagraph()
    {
        using MemoryStream input = CreateDocxWithBody("""
            <w:p><w:bookmarkStart w:id="8" w:name="note"/><w:r><w:t>Start</w:t></w:r></w:p>
            <w:p><w:r><w:t>Middle</w:t></w:r></w:p>
            <w:p><w:r><w:t>End</w:t></w:r><w:bookmarkEnd w:id="8"/></w:p>
            """);
        var result = new DocxEditor().Check(input, new StringReader("docxpatch 1\nop add-comment\ntarget M.P0002\nanchor-text Middle\ntext Note\nend\n"));
        Assert.Contains(result.Diagnostics, d => d.Code == "E4305");
    }

    [Fact]
    public static void SelectedCommentDoesNotSwallowInvisibleRunContent()
    {
        using MemoryStream input = CreateDocxWithBody("<w:p><w:r><w:t>Before</w:t><w:sym w:font=\"Symbol\" w:char=\"F020\"/><w:t>After</w:t></w:r></w:p>");
        var result = new DocxEditor().Check(input, new StringReader("docxpatch 1\nop add-comment\ntarget M.P0001\nanchor-text BeforeAfter\ntext Note\nend\n"));
        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, d => d.Code == "E4317");
    }

    [Fact]
    public static void CellWithNestedTableRefusesWithoutOutput()
    {
        string nested = CellBody.Replace("<w:p><w:r><w:t>First estimate.</w:t></w:r></w:p>", CellBody, StringComparison.Ordinal);
        using MemoryStream input = CreateDocxWithBody(nested);
        using var output = new MemoryStream();
        var result = new DocxEditor().Apply(input, new StringReader($"docxpatch 1\nop add-comment\ntarget {Cell}\ntext Note\nend\n"), output);
        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, d => d.Code == "E4317");
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public static void WholeCellCommentSpansAllDirectParagraphs()
    {
        using MemoryStream input = CreateDocxWithBody(CellBody);
        using MemoryStream output = Apply(input, $"docxpatch 1\nop add-comment\ntarget {Cell}\nexpect-text First estimate.Second estimate.\ntext Review cell.\nend\n");
        XDocument xml = XDocument.Parse(ReadDocumentXml(output));
        XElement cell = xml.Descendants(W + "tc").Single();
        Assert.NotNull(cell.Elements(W + "p").First().Element(W + "commentRangeStart"));
        Assert.NotNull(cell.Elements(W + "p").Last().Element(W + "commentRangeEnd"));
        AssertValid(output);
        output.Position = 0;
        Assert.Equal(Cell, Assert.Single(new DocxEditor().Changes(output).CommentSummary).AnchorTargetId);
    }

    [Fact]
    public static void CellOccurrenceSelectsAcrossParagraphs()
    {
        using MemoryStream input = CreateDocxWithBody(CellBody);
        using MemoryStream output = Apply(input, $"docxpatch 1\nop add-comment\ntarget {Cell}\nanchor-text estimate\noccurrence 2\ntext Review second estimate.\nend\n", TrackChangesMode.Require);
        XDocument xml = XDocument.Parse(ReadDocumentXml(output));
        XElement cell = xml.Descendants(W + "tc").Single();
        Assert.Null(cell.Elements(W + "p").First().Element(W + "commentRangeStart"));
        Assert.Equal("estimate", AnchoredText(cell.Elements(W + "p").Last(), "0"));
        AssertValid(output);
    }

    [Theory]
    [InlineData("anchor-text estimate", "E1202")]
    [InlineData("anchor-text estimate\noccurrence 3", "E4203")]
    [InlineData("anchor-text estimate.Second", "E4203")]
    [InlineData("expect-text stale", "E3201")]
    public static void CellCommentRejectsAmbiguousMissingCrossParagraphOrStaleAnchor(string fields, string code)
    {
        using MemoryStream input = CreateDocxWithBody(CellBody);
        var result = new DocxEditor().Check(input, new StringReader($"docxpatch 1\nop add-comment\ntarget {Cell}\n{fields}\ntext Note\nend\n"));
        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, d => d.Code == code);
    }

    [Theory]
    [InlineData("M.T0001.R01.C01", true)]
    [InlineData("M.T0001.MG0001", true)]
    [InlineData("M.T0001.R02.C01", false)]
    public static void MergedCellCommentRequiresRoot(string target, bool success)
    {
        using MemoryStream input = CreateDocxWithBody("""
            <w:tbl><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>
            <w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Root</w:t></w:r></w:p></w:tc></w:tr>
            <w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>
            """);
        var result = new DocxEditor().Check(input, new StringReader($"docxpatch 1\nop add-comment\ntarget {target}\ntext Note\nend\n"));
        Assert.Equal(success, result.Success);
        if (!success) Assert.Contains(result.Diagnostics, d => d.Code == "E4301");
    }

    private static string AnchoredText(XElement paragraph, string id)
    {
        XElement start = paragraph.Elements(W + "commentRangeStart").Single(e => (string?)e.Attribute(W + "id") == id);
        return string.Concat(start.ElementsAfterSelf().TakeWhile(e => e.Name != W + "commentRangeEnd" || (string?)e.Attribute(W + "id") != id)
            .DescendantsAndSelf(W + "t").Select(e => e.Value));
    }
}
