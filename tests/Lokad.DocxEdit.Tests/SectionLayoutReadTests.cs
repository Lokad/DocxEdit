using System.Text.Json;
using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

public static class SectionLayoutReadTests
{
    [Fact]
    public static void ReadsEachSectionsStoredLayoutInTextAndJson()
    {
        using var input = CreateDocxWithBody("""
            <w:p><w:pPr><w:sectPr>
              <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
              <w:pgMar w:top="-120" w:bottom="0" w:left="800" w:right="900" w:header="100" w:footer="200" w:gutter="300"/>
            </w:sectPr></w:pPr></w:p>
            <w:p/>
            <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440"/></w:sectPr>
            """);
        DocxReadResult read = new DocxEditor().Read(input);
        Assert.True(read.Success);
        Assert.Equal(2, read.Sections.Count);
        DocxSectionInfo first = read.Sections[0];
        Assert.Equal(15840, first.PageWidthTwips);
        Assert.Equal(12240, first.PageHeightTwips);
        Assert.Equal(-120, first.MarginTopTwips);
        Assert.Equal(0, first.MarginBottomTwips);
        Assert.Equal(800, first.MarginLeftTwips);
        Assert.Equal(900, first.MarginRightTwips);
        Assert.Equal(100, first.HeaderDistanceTwips);
        Assert.Equal(200, first.FooterDistanceTwips);
        Assert.Equal(300, first.GutterTwips);
        Assert.Equal(11906, read.Sections[1].PageWidthTwips);
        Assert.Equal(16838, read.Sections[1].PageHeightTwips);
        Assert.Null(read.Sections[1].MarginLeftTwips);
        Assert.Contains("page-width-twips=15840 page-height-twips=12240", read.Text);
        Assert.Contains("margin-top-twips=-120", read.Text);
        Assert.Contains("header-distance-twips=100 footer-distance-twips=200 gutter-twips=300", read.Text);

        using JsonDocument json = JsonDocument.Parse(JsonSerializer.Serialize(read, DocxJson.CreateOptions(false)));
        JsonElement sections = json.RootElement.GetProperty("Sections");
        Assert.Equal(15840, sections[0].GetProperty("PageWidthTwips").GetInt32());
        Assert.Equal(-120, sections[0].GetProperty("MarginTopTwips").GetInt32());
        Assert.Equal(JsonValueKind.Null, sections[1].GetProperty("MarginLeftTwips").ValueKind);
    }

    [Theory]
    [InlineData("<w:sectPr/>")]
    [InlineData("<w:sectPr><w:pgSz w:w=\"invalid\" w:h=\"2147483648\"/><w:pgMar w:left=\"\" w:top=\"invalid\"/></w:sectPr>")]
    public static void MissingOrMalformedMeasurementsRemainUnknown(string sectionXml)
    {
        using var input = CreateDocxWithBody("<w:p/>" + sectionXml);
        DocxReadResult read = new DocxEditor().Read(input);
        Assert.True(read.Success);
        DocxSectionInfo section = Assert.Single(read.Sections);
        Assert.Null(section.PageWidthTwips);
        Assert.Null(section.PageHeightTwips);
        Assert.Null(section.MarginTopTwips);
        Assert.Null(section.MarginBottomTwips);
        Assert.Null(section.MarginLeftTwips);
        Assert.Null(section.MarginRightTwips);
        Assert.Null(section.HeaderDistanceTwips);
        Assert.Null(section.FooterDistanceTwips);
        Assert.Null(section.GutterTwips);
        Assert.DoesNotContain("page-width-twips", read.Text);
    }

    [Fact]
    public static void ExistingOrientationEditReportsSwappedDimensions()
    {
        var editor = new DocxEditor();
        using var input = CreateDocxWithBody("""
            <w:p/>
            <w:sectPr>
              <w:pgSz w:w="11906" w:h="16838"/>
              <w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/>
            </w:sectPr>
            """);
        using var patch = new StringReader("""
            docxpatch 1
            op set-section-orientation
            target M.S0001
            orientation landscape
            end
            """);
        using var output = new MemoryStream();
        Assert.True(editor.Apply(input, patch, output).Success);
        output.Position = 0;
        DocxSectionInfo section = Assert.Single(editor.Read(output).Sections);
        Assert.Equal(DocxOrientation.Landscape, section.Orientation);
        Assert.Equal(16838, section.PageWidthTwips);
        Assert.Equal(11906, section.PageHeightTwips);
        Assert.Equal(1440, section.MarginLeftTwips);
    }
}
