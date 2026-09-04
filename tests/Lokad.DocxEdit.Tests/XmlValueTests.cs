using Lokad.DocxEdit.Ooxml;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Rendering;

namespace Lokad.DocxEdit.Tests;

public static class XmlValueTests
{
    [Theory]
    [InlineData("", "")]
    [InlineData("abc", "abc")]
    [InlineData("a\\b", "a\\\\b")]
    [InlineData("a\"b", "a\\\"b")]
    [InlineData("a\rb", "a\\rb")]
    [InlineData("a\nb", "a\\nb")]
    [InlineData("a\tb", "a\\tb")]
    public static void EscapeTextEscapesQuotingAndControlCharacters(string text, string expected)
    {
        Assert.Equal(expected, XmlValues.EscapeText(text));
    }

    [Fact]
    public static void TryReadIntReturnsNullForMissingValue()
    {
        Assert.Null(XmlValues.TryReadInt(null));
    }

    [Theory]
    [InlineData("", null)]
    [InlineData("12", 12)]
    [InlineData("-3", -3)]
    [InlineData(" 7 ", 7)]
    [InlineData("abc", null)]
    [InlineData("12.5", null)]
    public static void TryReadIntParsesOptionalIntegers(string? text, int? expected)
    {
        Assert.Equal(expected, XmlValues.TryReadInt(text));
    }

    [Fact]
    public static void RenderReadEscapesCarriageReturns()
    {
        var model = new DocxDocumentModel(
            [new DocxParagraphInfo("M.P0001", "main", "a\rb", null, null, [])],
            [], [], [], [], [], [], []);

        string output = TextRenderers.RenderRead(model, 4000);

        Assert.Contains("a\\rb", output, StringComparison.Ordinal);
        Assert.DoesNotContain("a\rb", output, StringComparison.Ordinal);
    }

    [Fact]
    public static void PublicStyleRenderingEscapesCarriageReturns()
    {
        var result = new DocxStylesResult
        {
            Success = true,
            Diagnostics = [],
            Styles = [new DocxStyleInfo("S1", "Name\rb", "paragraph", false)],
        };

        string output = DocxTextRenderer.RenderStyles(result);

        Assert.Contains("Name\\rb", output, StringComparison.Ordinal);
    }
}
