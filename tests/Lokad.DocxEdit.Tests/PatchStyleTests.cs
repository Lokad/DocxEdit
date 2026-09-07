using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml.Linq;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

public static class PatchStyleTests
{

    [Fact]
    public static void ApplySetStyleResolvesParagraphStyleByDisplayName()
    {
        using MemoryStream input = CreateDocxWithStyles("""
              <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
              <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/></w:style>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-style
            target M.P0001
            style Heading 2
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Contains("w:val=\"Heading2\"", ReadDocumentXml(output), StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckSetStyleFailsWhenRequestedStyleIsMissing()
    {
        using MemoryStream input = CreateDocxWithStyles("""
              <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-style
            target M.P0001
            style Missing Style
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E7101");
    }

    [Fact]
    public static void CheckSetStyleFailsWhenDisplayNameIsAmbiguous()
    {
        using MemoryStream input = CreateDocxWithStyles("""
              <w:style w:type="paragraph" w:styleId="Custom1"><w:name w:val="Custom"/></w:style>
              <w:style w:type="paragraph" w:styleId="Custom2"><w:name w:val="Custom"/></w:style>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-style
            target M.P0001
            style Custom
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E7102");
    }

    [Fact]
    public static void ApplySetStyleCanEditHeaderParagraph()
    {
        using MemoryStream input = CreateDocxWithHeaderFooterAndStyles();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-style
            target H001.P0001
            style Heading 2
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Contains("w:val=\"Heading2\"", ReadEntry(output, "word/header1.xml"), StringComparison.Ordinal);
        output.Position = 0;
        DocxParagraphInfo paragraph = new DocxEditor().Read(output, new DocxReadOptions { IncludeHeadersFooters = true }).Paragraphs.Single(paragraph => paragraph.Id.ToWireValue() == "H001.P0001");
        Assert.Equal(2, paragraph.HeadingLevel);
    }
}
