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

    [Fact]
    public static void ApplyInsertAfterWithStyleDisplayNameWritesStyleId()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="paragraph" w:styleId="CustomHeading"><w:name w:val="Example Heading"/></w:style>
            """, """
              <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text Inserted
            style Example Heading
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Contains("w:val=" + (char)34 + "CustomHeading" + (char)34, ReadDocumentXml(output), StringComparison.Ordinal);
        output.Position = 0;
        Assert.DoesNotContain(new DocxEditor().Validate(output).Diagnostics, static diagnostic => diagnostic.Code == "W9116");
    }

    [Fact]
    public static void ApplyReplaceParagraphWithStyleDisplayNameWritesStyleId()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="paragraph" w:styleId="CustomHeading"><w:name w:val="Example Heading"/></w:style>
            """, """
              <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-paragraph
            target M.P0001
            text Changed
            style Example Heading
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Contains("w:val=" + (char)34 + "CustomHeading" + (char)34, ReadDocumentXml(output), StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckInsertAfterWithUnknownStyleFailsBeforePublication()
    {
        using MemoryStream input = CreateDocxWithStyles("""
              <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text Inserted
            style No Such Style
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E7101");
    }

    [Fact]
    public static void ApplySetTableStyleResolvesTableStyleByDisplayName()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="table" w:styleId="TableGridCustom"><w:name w:val="Custom Grid"/></w:style>
            """, """
              <w:tbl>
                <w:tblPr><w:tblStyle w:val="TableNormal"/></w:tblPr>
                <w:tblGrid><w:gridCol/></w:tblGrid>
                <w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>
              </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-table-style
            target M.T0001
            style Custom Grid
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Contains("w:val=" + (char)34 + "TableGridCustom" + (char)34, ReadDocumentXml(output), StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckReplaceParagraphWithTableStyleIdFailsWrongKind()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="table" w:styleId="TableGridCustom"><w:name w:val="Custom Grid"/></w:style>
            """, """
              <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-paragraph
            target M.P0001
            text Changed
            style TableGridCustom
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E7103");
    }

    [Fact]
    public static void CheckSetTableStyleWithParagraphStyleNameFailsWrongKind()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="paragraph" w:styleId="CustomHeading"><w:name w:val="Example Heading"/></w:style>
              <w:style w:type="table" w:styleId="TableGridCustom"><w:name w:val="Custom Grid"/></w:style>
            """, """
              <w:tbl>
                <w:tblPr><w:tblStyle w:val="TableGridCustom"/></w:tblPr>
                <w:tblGrid><w:gridCol/></w:tblGrid>
                <w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>
              </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-table-style
            target M.T0001
            style Example Heading
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E7103");
    }

}
