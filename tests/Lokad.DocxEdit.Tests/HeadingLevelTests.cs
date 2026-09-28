using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D09: heading levels follow real outline semantics.
public static class HeadingLevelTests
{
    [Fact]
    public static void ReadRecognizesCustomStyleWithOutlineLevel()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="paragraph" w:styleId="CustomHeading"><w:name w:val="Example Heading"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
            """, """
              <w:p><w:pPr><w:pStyle w:val="CustomHeading"/></w:pPr><w:r><w:t>Actual heading</w:t></w:r></w:p>
            """);
        DocxReadResult read = new DocxEditor().Read(input);
        Assert.True(read.Success);
        Assert.Equal(1, Assert.Single(read.Paragraphs).HeadingLevel);
    }

    [Fact]
    public static void ReadIgnoresDigitBearingNonHeadingStyle()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="paragraph" w:styleId="Appendix2"><w:name w:val="Appendix 2"/></w:style>
            """, """
              <w:p><w:pPr><w:pStyle w:val="Appendix2"/></w:pPr><w:r><w:t>Appendix</w:t></w:r></w:p>
            """);
        DocxReadResult read = new DocxEditor().Read(input);
        Assert.True(read.Success);
        Assert.Null(Assert.Single(read.Paragraphs).HeadingLevel);
    }

    [Fact]
    public static void ReadResolvesHeadingLevelThroughBasedOn()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="paragraph" w:styleId="BaseHeading"><w:name w:val="Base Heading"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>
              <w:style w:type="paragraph" w:styleId="Derived"><w:name w:val="Derived"/><w:basedOn w:val="BaseHeading"/></w:style>
            """, """
              <w:p><w:pPr><w:pStyle w:val="Derived"/></w:pPr><w:r><w:t>Derived heading</w:t></w:r></w:p>
            """);
        DocxReadResult read = new DocxEditor().Read(input);
        Assert.True(read.Success);
        Assert.Equal(2, Assert.Single(read.Paragraphs).HeadingLevel);
    }

    [Fact]
    public static void ReadPrefersDirectOutlineLevelOverStyle()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="paragraph" w:styleId="CustomHeading"><w:name w:val="Example Heading"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
            """, """
              <w:p><w:pPr><w:pStyle w:val="CustomHeading"/><w:outlineLvl w:val="2"/></w:pPr><w:r><w:t>Leveled heading</w:t></w:r></w:p>
            """);
        DocxReadResult read = new DocxEditor().Read(input);
        Assert.True(read.Success);
        Assert.Equal(3, Assert.Single(read.Paragraphs).HeadingLevel);
    }

    [Fact]
    public static void ReadTreatsOutlineLevelNineAsBodyText()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="paragraph" w:styleId="CustomHeading"><w:name w:val="Example Heading"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
            """, """
              <w:p><w:pPr><w:pStyle w:val="CustomHeading"/><w:outlineLvl w:val="9"/></w:pPr><w:r><w:t>Body text</w:t></w:r></w:p>
            """);
        DocxReadResult read = new DocxEditor().Read(input);
        Assert.True(read.Success);
        Assert.Null(Assert.Single(read.Paragraphs).HeadingLevel);
    }

    [Fact]
    public static void ReadKeepsLatentBuiltInHeadingWithoutStylesPart()
    {
        using MemoryStream input = CreateDocxWithBody("""
              <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Latent heading</w:t></w:r></w:p>
            """);
        DocxReadResult read = new DocxEditor().Read(input);
        Assert.True(read.Success);
        Assert.Equal(1, Assert.Single(read.Paragraphs).HeadingLevel);
    }

    [Fact]
    public static void CheckHeadingSelectorFindsCustomOutlineHeading()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="paragraph" w:styleId="CustomHeading"><w:name w:val="Example Heading"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
            """, """
              <w:p><w:pPr><w:pStyle w:val="CustomHeading"/></w:pPr><w:r><w:t>Actual heading</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target heading:"Actual heading"
            find Actual
            with Renamed
            end
            """);
        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output);
        Assert.True(apply.Success, string.Join("|", apply.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Equal("Renamed heading", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void CheckHeadingSelectorRejectsDigitBearingNonHeading()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="paragraph" w:styleId="Appendix2"><w:name w:val="Appendix 2"/></w:style>
            """, """
              <w:p><w:pPr><w:pStyle w:val="Appendix2"/></w:pPr><w:r><w:t>Appendix</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target heading:"Appendix"
            find Appendix
            with Renamed
            end
            """);
        DocxCheckResult check = new DocxEditor().Check(input, patch);
        Assert.False(check.Success);
        Assert.Contains(check.Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }

    [Fact]
    public static void CheckHeadingSelectorWithLevelMatchesOutlineLevel()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="paragraph" w:styleId="CustomHeading"><w:name w:val="Example Heading"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>
            """, """
              <w:p><w:pPr><w:pStyle w:val="CustomHeading"/></w:pPr><w:r><w:t>Second heading</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target heading:2:"Second heading"
            find Second
            with Renamed
            end
            """);
        DocxCheckResult check = new DocxEditor().Check(input, patch);
        Assert.True(check.Success, string.Join("|", check.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
    }
}
