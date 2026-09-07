using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml.Linq;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

public static class PatchSectionTests
{

    [Fact]
    public static void CheckSetSectionRejectsFailedGuards()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>One</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:pgSz w:w="12240" w:h="15840"/>
                      <w:cols w:num="1"/>
                    </w:sectPr>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-section-orientation
            target M.S0001
            expect-columns 2
            orientation landscape
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E3201" && diagnostic.TargetId == "M.S0001");
    }

    [Fact]
    public static void ApplySetSectionColumnsAndOrientation()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>One</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:pgSz w:w="12240" w:h="15840"/>
                      <w:cols w:num="1"/>
                    </w:sectPr>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-section-columns
            target M.S0001
            count 2
            end

            op set-section-orientation
            target M.S0001
            orientation landscape
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxSectionInfo section = Assert.Single(new DocxEditor().Read(output).Sections);
        Assert.Equal(2, section.Columns);
        Assert.Equal(DocxOrientation.Landscape, section.Orientation);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("w:num=\"2\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:orient=\"landscape\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:w=\"15840\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:h=\"12240\"", xml, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("set-section-columns", """
        target M.S0001
        count 2
        """)]
    [InlineData("set-section-orientation", """
        target M.S0001
        orientation landscape
        """)]
    public static void CheckTrackChangesRequireAllowsSectionPropertyRevisions(string operationName, string operationFields)
    {
        using MemoryStream input = CreateDocxWithReferencedSection();
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            {operationFields}
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code is "E6001" or "E6002");
    }

    [Theory]
    [InlineData("set-section-columns", """
        target M.S0001
        count 2
        """, "w:num=\"2\"", "w:w=\"12240\"", "w:h=\"15840\"")]
    [InlineData("set-section-orientation", """
        target M.S0001
        orientation landscape
        """, "w:orient=\"landscape\"", "w:w=\"15840\"", "w:h=\"12240\"")]
    public static void ApplyTrackChangesSuggestGeneratesSectionPropertyRevisions(
        string operationName,
        string operationFields,
        string expectedXml,
        string expectedWidth,
        string expectedHeight)
    {
        using MemoryStream input = CreateDocxWithReferencedSection();
        using var output = new MemoryStream();
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            {operationFields}
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code is "W4001" or "W4002");
        Assert.Equal(["1"], Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains(expectedXml, xml, StringComparison.Ordinal);
        Assert.Contains(expectedWidth, xml, StringComparison.Ordinal);
        Assert.Contains(expectedHeight, xml, StringComparison.Ordinal);
        Assert.Contains("<w:headerReference w:type=\"default\" r:id=\"rHeader\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:footerReference w:type=\"default\" r:id=\"rFooter\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:sectPrChange w:id=\"1\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:ins", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:del", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains(new DocxEditor().Changes(output).Summary, summary => summary.Type == "section-properties-change" && summary.Count == 1);
    }

    [Theory]
    [InlineData("set-section-columns", """
        target M.S0001
        count 2
        """)]
    [InlineData("set-section-orientation", """
        target M.S0001
        orientation landscape
        """)]
    public static void CheckTrackChangesRequireRejectsExistingSectionPropertyRevisions(string operationName, string operationFields)
    {
        using MemoryStream input = CreateDocxWithExistingSectionPropertyRevision();
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            {operationFields}
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
        Assert.Contains("section already contains tracked section property revision markup", diagnostic.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("set-section-columns", """
        target M.S0001
        count 2
        """, "w:num=\"2\"")]
    [InlineData("set-section-orientation", """
        target M.S0001
        orientation landscape
        """, "w:orient=\"landscape\"")]
    public static void ApplyTrackChangesSuggestPreservesExistingSectionPropertyRevisionsByDirectFallback(
        string operationName,
        string operationFields,
        string expectedXml)
    {
        using MemoryStream input = CreateDocxWithExistingSectionPropertyRevision();
        using var output = new MemoryStream();
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            {operationFields}
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W4002");
        Assert.Empty(Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains(expectedXml, xml, StringComparison.Ordinal);
        Assert.Contains("<w:sectPrChange w:id=\"7\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:sectPrChange w:id=\"1\"", xml, StringComparison.Ordinal);
    }
}
