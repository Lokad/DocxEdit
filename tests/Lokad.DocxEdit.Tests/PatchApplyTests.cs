using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml.Linq;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

public static class PatchApplyTests
{
    [Fact]
    public static void CheckReplaceTextSucceedsForMatchingParagraph()
    {
        using MemoryStream input = CreateDocx("Revenue increased by 8.4%.");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            expect-text <<<
            Revenue increased by 8.4%.
            >>>
            find <<<
            8.4%
            >>>
            with <<<
            9.1%
            >>>
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success);
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.True(report.Success);
    }


    [Fact]
    public static void CheckReplaceTextFailsWhenGuardDoesNotMatch()
    {
        using MemoryStream input = CreateDocx("Revenue increased by 9.1%.");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            expect-text <<<
            Revenue increased by 8.4%.
            >>>
            find <<<
            8.4%
            >>>
            with <<<
            9.1%
            >>>
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E3201");
        Assert.False(Assert.Single(result.Operations).Success);
    }

    [Fact]
    public static void ApplyReplaceTextWritesNewDocumentWithoutMutatingInput()
    {
        using MemoryStream input = CreateDocx("Revenue increased by 8.4%.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            expect-text <<<
            Revenue increased by 8.4%.
            >>>
            find <<<
            8.4%
            >>>
            with <<<
            9.1%
            >>>
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        input.Position = 0;
        output.Position = 0;
        Assert.Equal("Revenue increased by 8.4%.", Assert.Single(new DocxEditor().Read(input).Paragraphs).Text);
        Assert.Equal("Revenue increased by 9.1%.", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void ApplyWorksWithMemoryStreamInputAndOutput()
    {
        using MemoryStream input = CreateDocx("Revenue increased.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find increased
            with rose
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("Revenue rose.", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }






















    [Fact]
    public static void ReadWorksWithNonSeekableInputStream()
    {
        using MemoryStream seekable = CreateDocx("Revenue increased.");
        using var nonSeekable = new NonSeekableReadStream(seekable.ToArray());

        DocxReadResult result = new DocxEditor().Read(nonSeekable);

        Assert.True(result.Success);
        Assert.Equal("Revenue increased.", Assert.Single(result.Paragraphs).Text);
    }

    [Fact]
    public static void ApplyReplaceTextWorksAcrossAdjacentRuns()
    {
        using MemoryStream input = CreateDocxWithRuns("Revenue ", "increased");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            expect-text <<<
            Revenue increased
            >>>
            find <<<
            Revenue increased
            >>>
            with <<<
            Revenue rose
            >>>
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("Revenue rose", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }




















































    [Fact]
    public static void ApplyReplaceTextFailsWhenFindTextIsMissing()
    {
        using MemoryStream input = CreateDocx("Revenue increased by 8.4%.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find missing
            with value
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E4203");
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public static void ApplyReplaceTextPreservesXmlSpaceWhenReplacementRequiresIt()
    {
        using MemoryStream input = CreateDocx("Revenue increased by 8.4%.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find 8.4%
            with <<<
             9.1% 
            >>>
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("Revenue increased by  9.1% .", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        output.Position = 0;
        Assert.Contains("xml:space=\"preserve\"", ReadDocumentXml(output), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyReplaceTextHonorsOccurrence()
    {
        using MemoryStream input = CreateDocx("foo foo foo");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find foo
            with bar
            occurrence 2
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("foo bar foo", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void ApplyReplaceTextPreservesTargetRunProperties()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:rPr><w:b/></w:rPr><w:t>Revenue </w:t></w:r>
                      <w:r><w:rPr><w:i/></w:rPr><w:t>increased</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find increased
            with rose
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("Revenue rose", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:b", xml, StringComparison.Ordinal);
        Assert.Contains("<w:i", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>rose</w:t>", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckReplaceTextRejectsProtectedBoundaries()
    {
        using MemoryStream hyperlinkInput = CreateDocxWithBody("""
                    <w:p>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rLink">
                        <w:r><w:t>Link text</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """);
        using var hyperlinkPatch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Link
            with Anchor
            end
            """);

        DocxCheckResult hyperlinkResult = new DocxEditor().Check(hyperlinkInput, hyperlinkPatch);

        Assert.False(hyperlinkResult.Success);
        Assert.Contains(hyperlinkResult.Diagnostics, diagnostic => diagnostic.Code == "E4305");

        using MemoryStream revisionInput = CreateDocxWithBody("""
                    <w:p>
                      <w:ins w:id="1" w:author="A">
                        <w:r><w:t>Inserted</w:t></w:r>
                      </w:ins>
                    </w:p>
            """);
        using var revisionPatch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Inserted
            with Edited
            end
            """);

        DocxCheckResult revisionResult = new DocxEditor().Check(revisionInput, revisionPatch);

        Assert.False(revisionResult.Success);
        Assert.Contains(revisionResult.Diagnostics, diagnostic => diagnostic.Code == "E4305");

    }

    [Fact]
    public static void CheckReplaceTextInsideCommentRangePreservesMarkers()
    {
        using MemoryStream input = CreateDocxWithCommentAnchoredParagraph();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Commented
            with Updated
            end
            """);

        DocxCheckResult check = new DocxEditor().Check(input, patch);

        Assert.True(check.Success, string.Join("|", check.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));

        using MemoryStream applyInput = CreateDocxWithCommentAnchoredParagraph();
        using var output = new MemoryStream();
        using var applyPatch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Commented
            with Updated
            end
            """);

        Assert.True(new DocxEditor().Apply(applyInput, applyPatch, output).Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:t>Updated</w:t>", xml, StringComparison.Ordinal);
        Assert.Contains("commentRangeStart", xml, StringComparison.Ordinal);
        Assert.Contains("commentRangeEnd", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyPreservesUnknownPartsAndUnrelatedMedia()
    {
        using MemoryStream input = CreateDocxWithBody(
            """
                    <w:p><w:r><w:t>Before</w:t></w:r></w:p>
            """,
            null,
            archive =>
            {
                AddEntry(archive, "custom/data.bin", "opaque");
                AddEntry(archive, "word/media/unrelated.png", "unrelated");
            });
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Before
            with After
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("opaque", ReadEntry(output, "custom/data.bin"));
        output.Position = 0;
        Assert.Equal("unrelated", ReadEntry(output, "word/media/unrelated.png"));
    }

    [Fact]
    public static void ApplyReplaceParagraphUpdatesTextAndStyle()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/></w:style>
            """, """
                    <w:p>
                      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
                      <w:r><w:t>Old heading</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-paragraph
            target M.P0001
            expect-text Old heading
            style Heading2
            text New heading
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("New heading", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        output.Position = 0;
        Assert.Contains("w:val=\"Heading2\"", ReadDocumentXml(output), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyReplaceTextCanTargetHeadingByTextAndLevel()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
                      <w:r><w:t>Old heading</w:t></w:r>
                    </w:p>
                    <w:p><w:r><w:t>Body</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target heading:2:"Old heading"
            find Old
            with New
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("New heading", new DocxEditor().Read(output).Paragraphs[0].Text);
    }

    [Fact]
    public static void ApplyInsertAfterCanTargetParagraphContainingText()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
                    <w:p><w:r><w:t>Needle paragraph</w:t></w:r></w:p>
                    <w:p><w:r><w:t>Omega</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target text:"Needle"
            text Inserted
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal(
            new[] { "Alpha", "Needle paragraph", "Inserted", "Omega" },
            new DocxEditor().Read(output).Paragraphs.Select(paragraph => paragraph.Text));
    }

    [Fact]
    public static void ApplyTextEditAndInsertAfterPreserveListParagraphProperties()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:pPr>
                        <w:pStyle w:val="ListParagraph"/>
                        <w:numPr>
                          <w:ilvl w:val="1"/>
                          <w:numId w:val="9"/>
                        </w:numPr>
                      </w:pPr>
                      <w:r><w:t>First item</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find First
            with Updated
            end

            op insert-after
            target M.P0001
            copy-paragraph-properties true
            text Second item
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal(new[] { "Updated item", "Second item" }, new DocxEditor().Read(output).Paragraphs.Select(paragraph => paragraph.Text));
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Equal(2, CountOccurrences(xml, "w:pStyle w:val=\"ListParagraph\""));
        Assert.Equal(2, CountOccurrences(xml, "w:numId w:val=\"9\""));
        Assert.Equal(2, CountOccurrences(xml, "w:ilvl w:val=\"1\""));
    }

















    [Fact]
    public static void CheckTextSelectorRejectsAmbiguousMatches()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>Revenue north</w:t></w:r></w:p>
                    <w:p><w:r><w:t>Revenue south</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target text:Revenue
            find Revenue
            with Sales
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E1202");
    }

    [Fact]
    public static void CheckTextSelectorZeroMatchStatesMainStoryScope()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target text:"Missing"
            find Missing
            with Updated
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal("E1201", diagnostic.Code);
        Assert.Contains("Semantic selectors search only the main story", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckTextSelectorSuggestsNearbyTargetsWithoutLeakingText()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
                      <w:r><w:t>Private heading text</w:t></w:r>
                    </w:p>
                    <w:p><w:r><w:t>Private paragraph text</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target text:"Missing"
            find Missing
            with Updated
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal("E1201", diagnostic.Code);
        Assert.Contains("Nearby paragraphs: M.P0001, M.P0002", diagnostic.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("Private", diagnostic.Message, StringComparison.Ordinal);
    }



























































    [Fact]
    public static void ApplyInsertBeforeParagraphAddsParagraphAtTargetPosition()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>One</w:t></w:r></w:p>
                    <w:p><w:r><w:t>Two</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-before
            target M.P0002
            text Inserted
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal(new[] { "One", "Inserted", "Two" }, new DocxEditor().Read(output).Paragraphs.Select(paragraph => paragraph.Text));
    }

    [Fact]
    public static void ApplyInsertAfterTableAddsParagraphAfterTable()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
            """, """
                    <w:p><w:r><w:t>One</w:t></w:r></w:p>
                    <w:tbl>
                      <w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:p><w:r><w:t>Two</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.T0001
            style Normal
            text Inserted
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal(new[] { "One", "Inserted", "Two" }, read.Paragraphs.Select(paragraph => paragraph.Text));
        output.Position = 0;
        Assert.Contains("w:val=\"Normal\"", ReadDocumentXml(output), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyDeleteBlockRemovesParagraphOrTable()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>One</w:t></w:r></w:p>
                    <w:tbl>
                      <w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:p><w:r><w:t>Two</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-block
            target M.T0001
            end

            op delete-block
            target M.P0001
            expect-text One
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Empty(read.Tables);
        Assert.Equal("Two", Assert.Single(read.Paragraphs).Text);
    }










    [Fact]
    public static void ApplyReplaceTextCanEditHeaderParagraph()
    {
        using MemoryStream input = CreateDocxWithHeaderFooter("Header text", "Footer text");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target H001.P0001
            find Header
            with Confidential
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output, new DocxReadOptions { IncludeHeadersFooters = true });
        Assert.Contains(read.Paragraphs, paragraph => paragraph.Id.ToWireValue() == "H001.P0001" && paragraph.Text == "Confidential text");
    }

    [Fact]
    public static void ApplyReplaceParagraphCanEditFooterParagraph()
    {
        using MemoryStream input = CreateDocxWithHeaderFooter("Header text", "Footer text");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-paragraph
            target F001.P0001
            expect-text Footer text
            text Page 2
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output, new DocxReadOptions { IncludeHeadersFooters = true });
        Assert.Contains(read.Paragraphs, paragraph => paragraph.Id.ToWireValue() == "F001.P0001" && paragraph.Text == "Page 2");
    }

    [Fact]
    public static void ApplyBlockOperationsCanEditHeaderAndFooterStories()
    {
        using MemoryStream input = CreateDocxWithHeaderFooter("Header text", "Footer text");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target H001.P0001
            text Header detail
            end

            op insert-after
            target F001.P0001
            text Footer detail
            end

            op delete-block
            target F001.P0001
            expect-text Footer text
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output, new DocxReadOptions { IncludeHeadersFooters = true });
        Assert.Contains(read.Paragraphs, paragraph => paragraph.Id.ToWireValue() == "H001.P0001" && paragraph.Text == "Header text");
        Assert.Contains(read.Paragraphs, paragraph => paragraph.Id.ToWireValue() == "H001.P0002" && paragraph.Text == "Header detail");
        Assert.Contains(read.Paragraphs, paragraph => paragraph.Id.ToWireValue() == "F001.P0001" && paragraph.Text == "Footer detail");
        Assert.DoesNotContain(read.Paragraphs, paragraph => paragraph.Story == "footer[1]" && paragraph.Text == "Footer text");
    }












































    [Fact]
    public static void ApplyGeneratedDocumentPassesStructuralValidation()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>Anchor</w:t></w:r></w:p>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            expect-text Anchor
            find Anchor
            with Updated anchor
            end

            op add-bookmark
            target M.P0001
            expect-text Updated anchor
            name ReviewAnchor
            end

            op set-table-metadata
            target M.T0001
            caption Review table
            description Generated validation fixture
            end

            op append-row
            target M.T0001
            cell South
            cell Profit
            end
            """);

        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output);

        Assert.True(apply.Success);
        output.Position = 0;
        DocxValidateResult validate = new DocxEditor().Validate(output);
        Assert.True(validate.Success, string.Join(Environment.NewLine, validate.Diagnostics.Select(diagnostic => $"{diagnostic.Code}: {diagnostic.Message}")));
        Assert.DoesNotContain(validate.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error);
    }




























    [Fact]
    public static void ApplyReplaceTextWithEscapedSelectorMatchesLineBreak()
    {
        using MemoryStream input = CreateDocxWithBody(
            """
            <w:p><w:r><w:t>AAA</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>BBB</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target text:"AAA\nBBB"
            find BBB
            with CCC
            end
            """);
        using var output = new MemoryStream();
        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxParagraphInfo paragraph = Assert.Single(new DocxEditor().Read(output).Paragraphs);
        Assert.Equal("AAA\nCCC", paragraph.Text);
    }































    private static byte[] ReadEntryBytes(Stream docx, string entryName)
    {
        using var archive = new ZipArchive(docx, ZipArchiveMode.Read, leaveOpen: true);
        ZipArchiveEntry entry = archive.GetEntry(entryName)
            ?? throw new InvalidDataException($"Missing {entryName}.");
        using Stream stream = entry.Open();
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }

    private static string ReadEntry(Stream docx, string entryName)
    {
        using var archive = new ZipArchive(docx, ZipArchiveMode.Read, leaveOpen: true);
        ZipArchiveEntry entry = archive.GetEntry(entryName)
            ?? throw new InvalidDataException($"Missing {entryName}.");
        using Stream stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    }



    private static void AddEntry(ZipArchive archive, string name, string text)
    {
        ZipArchiveEntry entry = archive.CreateEntry(name);
        using Stream stream = entry.Open();
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        stream.Write(bytes, 0, bytes.Length);
    }






    private sealed class NonSeekableReadStream : MemoryStream
    {
        public NonSeekableReadStream(byte[] bytes)
            : base(bytes, writable: false)
        {
        }

        public override bool CanSeek => false;

        public override long Seek(long offset, SeekOrigin loc)
        {
            throw new NotSupportedException();
        }

        public override long Position
        {
            get => base.Position;
            set => throw new NotSupportedException();
        }

        public override long Length => throw new NotSupportedException();
    }

    [Fact]
    public static void ApplyReplaceTextWithEmptyWithDeletesMatchedText()
    {
        using MemoryStream input = CreateDocx("Alpha Beta");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            with ""
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        DocxParagraphInfo paragraph = Assert.Single(new DocxEditor().Read(output).Paragraphs);
        Assert.Equal(" Beta", paragraph.Text);
        Assert.NotEmpty(paragraph.Runs);
    }

    [Fact]
    public static void ApplyReplaceTextWithEmptyHeredocDeletesMatchedText()
    {
        using MemoryStream input = CreateDocx("Alpha Beta");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            with <<<
            >>>
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Equal(" Beta", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void ApplyReplaceTextMissingWithStillFails()
    {
        using MemoryStream input = CreateDocx("Alpha Beta");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E4202");
    }

    [Fact]
    public static void ApplyReplaceTextEmptyFindStillFails()
    {
        using MemoryStream input = CreateDocx("Alpha Beta");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find ""
            with Omega
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E4202");
    }

    [Fact]
    public static void ApplyReplaceTextEmptyWithUnderRequireTracksDeletion()
    {
        var options = new DocxEditOptions { TrackChanges = TrackChangesMode.Require };
        using MemoryStream input = CreateDocx("Alpha Beta");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            with ""
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.NotEmpty(Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        Assert.Equal(" Beta", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void CheckReplaceTextEmptyWithAgreesWithApply()
    {
        const string patchText = "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith \"\"\nend\n";
        using MemoryStream checkInput = CreateDocx("Alpha Beta");
        DocxCheckResult check = new DocxEditor().Check(checkInput, new StringReader(patchText));
        using MemoryStream applyInput = CreateDocx("Alpha Beta");
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, new StringReader(patchText), output);
        Assert.Equal(apply.Success, check.Success);
        Assert.True(check.Success);
    }

    [Fact]
    public static void ApplyReplaceParagraphWithEmptyTextClearsParagraph()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-paragraph
            target M.P0001
            text ""
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Equal("", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }


    [Fact]
    public static void ApplyReplaceTextAmbiguousWithoutOccurrenceFails()
    {
        const string patchText = "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Omega\nend\n";
        using MemoryStream checkInput = CreateDocx("Alpha Alpha");
        DocxCheckResult check = new DocxEditor().Check(checkInput, new StringReader(patchText));
        Assert.False(check.Success);
        Assert.Contains(check.Diagnostics, static diagnostic => diagnostic.Code == "E1202");
        using MemoryStream applyInput = CreateDocx("Alpha Alpha");
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, new StringReader(patchText), output);
        Assert.False(apply.Success);
        DocxDiagnostic diagnostic = Assert.Single(apply.Diagnostics, static d => d.Code == "E1202");
        Assert.Contains("2 occurrences", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyReplaceTextOccurrenceSelectsSingleMatch()
    {
        using MemoryStream firstInput = CreateDocx("Alpha Alpha");
        using var firstOutput = new MemoryStream();
        using var firstPatch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            with Omega
            occurrence 1
            end
            """);
        Assert.True(new DocxEditor().Apply(firstInput, firstPatch, firstOutput).Success);
        firstOutput.Position = 0;
        Assert.Equal("Omega Alpha", Assert.Single(new DocxEditor().Read(firstOutput).Paragraphs).Text);

        using MemoryStream secondInput = CreateDocx("Alpha Alpha");
        using var secondOutput = new MemoryStream();
        using var secondPatch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            with Omega
            occurrence 2
            end
            """);
        Assert.True(new DocxEditor().Apply(secondInput, secondPatch, secondOutput).Success);
        secondOutput.Position = 0;
        Assert.Equal("Alpha Omega", Assert.Single(new DocxEditor().Read(secondOutput).Paragraphs).Text);
    }

    [Fact]
    public static void ApplyReplaceTextOccurrenceAllReplacesEveryMatch()
    {
        using MemoryStream input = CreateDocx("Alpha Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            with Omega
            occurrence all
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Equal("Omega Omega", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void ApplyReplaceTextOccurrenceBeyondMatchesFails()
    {
        using MemoryStream input = CreateDocx("Alpha Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            with Omega
            occurrence 3
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E4203");
    }


    [Fact]
    public static void ApplyReplaceParagraphWithMultilineTextWritesBreakAndTabNodes()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-paragraph
            target M.P0001
            text "Line one\nLine two\tTab"
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Equal(1, CountOccurrences(xml, "<w:br"));
        Assert.Equal(1, CountOccurrences(xml, "<w:tab"));
        Assert.DoesNotContain("Line one\nLine two", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Equal("Line one\nLine two\tTab", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void ApplyReplaceTextWithoutPreserveRunsWritesBreakNodes()
    {
        using MemoryStream input = CreateDocx("Alpha Beta");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            with "X\nY"
            preserve-runs false
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Equal(1, CountOccurrences(xml, "<w:br"));
        output.Position = 0;
        Assert.Equal("X\nY Beta", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void ApplyInsertAfterAndReplaceParagraphEmitEquivalentBreakStructure()
    {
        const string patchText = "docxpatch 1\n\nop OP\ntarget M.P0001\ntext LINE\nend\n";
        using MemoryStream insertInput = CreateDocx("Alpha");
        using var insertOutput = new MemoryStream();
        using var insertPatch = new StringReader(patchText.Replace("OP", "insert-after").Replace("LINE", "\"A\\nB\""));
        Assert.True(new DocxEditor().Apply(insertInput, insertPatch, insertOutput).Success);
        insertOutput.Position = 0;
        string insertXml = ReadDocumentXml(insertOutput);

        using MemoryStream replaceInput = CreateDocx("Alpha");
        using var replaceOutput = new MemoryStream();
        using var replacePatch = new StringReader(patchText.Replace("OP", "replace-paragraph").Replace("LINE", "\"A\\nB\""));
        Assert.True(new DocxEditor().Apply(replaceInput, replacePatch, replaceOutput).Success);
        replaceOutput.Position = 0;
        string replaceXml = ReadDocumentXml(replaceOutput);

        Assert.Equal(CountOccurrences(insertXml, "<w:br"), CountOccurrences(replaceXml, "<w:br"));
        insertOutput.Position = 0;
        replaceOutput.Position = 0;
        Assert.Equal("A\nB", new DocxEditor().Read(insertOutput).Paragraphs.Last().Text);
        Assert.Equal("A\nB", Assert.Single(new DocxEditor().Read(replaceOutput).Paragraphs).Text);
    }

    [Fact]
    public static void ApplyInsertAfterWithBlankLineKeepsSingleParagraph()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text <<<
            A

            B
            >>>
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        IReadOnlyList<DocxParagraphInfo> paragraphs = new DocxEditor().Read(output).Paragraphs;
        Assert.Equal(2, paragraphs.Count);
        Assert.Equal("A\n\nB", paragraphs[1].Text);
    }

    [Fact]
    public static void ApplyReplaceParagraphOnTableIdFailsWithWrongKindGuidance()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-paragraph
            target M.T0001
            text Changed
            end
            """);

        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output);

        Assert.False(apply.Success);
        DocxDiagnostic diagnostic = Assert.Single(apply.Diagnostics, static d => d.Code == "E1201");
        Assert.Contains("table ID", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("M.P0001", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckInsertAfterOnRowIdFailsWithWrongKindGuidance()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.T0001.R01
            text Inserted
            end
            """);

        DocxCheckResult check = new DocxEditor().Check(input, patch);

        Assert.False(check.Success);
        DocxDiagnostic diagnostic = Assert.Single(check.Diagnostics, static d => d.Code == "E1201");
        Assert.Contains("table row ID", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("M.T0001", diagnostic.Message, StringComparison.Ordinal);
    }


    [Fact]
    public static void ApplyReplaceTextIdenticalReplacementIsNoOpUnderRequire()
    {
        var options = new DocxEditOptions { TrackChanges = TrackChangesMode.Require };
        using MemoryStream probe = CreateDocx("Anchor");
        string before = ReadDocumentXml(probe);
        using MemoryStream input = CreateDocx("Anchor");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Anchor
            with Anchor
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal("I0001", diagnostic.Code);
        Assert.Equal(DocxSeverity.Info, diagnostic.Severity);
        Assert.Empty(Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        Assert.Equal(before, ReadDocumentXml(output));
    }

    [Fact]
    public static void ApplyReplaceTextIdenticalReplacementIsNoOp()
    {
        using MemoryStream probe = CreateDocx("Anchor");
        string before = ReadDocumentXml(probe);
        using MemoryStream input = CreateDocx("Anchor");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Anchor
            with Anchor
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Contains(result.Diagnostics, static d => d.Code == "I0001");
        output.Position = 0;
        Assert.Equal(before, ReadDocumentXml(output));
    }

    [Fact]
    public static void ApplyReplaceParagraphIdenticalTextIsNoOp()
    {
        using MemoryStream probe = CreateDocx("Anchor");
        string before = ReadDocumentXml(probe);
        using MemoryStream input = CreateDocx("Anchor");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-paragraph
            target M.P0001
            text Anchor
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Contains(result.Diagnostics, static d => d.Code == "I0001");
        output.Position = 0;
        Assert.Equal(before, ReadDocumentXml(output));
    }

    [Fact]
    public static void ApplyReplaceParagraphSameTextNewStyleStillEdits()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
            """, """
              <w:p><w:r><w:t>Anchor</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-paragraph
            target M.P0001
            text Anchor
            style Normal
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.DoesNotContain(result.Diagnostics, static d => d.Code == "I0001");
    }

    [Fact]
    public static void ApplyEmptyPatchSucceedsWithNoOperations()
    {
        using MemoryStream input = CreateDocx("Anchor");
        using var output = new MemoryStream();
        using var patch = new StringReader("docxpatch 1\n");

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Empty(result.Operations);
    }

    [Fact]
    public static void ApplyAllNoOpPatchDoesNotMarkFieldsDirty()
    {
        using MemoryStream input = CreateDocxWithRefField();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Acme Corp
            with Acme Corp
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Contains(result.Diagnostics, static d => d.Code == "I0001");
        Assert.DoesNotContain(result.Diagnostics, static d => d.Code == "W5103");
    }


    [Fact]
    public static void CheckTextSelectorMatchesAcrossWhitespaceRuns()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t xml:space="preserve">Alpha  Beta</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target text:"Alpha Beta"
            find Alpha
            with Omega
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Equal("Omega  Beta", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void CheckTextSelectorStaysCaseSensitive()
    {
        using MemoryStream input = CreateDocx("Alpha Beta");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target text:"alpha"
            find Alpha
            with Omega
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }

    [Fact]
    public static void CheckHeadingSelectorMatchesAcrossWhitespaceRuns()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
            """, """
              <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">Executive  Summary</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target heading:"Executive Summary"
            find Executive
            with Revised
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Equal("Revised  Summary", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void CheckGuardStaysExactWhenSelectorNormalizes()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t xml:space="preserve">Alpha  Beta</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target text:"Alpha Beta"
            expect-text Alpha Beta
            find Alpha
            with Omega
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E3201");
    }


    [Fact]
    public static void CheckSelectorErrorCarriesTargetFieldLocation()
    {
        using MemoryStream input = CreateDocx("Alpha Beta");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target text:"Nope"
            find Alpha
            with Omega
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, static d => d.Code == "E1201");
        Assert.Equal(4, diagnostic.Line);
        Assert.Equal(1, diagnostic.Column);
    }

    [Fact]
    public static void CheckGuardErrorCarriesGuardFieldLocation()
    {
        using MemoryStream input = CreateDocx("Alpha Beta");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            expect-text Stale text
            find Alpha
            with Omega
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, static d => d.Code == "E3201");
        Assert.Equal(5, diagnostic.Line);
        Assert.Equal(1, diagnostic.Column);
    }

    [Fact]
    public static void CheckAmbiguityErrorCarriesTargetFieldLocation()
    {
        using MemoryStream input = CreateDocxWithBody("""
              <w:p><w:r><w:t>Revenue north</w:t></w:r></w:p>
              <w:p><w:r><w:t>Revenue south</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target text:"Revenue"
            find Revenue
            with Sales
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, static d => d.Code == "E1202");
        Assert.Equal(4, diagnostic.Line);
        Assert.Equal(1, diagnostic.Column);
    }

    [Fact]
    public static void CheckProtectedErrorCarriesTargetFieldLocation()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships("""
                    <w:p>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rLink">
                        <w:r><w:t>Text</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """, """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/old" TargetMode="External" />
                </Relationships>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Text
            with Changed
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, static d => d.Code == "E4305");
        Assert.Equal(4, diagnostic.Line);
        Assert.Equal(1, diagnostic.Column);
    }

    [Fact]
    public static void CheckFindNotFoundCarriesFindFieldLocation()
    {
        using MemoryStream input = CreateDocx("Alpha Beta");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Missing
            with Omega
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, static d => d.Code == "E4203");
        Assert.Equal(5, diagnostic.Line);
        Assert.Equal(1, diagnostic.Column);
    }


    [Fact]
    public static void ApplyInsertAfterEmptyTextStillFails()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text ""
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E4202");
    }

    [Fact]
    public static void ApplyInsertAfterMultipleTextsPreserveFileOrder()
    {
        const string patchText = "docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext First\ntext Second\nend\n";
        using MemoryStream applyInput = CreateDocx("Alpha");
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, new StringReader(patchText), output);

        Assert.True(apply.Success, string.Join("|", apply.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Equal(new[] { "Alpha", "First", "Second" }, new DocxEditor().Read(output).Paragraphs.Select(static p => p.Text).ToArray());

        using MemoryStream checkInput = CreateDocx("Alpha");
        DocxCheckResult check = new DocxEditor().Check(checkInput, new StringReader(patchText));

        Assert.True(check.Success, string.Join("|", check.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Equal(new[] { "M.P0002", "M.P0003" }, Assert.Single(check.Operations).CreatedTargetIds);
    }

    [Fact]
    public static void ApplyInsertBeforeMultipleTextsPreserveFileOrder()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
                    <w:p><w:r><w:t>Beta</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-before
            target M.P0002
            text X
            text Y
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Equal(new[] { "Alpha", "X", "Y", "Beta" }, new DocxEditor().Read(output).Paragraphs.Select(static p => p.Text).ToArray());
    }

    [Fact]
    public static void SingleStyleAppliesToEveryInsertedParagraph()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody(
            """
              <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
              <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/></w:style>
            """,
            """
                    <w:p><w:r><w:t>Anchor</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            style Heading 2
            text First
            text Second
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Equal(2, CountOccurrences(ReadDocumentXml(output), "w:pStyle w:val=\"Heading2\""));
    }

    [Fact]
    public static void PositionalStylesPairWithTextsInOrder()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody(
            """
              <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
              <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/></w:style>
            """,
            """
                    <w:p><w:r><w:t>Anchor</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            style Normal
            style Heading 2
            text First
            text Second
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Equal(1, CountOccurrences(xml, "w:pStyle w:val=\"Normal\""));
        Assert.Equal(1, CountOccurrences(xml, "w:pStyle w:val=\"Heading2\""));
        Assert.True(
            xml.IndexOf("w:pStyle w:val=\"Normal\"", StringComparison.Ordinal) < xml.IndexOf("w:pStyle w:val=\"Heading2\"", StringComparison.Ordinal),
            "First inserted paragraph should carry Normal and second Heading 2.");
    }

    [Fact]
    public static void StyleCountMismatchFails()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            style Normal
            style Heading 2
            style Title
            text First
            text Second
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, static d => d.Code == "E4205");
        Assert.Equal("insert-after", diagnostic.HelpTopic);
    }

    [Fact]
    public static void CopyPropertiesApplyToEveryInsertedParagraph()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:pPr>
                        <w:pStyle w:val="ListParagraph"/>
                        <w:numPr>
                          <w:ilvl w:val="1"/>
                          <w:numId w:val="9"/>
                        </w:numPr>
                      </w:pPr>
                      <w:r><w:t>First item</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            copy-paragraph-properties true
            text Second item
            text Third item
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Equal(3, CountOccurrences(xml, "w:numId w:val=\"9\""));
    }
    [Fact]
    public static void CheckTextSelectorSuggestsTextuallyNearestParagraphs()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>Alpha unrelated</w:t></w:r></w:p>
                    <w:p><w:r><w:t>Zebra unrelated</w:t></w:r></w:p>
                    <w:p><w:r><w:t>Revenue quarterly results</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target text:"QUARTERLY"
            find Quarterly
            with Annual
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal("E1201", diagnostic.Code);
        Assert.Contains("Nearby paragraphs: M.P0003, M.P0001, M.P0002", diagnostic.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("Revenue", diagnostic.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("quarterly", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckHeadingSelectorSuggestsNearestHeading()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
                      <w:r><w:t>Alpha overview</w:t></w:r>
                    </w:p>
                    <w:p>
                      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
                      <w:r><w:t>Beta quarterly review</w:t></w:r>
                    </w:p>
                    <w:p><w:r><w:t>Quarterly body text</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target heading:"quarterly summary"
            find quarterly
            with annual
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal("E1201", diagnostic.Code);
        Assert.Contains("M.P0002 heading level=1", diagnostic.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("Beta", diagnostic.Message, StringComparison.Ordinal);
    }
}
