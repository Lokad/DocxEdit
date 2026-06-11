using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace DocxEdit.Tests;

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
    public static void ApplyMarksFieldsDirtyByDefault()
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
        string settings = ReadEntry(output, "word/settings.xml");
        Assert.Contains("updateFields", settings, StringComparison.Ordinal);
        Assert.Contains("w:val=\"true\"", settings, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains("relationships/settings", ReadEntry(output, "word/_rels/document.xml.rels"), StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains("wordprocessingml.settings+xml", ReadEntry(output, "[Content_Types].xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyCanSkipMarkingFieldsDirty()
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

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { MarkFieldsDirtyWhenEditing = false });

        Assert.True(result.Success);
        output.Position = 0;
        Assert.False(EntryExists(output, "word/settings.xml"));
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
    public static void CheckTrackChangesRequireFailsUnsupportedOperations()
    {
        using MemoryStream input = CreateDocx("Revenue increased.");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find increased
            with rose
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E6001");
    }

    [Fact]
    public static void ApplyTrackChangesSuggestWarnsAndAppliesDirectly()
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

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W4001" && diagnostic.Severity == DocxSeverity.Warning);
        output.Position = 0;
        Assert.Equal("Revenue rose.", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
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
    public static void ApplyPreservesUnknownPartsAndUnrelatedMedia()
    {
        using MemoryStream input = CreateDocxWithBody(
            """
                    <w:p><w:r><w:t>Before</w:t></w:r></w:p>
            """,
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
        using MemoryStream input = CreateDocxWithBody("""
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
    public static void ApplyInsertAfterCanTargetBookmarkSelectorAndPreserveBookmark()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="1" w:name="ReviewPoint"/>
                      <w:r><w:t>Bookmarked paragraph</w:t></w:r>
                      <w:bookmarkEnd w:id="1"/>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target bookmark:"ReviewPoint"
            text Inserted
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal(new[] { "Bookmarked paragraph", "Inserted" }, new DocxEditor().Read(output).Paragraphs.Select(paragraph => paragraph.Text));
        output.Position = 0;
        Assert.Contains("w:name=\"ReviewPoint\"", ReadDocumentXml(output), StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckContentControlSelectorResolvesAndRejectsProtectedTextEdit()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:alias w:val="Client Name"/>
                          <w:tag w:val="client-name"/>
                        </w:sdtPr>
                        <w:sdtContent>
                          <w:r><w:t>Client</w:t></w:r>
                        </w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target content-control:"client-name"
            find Client
            with Customer
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E4305" && diagnostic.TargetId == "content-control:\"client-name\"");
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
        Assert.Equal("landscape", section.Orientation);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("w:num=\"2\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:orient=\"landscape\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:w=\"15840\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:h=\"12240\"", xml, StringComparison.Ordinal);
    }

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
        Assert.Contains(read.Paragraphs, paragraph => paragraph.Id == "H001.P0001" && paragraph.Text == "Confidential text");
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
        Assert.Contains(read.Paragraphs, paragraph => paragraph.Id == "F001.P0001" && paragraph.Text == "Page 2");
    }

    [Fact]
    public static void ApplySetCellPreservesCellPropertiesAndDoesNotMutateInput()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>
                          <w:p>
                            <w:pPr><w:pStyle w:val="TableBody"/></w:pPr>
                            <w:r><w:t>Old</w:t></w:r>
                          </w:p>
                        </w:tc>
                        <w:tc><w:p><w:r><w:t>Other</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-cell
            target M.T0001.R01.C01
            expect-text Old
            text New
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        input.Position = 0;
        output.Position = 0;
        Assert.Equal("Old", new DocxEditor().Read(input).Tables[0].Cells[0].Text);
        Assert.Equal("New", new DocxEditor().Read(output).Tables[0].Cells[0].Text);

        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:tcW", xml, StringComparison.Ordinal);
        Assert.Contains("w:val=\"TableBody\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckSetCellRejectsComplexCellWithoutForce()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:p><w:r><w:t>One</w:t></w:r></w:p>
                          <w:p><w:r><w:t>Two</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-cell
            target M.T0001.R01.C01
            text Replacement
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E4302");
    }

    [Fact]
    public static void ApplyAppendRowClonesRowShapeAndCellProperties()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>
                      <w:tr>
                        <w:trPr><w:trHeight w:val="240"/></w:trPr>
                        <w:tc>
                          <w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>
                          <w:p>
                            <w:pPr><w:pStyle w:val="TableRegion"/></w:pPr>
                            <w:r><w:t>North</w:t></w:r>
                          </w:p>
                        </w:tc>
                        <w:tc>
                          <w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>
                          <w:p><w:r><w:t>Revenue</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op append-row
            target M.T0001
            cell South
            cell Profit
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal(2, table.RowCount);
        Assert.Equal("South", table.Cells.Single(cell => cell.RowIndex == 2 && cell.ColumnIndex == 1).Text);
        Assert.Equal("Profit", table.Cells.Single(cell => cell.RowIndex == 2 && cell.ColumnIndex == 2).Text);

        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Equal(2, CountOccurrences(xml, "<w:trHeight"));
        Assert.Equal(4, CountOccurrences(xml, "<w:tcW"));
        Assert.Equal(2, CountOccurrences(xml, "w:val=\"TableRegion\""));
    }

    [Fact]
    public static void CheckAppendRowRejectsCellCountMismatch()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op append-row
            target M.T0001
            cell South
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E4303");
    }

    [Fact]
    public static void CheckAppendRowRejectsFailedShapeGuards()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op append-row
            target M.T0001
            expect-row-count 2
            expect-column-count 2
            cell South
            cell Profit
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E3201" && diagnostic.TargetId == "M.T0001");
    }

    [Fact]
    public static void ApplyReplaceImageUsesAssetProviderForPng()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", "new-png", null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-image
            target M.I0001
            asset chart.png
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("new-png", ReadEntry(output, "word/media/image1.png"));
    }

    [Fact]
    public static void CheckReplaceImageRejectsFailedContentTypeGuard()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        var assets = new MemoryAssetProvider("chart.png", "new-png", null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-image
            target M.I0001
            expect-content-type image/jpeg
            asset chart.png
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { AssetProvider = assets });

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E3201" && diagnostic.TargetId == "M.I0001");
    }

    [Fact]
    public static void ApplyReplaceImageUsesAssetProviderForJpeg()
    {
        using MemoryStream input = CreateDocxWithImage("jpeg", "image/jpeg", "old-jpeg");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("photo.jpeg", "new-jpeg", null, "photo.jpeg");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-image
            target M.I0001
            asset photo.jpeg
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("new-jpeg", ReadEntry(output, "word/media/image1.jpeg"));
    }

    [Fact]
    public static void CheckReplaceImageFailsWithoutAssetProvider()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-image
            target M.I0001
            asset chart.png
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E5201");
    }

    [Fact]
    public static void CheckReplaceImageRejectsUnsupportedAssetType()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        var assets = new MemoryAssetProvider("asset.bin", "not-image", null, "asset.bin");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-image
            target M.I0001
            asset asset.bin
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { AssetProvider = assets });

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E5203");
    }

    [Fact]
    public static void ApplySetImageAltUpdatesDrawingProperties()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-alt
            target M.I0001
            alt Updated chart
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("descr=\"Updated chart\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("descr=\"Old chart\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyDeleteImageRemovesDrawingAndPreservesParagraph()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-image
            target M.I0001
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Empty(read.Images);
        Assert.Single(read.Paragraphs);
        output.Position = 0;
        Assert.DoesNotContain("<w:drawing>", ReadDocumentXml(output), StringComparison.Ordinal);
        output.Position = 0;
        Assert.False(EntryExists(output, "word/media/image1.png"));
        output.Position = 0;
        string relationships = ReadEntry(output, "word/_rels/document.xml.rels");
        Assert.DoesNotContain("rImage", relationships, StringComparison.Ordinal);
        Assert.DoesNotContain("media/image1.png", relationships, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyInsertImageAfterAddsInlinePngImage()
    {
        using MemoryStream input = CreateDocx("Intro");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", "new-png", null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset chart.png
            width 1in
            alt Inserted chart
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        DocxImageInfo image = Assert.Single(read.Images);
        Assert.Equal("/word/media/image1.png", image.PartName);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(2, read.Paragraphs.Count);
        output.Position = 0;
        Assert.Equal("new-png", ReadEntry(output, "word/media/image1.png"));
        output.Position = 0;
        Assert.Contains("ContentType=\"image/png\"", ReadEntry(output, "[Content_Types].xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyInsertImageAfterPreservesPngAspectRatioWhenOnlyWidthIsSupplied()
    {
        using MemoryStream input = CreateDocx("Intro");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", CreatePngBytes(200, 100), null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset chart.png
            width 2in
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        (long cx, long cy) = ReadFirstInlineImageExtent(output);
        Assert.Equal(1_828_800, cx);
        Assert.Equal(914_400, cy);
    }

    [Fact]
    public static void ApplyInsertImageAfterInfersPngDimensionsWhenNoneAreSupplied()
    {
        using MemoryStream input = CreateDocx("Intro");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", CreatePngBytes(200, 100), null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset chart.png
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        (long cx, long cy) = ReadFirstInlineImageExtent(output);
        Assert.Equal(1_905_000, cx);
        Assert.Equal(952_500, cy);
    }

    [Fact]
    public static void ApplyInsertImageAfterPreservesJpegAspectRatioWhenOnlyHeightIsSupplied()
    {
        using MemoryStream input = CreateDocx("Intro");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("photo.jpeg", CreateJpegBytes(120, 60), null, "photo.jpeg");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset photo.jpeg
            height 1in
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        (long cx, long cy) = ReadFirstInlineImageExtent(output);
        Assert.Equal(1_828_800, cx);
        Assert.Equal(914_400, cy);
    }

    [Fact]
    public static void ApplyInsertImageAfterAddsInlineJpegImage()
    {
        using MemoryStream input = CreateDocx("Intro");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("photo.jpeg", "new-jpeg", null, "photo.jpeg");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset photo.jpeg
            width 72pt
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        DocxImageInfo image = Assert.Single(new DocxEditor().Read(output).Images);
        Assert.Equal("/word/media/image1.jpeg", image.PartName);
        Assert.Equal("image/jpeg", image.ContentType);
        output.Position = 0;
        Assert.Equal("new-jpeg", ReadEntry(output, "word/media/image1.jpeg"));
    }

    [Fact]
    public static void ApplyInsertImageAfterAllocatesUniqueDocPrId()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", "new-png", null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset chart.png
            width 1in
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<wp:docPr id=\"1\"", xml, StringComparison.Ordinal);
        Assert.Contains("<wp:docPr id=\"2\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyInsertImageAfterTwiceRefreshesContentTypesAndRelationships()
    {
        using MemoryStream input = CreateDocx("Intro");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider(
            ("chart.png", "new-png", null, "chart.png"),
            ("logo.png", "new-logo", null, "logo.png"));
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset chart.png
            width 1in
            end

            op insert-image-after
            target M.P0001
            asset logo.png
            width 72pt
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal(2, read.Images.Count);
        Assert.Contains(read.Images, image => image.PartName == "/word/media/image1.png" && image.ContentType == "image/png");
        Assert.Contains(read.Images, image => image.PartName == "/word/media/image2.png" && image.ContentType == "image/png");

        output.Position = 0;
        Assert.Equal("new-png", ReadEntry(output, "word/media/image1.png"));
        output.Position = 0;
        Assert.Equal("new-logo", ReadEntry(output, "word/media/image2.png"));

        output.Position = 0;
        string contentTypes = ReadEntry(output, "[Content_Types].xml");
        Assert.Contains("PartName=\"/word/media/image1.png\"", contentTypes, StringComparison.Ordinal);
        Assert.Contains("ContentType=\"image/png\"", contentTypes, StringComparison.Ordinal);
        Assert.Contains("PartName=\"/word/media/image2.png\"", contentTypes, StringComparison.Ordinal);

        output.Position = 0;
        string relationships = ReadEntry(output, "word/_rels/document.xml.rels");
        Assert.Contains("Target=\"media/image1.png\"", relationships, StringComparison.Ordinal);
        Assert.Contains("Target=\"media/image2.png\"", relationships, StringComparison.Ordinal);
        Assert.Equal(2, CountOccurrences(relationships, "relationships/image"));
    }

    [Fact]
    public static void ApplyInsertRowBeforeClonesTargetRowShape()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:trPr><w:trHeight w:val="480"/></w:trPr>
                        <w:tc>
                          <w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>
                          <w:p><w:r><w:t>South</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                          <w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>
                          <w:p><w:r><w:t>Profit</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-row-before
            target M.T0001.R02
            cell East
            cell Margin
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal(3, table.RowCount);
        Assert.Equal("East", table.Cells.Single(cell => cell.RowIndex == 2 && cell.ColumnIndex == 1).Text);
        Assert.Equal("South", table.Cells.Single(cell => cell.RowIndex == 3 && cell.ColumnIndex == 1).Text);

        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Equal(2, CountOccurrences(xml, "<w:trHeight"));
        Assert.Equal(4, CountOccurrences(xml, "<w:tcW"));
    }

    [Fact]
    public static void ApplyInsertRowAfterPlacesRowAfterTarget()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-row-after
            target M.T0001.R01
            cell East
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal("North", table.Cells.Single(cell => cell.RowIndex == 1 && cell.ColumnIndex == 1).Text);
        Assert.Equal("East", table.Cells.Single(cell => cell.RowIndex == 2 && cell.ColumnIndex == 1).Text);
        Assert.Equal("South", table.Cells.Single(cell => cell.RowIndex == 3 && cell.ColumnIndex == 1).Text);
    }

    [Fact]
    public static void ApplyDeleteRowRemovesSelectedRow()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>East</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-row
            target M.T0001.R02
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal(2, table.RowCount);
        Assert.Equal("North", table.Cells.Single(cell => cell.RowIndex == 1 && cell.ColumnIndex == 1).Text);
        Assert.Equal("East", table.Cells.Single(cell => cell.RowIndex == 2 && cell.ColumnIndex == 1).Text);
    }

    [Fact]
    public static void CheckDeleteRowRejectsLastRow()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Only</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op delete-row
            target M.T0001.R01
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E4304");
    }

    private static MemoryStream CreateDocx(string paragraphText)
    {
        return CreateDocxWithRuns(paragraphText);
    }

    private static MemoryStream CreateDocxWithRuns(params string[] runTexts)
    {
        var stream = new MemoryStream();
        string runs = string.Concat(runTexts.Select(text => $"<w:r><w:t>{text}</w:t></w:r>"));
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """);
            AddEntry(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
                """);
            AddEntry(archive, "word/document.xml", $$"""
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>{{runs}}</w:p>
                  </w:body>
                </w:document>
                """);
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithBody(string bodyXml, Action<ZipArchive>? extra = null)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """);
            AddEntry(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
                """);
            string documentXml = """
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>

                """ + bodyXml + """

                  </w:body>
                </w:document>
                """;
            AddEntry(archive, "word/document.xml", documentXml);
            extra?.Invoke(archive);
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithImage(string extension, string contentType, string mediaBytes)
    {
        var stream = new MemoryStream();
        string partName = $"word/media/image1.{extension}";
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", $$"""
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="{{extension}}" ContentType="{{contentType}}"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """);
            AddEntry(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/_rels/document.xml.rels", $$"""
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.{{extension}}"/>
                </Relationships>
                """);
            AddEntry(archive, "word/document.xml", """
                <w:document
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:drawing>
                          <wp:inline>
                            <wp:docPr id="1" name="Picture 1" descr="Old chart"/>
                            <a:graphic>
                              <a:graphicData>
                                <pic:pic>
                                  <pic:blipFill>
                                    <a:blip r:embed="rImage"/>
                                  </pic:blipFill>
                                </pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                    </w:p>
                  </w:body>
                </w:document>
                """);
            AddEntry(archive, partName, mediaBytes);
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithHeaderFooter(string headerText, string footerText)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
                </Types>
                """);
            AddEntry(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                  <Relationship Id="rFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/document.xml", """
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>Main text</w:t></w:r></w:p>
                  </w:body>
                </w:document>
                """);
            AddEntry(archive, "word/header1.xml", $$"""
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:p><w:r><w:t>{{headerText}}</w:t></w:r></w:p>
                </w:hdr>
                """);
            AddEntry(archive, "word/footer1.xml", $$"""
                <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:p><w:r><w:t>{{footerText}}</w:t></w:r></w:p>
                </w:ftr>
                """);
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithStyles(string styleElements)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
                </Types>
                """);
            AddEntry(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/document.xml", """
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>Styled paragraph</w:t></w:r></w:p>
                  </w:body>
                </w:document>
                """);
            string stylesXml = """
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

                """ + styleElements + """

                </w:styles>
                """;
            AddEntry(archive, "word/styles.xml", stylesXml);
        }

        stream.Position = 0;
        return stream;
    }

    private static string ReadDocumentXml(Stream docx)
    {
        using var archive = new ZipArchive(docx, ZipArchiveMode.Read, leaveOpen: true);
        ZipArchiveEntry entry = archive.GetEntry("word/document.xml")
            ?? throw new InvalidDataException("Missing word/document.xml.");
        using Stream stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    }

    private static (long Cx, long Cy) ReadFirstInlineImageExtent(Stream docx)
    {
        string xml = ReadDocumentXml(docx);
        XElement extent = XDocument.Parse(xml)
            .Descendants(XName.Get("extent", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"))
            .First();
        return ((long)extent.Attribute("cx")!, (long)extent.Attribute("cy")!);
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

    private static bool EntryExists(Stream docx, string entryName)
    {
        using var archive = new ZipArchive(docx, ZipArchiveMode.Read, leaveOpen: true);
        return archive.GetEntry(entryName) is not null;
    }

    private static int CountOccurrences(string text, string value)
    {
        int count = 0;
        int index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += value.Length;
        }

        return count;
    }

    private static void AddEntry(ZipArchive archive, string name, string text)
    {
        ZipArchiveEntry entry = archive.CreateEntry(name);
        using Stream stream = entry.Open();
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static byte[] CreatePngBytes(int width, int height)
    {
        byte[] bytes =
        [
            0x89, 0x50, 0x4E, 0x47,
            0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x0D,
            0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00
        ];
        WriteBigEndianInt32(bytes, 16, width);
        WriteBigEndianInt32(bytes, 20, height);
        return bytes;
    }

    private static byte[] CreateJpegBytes(int width, int height)
    {
        byte[] bytes =
        [
            0xFF, 0xD8,
            0xFF, 0xC0,
            0x00, 0x11,
            0x08,
            0x00, 0x00,
            0x00, 0x00,
            0x03,
            0x01, 0x11, 0x00,
            0x02, 0x11, 0x00,
            0x03, 0x11, 0x00,
            0xFF, 0xD9
        ];
        WriteBigEndianUInt16(bytes, 7, height);
        WriteBigEndianUInt16(bytes, 9, width);
        return bytes;
    }

    private static void WriteBigEndianInt32(byte[] bytes, int offset, int value)
    {
        bytes[offset] = (byte)((value >> 24) & 0xFF);
        bytes[offset + 1] = (byte)((value >> 16) & 0xFF);
        bytes[offset + 2] = (byte)((value >> 8) & 0xFF);
        bytes[offset + 3] = (byte)(value & 0xFF);
    }

    private static void WriteBigEndianUInt16(byte[] bytes, int offset, int value)
    {
        bytes[offset] = (byte)((value >> 8) & 0xFF);
        bytes[offset + 1] = (byte)(value & 0xFF);
    }

    private sealed class MemoryAssetProvider : IDocxAssetProvider
    {
        private readonly Dictionary<string, MemoryAsset> assets;

        public MemoryAssetProvider(string reference, string text, string? contentTypeHint, string? fileNameHint)
            : this((reference, text, contentTypeHint, fileNameHint))
        {
        }

        public MemoryAssetProvider(string reference, byte[] bytes, string? contentTypeHint, string? fileNameHint)
            : this(new MemoryAsset(reference, bytes, contentTypeHint, fileNameHint))
        {
        }

        public MemoryAssetProvider(params (string Reference, string Text, string? ContentTypeHint, string? FileNameHint)[] assets)
            : this(assets.Select(asset => new MemoryAsset(
                asset.Reference,
                Encoding.UTF8.GetBytes(asset.Text),
                asset.ContentTypeHint,
                asset.FileNameHint)).ToArray())
        {
        }

        private MemoryAssetProvider(params MemoryAsset[] assets)
        {
            this.assets = assets.ToDictionary(
                asset => asset.Reference,
                asset => asset,
                StringComparer.Ordinal);
        }

        public bool TryOpen(string requestedReference, out Stream stream, out string? contentTypeHint, out string? fileNameHint)
        {
            if (!assets.TryGetValue(requestedReference, out MemoryAsset? asset))
            {
                stream = Stream.Null;
                contentTypeHint = null;
                fileNameHint = null;
                return false;
            }

            stream = new MemoryStream(asset.Bytes, writable: false);
            contentTypeHint = asset.ContentTypeHint;
            fileNameHint = asset.FileNameHint;
            return true;
        }

        private sealed record MemoryAsset(string Reference, byte[] Bytes, string? ContentTypeHint, string? FileNameHint);
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
}
