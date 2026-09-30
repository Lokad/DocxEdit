using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml.Linq;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

public static class PatchBookmarkTests
{

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
    public static void ApplyAddBookmarkWrapsParagraphContent()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:pPr><w:pStyle w:val="BodyText"/></w:pPr>
                      <w:r><w:t>Client paragraph</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op add-bookmark
            target M.P0001
            expect-text Client paragraph
            name ClientParagraph
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        DocxBookmarkInfo bookmark = Assert.Single(read.Bookmarks);
        Assert.Equal("ClientParagraph", bookmark.Name);
        Assert.Equal("M.P0001", bookmark.StartTargetId?.ToWireValue());
        Assert.Equal("M.P0001", bookmark.EndTargetId?.ToWireValue());
        Assert.Equal("Client paragraph", Assert.Single(read.Paragraphs).Text);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:pPr><w:pStyle w:val=\"BodyText\" /></w:pPr><w:bookmarkStart", xml, StringComparison.Ordinal);
        Assert.Contains("w:name=\"ClientParagraph\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:bookmarkEnd", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckAddBookmarkRejectsDuplicateName()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="4" w:name="ClientParagraph"/>
                      <w:r><w:t>First</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                    </w:p>
                    <w:p><w:r><w:t>Second</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op add-bookmark
            target M.P0002
            name ClientParagraph
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E4311" &&
            diagnostic.Message.Contains("already exists", StringComparison.Ordinal));
    }

    [Fact]
    public static void ApplyReplaceBookmarkTextPreservesMarkers()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:bookmarkStart w:id="4" w:name="ClientName"/>
                      <w:r><w:t>Old Client</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-bookmark-text
            target M.B0001
            text New Client
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("Before New Client After", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("w:bookmarkStart w:id=\"4\" w:name=\"ClientName\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:bookmarkEnd w:id=\"4\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("Old Client", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyReplaceBookmarkTextHandlesMultiRunSameParagraphRange()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:bookmarkStart w:id="4" w:name="ClientName"/>
                      <w:r><w:t>Old </w:t></w:r>
                      <w:r><w:t>Client</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-bookmark-text
            target M.B0001
            text New Client
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("Before New Client After", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("w:bookmarkStart w:id=\"4\" w:name=\"ClientName\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:bookmarkEnd w:id=\"4\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("Old ", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:r><w:t>Client</w:t></w:r>", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyReplaceBookmarkTextHandlesMultiParagraphRange()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="4" w:name="ClientName"/>
                      <w:r><w:t>Old first</w:t></w:r>
                    </w:p>
                    <w:p><w:r><w:t>Old middle</w:t></w:r></w:p>
                    <w:p>
                      <w:r><w:t>Old last</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-bookmark-text
            target M.B0001
            text <<<
            New first
            New last
            >>>
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal(["New first", "New last"], read.Paragraphs.Select(paragraph => paragraph.Text).ToArray());
        DocxBookmarkInfo bookmark = Assert.Single(read.Bookmarks);
        Assert.Equal("M.P0001", bookmark.StartTargetId?.ToWireValue());
        Assert.Equal("M.P0002", bookmark.EndTargetId?.ToWireValue());
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("w:bookmarkStart w:id=\"4\" w:name=\"ClientName\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:bookmarkEnd w:id=\"4\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("Old first", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("Old middle", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("Old last", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyReplaceBookmarkTextPreservesMultiParagraphBoundaryText()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:bookmarkStart w:id="4" w:name="ClientName"/>
                      <w:r><w:t>Old first</w:t></w:r>
                    </w:p>
                    <w:p><w:r><w:t>Old middle</w:t></w:r></w:p>
                    <w:p>
                      <w:r><w:t>Old last</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-bookmark-text
            target M.B0001
            text New Client
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("Before New Client After", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("w:bookmarkStart w:id=\"4\" w:name=\"ClientName\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:bookmarkEnd w:id=\"4\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("Old first", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("Old middle", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("Old last", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyRenameBookmarkUpdatesNameAndInternalHyperlinkAnchors()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="4" w:name="OldName"/>
                      <w:r><w:t>Target</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                    </w:p>
                    <w:p>
                      <w:hyperlink w:anchor="OldName"><w:r><w:t>Jump</w:t></w:r></w:hyperlink>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op rename-bookmark
            target M.B0001
            name NewName
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        DocxBookmarkInfo bookmark = Assert.Single(read.Bookmarks);
        Assert.Equal("NewName", bookmark.Name);
        DocxHyperlinkInfo hyperlink = Assert.Single(read.Hyperlinks);
        Assert.Equal("NewName", hyperlink.Anchor);
        Assert.False(hyperlink.IsAnchorMissing);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("w:name=\"NewName\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:anchor=\"NewName\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("OldName", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyDeleteBookmarkRemovesMarkersAndPreservesText()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:bookmarkStart w:id="4" w:name="ClientName"/>
                      <w:r><w:t>Bookmarked</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-bookmark
            target M.B0001
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Empty(read.Bookmarks);
        Assert.Equal("Before Bookmarked After", Assert.Single(read.Paragraphs).Text);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.DoesNotContain("bookmarkStart", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("bookmarkEnd", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyDeleteBookmarkGuardMismatchWritesNothing()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:bookmarkStart w:id="4" w:name="ClientName"/>
                      <w:r><w:t>Old Client</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-bookmark
            target M.B0001
            expect-name StaleName
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static d => d.Code == "E3201");
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public static void CheckDeleteBookmarkGuardMismatchFails()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:bookmarkStart w:id="4" w:name="ClientName"/>
                      <w:r><w:t>Old Client</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op delete-bookmark
            target M.B0001
            expect-name StaleName
            end
            """);

        DocxCheckResult check = new DocxEditor().Check(input, patch);
        Assert.False(check.Success);
        DocxDiagnostic failure = Assert.Single(check.Diagnostics, static d => d.Code == "E3201");
        Assert.Equal("delete-bookmark", failure.HelpTopic);
        Assert.Equal(5, failure.Line);
    }

    [Fact]
    public static void ApplyDeleteBookmarkGuardMatchSucceeds()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:bookmarkStart w:id="4" w:name="ClientName"/>
                      <w:r><w:t>Old Client</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-bookmark
            target M.B0001
            expect-name ClientName
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.DoesNotContain("bookmarkStart", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("bookmarkEnd", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Equal("Before Old Client After", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void CheckRenameBookmarkRejectsDuplicateName()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="1" w:name="First"/>
                      <w:r><w:t>First</w:t></w:r>
                      <w:bookmarkEnd w:id="1"/>
                    </w:p>
                    <w:p>
                      <w:bookmarkStart w:id="2" w:name="Second"/>
                      <w:r><w:t>Second</w:t></w:r>
                      <w:bookmarkEnd w:id="2"/>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op rename-bookmark
            target M.B0001
            name Second
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E4311");
    }

    [Fact]
    public static void ApplyReplaceBookmarkTextHandlesTableSpanningRangeWhenStructureIsPreserved()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:bookmarkStart w:id="4" w:name="ClientName"/>
                      <w:r><w:t>Old start</w:t></w:r>
                    </w:p>
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Old A</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Old B</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:p>
                      <w:r><w:t>Old end</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-bookmark-text
            target M.B0001
            text <<<
            New start
            New A
            New B
            New end
            >>>
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal(["Before New start", "New end After"], read.Paragraphs.Select(paragraph => paragraph.Text).ToArray());
        DocxTableInfo table = Assert.Single(read.Tables);
        Assert.Equal(["New A", "New B"], table.Cells.Select(cell => cell.Text).ToArray());
        DocxBookmarkInfo bookmark = Assert.Single(read.Bookmarks);
        Assert.Equal("M.P0001", bookmark.StartTargetId?.ToWireValue());
        Assert.Equal("M.P0002", bookmark.EndTargetId?.ToWireValue());
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:tbl>", xml, StringComparison.Ordinal);
        Assert.Contains("w:bookmarkStart w:id=\"4\" w:name=\"ClientName\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:bookmarkEnd w:id=\"4\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("Old start", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("Old A", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("Old B", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("Old end", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckReplaceBookmarkTextRejectsTableSpanningLineCountMismatch()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="4" w:name="ClientName"/>
                      <w:r><w:t>Old start</w:t></w:r>
                    </w:p>
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Old A</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Old B</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:p>
                      <w:r><w:t>Old end</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-bookmark-text
            target M.B0001
            text One line
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal("E4311", diagnostic.Code);
        Assert.Contains("requires 4 replacement lines", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckReplaceBookmarkTextRejectsTableSpanningProtectedBoundary()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="4" w:name="ClientName"/>
                      <w:r><w:t>Old start</w:t></w:r>
                    </w:p>
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:hyperlink w:anchor="Target"><w:r><w:t>Old link</w:t></w:r></w:hyperlink></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:p>
                      <w:r><w:t>Old end</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-bookmark-text
            target M.B0001
            text <<<
            New start
            New link
            New end
            >>>
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal("E4311", diagnostic.Code);
        Assert.Contains("protected OOXML boundary 'hyperlink'", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyReplaceBookmarkTextWithNameSelector()
    {
        using MemoryStream input = CreateDocxWithSingleBookmark();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-bookmark-text
            target bookmark:"ClientName"
            text New Client
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Equal("Before New Client After", new DocxEditor().Read(output).Paragraphs[0].Text);
    }

    [Fact]
    public static void CheckReplaceBookmarkTextWithUnknownNameFails()
    {
        using MemoryStream input = CreateDocxWithSingleBookmark();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-bookmark-text
            target bookmark:"Missing"
            text New Client
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }

    [Fact]
    public static void CheckReplaceBookmarkTextWithDuplicateNameFails()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="1" w:name="ClientName"/>
                      <w:r><w:t>First</w:t></w:r></w:p>
                    <w:p>
                      <w:bookmarkStart w:id="2" w:name="ClientName"/>
                      <w:r><w:t>Second</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-bookmark-text
            target bookmark:"ClientName"
            text New
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, static d => d.Code == "E1202");
        Assert.Contains("M.B0001", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("M.B0002", diagnostic.Message, StringComparison.Ordinal);
    }


    [Fact]
    public static void CheckAddBookmarkUnderRequireStillFails()
    {
        var options = new DocxEditOptions { TrackChanges = TrackChangesMode.Require };
        using MemoryStream input = CreateDocxWithBody("""
              <w:p><w:r><w:t>Client paragraph</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op add-bookmark
            target M.P0001
            expect-text Client paragraph
            name ClientParagraph
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E6001");
    }


    [Fact]
    public static void ApplyRenameBookmarkWithMatchingExpectNameSucceeds()
    {
        using MemoryStream input = CreateDocxWithSingleBookmark();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op rename-bookmark
            target M.B0001
            expect-name ClientName
            name RenamedClient
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Equal("RenamedClient", Assert.Single(new DocxEditor().Read(output).Bookmarks).Name);
    }

    [Fact]
    public static void CheckRenameBookmarkWithMismatchedExpectNameFails()
    {
        using MemoryStream input = CreateDocxWithSingleBookmark();
        using var patch = new StringReader("""
            docxpatch 1

            op rename-bookmark
            target M.B0001
            expect-name StaleName
            name RenamedClient
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, static d => d.Code == "E3201");
        Assert.Equal(5, diagnostic.Line);
        Assert.Equal(1, diagnostic.Column);
    }

    [Fact]
    public static void CheckReplaceBookmarkTextGuardMismatchFails()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:bookmarkStart w:id="4" w:name="ClientName"/>
                      <w:r><w:t>Old Client</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-bookmark-text
            target M.B0001
            expect-text Stale
            text New Client
            end
            """);

        DocxCheckResult check = new DocxEditor().Check(input, patch);
        Assert.False(check.Success);
        DocxDiagnostic failure = Assert.Single(check.Diagnostics, static d => d.Code == "E3201");
        Assert.Equal("replace-bookmark-text", failure.HelpTopic);
        Assert.Equal(5, failure.Line);

        using MemoryStream applyInput = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:bookmarkStart w:id="4" w:name="ClientName"/>
                      <w:r><w:t>Old Client</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var applyPatch = new StringReader("""
            docxpatch 1

            op replace-bookmark-text
            target M.B0001
            expect-text Stale
            text New Client
            end
            """);
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, applyPatch, output);
        Assert.False(apply.Success);
        Assert.Contains(apply.Diagnostics, static d => d.Code == "E3201");
    }

    [Fact]
    public static void ApplyReplaceBookmarkTextGuardMatchSucceeds()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:bookmarkStart w:id="4" w:name="ClientName"/>
                      <w:r><w:t>Old Client</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-bookmark-text
            target M.B0001
            expect-text Old Client
            text New Client
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("Before New Client After", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void DeleteDoesNotRetargetLaterExplicitBookmark()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="7" w:name="First"/>
                      <w:r><w:t>Alpha</w:t></w:r>
                      <w:bookmarkEnd w:id="7"/>
                    </w:p>
                    <w:p>
                      <w:bookmarkStart w:id="8" w:name="Second"/>
                      <w:r><w:t>Beta</w:t></w:r>
                      <w:bookmarkEnd w:id="8"/>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-block
            target M.P0001
            end

            op replace-bookmark-text
            target M.B0002
            text Changed
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        output.Position = 0;
        Assert.Equal("Changed", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }
}
