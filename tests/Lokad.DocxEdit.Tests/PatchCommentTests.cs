using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml.Linq;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

public static class PatchCommentTests
{

    [Fact]
    public static void ApplyAddCommentCreatesCommentsPartAndAnchorsParagraph()
    {
        using MemoryStream input = CreateDocx("Anchor paragraph.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op add-comment
            target M.P0001
            expect-text Anchor paragraph.
            text Review note
            author Reviewer
            initials RV
            date 2026-06-07T12:00:00Z
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxChangeInfo comment = Assert.Single(
            new DocxEditor().Changes(output, new DocxChangesOptions { IncludeCommentText = true }).Changes,
            change => change.Type == "comment");
        Assert.Equal("0", comment.CommentId);
        Assert.Equal("Reviewer", comment.CommentAuthor);
        Assert.Equal("RV", comment.CommentInitials);
        Assert.Equal("Review note", comment.CommentTextSnippet);
        Assert.Equal("M.P0001", comment.CommentAnchorTargetId);
        Assert.Equal("M.P0001", comment.CommentReferenceTargetId);

        output.Position = 0;
        string documentXml = ReadDocumentXml(output);
        Assert.Contains("<w:commentRangeStart w:id=\"0\"", documentXml, StringComparison.Ordinal);
        Assert.Contains("<w:commentRangeEnd w:id=\"0\"", documentXml, StringComparison.Ordinal);
        Assert.Contains("<w:commentReference w:id=\"0\"", documentXml, StringComparison.Ordinal);

        output.Position = 0;
        string commentsXml = ReadEntry(output, "word/comments.xml");
        Assert.Contains("w:comment w:id=\"0\"", commentsXml, StringComparison.Ordinal);
        Assert.Contains("w:author=\"Reviewer\"", commentsXml, StringComparison.Ordinal);
        Assert.Contains("Review note", commentsXml, StringComparison.Ordinal);

        output.Position = 0;
        Assert.Contains("relationships/comments", ReadEntry(output, "word/_rels/document.xml.rels"), StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains("wordprocessingml.comments+xml", ReadEntry(output, "[Content_Types].xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyAddCommentCanAnchorSelectedParagraphText()
    {
        using MemoryStream input = CreateDocx("Alpha beta gamma.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op add-comment
            target M.P0001
            expect-text Alpha beta gamma.
            anchor-text beta
            text Review beta only
            author Reviewer
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        string documentXml = ReadDocumentXml(output);
        int alphaIndex = documentXml.IndexOf(">Alpha ", StringComparison.Ordinal);
        int startIndex = documentXml.IndexOf("<w:commentRangeStart w:id=\"0\"", StringComparison.Ordinal);
        int betaIndex = documentXml.IndexOf(">beta<", StringComparison.Ordinal);
        int endIndex = documentXml.IndexOf("<w:commentRangeEnd w:id=\"0\"", StringComparison.Ordinal);
        int gammaIndex = documentXml.IndexOf("> gamma.<", StringComparison.Ordinal);
        Assert.True(alphaIndex < startIndex);
        Assert.True(startIndex < betaIndex);
        Assert.True(betaIndex < endIndex);
        Assert.True(endIndex < gammaIndex);
        Assert.Contains("<w:commentReference w:id=\"0\"", documentXml, StringComparison.Ordinal);
        output.Position = 0;
        DocxChangeInfo comment = Assert.Single(
            new DocxEditor().Changes(output, new DocxChangesOptions { IncludeCommentText = true }).Changes,
            change => change.Type == "comment");
        Assert.Equal("M.P0001", comment.CommentAnchorTargetId);
        Assert.Equal("Review beta only", comment.CommentTextSnippet);
        output.Position = 0;
        Assert.True(new DocxEditor().Validate(output).Success);
    }

    [Fact]
    public static void ApplyAddCommentSelectedTextRequiresOccurrenceWhenAmbiguous()
    {
        using MemoryStream input = CreateDocx("Alpha beta beta.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op add-comment
            target M.P0001
            anchor-text beta
            text Review beta
            author Reviewer
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E1202" &&
            diagnostic.Message.Contains("Specify occurrence", StringComparison.Ordinal));
    }

    [Fact]
    public static void ApplyAddCommentSelectedTextRejectsProtectedMarkupBoundaries()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="1" w:name="B"/>
                      <w:r><w:t>Alpha beta.</w:t></w:r>
                      <w:bookmarkEnd w:id="1"/>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op add-comment
            target M.P0001
            anchor-text beta
            text Review beta
            author Reviewer
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E4305" &&
            diagnostic.Message.Contains("protected OOXML boundary 'bookmark'", StringComparison.Ordinal));
    }

    [Fact]
    public static void ApplyAddCommentAllocatesNextExistingCommentId()
    {
        using MemoryStream input = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="3" w:author="Reviewer">
                    <w:p><w:r><w:t>Existing comment</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op add-comment
            target M.P0001
            text Follow-up note
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(
            input,
            patch,
            output,
            new DocxEditOptions
            {
                Author = "Second Reviewer",
                TimestampUtc = DateTimeOffset.Parse("2026-06-08T09:30:00Z").ToUniversalTime()
            });

        Assert.True(result.Success);
        output.Position = 0;
        IReadOnlyList<DocxCommentThreadSummary> summaries = new DocxEditor()
            .Changes(output, new DocxChangesOptions { IncludeCommentText = true })
            .CommentSummary;
        Assert.Contains(summaries, summary => summary.CommentId == "3" && summary.TextSnippet == "Existing comment");
        Assert.Contains(summaries, summary => summary.CommentId == "4" && summary.TextSnippet == "Follow-up note");

        output.Position = 0;
        string commentsXml = ReadEntry(output, "word/comments.xml");
        Assert.Contains("w:comment w:id=\"3\"", commentsXml, StringComparison.Ordinal);
        Assert.Contains("w:comment w:id=\"4\"", commentsXml, StringComparison.Ordinal);
        Assert.Contains("w:author=\"Second Reviewer\"", commentsXml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetCommentTextPreservesCommentMetadata()
    {
        using MemoryStream input = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="3" w:author="Reviewer" w:initials="RV" w:date="2026-06-07T12:00:00Z">
                    <w:p><w:r><w:t>Old comment</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-comment-text
            target comment:3
            text Updated comment
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxChangeInfo comment = Assert.Single(
            new DocxEditor().Changes(output, new DocxChangesOptions { IncludeCommentText = true }).Changes,
            change => change.Type == "comment");
        Assert.Equal("3", comment.CommentId);
        Assert.Equal("Reviewer", comment.CommentAuthor);
        Assert.Equal("RV", comment.CommentInitials);
        Assert.Equal("Updated comment", comment.CommentTextSnippet);
        output.Position = 0;
        string commentsXml = ReadEntry(output, "word/comments.xml");
        Assert.Contains("w:author=\"Reviewer\"", commentsXml, StringComparison.Ordinal);
        Assert.Contains("Updated comment", commentsXml, StringComparison.Ordinal);
        Assert.DoesNotContain("Old comment", commentsXml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyResolveAndReopenCommentUpdatesExtendedMetadata()
    {
        using MemoryStream input = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w:comment w:id="3" w:author="Reviewer">
                    <w:p w15:paraId="00ABCDEF"><w:r><w:t>Comment body</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """,
            """
                <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w15:commentEx w15:paraId="00ABCDEF" w15:done="0"/>
                </w15:commentsEx>
                """, null);
        using var resolvedOutput = new MemoryStream();
        using var resolvePatch = new StringReader("""
            docxpatch 1

            op resolve-comment
            target comment:3
            end
            """);

        DocxApplyResult resolveResult = new DocxEditor().Apply(input, resolvePatch, resolvedOutput);

        Assert.True(resolveResult.Success);
        resolvedOutput.Position = 0;
        Assert.True(Assert.Single(new DocxEditor().Changes(resolvedOutput).CommentSummary).Resolved);
        resolvedOutput.Position = 0;
        Assert.Contains("w15:done=\"1\"", ReadEntry(resolvedOutput, "word/commentsExtended.xml"), StringComparison.Ordinal);

        resolvedOutput.Position = 0;
        using var reopenedOutput = new MemoryStream();
        using var reopenPatch = new StringReader("""
            docxpatch 1

            op reopen-comment
            target C001.C0001
            end
            """);

        DocxApplyResult reopenResult = new DocxEditor().Apply(resolvedOutput, reopenPatch, reopenedOutput);

        Assert.True(reopenResult.Success);
        reopenedOutput.Position = 0;
        Assert.False(Assert.Single(new DocxEditor().Changes(reopenedOutput).CommentSummary).Resolved);
        reopenedOutput.Position = 0;
        Assert.Contains("w15:done=\"0\"", ReadEntry(reopenedOutput, "word/commentsExtended.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyResolveCommentCreatesExtendedMetadataWhenMissing()
    {
        using MemoryStream input = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="3" w:author="Reviewer">
                    <w:p><w:r><w:t>Comment body</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op resolve-comment
            target comment:3
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxCommentThreadSummary summary = Assert.Single(new DocxEditor().Changes(output).CommentSummary);
        Assert.Equal("3", summary.CommentId);
        Assert.True(summary.Resolved);
        Assert.Equal("00000001", summary.ParaId);

        output.Position = 0;
        string commentsXml = ReadEntry(output, "word/comments.xml");
        Assert.Contains("w15:paraId=\"00000001\"", commentsXml, StringComparison.Ordinal);

        output.Position = 0;
        string commentsExtendedXml = ReadEntry(output, "word/commentsExtended.xml");
        Assert.Contains("w15:commentEx w15:paraId=\"00000001\" w15:done=\"1\"", commentsExtendedXml, StringComparison.Ordinal);

        output.Position = 0;
        Assert.Contains("relationships/commentsExtended", ReadEntry(output, "word/_rels/document.xml.rels"), StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains("vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml", ReadEntry(output, "[Content_Types].xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyReopenCommentCreatesUnresolvedExtendedMetadataWhenMissing()
    {
        using MemoryStream input = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="3" w:author="Reviewer">
                    <w:p><w:r><w:t>Comment body</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op reopen-comment
            target comment:3
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxCommentThreadSummary summary = Assert.Single(new DocxEditor().Changes(output).CommentSummary);
        Assert.False(summary.Resolved);
        output.Position = 0;
        Assert.Contains("w15:done=\"0\"", ReadEntry(output, "word/commentsExtended.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyDeleteCommentRemovesExtendedMetadata()
    {
        using MemoryStream input = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w:comment w:id="3" w:author="Reviewer">
                    <w:p w15:paraId="00ABCDEF"><w:r><w:t>Comment body</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """,
            """
                <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w15:commentEx w15:paraId="00ABCDEF" w15:done="1"/>
                </w15:commentsEx>
                """, null);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-comment
            target comment:3
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Empty(new DocxEditor().Changes(output).Changes);
        output.Position = 0;
        string commentsExtendedXml = ReadEntry(output, "word/commentsExtended.xml");
        Assert.DoesNotContain("00ABCDEF", commentsExtendedXml, StringComparison.Ordinal);
        Assert.DoesNotContain("commentEx", commentsExtendedXml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyDeleteCommentRemovesCommentsIdsMetadata()
    {
        using MemoryStream input = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w:comment w:id="3" w:author="Reviewer">
                    <w:p w15:paraId="00ABCDEF"><w:r><w:t>Comment body</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """,
            commentsExtendedXml: null,
            commentsIdsXml: """
                <w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">
                  <w16cid:commentId w16cid:paraId="00ABCDEF" w16cid:durableId="7F0A11BC"/>
                </w16cid:commentsIds>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-comment
            target comment:3
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        string commentsIdsXml = ReadEntry(output, "word/commentsIds.xml");
        Assert.DoesNotContain("00ABCDEF", commentsIdsXml, StringComparison.Ordinal);
        Assert.DoesNotContain("commentId", commentsIdsXml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.True(new DocxEditor().Validate(output).Success);
    }

    [Fact]
    public static void ApplyDeleteCommentRemovesBodyAndAnchors()
    {
        using MemoryStream input = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="3" w:author="Reviewer" w:initials="RV" w:date="2026-06-07T12:00:00Z">
                    <w:p><w:r><w:t>Old comment</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-comment
            target C001.C0001
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Empty(new DocxEditor().Changes(output).Changes);
        output.Position = 0;
        string documentXml = ReadDocumentXml(output);
        Assert.DoesNotContain("commentRangeStart", documentXml, StringComparison.Ordinal);
        Assert.DoesNotContain("commentRangeEnd", documentXml, StringComparison.Ordinal);
        Assert.DoesNotContain("commentReference", documentXml, StringComparison.Ordinal);
        output.Position = 0;
        string commentsXml = ReadEntry(output, "word/comments.xml");
        Assert.DoesNotContain("w:comment w:id=\"3\"", commentsXml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Equal("Commented", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void ApplyDeleteCommentRemovesHeaderAnchorsAndExtendedMetadata()
    {
        using MemoryStream input = CreateDocxWithHeaderComment();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-comment
            target comment:3
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        string headerXml = ReadEntry(output, "word/header1.xml");
        Assert.DoesNotContain("commentRangeStart", headerXml, StringComparison.Ordinal);
        Assert.DoesNotContain("commentRangeEnd", headerXml, StringComparison.Ordinal);
        Assert.DoesNotContain("commentReference", headerXml, StringComparison.Ordinal);
        Assert.Contains("Header text", headerXml, StringComparison.Ordinal);
        output.Position = 0;
        string commentsExtendedXml = ReadEntry(output, "word/commentsExtended.xml");
        Assert.DoesNotContain("00ABCDEF", commentsExtendedXml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyAddCommentReplyCreatesThreadMetadata()
    {
        using MemoryStream input = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w:comment w:id="3" w:author="Reviewer">
                    <w:p w15:paraId="00ABCDEF"><w:r><w:t>Comment body</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """,
            """
                <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w15:commentEx w15:paraId="00ABCDEF" w15:done="0"/>
                </w15:commentsEx>
                """, null);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op add-comment-reply
            target comment:3
            text Reply
            author Second Reviewer
            initials SR
            date 2026-06-08T09:30:00Z
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        IReadOnlyList<DocxCommentThreadSummary> comments = new DocxEditor()
            .Changes(output, new DocxChangesOptions { IncludeCommentText = true })
            .CommentSummary;
        DocxCommentThreadSummary reply = Assert.Single(comments, comment => comment.CommentId == "4");
        Assert.Equal("00ABCDEF", reply.ParentParaId);
        Assert.Equal("00ABCDEF", reply.RootParaId);
        Assert.Equal("Reply", reply.TextSnippet);
        Assert.True(reply.IsReply);
        Assert.Equal("00000001", reply.DurableId);

        output.Position = 0;
        string commentsXml = ReadEntry(output, "word/comments.xml");
        Assert.Contains("w:comment w:id=\"4\"", commentsXml, StringComparison.Ordinal);
        Assert.Contains("w15:paraId=\"00000001\"", commentsXml, StringComparison.Ordinal);
        Assert.Contains("w:author=\"Second Reviewer\"", commentsXml, StringComparison.Ordinal);
        Assert.Contains("w:initials=\"SR\"", commentsXml, StringComparison.Ordinal);

        output.Position = 0;
        string commentsExtendedXml = ReadEntry(output, "word/commentsExtended.xml");
        Assert.Contains("w15:commentEx w15:paraId=\"00000001\" w15:paraIdParent=\"00ABCDEF\" w15:done=\"0\"", commentsExtendedXml, StringComparison.Ordinal);

        output.Position = 0;
        string commentsIdsXml = ReadEntry(output, "word/commentsIds.xml");
        Assert.Contains("w16cid:commentId w16cid:paraId=\"00000001\" w16cid:durableId=\"00000001\"", commentsIdsXml, StringComparison.Ordinal);

        output.Position = 0;
        Assert.True(new DocxEditor().Validate(output).Success);
    }

    [Fact]
    public static void ApplyDeleteCommentReplyRemovesReplyMetadataOnly()
    {
        using MemoryStream input = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w:comment w:id="3" w:author="Reviewer">
                    <w:p w15:paraId="00PARENT"><w:r><w:t>Comment body</w:t></w:r></w:p>
                  </w:comment>
                  <w:comment w:id="4" w:author="Second Reviewer">
                    <w:p w15:paraId="00REPLY1"><w:r><w:t>Reply body</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """,
            """
                <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w15:commentEx w15:paraId="00PARENT" w15:done="0"/>
                  <w15:commentEx w15:paraId="00REPLY1" w15:paraIdParent="00PARENT" w15:done="0"/>
                </w15:commentsEx>
                """,
            """
                <w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">
                  <w16cid:commentId w16cid:paraId="00REPLY1" w16cid:durableId="7F0A11BC"/>
                </w16cid:commentsIds>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-comment-reply
            target comment:3.reply:1
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        IReadOnlyList<DocxCommentThreadSummary> comments = new DocxEditor().Changes(output).CommentSummary;
        Assert.Single(comments);
        Assert.Equal("3", comments[0].CommentId);

        output.Position = 0;
        string commentsXml = ReadEntry(output, "word/comments.xml");
        Assert.Contains("w:comment w:id=\"3\"", commentsXml, StringComparison.Ordinal);
        Assert.DoesNotContain("w:comment w:id=\"4\"", commentsXml, StringComparison.Ordinal);

        output.Position = 0;
        string commentsExtendedXml = ReadEntry(output, "word/commentsExtended.xml");
        Assert.Contains("00PARENT", commentsExtendedXml, StringComparison.Ordinal);
        Assert.DoesNotContain("00REPLY1", commentsExtendedXml, StringComparison.Ordinal);

        output.Position = 0;
        string commentsIdsXml = ReadEntry(output, "word/commentsIds.xml");
        Assert.DoesNotContain("00REPLY1", commentsIdsXml, StringComparison.Ordinal);
        Assert.DoesNotContain("7F0A11BC", commentsIdsXml, StringComparison.Ordinal);

        output.Position = 0;
        Assert.True(new DocxEditor().Validate(output).Success);
    }

    [Fact]
    public static void ApplyTrackedTextChangePlusCommentUnderRequireSucceeds()
    {
        var options = new DocxEditOptions { TrackChanges = TrackChangesMode.Require };
        using MemoryStream input = CreateDocx("Anchor paragraph.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Anchor
            with Edited
            end

            op add-comment
            target M.P0001
            text Review note
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Empty(Assert.Single(result.Operations, static o => o.OperationName == "add-comment").GeneratedRevisionIds);
        Assert.NotEmpty(Assert.Single(result.Operations, static o => o.OperationName == "replace-text").GeneratedRevisionIds);
        output.Position = 0;
        string documentXml = ReadDocumentXml(output);
        Assert.Contains("<w:del", documentXml, StringComparison.Ordinal);
        Assert.Contains("<w:ins", documentXml, StringComparison.Ordinal);
        Assert.Contains("<w:commentRangeStart", documentXml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains("Review note", ReadEntry(output, "word/comments.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckTrackedTextChangePlusCommentUnderRequireSucceeds()
    {
        var options = new DocxEditOptions { TrackChanges = TrackChangesMode.Require };
        using MemoryStream input = CreateDocx("Anchor paragraph.");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Anchor
            with Edited
            end

            op add-comment
            target M.P0001
            text Review note
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
    }

    [Fact]
    public static void ApplyAddCommentAloneUnderRequireSucceeds()
    {
        var options = new DocxEditOptions { TrackChanges = TrackChangesMode.Require };
        using MemoryStream input = CreateDocx("Anchor paragraph.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op add-comment
            target M.P0001
            text Review note
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, options);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Contains("Review note", ReadEntry(output, "word/comments.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckSetCommentTextGuardMismatchFails()
    {
        using MemoryStream input = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="3" w:author="Reviewer" w:initials="RV" w:date="2026-06-07T12:00:00Z">
                    <w:p><w:r><w:t>Old comment</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-comment-text
            target comment:3
            expect-text Stale comment
            text Updated comment
            end
            """);

        DocxCheckResult check = new DocxEditor().Check(input, patch);
        Assert.False(check.Success);
        DocxDiagnostic failure = Assert.Single(check.Diagnostics, static d => d.Code == "E3201");
        Assert.Equal("set-comment-text", failure.HelpTopic);
        Assert.Equal(5, failure.Line);
    }

    [Fact]
    public static void ApplySetCommentTextGuardMatchSucceeds()
    {
        using MemoryStream input = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="3" w:author="Reviewer" w:initials="RV" w:date="2026-06-07T12:00:00Z">
                    <w:p><w:r><w:t>Old comment</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-comment-text
            target comment:3
            expect-text Old comment
            text Updated comment
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        string commentsXml = ReadEntry(output, "word/comments.xml");
        Assert.Contains("Updated comment", commentsXml, StringComparison.Ordinal);
        Assert.DoesNotContain("Old comment", commentsXml, StringComparison.Ordinal);
    }
    [Fact]
    public static void ApplyDeleteCommentGuardMismatchWritesNothing()
    {
        using MemoryStream input = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="3" w:author="Reviewer" w:initials="RV" w:date="2026-06-07T12:00:00Z">
                    <w:p><w:r><w:t>Old comment</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-comment
            target comment:3
            expect-text Stale comment
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static d => d.Code == "E3201");
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public static void CheckDeleteCommentGuardMismatchFails()
    {
        using MemoryStream input = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="3" w:author="Reviewer" w:initials="RV" w:date="2026-06-07T12:00:00Z">
                    <w:p><w:r><w:t>Old comment</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        using var patch = new StringReader("""
            docxpatch 1

            op delete-comment
            target comment:3
            expect-text Stale comment
            end
            """);

        DocxCheckResult check = new DocxEditor().Check(input, patch);
        Assert.False(check.Success);
        DocxDiagnostic failure = Assert.Single(check.Diagnostics, static d => d.Code == "E3201");
        Assert.Equal("delete-comment", failure.HelpTopic);
        Assert.Equal(5, failure.Line);
    }

    [Fact]
    public static void ApplyDeleteCommentGuardMatchSucceeds()
    {
        using MemoryStream input = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="3" w:author="Reviewer" w:initials="RV" w:date="2026-06-07T12:00:00Z">
                    <w:p><w:r><w:t>Old comment</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-comment
            target comment:3
            expect-text Old comment
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        string commentsXml = ReadEntry(output, "word/comments.xml");
        Assert.DoesNotContain("w:comment w:id=\"3\"", commentsXml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Equal("Commented", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void CheckTrackChangesRequirePermitsCommentReviewOperations()
    {
        var options = new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Require,
            Author = "Agent",
            TimestampUtc = new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero)
        };
        string[] patches =
        [
            "docxpatch 1\n\nop resolve-comment\ntarget comment:3\nend\n",
            "docxpatch 1\n\nop reopen-comment\ntarget comment:3\nend\n",
            "docxpatch 1\n\nop delete-comment\ntarget comment:3\nend\n",
        ];
        foreach (string patchText in patches)
        {
            using MemoryStream input = CreateDocxWithCommentAnchoredParagraph();
            using var patch = new StringReader(patchText);
            DocxCheckResult result = new DocxEditor().Check(input, patch, options);
            Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
            Assert.DoesNotContain(result.Diagnostics, static diagnostic => diagnostic.Code == "E6001");
        }
    }
}
