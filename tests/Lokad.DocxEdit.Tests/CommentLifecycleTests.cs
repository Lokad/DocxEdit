using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

public static class CommentLifecycleTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

    internal static MemoryStream CreateThread()
    {
        using var input = new MemoryStream();
        Assert.True(new DocxEditor().Create(input).Success);
        return Apply(input, """
            docxpatch 1
            op replace-paragraph
            target M.P0001
            text Alpha beta gamma.
            end
            op add-comment
            target M.P0001
            anchor-text beta
            text Please clarify beta.
            as review
            end
            op add-comment-reply
            target @review
            text Clarification is pending.
            end
            op resolve-comment
            target @review
            end
            """);
    }

    internal static MemoryStream Apply(MemoryStream input, string patch, TrackChangesMode mode = TrackChangesMode.Off)
    {
        input.Position = 0;
        var output = new MemoryStream();
        DocxApplyResult result = new DocxEditor().Apply(input, new StringReader(patch), output,
            new DocxEditOptions { TrackChanges = mode });
        Assert.True(result.Success, string.Join("; ", result.Diagnostics));
        output.Position = 0;
        return output;
    }

    internal static void AssertValid(MemoryStream document)
    {
        document.Position = 0;
        DocxValidateResult result = new DocxEditor().Validate(document);
        Assert.True(result.Success, string.Join("; ", result.Diagnostics));
        document.Position = 0;
        using var package = WordprocessingDocument.Open(document, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(package).ToArray();
        Assert.True(errors.Length == 0, string.Join("; ", errors.Select(e => e.Description + " " + e.Path?.XPath)));
    }

    [Fact]
    public static void CommentThreadUsesSchemaValidMetadata()
    {
        using MemoryStream thread = CreateThread();
        AssertValid(thread);
        thread.Position = 0;
        var comments = new DocxEditor().Changes(thread).CommentSummary;
        Assert.All(comments, comment => Assert.Equal("M.P0001", comment.AnchorTargetId));
        Assert.Equal(2, comments.Select(c => c.DurableId).Distinct().Count());
        thread.Position = 0;
        XDocument body = XDocument.Parse(ReadDocumentXml(thread));
        Assert.Equal(2, body.Descendants(W + "commentReference").Count());
        Assert.Equal(2, body.Descendants(W + "commentRangeStart").Count());
        Assert.Equal(2, body.Descendants(W + "commentRangeEnd").Count());
    }

    [Fact]
    public static void RepliesToRepliesAreRefused()
    {
        using MemoryStream thread = CreateThread();
        var result = new DocxEditor().Check(thread, new StringReader("docxpatch 1\nop add-comment-reply\ntarget comment:1\ntext Nested reply\nend\n"));
        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, d => d.Code == "E4314");
    }

    [Fact]
    public static void MultipleRepliesKeepInsertionOrderAndDistinctAnchors()
    {
        using MemoryStream thread = CreateThread();
        using MemoryStream added = Apply(thread, "docxpatch 1\nop add-comment-reply\ntarget comment:0\ntext Second reply.\nend\n");
        AssertValid(added);
        added.Position = 0;
        XDocument body = XDocument.Parse(ReadDocumentXml(added));
        Assert.Equal(new[] { "0", "1", "2" }, body.Descendants(W + "commentRangeStart").Select(e => (string?)e.Attribute(W + "id")));
        using MemoryStream deleted = Apply(added, "docxpatch 1\nop delete-comment-reply\ntarget comment:0.reply:1\nend\n");
        AssertValid(deleted);
        deleted.Position = 0;
        var remaining = new DocxEditor().Changes(deleted, new() { IncludeCommentText = true }).CommentSummary;
        Assert.Equal("Second reply.", Assert.Single(remaining, c => c.IsReply == true).TextSnippet);
    }

    [Theory]
    [InlineData("comment:0", TrackChangesMode.Off)]
    [InlineData("comment:1", TrackChangesMode.Off)]
    [InlineData("comment:0", TrackChangesMode.Require)]
    [InlineData("comment:1", TrackChangesMode.Require)]
    public static void EditingCommentKeepsThreadAndResolution(string target, TrackChangesMode mode)
    {
        using MemoryStream thread = CreateThread();
        using MemoryStream edited = Apply(thread, $"docxpatch 1\nop set-comment-text\ntarget {target}\ntext Updated note.\nend\n", mode);
        AssertValid(edited);
        edited.Position = 0;
        var comments = new DocxEditor().Changes(edited, new() { IncludeCommentText = true }).CommentSummary;
        Assert.Equal(2, comments.Count);
        Assert.True(Assert.Single(comments, c => c.CommentId == "0").Resolved);
        Assert.True(Assert.Single(comments, c => c.CommentId == "1").IsReply);
        Assert.Contains("Updated note.", Assert.Single(comments, c => "comment:" + c.CommentId == target).TextSnippet);
    }

    [Fact]
    public static void DeleteParentWithRepliesRefusesWithoutOutput()
    {
        using MemoryStream thread = CreateThread();
        const string patch = "docxpatch 1\nop delete-comment\ntarget comment:0\nend\n";
        var check = new DocxEditor().Check(thread, new StringReader(patch));
        Assert.False(check.Success);
        Assert.Contains(check.Diagnostics, d => d.Code == "E4314");
        thread.Position = 0;
        using var output = new MemoryStream();
        var result = new DocxEditor().Apply(thread, new StringReader(patch), output);
        Assert.False(result.Success);
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public static void DeleteLeafThenParentRemovesAllThreadMetadata()
    {
        using MemoryStream thread = CreateThread();
        using MemoryStream deleted = Apply(thread, """
            docxpatch 1
            op delete-comment-reply
            target comment:0.reply:1
            end
            op delete-comment
            target comment:0
            end
            """);
        AssertValid(deleted);
        deleted.Position = 0;
        Assert.Empty(new DocxEditor().Changes(deleted).CommentSummary);
    }

    [Fact]
    public static void MultiParagraphCommentUsesLastParagraphIdentity()
    {
        using MemoryStream thread = CreateThread();
        // Reproduce a Word-authored multi-paragraph body: the last paragraph
        // retains the ID referenced by commentsExtended and commentsIds.
        using (var zip = new System.IO.Compression.ZipArchive(thread, System.IO.Compression.ZipArchiveMode.Update, true))
        {
            var entry = zip.GetEntry("word/comments.xml")!;
            XDocument xml;
            using (Stream stream = entry.Open()) xml = XDocument.Load(stream);
            XElement parent = xml.Root!.Elements(W + "comment").First();
            parent.AddFirst(new XElement(W + "p", new XAttribute(W14 + "paraId", "7FFFFFFE"),
                new XElement(W + "r", new XElement(W + "t", "First paragraph."))));
            using (Stream stream = entry.Open()) { stream.SetLength(0); xml.Save(stream); }
        }
        thread.Position = 0;
        var before = new DocxEditor().Changes(thread).CommentSummary;
        Assert.True(Assert.Single(before, c => c.CommentId == "0").Resolved);
        using MemoryStream edited = Apply(thread, "docxpatch 1\nop set-comment-text\ntarget comment:0\ntext Replaced body.\nend\n");
        AssertValid(edited);
        edited.Position = 0;
        Assert.True(Assert.Single(new DocxEditor().Changes(edited).CommentSummary, c => c.CommentId == "0").Resolved);
    }
}
