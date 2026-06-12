using System.Text;

namespace DocxEdit;

public static class DocxTextRenderer
{
    public static string RenderRead(DocxReadResult result)
    {
        ArgumentNullException.ThrowIfNull(result);
        return result.Text;
    }

    public static string RenderReadSummary(DocxReadResult result)
    {
        ArgumentNullException.ThrowIfNull(result);
        var builder = new StringBuilder();
        builder.Append("parts count=").Append(result.PartNames.Count).AppendLine();
        builder.Append("main-document-part=").AppendLine(result.MainDocumentPartName ?? "unknown");
        builder.Append("paragraphs count=").Append(result.Paragraphs.Count).AppendLine();
        builder.Append("tables count=").Append(result.Tables.Count).AppendLine();
        builder.Append("images count=").Append(result.Images.Count).AppendLine();
        builder.Append("sections count=").Append(result.Sections.Count).AppendLine();
        builder.Append("bookmarks count=").Append(result.Bookmarks.Count).AppendLine();
        builder.Append("content-controls count=").Append(result.ContentControls.Count).AppendLine();
        builder.Append("fields count=").Append(result.Fields.Count).AppendLine();
        foreach (IGrouping<string, DocxParagraphInfo> group in result.Paragraphs
                     .GroupBy(paragraph => paragraph.Story)
                     .OrderBy(group => group.Key, StringComparer.Ordinal))
        {
            builder.Append("story=\"")
                .Append(EscapeText(group.Key))
                .Append("\" paragraphs=")
                .Append(group.Count())
                .AppendLine();
        }

        return builder.ToString();
    }

    public static string RenderOutline(DocxOutlineResult result)
    {
        ArgumentNullException.ThrowIfNull(result);
        return RenderLines(result.Lines);
    }

    public static string RenderFind(DocxFindResult result)
    {
        ArgumentNullException.ThrowIfNull(result);
        return RenderLines(result.Matches);
    }

    public static string RenderDump(DocxDumpResult result)
    {
        ArgumentNullException.ThrowIfNull(result);
        if (result.Text is null)
        {
            return string.Empty;
        }

        return result.Text.EndsWith('\n')
            ? result.Text
            : result.Text + Environment.NewLine;
    }

    public static string RenderContext(DocxContextResult result)
    {
        ArgumentNullException.ThrowIfNull(result);
        return result.Text;
    }

    public static string RenderStyles(DocxStylesResult result)
    {
        ArgumentNullException.ThrowIfNull(result);
        var builder = new StringBuilder();
        foreach (DocxStyleInfo style in result.Styles
                     .OrderBy(style => style.Type, StringComparer.Ordinal)
                     .ThenBy(style => style.StyleId, StringComparer.Ordinal))
        {
            string defaultText = style.IsDefault ? " default=true" : string.Empty;
            string basedOn = style.BasedOnStyleId is null ? string.Empty : $" based-on={EscapeText(style.BasedOnStyleId)}";
            string next = style.NextStyleId is null ? string.Empty : $" next={EscapeText(style.NextStyleId)}";
            string linked = style.LinkedStyleId is null ? string.Empty : $" linked={EscapeText(style.LinkedStyleId)}";
            string numbering = style.NumberingId is null
                ? string.Empty
                : $" numbering numId={EscapeText(style.NumberingId)} level={style.NumberingLevel ?? 0}";
            builder.Append(style.Type)
                .Append(" styleId=")
                .Append(style.StyleId)
                .Append(" name=\"")
                .Append(EscapeText(style.Name))
                .Append('"')
                .Append(defaultText)
                .Append(basedOn)
                .Append(next)
                .Append(linked)
                .Append(numbering)
                .AppendLine();
        }

        return builder.ToString();
    }

    public static string RenderMedia(DocxMediaResult result)
    {
        ArgumentNullException.ThrowIfNull(result);
        var builder = new StringBuilder();
        foreach (DocxImageInfo image in result.Images)
        {
            builder.Append(image.Id)
                .Append(' ')
                .Append(image.PartName)
                .Append(' ')
                .Append(image.ContentType ?? "unknown")
                .Append(' ')
                .Append(image.ByteLength)
                .AppendLine(" bytes");
        }

        return builder.ToString();
    }

    public static string RenderChanges(DocxChangesResult result)
    {
        ArgumentNullException.ThrowIfNull(result);
        var builder = new StringBuilder();
        foreach (DocxChangeSummary summary in result.Summary)
        {
            builder.Append(summary.Type).Append(" count=").Append(summary.Count).AppendLine();
        }

        foreach (DocxChangeGroupSummary summary in result.GroupSummary)
        {
            builder.Append("summary group=")
                .Append(summary.Group)
                .Append(" key=\"")
                .Append(EscapeText(summary.Key))
                .Append("\" type=")
                .Append(summary.Type)
                .Append(" count=")
                .Append(summary.Count)
                .AppendLine();
        }

        foreach (DocxChangeTargetSummary summary in result.TargetSummary)
        {
            builder.Append("target-summary target=")
                .Append(summary.TargetId)
                .Append(" count=")
                .Append(summary.Count)
                .Append(" types=\"")
                .Append(EscapeText(FormatTypeSummary(summary.Summary)))
                .AppendLine("\"");
        }

        foreach (DocxCommentThreadSummary summary in result.CommentSummary)
        {
            string anchorTarget = summary.AnchorTargetId is null ? "anchor-target=unknown" : $"anchor-target={summary.AnchorTargetId}";
            string referenceTarget = summary.ReferenceTargetId is null ? string.Empty : $" reference-target={summary.ReferenceTargetId}";
            string anchorStory = summary.AnchorStory is null ? string.Empty : $" anchor-story=\"{EscapeText(summary.AnchorStory)}\"";
            string anchorPart = summary.AnchorPartName is null ? string.Empty : $" anchor-part={summary.AnchorPartName}";
            string author = summary.Author is null ? string.Empty : $" author=\"{EscapeText(summary.Author)}\"";
            string timestamp = summary.TimestampUtc is null ? string.Empty : $" timestamp-utc={summary.TimestampUtc:O}";
            string initials = summary.Initials is null ? string.Empty : $" initials=\"{EscapeText(summary.Initials)}\"";
            string textLength = summary.TextLength is null ? string.Empty : $" comment-text-length={summary.TextLength}";
            string textSnippet = summary.TextSnippet is null ? string.Empty : $" comment-text=\"{EscapeText(summary.TextSnippet)}\"";
            string textTruncated = summary.TextTruncated ? " comment-text-truncated=true" : string.Empty;
            builder.Append("comment-summary comment-id=")
                .Append(EscapeText(summary.CommentId))
                .Append(' ')
                .Append(anchorTarget)
                .Append(referenceTarget)
                .Append(anchorStory)
                .Append(anchorPart)
                .Append(author)
                .Append(timestamp)
                .Append(initials)
                .Append(textLength)
                .Append(textSnippet)
                .Append(textTruncated)
                .Append(" count=")
                .Append(summary.Count)
                .Append(" types=\"")
                .Append(EscapeText(FormatTypeSummary(summary.Summary)))
                .AppendLine("\"");
        }

        foreach (DocxChangeInfo change in result.Changes)
        {
            string target = change.TargetId is null ? "target=unknown" : $"target={change.TargetId}";
            string targetSource = $" target-source={change.TargetSource}";
            string targetReason = change.TargetReason is null ? string.Empty : $" target-reason={change.TargetReason}";
            string nearestTarget = change.NearestTargetId is null ? string.Empty : $" nearest-target={change.NearestTargetId}";
            string targetNote = change.TargetNote is null ? string.Empty : $" target-note=\"{EscapeText(change.TargetNote)}\"";
            string pairedChange = change.PairedChangeId is null ? string.Empty : $" paired-change-id={change.PairedChangeId}";
            string revision = change.RevisionId is null ? string.Empty : $" revision-id={EscapeText(change.RevisionId)}";
            string author = change.Author is null ? string.Empty : $" author=\"{EscapeText(change.Author)}\"";
            string timestamp = change.TimestampUtc is null ? string.Empty : $" timestamp-utc={change.TimestampUtc:O}";
            string commentId = change.CommentId is null ? string.Empty : $" comment-id={EscapeText(change.CommentId)}";
            string commentAuthor = change.CommentAuthor is null ? string.Empty : $" comment-author=\"{EscapeText(change.CommentAuthor)}\"";
            string commentTimestamp = change.CommentTimestampUtc is null ? string.Empty : $" comment-timestamp-utc={change.CommentTimestampUtc:O}";
            string commentInitials = change.CommentInitials is null ? string.Empty : $" comment-initials=\"{EscapeText(change.CommentInitials)}\"";
            string commentAnchorTarget = change.CommentAnchorTargetId is null ? string.Empty : $" comment-anchor-target={change.CommentAnchorTargetId}";
            string commentReferenceTarget = change.CommentReferenceTargetId is null ? string.Empty : $" comment-reference-target={change.CommentReferenceTargetId}";
            string commentAnchorStory = change.CommentAnchorStory is null ? string.Empty : $" comment-anchor-story=\"{EscapeText(change.CommentAnchorStory)}\"";
            string commentAnchorPart = change.CommentAnchorPartName is null ? string.Empty : $" comment-anchor-part={change.CommentAnchorPartName}";
            string commentTextLength = change.CommentTextLength is null ? string.Empty : $" comment-text-length={change.CommentTextLength}";
            string commentTextSnippet = change.CommentTextSnippet is null ? string.Empty : $" comment-text=\"{EscapeText(change.CommentTextSnippet)}\"";
            string commentTextTruncated = change.CommentTextTruncated ? " comment-text-truncated=true" : string.Empty;
            builder.Append(change.Id)
                .Append(' ')
                .Append(change.Type)
                .Append(" story=\"")
                .Append(EscapeText(change.Story))
                .Append("\" part=")
                .Append(change.PartName)
                .Append(' ')
                .Append(target)
                .Append(" target-status=")
                .Append(change.TargetStatus)
                .Append(targetSource)
                .Append(targetReason)
                .Append(nearestTarget)
                .Append(targetNote)
                .Append(pairedChange)
                .Append(" text-length=")
                .Append(change.TextLength)
                .Append(" children=")
                .Append(change.ChildElementCount)
                .Append(revision)
                .Append(author)
                .Append(timestamp)
                .Append(commentId)
                .Append(commentAuthor)
                .Append(commentTimestamp)
                .Append(commentInitials)
                .Append(commentAnchorTarget)
                .Append(commentReferenceTarget)
                .Append(commentAnchorStory)
                .Append(commentAnchorPart)
                .Append(commentTextLength)
                .Append(commentTextSnippet)
                .Append(commentTextTruncated)
                .AppendLine();
        }

        return builder.ToString();
    }

    public static string RenderOperationSummary(IReadOnlyList<DocxPatchOperationReport> operations)
    {
        ArgumentNullException.ThrowIfNull(operations);
        var builder = new StringBuilder();
        foreach (DocxPatchOperationReport operation in operations)
        {
            string target = operation.Target is null ? "target=unknown" : $"target={operation.Target}";
            builder.Append("operation index=")
                .Append(operation.Index)
                .Append(" name=")
                .Append(operation.OperationName)
                .Append(' ')
                .Append(target)
                .Append(" success=")
                .Append(operation.Success)
                .AppendLine();
        }

        return builder.ToString();
    }

    private static string RenderLines(IReadOnlyList<string> lines)
    {
        var builder = new StringBuilder();
        foreach (string line in lines)
        {
            builder.AppendLine(line);
        }

        return builder.ToString();
    }

    private static string FormatTypeSummary(IReadOnlyList<DocxChangeSummary> summaries)
    {
        return string.Join(",", summaries.Select(summary => $"{summary.Type}:{summary.Count}"));
    }

    private static string EscapeText(string text)
    {
        return text
            .Replace("\\", "\\\\", StringComparison.Ordinal)
            .Replace("\"", "\\\"", StringComparison.Ordinal)
            .Replace("\r", "\\r", StringComparison.Ordinal)
            .Replace("\n", "\\n", StringComparison.Ordinal)
            .Replace("\t", "\\t", StringComparison.Ordinal);
    }
}
