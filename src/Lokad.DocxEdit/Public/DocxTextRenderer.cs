using System.Globalization;
using System.Text;

namespace Lokad.DocxEdit;

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
        builder.Append("hyperlinks count=").Append(result.Hyperlinks.Count).AppendLine();
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
            string relationshipId = image.RelationshipId is null ? string.Empty : $" relationship-id={EscapeText(image.RelationshipId)}";
            string target = image.ContainingTargetId is null ? " target=unknown" : $" target={image.ContainingTargetId}";
            string size = image.WidthEmu is null || image.HeightEmu is null ? string.Empty : $" size-emu={image.WidthEmu}x{image.HeightEmu}";
            string description = image.Description is null ? string.Empty : $" description=\"{EscapeText(image.Description)}\"";
            string wrap = image.WrapMode is null ? string.Empty : $" wrap={EscapeText(image.WrapMode)}";
            string behind = image.BehindDoc ? " behind-doc=true" : string.Empty;
            string layoutMetadata = RenderImageLayoutMetadata(image);
            string crop = RenderCrop(image);
            builder.Append(image.Id)
                .Append(" image layout=")
                .Append(EscapeText(image.LayoutKind))
                .Append(" part=")
                .Append(image.PartName)
                .Append(" content-type=")
                .Append(image.ContentType ?? "unknown")
                .Append(" bytes=")
                .Append(image.ByteLength)
                .Append(relationshipId)
                .Append(target)
                .Append(size)
                .Append(description)
                .Append(wrap)
                .Append(behind)
                .Append(layoutMetadata)
                .Append(crop)
                .AppendLine();
        }

        return builder.ToString();
    }

    private static string RenderImageLayoutMetadata(DocxImageInfo image)
    {
        return RenderLong("wrap-dist-top-emu", image.WrapDistanceTopEmu) +
            RenderLong("wrap-dist-bottom-emu", image.WrapDistanceBottomEmu) +
            RenderLong("wrap-dist-left-emu", image.WrapDistanceLeftEmu) +
            RenderLong("wrap-dist-right-emu", image.WrapDistanceRightEmu) +
            RenderLong("relative-height", image.RelativeHeight) +
            RenderBool("allow-overlap", image.AllowOverlap) +
            RenderBool("lock-aspect", image.LockAspectRatio) +
            RenderString("position-h-relative", image.HorizontalPositionRelativeFrom) +
            RenderLong("position-h-offset-emu", image.HorizontalPositionOffsetEmu) +
            RenderString("position-h-align", image.HorizontalPositionAlign) +
            RenderString("position-v-relative", image.VerticalPositionRelativeFrom) +
            RenderLong("position-v-offset-emu", image.VerticalPositionOffsetEmu) +
            RenderString("position-v-align", image.VerticalPositionAlign);
    }

    private static string RenderCrop(DocxImageInfo image)
    {
        return RenderPercent("crop-left-percent", image.CropLeftPercent) +
            RenderPercent("crop-top-percent", image.CropTopPercent) +
            RenderPercent("crop-right-percent", image.CropRightPercent) +
            RenderPercent("crop-bottom-percent", image.CropBottomPercent);
    }

    private static string RenderPercent(string name, decimal? value)
    {
        return value is null
            ? string.Empty
            : $" {name}={value.Value.ToString("0.###", CultureInfo.InvariantCulture)}";
    }

    private static string RenderLong(string name, long? value)
    {
        return value is null ? string.Empty : $" {name}={value}";
    }

    private static string RenderBool(string name, bool? value)
    {
        return value is null ? string.Empty : $" {name}={value.Value.ToString().ToLowerInvariant()}";
    }

    private static string RenderString(string name, string? value)
    {
        return value is null ? string.Empty : $" {name}={EscapeText(value)}";
    }

    public static string RenderValidate(DocxValidateResult result)
    {
        ArgumentNullException.ThrowIfNull(result);
        var builder = new StringBuilder();
        builder.AppendLine(result.Success ? "docxedit validate: OK" : "docxedit validate: FAILED");
        builder.Append("profile=")
            .Append(result.Profile.ToString().ToLowerInvariant())
            .AppendLine();
        foreach (DocxDiagnostic diagnostic in result.Diagnostics)
        {
            builder.Append(diagnostic.Severity)
                .Append(' ')
                .Append(diagnostic.Code)
                .Append(" part=")
                .Append(diagnostic.PartName ?? "unknown")
                .Append(' ')
                .AppendLine(diagnostic.Message);
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
            string paraId = summary.ParaId is null ? string.Empty : $" para-id={EscapeText(summary.ParaId)}";
            string parentParaId = summary.ParentParaId is null ? string.Empty : $" parent-para-id={EscapeText(summary.ParentParaId)}";
            string rootParaId = summary.RootParaId is null ? string.Empty : $" root-para-id={EscapeText(summary.RootParaId)}";
            string isReply = summary.IsReply is null ? string.Empty : $" is-reply={summary.IsReply.Value.ToString().ToLowerInvariant()}";
            string resolved = summary.Resolved is null ? string.Empty : $" resolved={summary.Resolved.Value.ToString().ToLowerInvariant()}";
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
                .Append(paraId)
                .Append(parentParaId)
                .Append(rootParaId)
                .Append(isReply)
                .Append(resolved)
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
            string commentParaId = change.CommentParaId is null ? string.Empty : $" comment-para-id={EscapeText(change.CommentParaId)}";
            string commentParentParaId = change.CommentParentParaId is null ? string.Empty : $" comment-parent-para-id={EscapeText(change.CommentParentParaId)}";
            string commentRootParaId = change.CommentRootParaId is null ? string.Empty : $" comment-root-para-id={EscapeText(change.CommentRootParaId)}";
            string commentIsReply = change.CommentIsReply is null ? string.Empty : $" comment-is-reply={change.CommentIsReply.Value.ToString().ToLowerInvariant()}";
            string commentResolved = change.CommentResolved is null ? string.Empty : $" comment-resolved={change.CommentResolved.Value.ToString().ToLowerInvariant()}";
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
                .Append(commentParaId)
                .Append(commentParentParaId)
                .Append(commentRootParaId)
                .Append(commentIsReply)
                .Append(commentResolved)
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
            string revisionIds = operation.GeneratedRevisionIds.Count == 0
                ? string.Empty
                : $" generated-revision-ids={EscapeText(string.Join(",", operation.GeneratedRevisionIds))}";
            builder.Append("operation index=")
                .Append(operation.Index)
                .Append(" name=")
                .Append(operation.OperationName)
                .Append(' ')
                .Append(target)
                .Append(" success=")
                .Append(operation.Success)
                .Append(revisionIds)
                .AppendLine();
            foreach (DocxPatchAffectedTarget affected in operation.AffectedTargets)
            {
                string parent = affected.ParentId is null ? string.Empty : $" parent={affected.ParentId}";
                string row = affected.RowIndex is null ? string.Empty : $" row={affected.RowIndex}";
                string column = affected.ColumnIndex is null ? string.Empty : $" column={affected.ColumnIndex}";
                string rowsBefore = affected.RowCountBefore is null ? string.Empty : $" rows-before={affected.RowCountBefore}";
                string rowsAfter = affected.RowCountAfter is null ? string.Empty : $" rows-after={affected.RowCountAfter}";
                string columns = affected.ColumnCount is null ? string.Empty : $" columns={affected.ColumnCount}";
                string cells = affected.CellCount is null ? string.Empty : $" cells={affected.CellCount}";
                builder.Append("  affected id=")
                    .Append(affected.Id)
                    .Append(" kind=")
                    .Append(affected.Kind)
                    .Append(" action=")
                    .Append(affected.Action)
                    .Append(parent)
                    .Append(row)
                    .Append(column)
                    .Append(rowsBefore)
                    .Append(rowsAfter)
                    .Append(columns)
                    .Append(cells)
                    .AppendLine();
            }
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
