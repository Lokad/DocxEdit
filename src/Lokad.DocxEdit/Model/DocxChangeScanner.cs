using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

internal static partial class DocxChangeScanner
{
    private static readonly IReadOnlyDictionary<string, string> ChangeTypes = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        ["ins"] = "inserted-run",
        ["del"] = "deleted-run",
        ["moveFrom"] = "move-from-run",
        ["moveTo"] = "move-to-run",
        ["moveFromRangeStart"] = "move-from-range-start",
        ["moveFromRangeEnd"] = "move-from-range-end",
        ["moveToRangeStart"] = "move-to-range-start",
        ["moveToRangeEnd"] = "move-to-range-end",
        ["rPrChange"] = "run-properties-change",
        ["pPrChange"] = "paragraph-properties-change",
        ["tblPrChange"] = "table-properties-change",
        ["trPrChange"] = "row-properties-change",
        ["tcPrChange"] = "cell-properties-change",
        ["sectPrChange"] = "section-properties-change",
        ["cellIns"] = "cell-inserted",
        ["cellDel"] = "cell-deleted",
        ["cellMerge"] = "cell-merge-change",
        ["customXmlInsRangeStart"] = "custom-xml-insert-range-start",
        ["customXmlInsRangeEnd"] = "custom-xml-insert-range-end",
        ["customXmlDelRangeStart"] = "custom-xml-delete-range-start",
        ["customXmlDelRangeEnd"] = "custom-xml-delete-range-end",
        ["commentRangeStart"] = "comment-range-start",
        ["commentRangeEnd"] = "comment-range-end",
        ["commentReference"] = "comment-reference",
        ["comment"] = "comment"
    };

    public static IReadOnlyList<DocxChangeInfo> Scan(
        OoxmlPackage package,
        bool includeCommentText,
        int maxCommentText,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var changes = new List<DocxChangeInfo>();
        IReadOnlyDictionary<string, CommentMetadata> comments = BuildCommentMap(package, includeCommentText, maxCommentText, cancellationToken);
        IReadOnlyDictionary<string, CommentAnchorMetadata> commentAnchors = BuildCommentAnchorMap(package, cancellationToken);
        int fallbackPartIndex = 1;
        foreach (OoxmlPart part in package.Parts.Values
            .Where(IsWordXmlPart)
            .OrderBy(part => part.Name, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            string story = GetStory(package, part.Name, fallbackPartIndex, cancellationToken);
            string prefix = GetIdPrefix(package, part.Name, ref fallbackPartIndex, cancellationToken);
            IReadOnlyDictionary<XElement, string> targets = BuildTargetMap(document, prefix);

            int index = 1;
            foreach (XElement element in document.Descendants().Where(IsChangeElement))
            {
                cancellationToken.ThrowIfCancellationRequested();
                string? commentId = ReadCommentId(element);
                comments.TryGetValue(commentId ?? string.Empty, out CommentMetadata? comment);
                commentAnchors.TryGetValue(commentId ?? string.Empty, out CommentAnchorMetadata? commentAnchor);
                TargetMetadata target = FindTarget(element, targets, commentAnchor?.AnchorTargetId);
                bool isCommentBody = element.Name.LocalName == "comment";
                changes.Add(new DocxChangeInfo
                {
                    Id = $"{prefix}.CH{index++:0000}",
                    Type = ReadChangeType(element),
                    Story = story,
                    PartName = part.Name,
                    ParentType = ReadParentType(element),
                    TargetId = target.TargetId,
                    Author = ReadRevisionAuthor(element),
                    TimestampUtc = ReadRevisionTimestamp(element),
                    RevisionId = ReadRevisionId(element),
                    TextLength = ReadRevisionTextLength(element),
                    ChildElementCount = element.Elements().Count(),
                    CommentId = commentId,
                    CommentAuthor = comment?.Author,
                    CommentTimestampUtc = comment?.TimestampUtc,
                    CommentInitials = comment?.Initials,
                    CommentParaId = comment?.ParaId,
                    CommentParentParaId = comment?.ParentParaId,
                    CommentRootParaId = comment?.RootParaId,
                    CommentDurableId = comment?.DurableId,
                    CommentIsReply = comment?.IsReply,
                    CommentResolved = comment?.Resolved,
                    CommentAnchorTargetId = commentAnchor?.AnchorTargetId,
                    CommentReferenceTargetId = commentAnchor?.ReferenceTargetId,
                    CommentAnchorStory = commentAnchor?.Story,
                    CommentAnchorPartName = commentAnchor?.PartName,
                    CommentTextLength = isCommentBody ? comment?.TextLength : null,
                    CommentTextSnippet = isCommentBody ? comment?.TextSnippet : null,
                    CommentTextTruncated = isCommentBody && comment?.TextTruncated == true,
                    TargetStatus = target.Status,
                    TargetSource = target.Source,
                    TargetReason = target.Reason,
                    NearestTargetId = target.NearestTargetId,
                    TargetNote = target.Note
                });
            }
        }

        return PairRangeBoundaries(changes);
    }

    private static bool IsWordXmlPart(OoxmlPart part)
    {
        return part.Name.StartsWith("/word/", StringComparison.OrdinalIgnoreCase) &&
            part.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
            !part.Name.Contains("/_rels/", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsChangeElement(XElement element)
    {
        return element.Name.Namespace == OoxmlNs.W &&
            ChangeTypes.ContainsKey(element.Name.LocalName);
    }

    private static string GetStory(OoxmlPackage package, string partName, int fallbackPartIndex, CancellationToken cancellationToken)
    {
        if (string.Equals(partName, package.MainDocumentPartName, StringComparison.OrdinalIgnoreCase))
        {
            return "main";
        }

        IReadOnlyList<ResolvedOoxmlRelationship> relationships = package.GetResolvedRelationships(package.MainDocumentPartName, cancellationToken);
        int commentsIndex = 1;
        foreach (ResolvedOoxmlRelationship relationship in relationships
            .Where(relationship => relationship.Type == OoxmlRelTypes.Comments)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
        {
            if (string.Equals(relationship.ResolvedTarget, partName, StringComparison.OrdinalIgnoreCase))
            {
                return $"comments[{commentsIndex}]";
            }

            commentsIndex++;
        }

        int headerIndex = 1;
        foreach (ResolvedOoxmlRelationship relationship in relationships
            .Where(relationship => relationship.Type == OoxmlRelTypes.Header)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
        {
            if (string.Equals(relationship.ResolvedTarget, partName, StringComparison.OrdinalIgnoreCase))
            {
                return $"header[{headerIndex}]";
            }

            headerIndex++;
        }

        int footerIndex = 1;
        foreach (ResolvedOoxmlRelationship relationship in relationships
            .Where(relationship => relationship.Type == OoxmlRelTypes.Footer)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
        {
            if (string.Equals(relationship.ResolvedTarget, partName, StringComparison.OrdinalIgnoreCase))
            {
                return $"footer[{footerIndex}]";
            }

            footerIndex++;
        }

        return $"part[{fallbackPartIndex}]";
    }

    private static string GetIdPrefix(OoxmlPackage package, string partName, ref int fallbackPartIndex, CancellationToken cancellationToken)
    {
        if (string.Equals(partName, package.MainDocumentPartName, StringComparison.OrdinalIgnoreCase))
        {
            return "M";
        }

        IReadOnlyList<ResolvedOoxmlRelationship> relationships = package.GetResolvedRelationships(package.MainDocumentPartName, cancellationToken);
        int commentsIndex = 1;
        foreach (ResolvedOoxmlRelationship relationship in relationships
            .Where(relationship => relationship.Type == OoxmlRelTypes.Comments)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
        {
            if (string.Equals(relationship.ResolvedTarget, partName, StringComparison.OrdinalIgnoreCase))
            {
                return $"C{commentsIndex:000}";
            }

            commentsIndex++;
        }

        int headerIndex = 1;
        foreach (ResolvedOoxmlRelationship relationship in relationships
            .Where(relationship => relationship.Type == OoxmlRelTypes.Header)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
        {
            if (string.Equals(relationship.ResolvedTarget, partName, StringComparison.OrdinalIgnoreCase))
            {
                return $"H{headerIndex:000}";
            }

            headerIndex++;
        }

        int footerIndex = 1;
        foreach (ResolvedOoxmlRelationship relationship in relationships
            .Where(relationship => relationship.Type == OoxmlRelTypes.Footer)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
        {
            if (string.Equals(relationship.ResolvedTarget, partName, StringComparison.OrdinalIgnoreCase))
            {
                return $"F{footerIndex:000}";
            }

            footerIndex++;
        }

        return $"P{fallbackPartIndex++:000}";
    }

    private static bool IsRangeBoundaryElement(XElement element)
    {
        if (element.Name.Namespace != OoxmlNs.W)
        {
            return false;
        }

        return element.Name.LocalName.EndsWith("RangeStart", StringComparison.Ordinal) ||
            element.Name.LocalName.EndsWith("RangeEnd", StringComparison.Ordinal) ||
            element.Name.LocalName is "commentRangeStart" or "commentRangeEnd";
    }

    private static bool IsCommentsPart(OoxmlPackage package, string partName, CancellationToken cancellationToken)
    {

        return package.GetRelationships(package.MainDocumentPartName, cancellationToken)
            .Any(relationship => !relationship.IsExternal &&
                relationship.Type == OoxmlRelTypes.Comments &&
                relationship.ResolvedTarget is not null &&
                string.Equals(relationship.ResolvedTarget, partName, StringComparison.OrdinalIgnoreCase));
    }

    private static string? ReadRevisionAuthor(XElement element)
    {
        return IsCommentElement(element)
            ? null
            : (string?)element.Attribute(OoxmlNs.W + "author");
    }

    private static string ReadChangeType(XElement element)
    {
        if (element.Name == OoxmlNs.W + "ins" &&
            element.Parent?.Name == OoxmlNs.W + "trPr")
        {
            return "row-inserted";
        }

        if (element.Name == OoxmlNs.W + "del" &&
            element.Parent?.Name == OoxmlNs.W + "trPr")
        {
            return "row-deleted";
        }

        return ChangeTypes[element.Name.LocalName];
    }

    private static string? ReadParentType(XElement element)
    {
        XElement? parent = element.Parent;
        if (parent is null)
        {
            return null;
        }

        if (parent.Name.Namespace != OoxmlNs.W)
        {
            return parent.Name.LocalName;
        }

        return parent.Name.LocalName switch
        {
            "body" => "body",
            "comments" => "comments",
            "comment" => "comment",
            "p" => "paragraph",
            "pPr" => "paragraph-properties",
            "r" => "run",
            "rPr" => "run-properties",
            "tbl" => "table",
            "tblPr" => "table-properties",
            "tr" => "row",
            "trPr" => "row-properties",
            "tc" => "cell",
            "tcPr" => "cell-properties",
            "sectPr" => "section-properties",
            _ => parent.Name.LocalName
        };
    }

    private static DateTimeOffset? ReadRevisionTimestamp(XElement element)
    {
        return IsCommentElement(element)
            ? null
            : ParseDate((string?)element.Attribute(OoxmlNs.W + "date"));
    }

    private static string? ReadRevisionId(XElement element)
    {
        return IsCommentElement(element)
            ? null
            : (string?)element.Attribute(OoxmlNs.W + "id");
    }

    private static string? ReadCommentId(XElement element)
    {
        return IsCommentElement(element)
            ? (string?)element.Attribute(OoxmlNs.W + "id")
            : null;
    }

    private static string? ReadCommentParaId(XElement comment)
    {
        return (string?)comment
            .Elements(OoxmlNs.W + "p")
            .FirstOrDefault()
            ?.Attribute(OoxmlNs.W15 + "paraId");
    }

    private static bool? ParseBoolean(string? value)
    {
        return value?.Trim() switch
        {
            "1" => true,
            "true" => true,
            "0" => false,
            "false" => false,
            null => null,
            "" => null,
            _ => null
        };
    }

    private static bool IsCommentElement(XElement element)
    {
        return element.Name.Namespace == OoxmlNs.W &&
            element.Name.LocalName is "comment" or "commentRangeStart" or "commentRangeEnd" or "commentReference";
    }

    private static DateTimeOffset? ParseDate(string? value)
    {
        return DateTimeOffset.TryParse(value, out DateTimeOffset parsed)
            ? parsed.ToUniversalTime()
            : null;
    }

    private static int ReadRevisionTextLength(XElement element)
    {
        int length = 0;
        foreach (XElement descendant in element.Descendants())
        {
            if (descendant.Name == OoxmlNs.W + "t" ||
                descendant.Name == OoxmlNs.W + "delText" ||
                descendant.Name == OoxmlNs.W + "instrText")
            {
                length += descendant.Value.Length;
            }
            else if (descendant.Name == OoxmlNs.W + "tab" ||
                descendant.Name == OoxmlNs.W + "br" ||
                descendant.Name == OoxmlNs.W + "cr")
            {
                length++;
            }
        }

        return length;
    }

    private static string ReadCommentText(XElement comment)
    {
        var builder = new System.Text.StringBuilder();
        bool wroteParagraph = false;
        foreach (XElement paragraph in comment.Elements(OoxmlNs.W + "p"))
        {
            if (wroteParagraph)
            {
                builder.Append('\n');
            }

            AppendText(paragraph, builder);
            wroteParagraph = true;
        }

        if (wroteParagraph)
        {
            return builder.ToString();
        }

        AppendText(comment, builder);
        return builder.ToString();
    }

    private static void AppendText(XElement container, System.Text.StringBuilder builder)
    {
        foreach (XElement descendant in container.Descendants())
        {
            if (descendant.Name == OoxmlNs.W + "t" ||
                descendant.Name == OoxmlNs.W + "delText" ||
                descendant.Name == OoxmlNs.W + "instrText")
            {
                builder.Append(descendant.Value);
            }
            else if (descendant.Name == OoxmlNs.W + "tab")
            {
                builder.Append('\t');
            }
            else if (descendant.Name == OoxmlNs.W + "br" ||
                descendant.Name == OoxmlNs.W + "cr")
            {
                builder.Append('\n');
            }
        }
    }

    private static string Truncate(string text, int maxText)
    {
        if (maxText < 0 || text.Length <= maxText)
        {
            return text;
        }

        return text[..maxText];
    }
}

internal sealed record CommentMetadata(
    string? Author,
    DateTimeOffset? TimestampUtc,
    string? Initials,
    string? ParaId,
    string? ParentParaId,
    string? RootParaId,
    string? DurableId,
    bool? IsReply,
    bool? Resolved,
    int? TextLength,
    string? TextSnippet,
    bool TextTruncated);

internal sealed record CommentExtensionMetadata(
    string? ParentParaId,
    bool? Resolved);

internal sealed record CommentIdMetadata(string? DurableId);

internal sealed record CommentAnchorMetadata(
    string? AnchorTargetId,
    string? ReferenceTargetId,
    string? Story,
    string? PartName);

internal sealed record TargetMetadata(
    string? TargetId,
    string Status,
    string Source,
    string? Reason,
    string? NearestTargetId,
    string? Note);

internal sealed record RangeBoundaryMetadata(RangeBoundaryKey Key, bool IsStart);

internal sealed record RangeBoundaryKey(string PartName, string Kind, string Id);

internal sealed class CommentAnchorBuilder(string? story, string? partName)
{
    public string? Story { get; set; } = story;
    public string? PartName { get; set; } = partName;
    public string? RangeStartTargetId { get; set; }
    public string? RangeEndTargetId { get; set; }
    public string? ReferenceTargetId { get; set; }
}
