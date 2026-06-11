using System.Xml.Linq;
using DocxEdit.Ooxml;

namespace DocxEdit.Model;

internal static class DocxChangeScanner
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
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var changes = new List<DocxChangeInfo>();
        IReadOnlyDictionary<string, CommentMetadata> comments = BuildCommentMap(package, cancellationToken);
        IReadOnlyDictionary<string, CommentAnchorMetadata> commentAnchors = BuildCommentAnchorMap(package, cancellationToken);
        int fallbackPartIndex = 1;
        foreach (OoxmlPart part in package.Parts.Values
            .Where(IsWordXmlPart)
            .OrderBy(part => part.Name, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            string story = GetStory(package, part.Name, fallbackPartIndex);
            string prefix = GetIdPrefix(package, part.Name, ref fallbackPartIndex);
            IReadOnlyDictionary<XElement, string> targets = BuildTargetMap(document, prefix);

            int index = 1;
            foreach (XElement element in document.Descendants().Where(IsChangeElement))
            {
                cancellationToken.ThrowIfCancellationRequested();
                string? commentId = ReadCommentId(element);
                comments.TryGetValue(commentId ?? string.Empty, out CommentMetadata? comment);
                commentAnchors.TryGetValue(commentId ?? string.Empty, out CommentAnchorMetadata? commentAnchor);
                string? targetId = FindTarget(element, targets);
                changes.Add(new DocxChangeInfo
                {
                    Id = $"{prefix}.CH{index++:0000}",
                    Type = ChangeTypes[element.Name.LocalName],
                    Story = story,
                    PartName = part.Name,
                    TargetId = targetId,
                    Author = ReadRevisionAuthor(element),
                    TimestampUtc = ReadRevisionTimestamp(element),
                    RevisionId = ReadRevisionId(element),
                    TextLength = ReadRevisionTextLength(element),
                    ChildElementCount = element.Elements().Count(),
                    CommentId = commentId,
                    CommentAuthor = comment?.Author,
                    CommentTimestampUtc = comment?.TimestampUtc,
                    CommentInitials = comment?.Initials,
                    CommentAnchorTargetId = commentAnchor?.AnchorTargetId,
                    CommentReferenceTargetId = commentAnchor?.ReferenceTargetId,
                    CommentAnchorStory = commentAnchor?.Story,
                    CommentAnchorPartName = commentAnchor?.PartName,
                    TargetStatus = GetTargetStatus(targetId, commentAnchor?.AnchorTargetId),
                    TargetNote = GetTargetNote(targetId, commentAnchor?.AnchorTargetId)
                });
            }
        }

        return changes;
    }

    public static IReadOnlyList<DocxChangeSummary> Summarize(IReadOnlyList<DocxChangeInfo> changes)
    {
        return changes
            .GroupBy(change => change.Type, StringComparer.Ordinal)
            .OrderBy(group => group.Key, StringComparer.Ordinal)
            .Select(group => new DocxChangeSummary(group.Key, group.Count()))
            .ToArray();
    }

    public static IReadOnlyList<DocxChangeTargetSummary> SummarizeTargets(IReadOnlyList<DocxChangeInfo> changes)
    {
        return changes
            .GroupBy(change => GetGroupTargetKey(change) ?? "(none)", StringComparer.Ordinal)
            .OrderBy(group => group.Key, StringComparer.Ordinal)
            .Select(group => new DocxChangeTargetSummary
            {
                TargetId = group.Key,
                Count = group.Count(),
                Summary = Summarize(group.ToArray())
            })
            .ToArray();
    }

    public static IReadOnlyList<DocxCommentThreadSummary> SummarizeComments(IReadOnlyList<DocxChangeInfo> changes)
    {
        return changes
            .Where(change => !string.IsNullOrWhiteSpace(change.CommentId))
            .GroupBy(change => change.CommentId!, StringComparer.Ordinal)
            .OrderBy(group => group.Key, StringComparer.Ordinal)
            .Select(group =>
            {
                DocxChangeInfo? anchor = group.FirstOrDefault(change => change.CommentAnchorTargetId is not null);
                DocxChangeInfo? metadata = group.FirstOrDefault(change => change.CommentAuthor is not null || change.CommentTimestampUtc is not null);
                return new DocxCommentThreadSummary
                {
                    CommentId = group.Key,
                    AnchorTargetId = anchor?.CommentAnchorTargetId,
                    ReferenceTargetId = anchor?.CommentReferenceTargetId,
                    AnchorStory = anchor?.CommentAnchorStory,
                    AnchorPartName = anchor?.CommentAnchorPartName,
                    Author = metadata?.CommentAuthor,
                    TimestampUtc = metadata?.CommentTimestampUtc,
                    Count = group.Count(),
                    Summary = Summarize(group.ToArray())
                };
            })
            .ToArray();
    }

    public static IReadOnlyList<DocxChangeGroupSummary> SummarizeGroups(IReadOnlyList<DocxChangeInfo> changes)
    {
        return EnumerateGroupValues(changes)
            .GroupBy(item => (item.Group, item.Key, item.Type))
            .OrderBy(group => group.Key.Group, StringComparer.Ordinal)
            .ThenBy(group => group.Key.Key, StringComparer.Ordinal)
            .ThenBy(group => group.Key.Type, StringComparer.Ordinal)
            .Select(group => new DocxChangeGroupSummary(group.Key.Group, group.Key.Key, group.Key.Type, group.Count()))
            .ToArray();
    }

    private static IEnumerable<(string Group, string Key, string Type)> EnumerateGroupValues(IReadOnlyList<DocxChangeInfo> changes)
    {
        foreach (DocxChangeInfo change in changes)
        {
            yield return ("story", NormalizeGroupKey(change.Story), change.Type);
            yield return ("part", NormalizeGroupKey(change.PartName), change.Type);
            yield return ("author", NormalizeGroupKey(change.Author ?? change.CommentAuthor), change.Type);
            yield return ("target", NormalizeGroupKey(GetGroupTargetKey(change)), change.Type);
        }
    }

    private static string? GetGroupTargetKey(DocxChangeInfo change)
    {
        return change.Type == "comment" && change.CommentAnchorTargetId is not null
            ? change.CommentAnchorTargetId
            : change.TargetId ?? change.CommentAnchorTargetId;
    }

    private static string GetTargetStatus(string? targetId, string? commentAnchorTargetId)
    {
        if (targetId is not null)
        {
            return "targeted";
        }

        if (commentAnchorTargetId is not null)
        {
            return "comment-anchor";
        }

        return "targetless";
    }

    private static string? GetTargetNote(string? targetId, string? commentAnchorTargetId)
    {
        if (targetId is not null)
        {
            return null;
        }

        if (commentAnchorTargetId is not null)
        {
            return "Linked through matching comment anchor metadata.";
        }

        return "No modeled paragraph, table, cell, section, or comment anchor target was found; the markup may be body-level, package-level, or inside an unsupported structure.";
    }

    private static string NormalizeGroupKey(string? key)
    {
        return string.IsNullOrWhiteSpace(key) ? "(none)" : key;
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

    private static string GetStory(OoxmlPackage package, string partName, int fallbackPartIndex)
    {
        if (string.Equals(partName, package.MainDocumentPartName, StringComparison.OrdinalIgnoreCase))
        {
            return "main";
        }

        if (package.MainDocumentPartName is not null)
        {
            IReadOnlyList<OoxmlRelationship> relationships = package.GetRelationships(package.MainDocumentPartName);
            int commentsIndex = 1;
            foreach (OoxmlRelationship relationship in relationships
                .Where(relationship => !relationship.IsExternal && relationship.Type == OoxmlRelTypes.Comments && relationship.ResolvedTarget is not null)
                .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
            {
                if (string.Equals(relationship.ResolvedTarget, partName, StringComparison.OrdinalIgnoreCase))
                {
                    return $"comments[{commentsIndex}]";
                }

                commentsIndex++;
            }

            int headerIndex = 1;
            foreach (OoxmlRelationship relationship in relationships
                .Where(relationship => !relationship.IsExternal && relationship.Type == OoxmlRelTypes.Header && relationship.ResolvedTarget is not null)
                .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
            {
                if (string.Equals(relationship.ResolvedTarget, partName, StringComparison.OrdinalIgnoreCase))
                {
                    return $"header[{headerIndex}]";
                }

                headerIndex++;
            }

            int footerIndex = 1;
            foreach (OoxmlRelationship relationship in relationships
                .Where(relationship => !relationship.IsExternal && relationship.Type == OoxmlRelTypes.Footer && relationship.ResolvedTarget is not null)
                .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
            {
                if (string.Equals(relationship.ResolvedTarget, partName, StringComparison.OrdinalIgnoreCase))
                {
                    return $"footer[{footerIndex}]";
                }

                footerIndex++;
            }
        }

        return $"part[{fallbackPartIndex}]";
    }

    private static string GetIdPrefix(OoxmlPackage package, string partName, ref int fallbackPartIndex)
    {
        if (string.Equals(partName, package.MainDocumentPartName, StringComparison.OrdinalIgnoreCase))
        {
            return "M";
        }

        if (package.MainDocumentPartName is not null)
        {
            IReadOnlyList<OoxmlRelationship> relationships = package.GetRelationships(package.MainDocumentPartName);
            int commentsIndex = 1;
            foreach (OoxmlRelationship relationship in relationships
                .Where(relationship => !relationship.IsExternal && relationship.Type == OoxmlRelTypes.Comments && relationship.ResolvedTarget is not null)
                .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
            {
                if (string.Equals(relationship.ResolvedTarget, partName, StringComparison.OrdinalIgnoreCase))
                {
                    return $"C{commentsIndex:000}";
                }

                commentsIndex++;
            }

            int headerIndex = 1;
            foreach (OoxmlRelationship relationship in relationships
                .Where(relationship => !relationship.IsExternal && relationship.Type == OoxmlRelTypes.Header && relationship.ResolvedTarget is not null)
                .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
            {
                if (string.Equals(relationship.ResolvedTarget, partName, StringComparison.OrdinalIgnoreCase))
                {
                    return $"H{headerIndex:000}";
                }

                headerIndex++;
            }

            int footerIndex = 1;
            foreach (OoxmlRelationship relationship in relationships
                .Where(relationship => !relationship.IsExternal && relationship.Type == OoxmlRelTypes.Footer && relationship.ResolvedTarget is not null)
                .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
            {
                if (string.Equals(relationship.ResolvedTarget, partName, StringComparison.OrdinalIgnoreCase))
                {
                    return $"F{footerIndex:000}";
                }

                footerIndex++;
            }
        }

        return $"P{fallbackPartIndex++:000}";
    }

    private static IReadOnlyDictionary<XElement, string> BuildTargetMap(XDocument document, string prefix)
    {
        var targets = new Dictionary<XElement, string>();
        XElement root = document.Root ?? new XElement("empty");
        if (root.Name == OoxmlNs.W + "comments")
        {
            int commentIndex = 1;
            foreach (XElement comment in root.Elements(OoxmlNs.W + "comment"))
            {
                targets[comment] = $"{prefix}.C{commentIndex++:0000}";
            }

            return targets;
        }

        XElement body = root.Element(OoxmlNs.W + "body") ?? root;
        int paragraphIndex = 1;
        int tableIndex = 1;
        int sectionIndex = 1;

        foreach (XElement block in body.Elements())
        {
            if (block.Name == OoxmlNs.W + "p")
            {
                targets[block] = $"{prefix}.P{paragraphIndex++:0000}";
                XElement? sectionProperties = block.Element(OoxmlNs.W + "pPr")?.Element(OoxmlNs.W + "sectPr");
                if (sectionProperties is not null)
                {
                    targets[sectionProperties] = $"{prefix}.S{sectionIndex++:0000}";
                }
            }
            else if (block.Name == OoxmlNs.W + "tbl")
            {
                string tableId = $"{prefix}.T{tableIndex++:0000}";
                targets[block] = tableId;
                int rowIndex = 1;
                foreach (XElement row in block.Elements(OoxmlNs.W + "tr"))
                {
                    string rowId = $"{tableId}.R{rowIndex:00}";
                    targets[row] = rowId;
                    int cellIndex = 1;
                    foreach (XElement cell in row.Elements(OoxmlNs.W + "tc"))
                    {
                        targets[cell] = $"{rowId}.C{cellIndex++:00}";
                    }

                    rowIndex++;
                }
            }
            else if (block.Name == OoxmlNs.W + "sectPr")
            {
                targets[block] = $"{prefix}.S{sectionIndex++:0000}";
            }
        }

        return targets;
    }

    private static string? FindTarget(XElement element, IReadOnlyDictionary<XElement, string> targets)
    {
        return FindNearestTarget(element, targets) ??
            FindAdjacentRangeTarget(element, targets);
    }

    private static string? FindNearestTarget(XElement element, IReadOnlyDictionary<XElement, string> targets)
    {
        foreach (XElement candidate in element.AncestorsAndSelf())
        {
            if (targets.TryGetValue(candidate, out string? id))
            {
                return id;
            }
        }

        return null;
    }

    private static string? FindAdjacentRangeTarget(XElement element, IReadOnlyDictionary<XElement, string> targets)
    {
        if (!IsRangeBoundaryElement(element))
        {
            return null;
        }

        bool isEnd = element.Name.LocalName.EndsWith("End", StringComparison.Ordinal);
        IEnumerable<XElement> siblings = isEnd
            ? element.ElementsBeforeSelf().Reverse()
            : element.ElementsAfterSelf();
        foreach (XElement sibling in siblings)
        {
            if (TryFindTargetInSubtree(sibling, targets, preferLast: isEnd, out string? targetId))
            {
                return targetId;
            }
        }

        siblings = isEnd
            ? element.ElementsAfterSelf()
            : element.ElementsBeforeSelf().Reverse();
        foreach (XElement sibling in siblings)
        {
            if (TryFindTargetInSubtree(sibling, targets, preferLast: !isEnd, out string? targetId))
            {
                return targetId;
            }
        }

        return null;
    }

    private static bool TryFindTargetInSubtree(
        XElement element,
        IReadOnlyDictionary<XElement, string> targets,
        bool preferLast,
        out string? targetId)
    {
        if (targets.TryGetValue(element, out targetId))
        {
            return true;
        }

        IEnumerable<XElement> descendants = preferLast
            ? element.Descendants().Reverse()
            : element.Descendants();
        foreach (XElement descendant in descendants)
        {
            if (targets.TryGetValue(descendant, out targetId))
            {
                return true;
            }
        }

        targetId = null;
        return false;
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

    private static IReadOnlyDictionary<string, CommentMetadata> BuildCommentMap(
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        if (package.MainDocumentPartName is null)
        {
            return new Dictionary<string, CommentMetadata>(StringComparer.Ordinal);
        }

        var comments = new Dictionary<string, CommentMetadata>(StringComparer.Ordinal);
        foreach (OoxmlRelationship relationship in package
            .GetRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => !relationship.IsExternal && relationship.Type == OoxmlRelTypes.Comments && relationship.ResolvedTarget is not null)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            OoxmlPart? part = package.GetPart(relationship.ResolvedTarget!);
            if (part is null)
            {
                continue;
            }

            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            foreach (XElement comment in document.Descendants(OoxmlNs.W + "comment"))
            {
                string? id = (string?)comment.Attribute(OoxmlNs.W + "id");
                if (string.IsNullOrWhiteSpace(id))
                {
                    continue;
                }

                comments[id] = new CommentMetadata(
                    (string?)comment.Attribute(OoxmlNs.W + "author"),
                    ParseDate((string?)comment.Attribute(OoxmlNs.W + "date")),
                    (string?)comment.Attribute(OoxmlNs.W + "initials"));
            }
        }

        return comments;
    }

    private static IReadOnlyDictionary<string, CommentAnchorMetadata> BuildCommentAnchorMap(
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        var anchors = new Dictionary<string, CommentAnchorBuilder>(StringComparer.Ordinal);
        int fallbackPartIndex = 1;
        foreach (OoxmlPart part in package.Parts.Values
            .Where(part => IsWordXmlPart(part) && !IsCommentsPart(package, part.Name))
            .OrderBy(part => part.Name, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            string story = GetStory(package, part.Name, fallbackPartIndex);
            string prefix = GetIdPrefix(package, part.Name, ref fallbackPartIndex);
            IReadOnlyDictionary<XElement, string> targets = BuildTargetMap(document, prefix);

            foreach (XElement element in document.Descendants().Where(IsCommentElement))
            {
                string? commentId = ReadCommentId(element);
                if (string.IsNullOrWhiteSpace(commentId))
                {
                    continue;
                }

                string? targetId = FindTarget(element, targets);
                if (!anchors.TryGetValue(commentId, out CommentAnchorBuilder? anchor))
                {
                    anchor = new CommentAnchorBuilder(story, part.Name);
                    anchors[commentId] = anchor;
                }

                anchor.Story ??= story;
                anchor.PartName ??= part.Name;
                switch (element.Name.LocalName)
                {
                    case "commentRangeStart":
                        anchor.RangeStartTargetId ??= targetId;
                        break;
                    case "commentRangeEnd":
                        anchor.RangeEndTargetId ??= targetId;
                        break;
                    case "commentReference":
                        anchor.ReferenceTargetId ??= targetId;
                        break;
                }
            }
        }

        return anchors.ToDictionary(
            pair => pair.Key,
            pair => new CommentAnchorMetadata(
                pair.Value.RangeStartTargetId ?? pair.Value.ReferenceTargetId ?? pair.Value.RangeEndTargetId,
                pair.Value.ReferenceTargetId,
                pair.Value.Story,
                pair.Value.PartName),
            StringComparer.Ordinal);
    }

    private static bool IsCommentsPart(OoxmlPackage package, string partName)
    {
        if (package.MainDocumentPartName is null)
        {
            return false;
        }

        return package.GetRelationships(package.MainDocumentPartName)
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
}

internal sealed record CommentMetadata(
    string? Author,
    DateTimeOffset? TimestampUtc,
    string? Initials);

internal sealed record CommentAnchorMetadata(
    string? AnchorTargetId,
    string? ReferenceTargetId,
    string? Story,
    string? PartName);

internal sealed class CommentAnchorBuilder(string? story, string? partName)
{
    public string? Story { get; set; } = story;
    public string? PartName { get; set; } = partName;
    public string? RangeStartTargetId { get; set; }
    public string? RangeEndTargetId { get; set; }
    public string? ReferenceTargetId { get; set; }
}
