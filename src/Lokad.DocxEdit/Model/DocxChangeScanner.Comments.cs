using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

internal static partial class DocxChangeScanner
{
    private static IReadOnlyDictionary<string, CommentMetadata> BuildCommentMap(
        OoxmlPackage package,
        bool includeCommentText,
        int maxCommentText,
        CancellationToken cancellationToken)
    {

        var comments = new Dictionary<string, CommentMetadata>(StringComparer.Ordinal);
        IReadOnlyDictionary<string, CommentExtensionMetadata> commentExtensions = BuildCommentExtensionMap(package, cancellationToken);
        IReadOnlyDictionary<string, CommentIdMetadata> commentIds = BuildCommentIdMap(package, cancellationToken);
        foreach (ResolvedOoxmlRelationship relationship in package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => relationship.Type == OoxmlRelTypes.Comments)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            OoxmlPart? part = package.GetPart(relationship.ResolvedTarget);
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

                string? paraId = ReadCommentParaId(comment);
                commentExtensions.TryGetValue(paraId ?? string.Empty, out CommentExtensionMetadata? extension);
                commentIds.TryGetValue(paraId ?? string.Empty, out CommentIdMetadata? commentId);
                string? rootParaId = ResolveCommentRootParaId(paraId, commentExtensions);
                string? text = includeCommentText ? ReadCommentText(comment) : null;
                comments[id] = new CommentMetadata(
                    (string?)comment.Attribute(OoxmlNs.W + "author"),
                    ParseDate((string?)comment.Attribute(OoxmlNs.W + "date")),
                    (string?)comment.Attribute(OoxmlNs.W + "initials"),
                    paraId,
                    extension?.ParentParaId,
                    rootParaId,
                    commentId?.DurableId,
                    paraId is null ? null : !string.IsNullOrWhiteSpace(extension?.ParentParaId),
                    extension?.Resolved,
                    text?.Length,
                    text is null ? null : Truncate(text, maxCommentText),
                    text is not null && maxCommentText >= 0 && text.Length > maxCommentText);
            }
        }

        return comments;
    }

    private static IReadOnlyDictionary<string, CommentExtensionMetadata> BuildCommentExtensionMap(
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {

        var extensions = new Dictionary<string, CommentExtensionMetadata>(StringComparer.Ordinal);
        foreach (ResolvedOoxmlRelationship relationship in package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => relationship.Type == OoxmlRelTypes.CommentsExtended)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            OoxmlPart? part = package.GetPart(relationship.ResolvedTarget);
            if (part is null)
            {
                continue;
            }

            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            foreach (XElement commentExtension in document.Descendants(OoxmlNs.W15 + "commentEx"))
            {
                string? paraId = (string?)commentExtension.Attribute(OoxmlNs.W15 + "paraId");
                if (string.IsNullOrWhiteSpace(paraId))
                {
                    continue;
                }

                extensions[paraId] = new CommentExtensionMetadata(
                    (string?)commentExtension.Attribute(OoxmlNs.W15 + "paraIdParent"),
                    ParseBoolean((string?)commentExtension.Attribute(OoxmlNs.W15 + "done")));
            }
        }

        return extensions;
    }

    private static IReadOnlyDictionary<string, CommentIdMetadata> BuildCommentIdMap(
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        var ids = new Dictionary<string, CommentIdMetadata>(StringComparer.Ordinal);
        foreach (string partName in GetCommentsIdsPartNames(package, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            foreach (XElement commentId in document.Descendants(OoxmlNs.W16Cid + "commentId"))
            {
                string? paraId = (string?)commentId.Attribute(OoxmlNs.W16Cid + "paraId");
                if (string.IsNullOrWhiteSpace(paraId))
                {
                    continue;
                }

                ids[paraId] = new CommentIdMetadata((string?)commentId.Attribute(OoxmlNs.W16Cid + "durableId"));
            }
        }

        return ids;
    }

    private static IReadOnlyList<string> GetCommentsIdsPartNames(
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {

        return package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => relationship.Type == OoxmlRelTypes.CommentsIds)
            .Select(relationship => relationship.ResolvedTarget)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string? ResolveCommentRootParaId(
        string? paraId,
        IReadOnlyDictionary<string, CommentExtensionMetadata> extensions)
    {
        if (string.IsNullOrWhiteSpace(paraId))
        {
            return null;
        }

        var seen = new HashSet<string>(StringComparer.Ordinal);
        string current = paraId;
        string root = current;
        while (seen.Add(current))
        {
            root = current;
            if (!extensions.TryGetValue(current, out CommentExtensionMetadata? extension) ||
                string.IsNullOrWhiteSpace(extension.ParentParaId))
            {
                return root;
            }

            current = extension.ParentParaId;
        }

        return root;
    }

    private static IReadOnlyDictionary<string, CommentAnchorMetadata> BuildCommentAnchorMap(
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        var anchors = new Dictionary<string, CommentAnchorBuilder>(StringComparer.Ordinal);
        int fallbackPartIndex = 1;
        HashSet<string> wordParts = DocxPartRoles.GetWordProcessingParts(package, cancellationToken).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (OoxmlPart part in package.Parts.Values
            .Where(part => (IsWordXmlPart(part) || wordParts.Contains(part.Name)) && !IsCommentsPart(package, part.Name, cancellationToken))
            .OrderBy(part => part.Name, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            string story = GetStory(package, part.Name, fallbackPartIndex, cancellationToken);
            string prefix = GetIdPrefix(package, part.Name, ref fallbackPartIndex, cancellationToken);
            IReadOnlyDictionary<XElement, string> targets = BuildTargetMap(document, prefix);

            foreach (XElement element in document.Descendants().Where(IsCommentElement))
            {
                string? commentId = ReadCommentId(element);
                if (string.IsNullOrWhiteSpace(commentId))
                {
                    continue;
                }

                string? targetId = FindTarget(element, targets, commentAnchorTargetId: null).TargetId;
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
}
