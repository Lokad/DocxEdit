using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

internal static partial class DocxChangeScanner
{
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
            .WithNonBlankKey(change => change.CommentId)
            .GroupBy(pair => pair.Key, pair => pair.Item, StringComparer.Ordinal)
            .OrderBy(group => group.Key, StringComparer.Ordinal)
            .Select(group =>
            {
                DocxChangeInfo? anchor = group.FirstOrDefault(change => change.CommentAnchorTargetId is not null);
                DocxChangeInfo? metadata = group.FirstOrDefault(change => change.CommentAuthor is not null || change.CommentTimestampUtc is not null);
                DocxChangeInfo? text = group.FirstOrDefault(change => change.CommentTextLength is not null || change.CommentTextSnippet is not null);
                return new DocxCommentThreadSummary
                {
                    CommentId = group.Key,
                    AnchorTargetId = anchor?.CommentAnchorTargetId,
                    ReferenceTargetId = anchor?.CommentReferenceTargetId,
                    AnchorStory = anchor?.CommentAnchorStory,
                    AnchorPartName = anchor?.CommentAnchorPartName,
                    Author = metadata?.CommentAuthor,
                    TimestampUtc = metadata?.CommentTimestampUtc,
                    Initials = metadata?.CommentInitials,
                    ParaId = metadata?.CommentParaId,
                    ParentParaId = metadata?.CommentParentParaId,
                    RootParaId = metadata?.CommentRootParaId,
                    DurableId = metadata?.CommentDurableId,
                    IsReply = metadata?.CommentIsReply,
                    Resolved = metadata?.CommentResolved,
                    TextLength = text?.CommentTextLength,
                    TextSnippet = text?.CommentTextSnippet,
                    TextTruncated = text?.CommentTextTruncated == true,
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

    private static IReadOnlyList<DocxChangeInfo> PairRangeBoundaries(IReadOnlyList<DocxChangeInfo> changes)
    {
        var starts = new Dictionary<RangeBoundaryKey, Queue<DocxChangeInfo>>();
        var paired = new Dictionary<string, string>(StringComparer.Ordinal);

        foreach (DocxChangeInfo change in changes)
        {
            RangeBoundaryMetadata? boundary = GetRangeBoundaryMetadata(change);
            if (boundary is null)
            {
                continue;
            }

            if (boundary.IsStart)
            {
                if (!starts.TryGetValue(boundary.Key, out Queue<DocxChangeInfo>? queue))
                {
                    queue = new Queue<DocxChangeInfo>();
                    starts[boundary.Key] = queue;
                }

                queue.Enqueue(change);
                continue;
            }

            if (starts.TryGetValue(boundary.Key, out Queue<DocxChangeInfo>? matchingStarts) &&
                matchingStarts.Count > 0)
            {
                DocxChangeInfo start = matchingStarts.Dequeue();
                paired[start.Id] = change.Id;
                paired[change.Id] = start.Id;
            }
        }

        if (paired.Count == 0)
        {
            return changes;
        }

        return changes
            .Select(change => paired.TryGetValue(change.Id, out string? pairedChangeId)
                ? change with { PairedChangeId = pairedChangeId }
                : change)
            .ToArray();
    }

    private static RangeBoundaryMetadata? GetRangeBoundaryMetadata(DocxChangeInfo change)
    {
        string? key = change.Type switch
        {
            "comment-range-start" or "comment-range-end" => change.CommentId,
            "move-from-range-start" or "move-from-range-end" => change.RevisionId,
            "move-to-range-start" or "move-to-range-end" => change.RevisionId,
            "custom-xml-insert-range-start" or "custom-xml-insert-range-end" => change.RevisionId,
            "custom-xml-delete-range-start" or "custom-xml-delete-range-end" => change.RevisionId,
            _ => null
        };

        if (string.IsNullOrWhiteSpace(key))
        {
            return null;
        }

        string pairKind = change.Type switch
        {
            "comment-range-start" or "comment-range-end" => "comment-range",
            "move-from-range-start" or "move-from-range-end" => "move-from-range",
            "move-to-range-start" or "move-to-range-end" => "move-to-range",
            "custom-xml-insert-range-start" or "custom-xml-insert-range-end" => "custom-xml-insert-range",
            "custom-xml-delete-range-start" or "custom-xml-delete-range-end" => "custom-xml-delete-range",
            _ => change.Type
        };
        bool isStart = change.Type.EndsWith("-start", StringComparison.Ordinal);
        return new RangeBoundaryMetadata(new RangeBoundaryKey(change.PartName, pairKind, key), isStart);
    }

    private static string NormalizeGroupKey(string? key)
    {
        return string.IsNullOrWhiteSpace(key) ? "(none)" : key;
    }
}
