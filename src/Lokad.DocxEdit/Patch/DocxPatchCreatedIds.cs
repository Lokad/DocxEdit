using System.Globalization;
using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

// D13: created target IDs for operation reports. Inserted paragraphs are
// identified by adjacency to the anchor element at report time, so earlier
// structural edits cannot shift the answer; new comments use allocated IDs,
// which are stable. Rows already report inserted IDs through affected targets.
internal static partial class DocxPatchEngine
{
    private static IReadOnlyList<string> BuildCreatedTargetIds(
        DocxPatchOperation operation,
        OoxmlPackage package,
        IReadOnlySet<string>? commentsBefore,
        IReadOnlyDictionary<string, int>? bookmarkCountsBefore,
        CancellationToken cancellationToken)
    {
        if (operation.OperationName is "insert-before" or "insert-after" or "insert-image-after")
        {
            return CreatedInsertParagraphId(operation, package, cancellationToken);
        }

        if ((operation.OperationName == "add-comment" || operation.OperationName == "add-comment-reply") && commentsBefore is not null)
        {
            return CreatedCommentIds(package, commentsBefore, cancellationToken);
        }

        if (operation.OperationName == "add-bookmark" && bookmarkCountsBefore is not null)
        {
            return CreatedBookmarkIds(package, bookmarkCountsBefore, cancellationToken);
        }

        return [];
    }

    private static IReadOnlyList<string> CreatedInsertParagraphId(
        DocxPatchOperation operation,
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        BlockTarget? anchor = ResolveInsertAnchor(operation, package, cancellationToken);
        if (anchor is null)
        {
            return [];
        }

        bool insertAfter = !string.Equals(operation.OperationName, "insert-before", StringComparison.Ordinal);
        int count = Math.Max(1, operation.FieldValues.Count(static field => field.Name == "text"));
        List<XElement> created = FindAdjacentInsertParagraphs(anchor.Block, insertAfter, count);
        if (created.Count != count)
        {
            return [];
        }
        if (!DocxPartRoles.GetStoryPrefixes(package, cancellationToken).TryGetValue(anchor.PartName, out string? prefix) || !TryParseStoryPrefix(prefix, out char story, out int storyPart))
        {
            return [];
        }

        return created.Select(element => new DocxTargetId(story, storyPart, DocxTargetKind.Paragraph, element.ElementsBeforeSelf(OoxmlNs.W + "p").Count() + 1, 0, 0).ToWireValue()).ToArray();
    }

    private static IReadOnlyList<string> CreatedCommentIds(
        OoxmlPackage package,
        IReadOnlySet<string> commentsBefore,
        CancellationToken cancellationToken)
    {
        HashSet<string> after = ReadCommentIds(package, cancellationToken);
        after.ExceptWith(commentsBefore);
        return after
            .OrderBy(static id => int.TryParse(id, NumberStyles.Integer, CultureInfo.InvariantCulture, out int n) ? n : int.MaxValue)
            .Select(static id => "comment:" + id)
            .ToArray();
    }

    private static HashSet<string> ReadCommentIds(OoxmlPackage package, CancellationToken cancellationToken)
    {
        HashSet<string> ids = new(StringComparer.Ordinal);
        foreach (string partName in GetCommentsPartNames(package, cancellationToken))
        {
            if (package.GetPart(partName) is null)
            {
                continue;
            }

            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
            foreach (XElement comment in document.Descendants(OoxmlNs.W + "comment"))
            {
                if (comment.Attribute(OoxmlNs.W + "id")?.Value is string id)
                {
                    ids.Add(id);
                }
            }
        }

        return ids;
    }

    private static bool TryParseStoryPrefix(string? prefix, out char story, out int storyPart)
    {
        story = (char)77;
        storyPart = 0;
        if (string.IsNullOrEmpty(prefix))
        {
            return false;
        }

        story = prefix[0];
        if (prefix.Length > 1 && !int.TryParse(prefix.Substring(1), NumberStyles.Integer, CultureInfo.InvariantCulture, out storyPart))
        {
            return false;
        }

        return true;
    }

    private static Dictionary<string, int> CountBookmarkStarts(OoxmlPackage package, CancellationToken cancellationToken)
    {
        Dictionary<string, int> counts = new(StringComparer.Ordinal);
        foreach (string partName in GetEditableStoryPartNames(package, cancellationToken))
        {
            if (package.GetPart(partName) is null)
            {
                continue;
            }

            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
            counts[partName] = document.Descendants(OoxmlNs.W + "bookmarkStart").Count();
        }

        return counts;
    }

    private static IReadOnlyList<string> CreatedBookmarkIds(
        OoxmlPackage package,
        IReadOnlyDictionary<string, int> bookmarksBefore,
        CancellationToken cancellationToken)
    {
        Dictionary<string, int> after = CountBookmarkStarts(package, cancellationToken);
        IReadOnlyDictionary<string, string> prefixes = DocxPartRoles.GetStoryPrefixes(package, cancellationToken);
        List<string> created = [];
        foreach ((string partName, int count) in after)
        {
            if (count - bookmarksBefore.GetValueOrDefault(partName) != 1)
            {
                continue;
            }

            if (!prefixes.TryGetValue(partName, out string? prefix) || !TryParseStoryPrefix(prefix, out char story, out int storyPart))
            {
                continue;
            }

            created.Add(new DocxTargetId(story, storyPart, DocxTargetKind.Bookmark, count, 0, 0).ToWireValue());
        }

        return created;
    }
}
