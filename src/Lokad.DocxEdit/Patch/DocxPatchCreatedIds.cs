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
        IReadOnlyDictionary<string, HashSet<string>>? bookmarkIdsBefore,
        CancellationToken cancellationToken)
    {
        if (operation.OperationName is "insert-before" or "insert-after" or "insert-image-after" or "insert-hyperlink-after")
        {
            return CreatedInsertParagraphId(operation, package, cancellationToken);
        }

        if ((operation.OperationName == "add-comment" || operation.OperationName == "add-comment-reply") && commentsBefore is not null)
        {
            return CreatedCommentIds(package, commentsBefore, cancellationToken);
        }

        if (operation.OperationName == "add-bookmark" && bookmarkIdsBefore is not null)
        {
            return CreatedBookmarkIds(operation, package, bookmarkIdsBefore, cancellationToken);
        }

        return [];
    }

    private static IReadOnlyList<string> CreatedInsertParagraphId(
        DocxPatchOperation operation,
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        var wireIds = new List<string>();
        var prefixes = DocxPartRoles.GetStoryPrefixes(package, cancellationToken);
        int createdIndex = 0;
        while (true)
        {
            string mark = CreatedMarkValue(operation, createdIndex);
            bool found = false;
            foreach (var entry in prefixes)
            {
                if (!TryParseStoryPrefix(entry.Value, out char story, out int storyPart))
                {
                    continue;
                }
                XDocument document = LoadDocumentPart(package, entry.Key, cancellationToken, out XElement _);
                XElement? created = null;
                foreach (XElement element in document.Descendants())
                {
                    if (string.Equals((string?)element.Attribute(SnapshotCreatedName), mark, StringComparison.Ordinal))
                    {
                        created = element;
                        break;
                    }
                }
                if (created is null)
                {
                    continue;
                }
                int? ordinal = PhysicalParagraphOrdinal(document, created);
                if (ordinal is null)
                {
                    return [];
                }
                wireIds.Add(new DocxTargetId(story, storyPart, DocxTargetKind.Paragraph, ordinal.Value, 0, 0).ToWireValue());
                found = true;
                break;
            }
            if (!found)
            {
                break;
            }
            createdIndex++;
        }
        return wireIds;
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

    private static Dictionary<string, HashSet<string>> CollectBookmarkOoxmlIds(OoxmlPackage package, CancellationToken cancellationToken)
    {
        Dictionary<string, HashSet<string>> ids = new(StringComparer.Ordinal);
        foreach (string partName in GetEditableStoryPartNames(package, cancellationToken))
        {
            if (package.GetPart(partName) is null)
            {
                continue;
            }

            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
            ids[partName] = document
                .Descendants(OoxmlNs.W + "bookmarkStart")
                .Select(start => (string?)start.Attribute(OoxmlNs.W + "id"))
                .Where(id => id is not null)
                .Cast<string>()
                .ToHashSet(StringComparer.Ordinal);
        }

        return ids;
    }

    private static IReadOnlyList<string> CreatedBookmarkIds(
        DocxPatchOperation operation,
        OoxmlPackage package,
        IReadOnlyDictionary<string, HashSet<string>> bookmarksBefore,
        CancellationToken cancellationToken)
    {
        Dictionary<string, HashSet<string>> after = CollectBookmarkOoxmlIds(package, cancellationToken);
        var touched = new Dictionary<string, XDocument>(StringComparer.OrdinalIgnoreCase);
        IReadOnlyDictionary<string, string> prefixes = DocxPartRoles.GetStoryPrefixes(package, cancellationToken);
        List<string> created = [];
        foreach ((string partName, HashSet<string> ids) in after)
        {
            bookmarksBefore.TryGetValue(partName, out HashSet<string>? before);
            List<string> added = ids.Where(id => before is null || !before.Contains(id)).ToList();
            if (added.Count != 1)
            {
                continue;
            }

            if (!prefixes.TryGetValue(partName, out string? prefix) || !TryParseStoryPrefix(prefix, out char story, out int storyPart))
            {
                continue;
            }

            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
            int ordinal = 0;
            foreach (XElement start in document.Descendants(OoxmlNs.W + "bookmarkStart"))
            {
                if (string.IsNullOrWhiteSpace((string?)start.Attribute(OoxmlNs.W + "name")))
                {
                    continue;
                }

                ordinal++;
                if (string.Equals((string?)start.Attribute(OoxmlNs.W + "id"), added[0], StringComparison.Ordinal))
                {
                    start.SetAttributeValue(SnapshotCreatedName, CreatedMarkValue(operation, created.Count));
                    touched[partName] = document;
                    created.Add(new DocxTargetId(story, storyPart, DocxTargetKind.Bookmark, ordinal, 0, 0).ToWireValue());
                    break;
                }
            }
        }

        foreach ((string partName, XDocument document) in touched)
        {
            SaveDocumentPart(package, partName, document);
        }

        return created;
    }
}
