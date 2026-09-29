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
        CancellationToken cancellationToken)
    {
        if (operation.OperationName is "insert-before" or "insert-after")
        {
            return CreatedInsertParagraphId(operation, package, cancellationToken);
        }

        if (operation.OperationName == "add-comment" && commentsBefore is not null)
        {
            return CreatedCommentIds(package, commentsBefore, cancellationToken);
        }

        return [];
    }

    private static IReadOnlyList<string> CreatedInsertParagraphId(
        DocxPatchOperation operation,
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        string? target = operation.Fields.GetValueOrDefault("target");
        if (target is null)
        {
            return [];
        }

        BlockTarget? anchor = ResolveBlockTarget(package, operation, target, cancellationToken, out _);
        if (anchor is null || (anchor.Block.Name != OoxmlNs.W + "p" && anchor.Block.Name != OoxmlNs.W + "tbl"))
        {
            return [];
        }

        bool insertAfter = !string.Equals(operation.OperationName, "insert-before", StringComparison.Ordinal);
        XElement? created = insertAfter
            ? anchor.Block.ElementsAfterSelf(OoxmlNs.W + "p").FirstOrDefault()
            : anchor.Block.ElementsBeforeSelf(OoxmlNs.W + "p").LastOrDefault();
        if (created is null)
        {
            return [];
        }

        int ordinal = created.ElementsBeforeSelf(OoxmlNs.W + "p").Count() + 1;
        if (!DocxPartRoles.GetStoryPrefixes(package, cancellationToken).TryGetValue(anchor.PartName, out string? prefix) || prefix is null || prefix.Length == 0)
        {
            return [];
        }

        int storyPart = 0;
        if (prefix.Length > 1 && !int.TryParse(prefix.Substring(1), NumberStyles.Integer, CultureInfo.InvariantCulture, out storyPart))
        {
            return [];
        }

        return [new DocxTargetId(prefix[0], storyPart, DocxTargetKind.Paragraph, ordinal, 0, 0).ToWireValue()];
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
}
