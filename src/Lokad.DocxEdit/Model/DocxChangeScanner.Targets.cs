using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

internal static partial class DocxChangeScanner
{
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

    private static TargetMetadata FindTarget(
        XElement element,
        IReadOnlyDictionary<XElement, string> targets,
        string? commentAnchorTargetId)
    {
        string? ancestorTargetId = FindAncestorTarget(element, targets);
        if (ancestorTargetId is not null)
        {
            return new TargetMetadata(ancestorTargetId, "targeted", "ancestor", null, null, null);
        }

        string? adjacentRangeTargetId = FindAdjacentRangeTarget(element, targets);
        if (adjacentRangeTargetId is not null)
        {
            return new TargetMetadata(
                adjacentRangeTargetId,
                "targeted",
                "adjacent-range",
                "range-boundary",
                null,
                "Range boundary is outside a modeled block; target is the nearest adjacent modeled block and should be treated as context.");
        }

        if (commentAnchorTargetId is not null)
        {
            return new TargetMetadata(
                null,
                "comment-anchor",
                "comment-anchor",
                null,
                null,
                "Linked through matching comment anchor metadata.");
        }

        string? nearestTargetId = FindNearestSurroundingTarget(element, targets);
        string reason = GetTargetlessReason(element);
        return new TargetMetadata(
            null,
            "targetless",
            "none",
            reason,
            nearestTargetId,
            GetTargetlessNote(reason, nearestTargetId));
    }

    private static string? FindAncestorTarget(XElement element, IReadOnlyDictionary<XElement, string> targets)
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

    private static string? FindNearestSurroundingTarget(XElement element, IReadOnlyDictionary<XElement, string> targets)
    {
        string? bestTargetId = null;
        int bestDistance = int.MaxValue;

        int distance = 1;
        foreach (XElement sibling in element.ElementsBeforeSelf().Reverse())
        {
            if (TryFindTargetInSubtree(sibling, targets, preferLast: true, out string? targetId) &&
                distance < bestDistance)
            {
                bestTargetId = targetId;
                bestDistance = distance;
            }

            distance++;
        }

        distance = 1;
        foreach (XElement sibling in element.ElementsAfterSelf())
        {
            if (TryFindTargetInSubtree(sibling, targets, preferLast: false, out string? targetId) &&
                distance < bestDistance)
            {
                bestTargetId = targetId;
                bestDistance = distance;
            }

            distance++;
        }

        return bestTargetId;
    }

    private static string GetTargetlessReason(XElement element)
    {
        if (IsRangeBoundaryElement(element))
        {
            return "range-boundary-no-adjacent-target";
        }

        if (element.Parent?.Name == OoxmlNs.W + "body")
        {
            return "body-level-markup";
        }

        if (element.Ancestors(OoxmlNs.W + "comments").Any())
        {
            return "comment-story";
        }

        if (element.Ancestors(OoxmlNs.W + "body").Any())
        {
            return "unmodeled-body-structure";
        }

        return "unmodeled-word-part";
    }

    private static string GetTargetlessNote(string reason, string? nearestTargetId)
    {
        string note = reason switch
        {
            "range-boundary-no-adjacent-target" => "Range boundary is outside a modeled block and no adjacent modeled block was found.",
            "body-level-markup" => "Markup is a direct child of the document body rather than a modeled paragraph, table, cell, or section.",
            "comment-story" => "Markup is in a comment story without a modeled comment or main-story anchor target.",
            "unmodeled-body-structure" => "Markup is inside a body structure that DocxEdit does not model as an edit target.",
            _ => "Markup is in a Word XML part but outside the currently modeled edit-target structures."
        };

        return nearestTargetId is null
            ? note
            : $"{note} nearest-target={nearestTargetId} is context only, not exact ownership.";
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
}
