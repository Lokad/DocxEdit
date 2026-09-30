using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

// D05: span-scoped text replacement for direct run-preserving edits. The map
// below mirrors ReadVisibleText exactly (deleted and moved-from subtrees stay
// invisible), while recording which characters live in direct paragraph runs.
// A match that stays inside directly editable text and crosses no protected
// boundary can be edited in place even when the paragraph carries protected
// content elsewhere; paragraph rewrites and tracked output keep the
// whole-paragraph gate because they rebuild the container.
internal static partial class DocxPatchEngine
{
    internal sealed record VisibleCharEntry(XElement TextElement, int OffsetInElement, bool InDirectRun, bool IsText);

    internal sealed record SpanMarkerEvent(XElement Element, int Position);

    internal sealed record ProtectedSpanInterval(int Start, int End, string Feature);

    // R02: sibling range markers (bookmarkStart/End, complex-field begin/end)
    // are zero-width, so ancestor checks cannot tell whether a match crosses
    // them. Each marker is recorded at its visible-text position during the map
    // walk and paired into half-open intervals, including ranges that start
    // before or end after this paragraph. A match overlapping a non-empty
    // interval is refused; touching an endpoint is allowed and empty ranges
    // never refuse.
    internal static string? FindProtectedSpanOverlap(
        XElement paragraph,
        int visibleLength,
        List<SpanMarkerEvent> markers,
        IReadOnlyList<TextRange> matches)
    {
        var intervals = new List<ProtectedSpanInterval>();
        var bookmarkStarts = new Dictionary<string, Queue<int>>(StringComparer.Ordinal);
        var fieldBegins = new Stack<int>();
        foreach (SpanMarkerEvent marker in markers)
        {
            if (marker.Element.Name == OoxmlNs.W + "bookmarkStart")
            {
                string? id = (string?)marker.Element.Attribute(OoxmlNs.W + "id");
                if (id is null)
                {
                    continue;
                }

                if (!bookmarkStarts.TryGetValue(id, out Queue<int>? starts))
                {
                    starts = new Queue<int>();
                    bookmarkStarts[id] = starts;
                }

                starts.Enqueue(marker.Position);
            }
            else if (marker.Element.Name == OoxmlNs.W + "bookmarkEnd")
            {
                string? id = (string?)marker.Element.Attribute(OoxmlNs.W + "id");
                if (id is null)
                {
                    continue;
                }

                if (bookmarkStarts.TryGetValue(id, out Queue<int>? starts) && starts.Count > 0)
                {
                    intervals.Add(new ProtectedSpanInterval(starts.Dequeue(), marker.Position, "bookmark"));
                }
                else
                {
                    intervals.Add(new ProtectedSpanInterval(0, marker.Position, "bookmark"));
                }
            }
            else if (marker.Element.Name == OoxmlNs.W + "fldChar")
            {
                string? kind = (string?)marker.Element.Attribute(OoxmlNs.W + "fldCharType");
                if (string.Equals(kind, "begin", StringComparison.Ordinal))
                {
                    fieldBegins.Push(marker.Position);
                }
                else if (string.Equals(kind, "end", StringComparison.Ordinal))
                {
                    intervals.Add(fieldBegins.Count > 0
                        ? new ProtectedSpanInterval(fieldBegins.Pop(), marker.Position, "field")
                        : new ProtectedSpanInterval(0, marker.Position, "field"));
                }
            }
        }

        foreach (KeyValuePair<string, Queue<int>> pending in bookmarkStarts)
        {
            while (pending.Value.Count > 0)
            {
                intervals.Add(new ProtectedSpanInterval(pending.Value.Dequeue(), visibleLength, "bookmark"));
            }
        }

        while (fieldBegins.Count > 0)
        {
            intervals.Add(new ProtectedSpanInterval(fieldBegins.Pop(), visibleLength, "field"));
        }

        foreach (TextRange match in matches)
        {
            foreach (ProtectedSpanInterval interval in intervals)
            {
                if (interval.Start < interval.End &&
                    match.Start < interval.End &&
                    interval.Start < match.Start + match.Length)
                {
                    return interval.Feature;
                }
            }
        }

        return null;
    }

    internal sealed class SpanEditPlan
    {
        public static SpanEditPlan Legacy { get; } = new SpanEditPlan { UseLegacyGate = true };

        public bool UseLegacyGate { get; init; }

        public string? ProtectedFeature { get; init; }

        public TextPosition?[]? Positions { get; init; }

        public IReadOnlyList<TextRange>? DirectMatches { get; init; }

        public List<VisibleCharEntry>? Map { get; init; }
    }

    internal static SpanEditPlan PlanSpanEdit(
        XElement paragraph,
        string current,
        IReadOnlyList<TextRange> matches)
    {
        var markers = new List<SpanMarkerEvent>();
        List<VisibleCharEntry> map = BuildVisibleTextMap(paragraph, markers);
        if (map.Count != current.Length)
        {
            return SpanEditPlan.Legacy;
        }

        var directPositions = new List<TextPosition?>();
        var directIndexByVisible = new int[map.Count];
        for (int i = 0; i < map.Count; i++)
        {
            VisibleCharEntry entry = map[i];
            if (!entry.InDirectRun)
            {
                directIndexByVisible[i] = -1;
                continue;
            }

            directIndexByVisible[i] = directPositions.Count;
            directPositions.Add(entry.IsText ? new TextPosition(entry.TextElement, entry.OffsetInElement) : null);
        }

        bool allDirect = true;
        foreach (TextRange match in matches)
        {
            if (match.Start < 0 || match.Length <= 0 || match.Start + match.Length > map.Count)
            {
                return SpanEditPlan.Legacy;
            }

            for (int i = match.Start; i < match.Start + match.Length; i++)
            {
                VisibleCharEntry entry = map[i];
                if (!entry.InDirectRun)
                {
                    allDirect = false;
                }

                foreach (XElement ancestor in entry.TextElement.Ancestors())
                {
                    if (ReferenceEquals(ancestor, paragraph))
                    {
                        break;
                    }

                    if (ProtectedTextEditElements.TryGetValue(ancestor.Name, out string? found) && found is not null)
                    {
                        return new SpanEditPlan { ProtectedFeature = found };
                    }
                }
            }
        }

        string? overlap = FindProtectedSpanOverlap(paragraph, map.Count, markers, matches);
        if (overlap is not null)
        {
            return new SpanEditPlan { ProtectedFeature = overlap };
        }

        if (!allDirect)
        {
            return SpanEditPlan.Legacy;
        }

        var directMatches = new List<TextRange>();
        foreach (TextRange match in matches)
        {
            directMatches.Add(new TextRange(directIndexByVisible[match.Start], match.Length));
        }

        return new SpanEditPlan { Positions = directPositions.ToArray(), DirectMatches = directMatches, Map = map };
    }

    internal static List<VisibleCharEntry> BuildVisibleTextMap(XElement paragraph, List<SpanMarkerEvent>? markers = null)
    {
        var map = new List<VisibleCharEntry>();
        foreach (XElement element in paragraph.Descendants())
        {
            if (element.Ancestors(OoxmlNs.W + "del").Any() ||
                element.Ancestors(OoxmlNs.W + "moveFrom").Any())
            {
                continue;
            }

            if (element.Name == OoxmlNs.W + "bookmarkStart" ||
                element.Name == OoxmlNs.W + "bookmarkEnd" ||
                element.Name == OoxmlNs.W + "fldChar")
            {
                markers?.Add(new SpanMarkerEvent(element, map.Count));
            }

            if (element.Name == OoxmlNs.W + "t")
            {
                XElement? run = element.Parent;
                bool direct = run is not null && run.Name == OoxmlNs.W + "r" && ReferenceEquals(run.Parent, paragraph);
                string value = element.Value;
                for (int i = 0; i < value.Length; i++)
                {
                    map.Add(new VisibleCharEntry(element, i, direct, true));
                }
            }
            else if (element.Name == OoxmlNs.W + "tab" || element.Name == OoxmlNs.W + "br")
            {
                XElement? run = element.Parent;
                bool direct = run is not null && run.Name == OoxmlNs.W + "r" && ReferenceEquals(run.Parent, paragraph);
                map.Add(new VisibleCharEntry(element, 0, direct, false));
            }
        }

        return map;
    }
}
