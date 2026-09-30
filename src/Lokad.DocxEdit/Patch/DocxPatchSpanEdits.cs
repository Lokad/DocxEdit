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

    private static readonly (XName Start, XName End, string Feature)[] SpanRangeMarkers =
    [
        (OoxmlNs.W + "bookmarkStart", OoxmlNs.W + "bookmarkEnd", "bookmark"),
        (OoxmlNs.W + "commentRangeStart", OoxmlNs.W + "commentRangeEnd", "comment"),
        (OoxmlNs.W + "moveFromRangeStart", OoxmlNs.W + "moveFromRangeEnd", "tracked-move-from-range"),
        (OoxmlNs.W + "moveToRangeStart", OoxmlNs.W + "moveToRangeEnd", "tracked-move-to-range"),
        (OoxmlNs.W + "customXmlInsRangeStart", OoxmlNs.W + "customXmlInsRangeEnd", "tracked-custom-xml-insertion"),
        (OoxmlNs.W + "customXmlDelRangeStart", OoxmlNs.W + "customXmlDelRangeEnd", "tracked-custom-xml-deletion"),
        (OoxmlNs.W + "customXmlMoveFromRangeStart", OoxmlNs.W + "customXmlMoveFromRangeEnd", "tracked-custom-xml-move-from"),
        (OoxmlNs.W + "customXmlMoveToRangeStart", OoxmlNs.W + "customXmlMoveToRangeEnd", "tracked-custom-xml-move-to"),
    ];

    private static bool IsSpanRangeMarker(XName name)
    {
        foreach ((XName start, XName end, string _) in SpanRangeMarkers)
        {
            if (name == start || name == end)
            {
                return true;
            }
        }

        return false;
    }

    // C01: story-wide range membership. Walk from the story root in document
    // order up to the target element and track open bookmark, comment, move,
    // custom-XML ranges plus complex-field depth. Markers inside deleted or
    // moved-from subtrees stay invisible and do not contribute. Unmatched ends
    // before the target do not carry forward. The result describes ranges that
    // are already open when the target starts, so an interior paragraph with
    // no local markers is still recognized as protected.
    internal static Dictionary<(int Row, string Id), int> GetEnteringRangeCounts(
        XElement target,
        out int enteringFieldDepth)
    {
        var entering = new Dictionary<(int Row, string Id), int>();
        enteringFieldDepth = 0;
        XElement root = target;
        while (root.Parent is not null)
        {
            root = root.Parent;
        }

        foreach (XElement element in root.Descendants())
        {
            if (ReferenceEquals(element, target))
            {
                break;
            }

            if (element.Ancestors(OoxmlNs.W + "del").Any() ||
                element.Ancestors(OoxmlNs.W + "moveFrom").Any())
            {
                continue;
            }

            int row = -1;
            for (int r = 0; r < SpanRangeMarkers.Length; r++)
            {
                if (element.Name == SpanRangeMarkers[r].Start || element.Name == SpanRangeMarkers[r].End)
                {
                    row = r;
                    break;
                }
            }

            if (row >= 0)
            {
                string? id = (string?)element.Attribute(OoxmlNs.W + "id");
                if (id is null)
                {
                    continue;
                }
                var key = (row, id);
                bool isStart = element.Name == SpanRangeMarkers[row].Start;
                if (isStart)
                {
                    if (!entering.TryGetValue(key, out int count))
                    {
                        count = 0;
                    }

                    entering[key] = count + 1;
                }
                else
                {
                    if (entering.TryGetValue(key, out int count) && count > 0)
                    {
                        if (count == 1)
                        {
                            entering.Remove(key);
                        }
                        else
                        {
                            entering[key] = count - 1;
                        }
                    }
                }

                continue;
            }

            if (element.Name == OoxmlNs.W + "fldChar")
            {
                string? kind = (string?)element.Attribute(OoxmlNs.W + "fldCharType");
                if (string.Equals(kind, "begin", StringComparison.Ordinal))
                {
                    enteringFieldDepth++;
                }
                else if (string.Equals(kind, "end", StringComparison.Ordinal) && enteringFieldDepth > 0)
                {
                    enteringFieldDepth--;
                }
            }
        }

        var startCounts = new Dictionary<(int Row, string Id), int>();
        var endCounts = new Dictionary<(int Row, string Id), int>();
        int totalBegins = 0;
        int totalEnds = 0;
        foreach (XElement element in root.Descendants())
        {
            int hrow = -1;
            for (int h = 0; h < SpanRangeMarkers.Length; h++)
            {
                if (element.Name == SpanRangeMarkers[h].Start || element.Name == SpanRangeMarkers[h].End)
                {
                    hrow = h;
                    break;
                }
            }
            if (hrow >= 0)
            {
                string? hid = (string?)element.Attribute(OoxmlNs.W + "id");
                if (hid is null)
                {
                    continue;
                }
                var hkey = (hrow, hid);
                bool hisStart = element.Name == SpanRangeMarkers[hrow].Start;
                if (hisStart)
                {
                    int hc = 0;
                    startCounts.TryGetValue(hkey, out hc);
                    startCounts[hkey] = hc + 1;
                }
                else
                {
                    int hc2 = 0;
                    endCounts.TryGetValue(hkey, out hc2);
                    endCounts[hkey] = hc2 + 1;
                }
                continue;
            }
            if (element.Name == OoxmlNs.W + "fldChar")
            {
                string? hkind = (string?)element.Attribute(OoxmlNs.W + "fldCharType");
                if (string.Equals(hkind, "begin", StringComparison.Ordinal))
                {
                    totalBegins++;
                }
                else if (string.Equals(hkind, "end", StringComparison.Ordinal))
                {
                    totalEnds++;
                }
            }
        }

        var unhealthy = new List<(int Row, string Id)>();
        foreach (KeyValuePair<(int Row, string Id), int> kv in entering)
        {
            int sc = 0;
            int ec = 0;
            startCounts.TryGetValue(kv.Key, out sc);
            endCounts.TryGetValue(kv.Key, out ec);
            if (sc != 1 || ec != 1)
            {
                unhealthy.Add(kv.Key);
            }
        }
        foreach ((int Row, string Id) key in unhealthy)
        {
            entering.Remove(key);
        }
        if (totalBegins != totalEnds)
        {
            enteringFieldDepth = 0;
        }

        return entering;
    }

    internal static bool TryGetStoryEnteringProtectedFeature(
        XElement target,
        out string feature)
    {
        Dictionary<(int Row, string Id), int> entering = GetEnteringRangeCounts(target, out int fieldDepth);
        foreach (KeyValuePair<(int Row, string Id), int> entry in entering)
        {
            if (entry.Value > 0)
            {
                feature = SpanRangeMarkers[entry.Key.Row].Feature;
                return true;
            }
        }

        if (fieldDepth > 0)
        {
            feature = "field";
            return true;
        }

        feature = string.Empty;
        return false;
    }

    // Shared gate for whole-paragraph operations. Entering ranges and fields
    // come first so interior targets with no local markers are still refused.
    // Otherwise fall back to the paragraph-only protected element scan.
    internal static bool TryGetStoryProtectedTextEditFeature(
        XElement paragraph,
        out string feature)
    {
        if (TryGetStoryEnteringProtectedFeature(paragraph, out string entering))
        {
            feature = entering;
            return true;
        }

        return TryGetProtectedTextEditFeature(paragraph, out feature);
    }


    // R02: sibling range markers (bookmark, comment, move, and custom-XML ranges plus complex-field begin/end)
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
        var openRanges = new Dictionary<(int Row, string Id), Queue<int>>();
        var fieldBegins = new Stack<int>();
        Dictionary<(int Row, string Id), int> entering = GetEnteringRangeCounts(paragraph, out int enteringFieldDepth);
        foreach (KeyValuePair<(int Row, string Id), int> entry in entering)
        {
            var starts = new Queue<int>();
            for (int i = 0; i < entry.Value; i++)
            {
                starts.Enqueue(0);
            }
            openRanges[entry.Key] = starts;
        }

        for (int i = 0; i < enteringFieldDepth; i++)
        {
            fieldBegins.Push(0);
        }

        foreach (SpanMarkerEvent marker in markers)
        {
            int row = -1;
            for (int r = 0; r < SpanRangeMarkers.Length; r++)
            {
                if (marker.Element.Name == SpanRangeMarkers[r].Start || marker.Element.Name == SpanRangeMarkers[r].End)
                {
                    row = r;
                    break;
                }
            }

            if (row >= 0 && marker.Element.Name == SpanRangeMarkers[row].Start)
            {
                string? id = (string?)marker.Element.Attribute(OoxmlNs.W + "id");
                if (id is null)
                {
                    continue;
                }

                var key = (row, id);
                if (!openRanges.TryGetValue(key, out Queue<int>? starts))
                {
                    starts = new Queue<int>();
                    openRanges[key] = starts;
                }

                starts.Enqueue(marker.Position);
            }
            else if (row >= 0)
            {
                string? id = (string?)marker.Element.Attribute(OoxmlNs.W + "id");
                if (id is null)
                {
                    continue;
                }

                var key = (row, id);
                if (openRanges.TryGetValue(key, out Queue<int>? starts) && starts.Count > 0)
                {
                    intervals.Add(new ProtectedSpanInterval(starts.Dequeue(), marker.Position, SpanRangeMarkers[row].Feature));
                }
                else
                {
                    intervals.Add(new ProtectedSpanInterval(0, marker.Position, SpanRangeMarkers[row].Feature));
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

        foreach (KeyValuePair<(int Row, string Id), Queue<int>> pending in openRanges)
        {
            while (pending.Value.Count > 0)
            {
                intervals.Add(new ProtectedSpanInterval(pending.Value.Dequeue(), visibleLength, SpanRangeMarkers[pending.Key.Row].Feature));
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

            if (element.Name == OoxmlNs.W + "fldChar" || IsSpanRangeMarker(element.Name))
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
