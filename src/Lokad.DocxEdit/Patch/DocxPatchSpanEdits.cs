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

    internal sealed class SpanEditPlan
    {
        public static SpanEditPlan Legacy { get; } = new SpanEditPlan { UseLegacyGate = true };

        public bool UseLegacyGate { get; init; }

        public string? ProtectedFeature { get; init; }

        public TextPosition?[]? Positions { get; init; }

        public IReadOnlyList<TextRange>? DirectMatches { get; init; }
    }

    internal static SpanEditPlan PlanSpanEdit(
        XElement paragraph,
        string current,
        IReadOnlyList<TextRange> matches)
    {
        List<VisibleCharEntry> map = BuildVisibleTextMap(paragraph);
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

        if (!allDirect)
        {
            return SpanEditPlan.Legacy;
        }

        var directMatches = new List<TextRange>();
        foreach (TextRange match in matches)
        {
            directMatches.Add(new TextRange(directIndexByVisible[match.Start], match.Length));
        }

        return new SpanEditPlan { Positions = directPositions.ToArray(), DirectMatches = directMatches };
    }

    internal static List<VisibleCharEntry> BuildVisibleTextMap(XElement paragraph)
    {
        var map = new List<VisibleCharEntry>();
        foreach (XElement element in paragraph.Descendants())
        {
            if (element.Ancestors(OoxmlNs.W + "del").Any() ||
                element.Ancestors(OoxmlNs.W + "moveFrom").Any())
            {
                continue;
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
