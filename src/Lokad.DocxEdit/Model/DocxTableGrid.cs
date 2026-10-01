using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

// Single owner for the small non-validating table-grid primitives: row grid
// offsets and cell column spans. Scanners, placement enumeration, and patch
// resolution share these; validating readers that emit diagnostics stay where
// they are.
internal static class DocxTableGrid
{
    public static int ReadGridOffset(XElement row, string localName)
    {
        string? value = (string?)row
            .Element(OoxmlNs.W + "trPr")
            ?.Element(OoxmlNs.W + localName)
            ?.Attribute(OoxmlNs.W + "val");
        return int.TryParse(value, out int parsed) && parsed > 0 ? parsed : 0;
    }

    public static int ReadColumnSpan(XElement cell)
    {
        string? spanText = (string?)cell
            .Element(OoxmlNs.W + "tcPr")
            ?.Element(OoxmlNs.W + "gridSpan")
            ?.Attribute(OoxmlNs.W + "val");
        return int.TryParse(spanText, out int span) && span > 0 ? span : 1;
    }
}
