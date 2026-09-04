using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

internal static class DocxStyleScanner
{
    private static readonly HashSet<string> SupportedStyleTypes = new(StringComparer.Ordinal)
    {
        "paragraph",
        "character",
        "table"
    };

    public static IReadOnlyList<DocxStyleInfo> Scan(OoxmlPackage package, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (package.MainDocumentPartName is null)
        {
            return [];
        }

        OoxmlRelationship? relationship = package
            .GetRelationships(package.MainDocumentPartName, cancellationToken)
            .FirstOrDefault(relationship =>
                !relationship.IsExternal &&
                relationship.Type == OoxmlRelTypes.Styles &&
                relationship.ResolvedTarget is not null);
        if (relationship?.ResolvedTarget is null)
        {
            return [];
        }

        OoxmlPart? stylesPart = package.GetPart(relationship.ResolvedTarget);
        if (stylesPart is null)
        {
            return [];
        }

        using Stream stream = stylesPart.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        var styles = new List<DocxStyleInfo>();
        foreach (XElement style in document.Root?.Elements(OoxmlNs.W + "style") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? styleId = (string?)style.Attribute(OoxmlNs.W + "styleId");
            string? type = (string?)style.Attribute(OoxmlNs.W + "type");
            if (string.IsNullOrWhiteSpace(styleId) ||
                string.IsNullOrWhiteSpace(type) ||
                !SupportedStyleTypes.Contains(type))
            {
                continue;
            }

            string name = (string?)style.Element(OoxmlNs.W + "name")?.Attribute(OoxmlNs.W + "val")
                ?? styleId;
            bool isDefault = WmlBoolean.IsTrue((string?)style.Attribute(OoxmlNs.W + "default"), valueWhenMissing: false);
            XElement? numberingProperties = style
                .Element(OoxmlNs.W + "pPr")
                ?.Element(OoxmlNs.W + "numPr");
            string? numberingId = (string?)numberingProperties
                ?.Element(OoxmlNs.W + "numId")
                ?.Attribute(OoxmlNs.W + "val");
            int? numberingLevel = XmlValues.TryReadInt((string?)numberingProperties
                ?.Element(OoxmlNs.W + "ilvl")
                ?.Attribute(OoxmlNs.W + "val"));
            styles.Add(new DocxStyleInfo(styleId, name, type, isDefault)
            {
                BasedOnStyleId = (string?)style.Element(OoxmlNs.W + "basedOn")?.Attribute(OoxmlNs.W + "val"),
                NextStyleId = (string?)style.Element(OoxmlNs.W + "next")?.Attribute(OoxmlNs.W + "val"),
                LinkedStyleId = (string?)style.Element(OoxmlNs.W + "link")?.Attribute(OoxmlNs.W + "val"),
                NumberingId = numberingId,
                NumberingLevel = numberingLevel
            });
        }

        return styles;
    }
}
