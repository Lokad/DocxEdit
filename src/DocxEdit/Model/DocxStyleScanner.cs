using System.Xml.Linq;
using DocxEdit.Ooxml;

namespace DocxEdit.Model;

internal static class DocxStyleScanner
{
    private static readonly HashSet<string> SupportedStyleTypes = new(StringComparer.Ordinal)
    {
        "paragraph",
        "character",
        "table"
    };

    public static IReadOnlyList<DocxStyleInfo> Scan(OoxmlPackage package, CancellationToken cancellationToken = default)
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
            bool isDefault = IsTrue((string?)style.Attribute(OoxmlNs.W + "default"));
            styles.Add(new DocxStyleInfo(styleId, name, type, isDefault));
        }

        return styles;
    }

    private static bool IsTrue(string? value)
    {
        return value is "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
    }
}
