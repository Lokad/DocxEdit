using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

internal static partial class DocxPackageValidator
{
    private static void ValidateParagraphStyleReferences(
        OoxmlPackage package,
        string partName,
        XDocument document,
        List<DocxDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        HashSet<string>? paragraphStyleIds = ReadParagraphStyleIds(package, cancellationToken);
        if (paragraphStyleIds is null)
        {
            return;
        }

        foreach (string styleId in document
            .Descendants(OoxmlNs.W + "pStyle")
            .Select(style => (string?)style.Attribute(OoxmlNs.W + "val"))
            .Where(styleId => !string.IsNullOrWhiteSpace(styleId))
            .OfType<string>()
            .Distinct(StringComparer.Ordinal)
            .Order(StringComparer.Ordinal))
        {
            if (!paragraphStyleIds.Contains(styleId))
            {
                diagnostics.Add(Warning(
                    "W9116",
                    $"Paragraph style reference '{styleId}' is not defined in /word/styles.xml.",
                    partName,
                    "style",
                    "missing-style-definition"));
            }
        }
    }

    private static HashSet<string>? ReadParagraphStyleIds(
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        OoxmlPart? stylesPart = package.GetPart("/word/styles.xml");
        if (stylesPart is null)
        {
            return null;
        }

        using Stream stream = stylesPart.OpenRead();
        XDocument styles = SafeXml.Load(stream, cancellationToken);
        return styles
            .Descendants(OoxmlNs.W + "style")
            .Where(style => string.Equals((string?)style.Attribute(OoxmlNs.W + "type"), "paragraph", StringComparison.Ordinal))
            .Select(style => (string?)style.Attribute(OoxmlNs.W + "styleId"))
            .Where(styleId => !string.IsNullOrWhiteSpace(styleId))
            .OfType<string>()
            .ToHashSet(StringComparer.Ordinal);
    }

    private static void ValidateNumberingReferences(
        OoxmlPackage package,
        string partName,
        XDocument document,
        List<DocxDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        if (string.Equals(partName, "/word/numbering.xml", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        HashSet<string>? numberingIds = ReadNumberingIds(package, cancellationToken);
        if (numberingIds is null)
        {
            return;
        }

        foreach (string numberingId in document
            .Descendants(OoxmlNs.W + "numPr")
            .Elements(OoxmlNs.W + "numId")
            .Select(numId => (string?)numId.Attribute(OoxmlNs.W + "val"))
            .Where(numberingId => !string.IsNullOrWhiteSpace(numberingId) && numberingId != "0")
            .OfType<string>()
            .Distinct(StringComparer.Ordinal)
            .Order(StringComparer.Ordinal))
        {
            if (!numberingIds.Contains(numberingId))
            {
                diagnostics.Add(Warning(
                    "W9117",
                    $"Numbering reference '{numberingId}' is not defined in /word/numbering.xml.",
                    partName,
                    "numbering",
                    "missing-numbering-definition"));
            }
        }
    }

    private static void ValidateNumberingDefinitions(
        string partName,
        XDocument document,
        List<DocxDiagnostic> diagnostics)
    {
        if (!string.Equals(partName, "/word/numbering.xml", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var abstractIds = document
            .Descendants(OoxmlNs.W + "abstractNum")
            .Select(abstractNum => (string?)abstractNum.Attribute(OoxmlNs.W + "abstractNumId"))
            .Where(abstractId => !string.IsNullOrWhiteSpace(abstractId))
            .OfType<string>()
            .ToHashSet(StringComparer.Ordinal);

        foreach (string abstractId in document
            .Descendants(OoxmlNs.W + "num")
            .Elements(OoxmlNs.W + "abstractNumId")
            .Select(abstractNumId => (string?)abstractNumId.Attribute(OoxmlNs.W + "val"))
            .Where(abstractId => !string.IsNullOrWhiteSpace(abstractId))
            .OfType<string>()
            .Distinct(StringComparer.Ordinal)
            .Order(StringComparer.Ordinal))
        {
            if (!abstractIds.Contains(abstractId))
            {
                diagnostics.Add(Warning(
                    "W9117",
                    $"Numbering definition references missing abstractNumId '{abstractId}'.",
                    partName,
                    "numbering",
                    "missing-abstract-numbering-definition"));
            }
        }
    }
}
