using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

// Single owner for Word heading-level resolution (D09). A paragraph is a
// heading when it carries outline semantics, not when its style ID happens
// to contain a digit: direct w:pPr/w:outlineLvl wins (0-based, so 0 means
// level 1), then the paragraph style chain (w:basedOn) in styles.xml, then
// the built-in latent Heading1..Heading9 levels when the style is not
// defined at all. An explicitly defined style without outlineLvl, and an
// outlineLvl of 9 or more (body text), mean not-a-heading.
internal static class DocxHeadingLevels
{
    public static int? GetHeadingLevel(OoxmlPackage package, XElement paragraph, CancellationToken cancellationToken)
    {
        XElement? properties = paragraph.Element(OoxmlNs.W + "pPr");
        if (TryReadOutlineLevel(properties, out int? directLevel))
        {
            return directLevel;
        }

        string? styleId = (string?)properties
            ?.Element(OoxmlNs.W + "pStyle")
            ?.Attribute(OoxmlNs.W + "val");
        if (string.IsNullOrEmpty(styleId))
        {
            return null;
        }

        if (TryReadStyleChainLevel(package, styleId, cancellationToken, out int? chainLevel, out bool defined))
        {
            return chainLevel;
        }

        if (!defined && TryReadBuiltInLevel(styleId, out int builtInLevel))
        {
            return builtInLevel;
        }

        return null;
    }

    private static bool TryReadOutlineLevel(XElement? paragraphProperties, out int? level)
    {
        level = null;
        string? value = (string?)paragraphProperties
            ?.Element(OoxmlNs.W + "outlineLvl")
            ?.Attribute(OoxmlNs.W + "val");
        if (!int.TryParse(value, out int parsed))
        {
            return false;
        }

        level = parsed is >= 0 and <= 8 ? parsed + 1 : null;
        return true;
    }

    private static bool TryReadStyleChainLevel(
        OoxmlPackage package,
        string styleId,
        CancellationToken cancellationToken,
        out int? level,
        out bool defined)
    {
        level = null;
        defined = false;
        XElement? stylesRoot = LoadStylesRoot(package, cancellationToken);
        if (stylesRoot is null)
        {
            return false;
        }

        var visited = new HashSet<string>(StringComparer.Ordinal);
        string? current = styleId;
        while (!string.IsNullOrEmpty(current) && visited.Add(current))
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement? style = stylesRoot
                .Elements(OoxmlNs.W + "style")
                .FirstOrDefault(element => string.Equals((string?)element.Attribute(OoxmlNs.W + "styleId"), current, StringComparison.Ordinal));
            if (style is null)
            {
                return defined;
            }

            defined = true;
            if (TryReadOutlineLevel(style.Element(OoxmlNs.W + "pPr"), out int? styleLevel))
            {
                level = styleLevel;
                return true;
            }

            current = (string?)style.Element(OoxmlNs.W + "basedOn")?.Attribute(OoxmlNs.W + "val");
        }

        return true;
    }

    private static XElement? LoadStylesRoot(OoxmlPackage package, CancellationToken cancellationToken)
    {
        string? stylesPartName = DocxPartRoles.FindStylesPartName(package, cancellationToken);
        OoxmlPart? stylesPart = stylesPartName is null ? null : package.GetPart(stylesPartName);
        if (stylesPart is null)
        {
            return null;
        }

        using Stream stream = stylesPart.OpenRead();
        return SafeXml.Load(stream, cancellationToken).Root;
    }

    private static bool TryReadBuiltInLevel(string styleId, out int level)
    {
        level = 0;
        if (styleId.Length != 8 || !styleId.StartsWith("Heading", StringComparison.Ordinal))
        {
            return false;
        }

        char digit = styleId[7];
        if (digit < (char)49 || digit > (char)57)
        {
            return false;
        }

        level = digit - (char)48;
        return true;
    }
}
