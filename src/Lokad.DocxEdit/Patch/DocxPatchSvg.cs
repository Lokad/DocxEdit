using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

internal static partial class DocxPatchEngine
{
    private const string SvgExtensionUri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
    private static readonly XNamespace SvgNs = "http://www.w3.org/2000/svg";

    private static bool TryReadSvg(byte[] bytes, out XElement? svg, out string? reason, bool validateContent = true)
    {
        svg = null;
        reason = null;
        try
        {
            using var stream = new MemoryStream(bytes, writable: false);
            svg = SafeXml.Load(stream, CancellationToken.None).Root;
        }
        catch (XmlException)
        {
            reason = "malformed XML or prohibited DTD";
            return false;
        }
        if (svg?.Name != SvgNs + "svg")
        {
            reason = "expected an SVG root in the SVG namespace";
            return false;
        }
        if (!validateContent) return true;
        foreach (XElement element in svg.DescendantsAndSelf())
        {
            if (element.Name == SvgNs + "script" || element.Name == SvgNs + "foreignObject")
            {
                reason = "scripts and foreignObject are unsupported";
                return false;
            }
            foreach (XAttribute attribute in element.Attributes())
            {
                if (attribute.Name.LocalName == "href" &&
                    !attribute.Value.StartsWith('#') &&
                    !attribute.Value.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
                {
                    reason = "external href references are unsupported; embed resources or use fragments";
                    return false;
                }
            }
        }
        return true;
    }

    // Absolute root lengths establish the viewport ratio. Otherwise viewBox
    // supplies the ratio and default pixel dimensions, without running CSS.
    private static bool TryReadSvgSize(byte[] bytes, out double width, out double height)
    {
        width = height = 0;
        if (!TryReadSvg(bytes, out XElement? svg, out _) || svg is null) return false;
        bool hasWidth = TryReadSvgLength((string?)svg.Attribute("width"), out width);
        bool hasHeight = TryReadSvgLength((string?)svg.Attribute("height"), out height);
        if (hasWidth && hasHeight) return true;
        string[]? viewBox = ((string?)svg.Attribute("viewBox"))?.Split(
            [' ', '\t', '\r', '\n', ','], StringSplitOptions.RemoveEmptyEntries);
        if (viewBox is not { Length: 4 } ||
            !viewBox.All(value => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) && double.IsFinite(number)) ||
            !double.TryParse(viewBox[2], NumberStyles.Float, CultureInfo.InvariantCulture, out double boxWidth) ||
            !double.TryParse(viewBox[3], NumberStyles.Float, CultureInfo.InvariantCulture, out double boxHeight) ||
            boxWidth <= 0 || boxHeight <= 0) return false;
        if (hasWidth) height = width * boxHeight / boxWidth;
        else if (hasHeight) width = height * boxWidth / boxHeight;
        else { width = boxWidth; height = boxHeight; }
        return double.IsFinite(width) && double.IsFinite(height) && width > 0 && height > 0;
    }

    private static bool TryReadSvgLength(string? text, out double pixels)
    {
        pixels = 0;
        if (text is null) return false;
        text = text.Trim();
        double factor = 1;
        foreach (var unit in new (string Suffix, double Factor)[]
        {
            ("px", 1), ("in", 96), ("cm", 96 / 2.54), ("mm", 96 / 25.4),
            ("pt", 96 / 72.0), ("pc", 16), ("q", 96 / 101.6)
        })
        {
            if (!text.EndsWith(unit.Suffix, StringComparison.Ordinal)) continue;
            text = text[..^unit.Suffix.Length];
            factor = unit.Factor;
            break;
        }
        if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double length)) return false;
        pixels = length * factor;
        return double.IsFinite(pixels) && pixels > 0;
    }

    private static void SetImageAssetReference(XElement blip, string relationshipId, string contentType)
    {
        blip.SetAttributeValue(OoxmlNs.R + "embed", relationshipId);
        blip.SetAttributeValue(OoxmlNs.R + "link", null);
        // A previous SVG must never continue to override a raster replacement.
        foreach (XElement svgBlip in blip.Descendants(OoxmlNs.Asvg + "svgBlip").ToArray())
        {
            XElement? extension = svgBlip.Parent;
            svgBlip.Remove();
            if (extension?.Name == OoxmlNs.A + "ext" && !extension.HasElements) extension.Remove();
        }
        XElement? extensions = blip.Element(OoxmlNs.A + "extLst");
        if (contentType == "image/svg+xml")
        {
            if (extensions is null) { extensions = new XElement(OoxmlNs.A + "extLst"); blip.Add(extensions); }
            extensions.Add(new XElement(OoxmlNs.A + "ext", new XAttribute("uri", SvgExtensionUri),
                new XElement(OoxmlNs.Asvg + "svgBlip", new XAttribute(XNamespace.Xmlns + "asvg", OoxmlNs.Asvg),
                    new XAttribute(OoxmlNs.R + "embed", relationshipId))));
        }
        else if (extensions is not null && !extensions.HasElements) extensions.Remove();
    }
}
