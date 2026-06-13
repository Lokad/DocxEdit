using System.Xml.Linq;

namespace Lokad.DocxEdit.Ooxml;

internal sealed class OoxmlContentTypes
{
    private static readonly XNamespace ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";

    private readonly Dictionary<string, string> defaults;
    private readonly Dictionary<string, string> overrides;

    private OoxmlContentTypes(Dictionary<string, string> defaults, Dictionary<string, string> overrides)
    {
        this.defaults = defaults;
        this.overrides = overrides;
    }

    public static OoxmlContentTypes Parse(Stream stream, CancellationToken cancellationToken = default)
    {
        XDocument document = SafeXml.Load(stream, cancellationToken);
        var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var overrides = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (XElement element in document.Root?.Elements() ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (element.Name == ContentTypesNamespace + "Default")
            {
                string extension = RequiredAttribute(element, "Extension").TrimStart('.');
                defaults[extension] = RequiredAttribute(element, "ContentType");
            }
            else if (element.Name == ContentTypesNamespace + "Override")
            {
                string partName = OoxmlPath.NormalizePartName(RequiredAttribute(element, "PartName"));
                overrides[partName] = RequiredAttribute(element, "ContentType");
            }
        }

        return new OoxmlContentTypes(defaults, overrides);
    }

    public string? GetContentType(string partName)
    {
        string normalized = OoxmlPath.NormalizePartName(partName);
        if (overrides.TryGetValue(normalized, out string? contentType))
        {
            return contentType;
        }

        string extension = Path.GetExtension(normalized).TrimStart('.');
        return defaults.TryGetValue(extension, out contentType) ? contentType : null;
    }

    private static string RequiredAttribute(XElement element, string name)
    {
        return (string?)element.Attribute(name)
            ?? throw new InvalidDataException($"Missing required {element.Name.LocalName} attribute '{name}'.");
    }
}

