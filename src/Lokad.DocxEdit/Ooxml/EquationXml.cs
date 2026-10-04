using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace Lokad.DocxEdit.Ooxml;

internal static class EquationXml
{
    internal static bool IsVisible(XElement math, DocxTextView view) => !math.Ancestors().Any(e =>
        (view == DocxTextView.Final && (e.Name == OoxmlNs.W + "del" || e.Name == OoxmlNs.W + "moveFrom")) ||
        (view == DocxTextView.Original && (e.Name == OoxmlNs.W + "ins" || e.Name == OoxmlNs.W + "moveTo")));

    internal static string Export(XElement math)
    {
        var copy = new XElement(math);
        copy.DescendantsAndSelf().Attributes().Where(a => a.Name.NamespaceName == "http://schemas.lokad.com/docxedit/snapshot").Remove();
        return copy.ToString(SaveOptions.DisableFormatting);
    }

    internal static string Hash(XElement math)
    {
        // Length-prefix expanded names and values so prefixes, attribute order,
        // indentation, and snapshot serialization cannot change the guard.
        var builder = new StringBuilder();
        foreach (XElement element in math.DescendantsAndSelf())
        {
            Add(element.Name.ToString());
            builder.Append('[');
            foreach (XAttribute attribute in element.Attributes().Where(a => !a.IsNamespaceDeclaration &&
                a.Name.NamespaceName != "http://schemas.lokad.com/docxedit/snapshot").OrderBy(a => a.Name.ToString(), StringComparer.Ordinal))
            {
                Add(attribute.Name.ToString()); Add(attribute.Value);
            }
            builder.Append(']');
            Add(element.Parent == null || ReferenceEquals(element, math) ? "0" : element.Ancestors().TakeWhile(e => !ReferenceEquals(e, math)).Count().ToString(System.Globalization.CultureInfo.InvariantCulture));
            foreach (XText text in element.Nodes().OfType<XText>())
                if (!string.IsNullOrWhiteSpace(text.Value) || element.Name.LocalName is "t" or "delText") Add(text.Value);
        }
        return Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(builder.ToString())));
        void Add(string value) { builder.Append(value.Length).Append(':').Append(value); }
    }
}
