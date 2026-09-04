using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

internal static partial class DocxPackageValidator
{
    private static void ValidateFieldBalance(XDocument document, string partName, List<DocxDiagnostic> diagnostics)
    {
        var stack = new Stack<ComplexFieldValidationState>();
        foreach (XElement element in document.Descendants())
        {
            if (element.Name == OoxmlNs.W + "instrText" && stack.Count == 0)
            {
                diagnostics.Add(Error("E9112", "Field instruction text appears outside a complex field.", partName));
                continue;
            }

            if (element.Name == OoxmlNs.W + "t" &&
                stack.Count > 0 &&
                !stack.Peek().HasSeparate &&
                !string.IsNullOrWhiteSpace(element.Value))
            {
                diagnostics.Add(Error("E9112", "Field result text appears before the complex field separate marker.", partName));
                continue;
            }

            if (element.Name != OoxmlNs.W + "fldChar")
            {
                continue;
            }

            string? type = (string?)element.Attribute(OoxmlNs.W + "fldCharType");
            if (string.Equals(type, "begin", StringComparison.Ordinal))
            {
                stack.Push(new ComplexFieldValidationState());
            }
            else if (string.Equals(type, "separate", StringComparison.Ordinal))
            {
                if (stack.Count == 0)
                {
                    diagnostics.Add(Error("E9104", "Complex field separate appears without a matching begin.", partName));
                }
                else if (stack.Peek().HasSeparate)
                {
                    diagnostics.Add(Error("E9104", "Complex field has duplicate separate markers.", partName));
                }
                else
                {
                    stack.Peek().HasSeparate = true;
                }
            }
            else if (string.Equals(type, "end", StringComparison.Ordinal))
            {
                if (stack.Count == 0)
                {
                    diagnostics.Add(Error("E9104", "Complex field end appears without a matching begin.", partName));
                }
                else
                {
                    stack.Pop();
                }
            }
            else
            {
                diagnostics.Add(Error("E9104", $"Complex field has invalid fldCharType '{type ?? string.Empty}'.", partName));
            }
        }

        if (stack.Count > 0)
        {
            diagnostics.Add(Error("E9104", $"Complex field has {stack.Count} unclosed begin marker(s).", partName));
        }
    }

    private static void ValidateFieldFlags(XDocument document, string partName, List<DocxDiagnostic> diagnostics)
    {
        foreach (XElement element in document.Descendants().Where(element =>
            element.Name == OoxmlNs.W + "fldSimple" ||
            element.Name == OoxmlNs.W + "fldChar"))
        {
            ValidateFieldOnOffAttribute(element, "dirty", partName, diagnostics);
            ValidateFieldOnOffAttribute(element, "fldLock", partName, diagnostics);
        }
    }

    private static void ValidateFieldOnOffAttribute(
        XElement element,
        string localName,
        string partName,
        List<DocxDiagnostic> diagnostics)
    {
        string? value = (string?)element.Attribute(OoxmlNs.W + localName);
        if (value is null)
        {
            return;
        }

        if (value is "true" or "false" or "1" or "0" or "on" or "off")
        {
            return;
        }

        diagnostics.Add(Error("E9112", $"Field attribute w:{localName} has invalid OnOff value '{value}'.", partName));
    }

    private static void ValidateContentControls(XDocument document, string partName, List<DocxDiagnostic> diagnostics)
    {
        var controlIds = new List<string>();
        foreach (XElement properties in document.Descendants(OoxmlNs.W + "sdtPr"))
        {
            XElement? idElement = properties.Element(OoxmlNs.W + "id");
            string? id = (string?)idElement?.Attribute(OoxmlNs.W + "val");
            if (id is not null)
            {
                if (!int.TryParse(id, out _))
                {
                    diagnostics.Add(Error("E9115", $"Content control w:id has invalid integer value '{id}'.", partName));
                }

                controlIds.Add(id);
            }

            string? lockValue = (string?)properties.Element(OoxmlNs.W + "lock")?.Attribute(OoxmlNs.W + "val");
            if (lockValue is not null && lockValue is not ("unlocked" or "sdtLocked" or "contentLocked" or "sdtContentLocked"))
            {
                diagnostics.Add(Error("E9115", $"Content control w:lock has invalid value '{lockValue}'.", partName));
            }

            XElement? checkedElement = properties
                .Element(OoxmlNs.W + "checkBox")
                ?.Element(OoxmlNs.W + "checked");
            if (checkedElement is not null)
            {
                ValidateContentControlOnOffAttribute(checkedElement, "checked", partName, diagnostics);
            }
        }

        foreach (IGrouping<string, string> group in controlIds
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .GroupBy(id => id, StringComparer.Ordinal)
            .Where(group => group.Count() > 1)
            .OrderBy(group => group.Key, StringComparer.Ordinal))
        {
            diagnostics.Add(Error("E9115", $"Duplicate content control w:id '{group.Key}' appears {group.Count()} times.", partName));
        }
    }

    private static void ValidateContentControlOnOffAttribute(
        XElement element,
        string localName,
        string partName,
        List<DocxDiagnostic> diagnostics)
    {
        string? value = (string?)element.Attribute(OoxmlNs.W + "val");
        if (value is null || value is "0" or "1" or "true" or "false" or "on" or "off")
        {
            return;
        }

        diagnostics.Add(Error("E9115", $"Content control w:{localName} has invalid OnOff value '{value}'.", partName));
    }
}
