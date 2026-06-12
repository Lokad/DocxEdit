using System.Xml.Linq;
using DocxEdit.Ooxml;

namespace DocxEdit.Model;

internal static class DocxPackageValidator
{
    public static IReadOnlyList<DocxDiagnostic> Validate(
        OoxmlPackage package,
        CancellationToken cancellationToken = default)
    {
        var diagnostics = new List<DocxDiagnostic>();
        foreach (OoxmlPart part in package.Parts.Values
            .Where(part => part.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(part => part.Name, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            ValidateRoot(part.Name, document.Root, diagnostics);
            if (part.Name.StartsWith("/word/", StringComparison.OrdinalIgnoreCase))
            {
                ValidateWordPart(package, part.Name, document, diagnostics, cancellationToken);
            }
        }

        return diagnostics;
    }

    private static void ValidateRoot(string partName, XElement? root, List<DocxDiagnostic> diagnostics)
    {
        if (root is null)
        {
            diagnostics.Add(Error("E9101", "XML part has no root element.", partName));
            return;
        }

        XName? expected = partName switch
        {
            "/word/document.xml" => OoxmlNs.W + "document",
            "/word/styles.xml" => OoxmlNs.W + "styles",
            "/word/numbering.xml" => OoxmlNs.W + "numbering",
            "/word/settings.xml" => OoxmlNs.W + "settings",
            "/word/comments.xml" => OoxmlNs.W + "comments",
            "/word/commentsExtended.xml" => OoxmlNs.W15 + "commentsEx",
            _ when partName.StartsWith("/word/header", StringComparison.OrdinalIgnoreCase) => OoxmlNs.W + "hdr",
            _ when partName.StartsWith("/word/footer", StringComparison.OrdinalIgnoreCase) => OoxmlNs.W + "ftr",
            _ => null
        };
        if (expected is not null && root.Name != expected)
        {
            diagnostics.Add(Error("E9102", $"Expected root '{expected.LocalName}', found '{root.Name.LocalName}'.", partName));
        }
    }

    private static void ValidateWordPart(
        OoxmlPackage package,
        string partName,
        XDocument document,
        List<DocxDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        ValidatePairedIds(document, OoxmlNs.W + "bookmarkStart", OoxmlNs.W + "bookmarkEnd", "bookmark", partName, diagnostics);
        ValidatePairedIds(document, OoxmlNs.W + "commentRangeStart", OoxmlNs.W + "commentRangeEnd", "comment range", partName, diagnostics);
        ValidateFieldBalance(document, partName, diagnostics);
        ValidateDrawingRelationships(package, partName, document, diagnostics, cancellationToken);
        ValidateDrawingProperties(document, partName, diagnostics);
        if (string.Equals(partName, "/word/commentsExtended.xml", StringComparison.OrdinalIgnoreCase))
        {
            ValidateCommentsExtended(package, partName, document, diagnostics, cancellationToken);
        }

        ValidateTables(document, partName, diagnostics);
    }

    private static void ValidatePairedIds(
        XDocument document,
        XName startName,
        XName endName,
        string label,
        string partName,
        List<DocxDiagnostic> diagnostics)
    {
        var starts = document.Descendants(startName)
            .Select(element => (string?)element.Attribute(OoxmlNs.W + "id"))
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .GroupBy(id => id!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        var ends = document.Descendants(endName)
            .Select(element => (string?)element.Attribute(OoxmlNs.W + "id"))
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .GroupBy(id => id!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);

        foreach (string id in starts.Keys.Concat(ends.Keys).Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal))
        {
            starts.TryGetValue(id, out int startCount);
            ends.TryGetValue(id, out int endCount);
            if (startCount != endCount)
            {
                diagnostics.Add(Error("E9103", $"Unbalanced {label} id '{id}': starts={startCount}, ends={endCount}.", partName));
            }
        }
    }

    private static void ValidateFieldBalance(XDocument document, string partName, List<DocxDiagnostic> diagnostics)
    {
        int depth = 0;
        foreach (XElement fieldChar in document.Descendants(OoxmlNs.W + "fldChar"))
        {
            string? type = (string?)fieldChar.Attribute(OoxmlNs.W + "fldCharType");
            if (string.Equals(type, "begin", StringComparison.Ordinal))
            {
                depth++;
            }
            else if (string.Equals(type, "end", StringComparison.Ordinal))
            {
                if (depth == 0)
                {
                    diagnostics.Add(Error("E9104", "Complex field end appears without a matching begin.", partName));
                }
                else
                {
                    depth--;
                }
            }
        }

        if (depth > 0)
        {
            diagnostics.Add(Error("E9104", $"Complex field has {depth} unclosed begin marker(s).", partName));
        }
    }

    private static void ValidateDrawingRelationships(
        OoxmlPackage package,
        string partName,
        XDocument document,
        List<DocxDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        var relationships = package.GetRelationships(partName, cancellationToken)
            .ToDictionary(relationship => relationship.Id, StringComparer.Ordinal);
        foreach (XElement blip in document.Descendants(OoxmlNs.A + "blip"))
        {
            string? relationshipId = (string?)blip.Attribute(OoxmlNs.R + "embed") ?? (string?)blip.Attribute(OoxmlNs.R + "link");
            if (!string.IsNullOrWhiteSpace(relationshipId) && !relationships.ContainsKey(relationshipId))
            {
                diagnostics.Add(Error("E9105", $"Drawing references missing relationship '{relationshipId}'.", partName));
            }
        }
    }

    private static void ValidateDrawingProperties(XDocument document, string partName, List<DocxDiagnostic> diagnostics)
    {
        var docPrIds = document
            .Descendants(OoxmlNs.Wp + "docPr")
            .Select(element => (string?)element.Attribute("id"))
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .Select(id => id!)
            .GroupBy(id => id, StringComparer.Ordinal)
            .Where(group => group.Count() > 1)
            .OrderBy(group => group.Key, StringComparer.Ordinal);
        foreach (IGrouping<string, string> group in docPrIds)
        {
            diagnostics.Add(Error("E9107", $"Duplicate drawing docPr id '{group.Key}' appears {group.Count()} times.", partName));
        }
    }

    private static void ValidateCommentsExtended(
        OoxmlPackage package,
        string partName,
        XDocument document,
        List<DocxDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        HashSet<string> commentParaIds = ReadCommentParaIds(package, cancellationToken);
        var extensionParaIds = document
            .Descendants(OoxmlNs.W15 + "commentEx")
            .Select(element => (ParaId: (string?)element.Attribute(OoxmlNs.W15 + "paraId"), Element: element))
            .ToArray();

        foreach ((string? paraId, XElement _) in extensionParaIds.Where(item => string.IsNullOrWhiteSpace(item.ParaId)))
        {
            diagnostics.Add(Error("E9108", "commentsExtended commentEx is missing w15:paraId.", partName));
        }

        foreach (IGrouping<string, string> group in extensionParaIds
            .Select(item => item.ParaId)
            .Where(paraId => !string.IsNullOrWhiteSpace(paraId))
            .Select(paraId => paraId!)
            .GroupBy(paraId => paraId, StringComparer.Ordinal)
            .Where(group => group.Count() > 1)
            .OrderBy(group => group.Key, StringComparer.Ordinal))
        {
            diagnostics.Add(Error("E9108", $"Duplicate commentsExtended paraId '{group.Key}' appears {group.Count()} times.", partName));
        }

        foreach (string paraId in extensionParaIds
            .Select(item => item.ParaId)
            .Where(paraId => !string.IsNullOrWhiteSpace(paraId))
            .Select(paraId => paraId!)
            .Distinct(StringComparer.Ordinal)
            .Order(StringComparer.Ordinal))
        {
            if (!commentParaIds.Contains(paraId))
            {
                diagnostics.Add(Error("E9108", $"commentsExtended paraId '{paraId}' has no matching comment paragraph.", partName));
            }
        }
    }

    private static HashSet<string> ReadCommentParaIds(OoxmlPackage package, CancellationToken cancellationToken)
    {
        var paraIds = new HashSet<string>(StringComparer.Ordinal);
        OoxmlPart? commentsPart = package.GetPart("/word/comments.xml");
        if (commentsPart is null)
        {
            return paraIds;
        }

        using Stream stream = commentsPart.OpenRead();
        XDocument comments = SafeXml.Load(stream, cancellationToken);
        foreach (string paraId in comments
            .Descendants(OoxmlNs.W + "comment")
            .Descendants(OoxmlNs.W + "p")
            .Select(paragraph => (string?)paragraph.Attribute(OoxmlNs.W15 + "paraId"))
            .Where(paraId => !string.IsNullOrWhiteSpace(paraId))
            .Select(paraId => paraId!))
        {
            paraIds.Add(paraId);
        }

        return paraIds;
    }

    private static void ValidateTables(XDocument document, string partName, List<DocxDiagnostic> diagnostics)
    {
        foreach (XElement table in document.Descendants(OoxmlNs.W + "tbl"))
        {
            if (!table.Elements(OoxmlNs.W + "tr").Any())
            {
                diagnostics.Add(Error("E9106", "Table has no rows.", partName));
            }

            foreach (XElement row in table.Elements(OoxmlNs.W + "tr"))
            {
                if (!row.Elements(OoxmlNs.W + "tc").Any())
                {
                    diagnostics.Add(Error("E9106", "Table row has no cells.", partName));
                }
            }
        }
    }

    private static DocxDiagnostic Error(string code, string message, string partName)
    {
        return new DocxDiagnostic(DocxSeverity.Error, code, message, PartName: partName);
    }
}
