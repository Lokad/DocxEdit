using System.Xml.Linq;
using DocxEdit.Ooxml;

namespace DocxEdit.Model;

internal static class DocxUnsupportedFeatureScanner
{
    private static readonly XName[] RevisionElements =
    [
        OoxmlNs.W + "ins",
        OoxmlNs.W + "del",
        OoxmlNs.W + "moveFrom",
        OoxmlNs.W + "moveTo",
        OoxmlNs.W + "moveFromRangeStart",
        OoxmlNs.W + "moveFromRangeEnd",
        OoxmlNs.W + "moveToRangeStart",
        OoxmlNs.W + "moveToRangeEnd",
        OoxmlNs.W + "rPrChange",
        OoxmlNs.W + "pPrChange",
        OoxmlNs.W + "tblPrChange",
        OoxmlNs.W + "trPrChange",
        OoxmlNs.W + "tcPrChange",
        OoxmlNs.W + "sectPrChange",
        OoxmlNs.W + "cellIns",
        OoxmlNs.W + "cellDel",
        OoxmlNs.W + "cellMerge",
        OoxmlNs.W + "customXmlInsRangeStart",
        OoxmlNs.W + "customXmlInsRangeEnd",
        OoxmlNs.W + "customXmlDelRangeStart",
        OoxmlNs.W + "customXmlDelRangeEnd"
    ];

    public static IReadOnlyList<DocxDiagnostic> Scan(
        OoxmlPackage package,
        bool includeHeadersFooters,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (package.MainDocumentPartName is null)
        {
            return [];
        }

        var diagnostics = new List<DocxDiagnostic>();
        ScanStory(package, package.MainDocumentPartName, "main", diagnostics, cancellationToken);
        if (includeHeadersFooters)
        {
            int headerIndex = 1;
            foreach (OoxmlRelationship relationship in package
                .GetRelationships(package.MainDocumentPartName, cancellationToken)
                .Where(relationship => !relationship.IsExternal && relationship.Type == OoxmlRelTypes.Header && relationship.ResolvedTarget is not null)
                .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
            {
                ScanStory(package, relationship.ResolvedTarget!, $"header[{headerIndex++}]", diagnostics, cancellationToken);
            }

            int footerIndex = 1;
            foreach (OoxmlRelationship relationship in package
                .GetRelationships(package.MainDocumentPartName, cancellationToken)
                .Where(relationship => !relationship.IsExternal && relationship.Type == OoxmlRelTypes.Footer && relationship.ResolvedTarget is not null)
                .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
            {
                ScanStory(package, relationship.ResolvedTarget!, $"footer[{footerIndex++}]", diagnostics, cancellationToken);
            }
        }

        return diagnostics;
    }

    private static void ScanStory(
        OoxmlPackage package,
        string partName,
        string story,
        List<DocxDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        OoxmlPart? part = package.GetPart(partName);
        if (part is null)
        {
            return;
        }

        using Stream stream = part.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        AddWarningIfAny(diagnostics, "W1001", "tracked-changes", CountAny(document, RevisionElements), partName, story, "Tracked-change markup is present; read output uses a final-view approximation. Use the changes command for markup metadata.");
        AddWarningIfAny(diagnostics, "W1002", "hyperlink", Count(document, OoxmlNs.W + "hyperlink"), partName, story, "Hyperlinks are preserved as text but are not modeled as link metadata.");
        AddWarningIfAny(diagnostics, "W1003", "field", CountAny(document, [OoxmlNs.W + "fldSimple", OoxmlNs.W + "fldChar", OoxmlNs.W + "instrText"]), partName, story, "Fields are preserved as text but are not modeled as field metadata.");
        AddWarningIfAny(diagnostics, "W1004", "comment", CountAny(document, [OoxmlNs.W + "commentRangeStart", OoxmlNs.W + "commentRangeEnd", OoxmlNs.W + "commentReference"]), partName, story, "Comment anchors are preserved but comments are not modeled.");
        AddWarningIfAny(diagnostics, "W1005", "bookmark", CountAny(document, [OoxmlNs.W + "bookmarkStart", OoxmlNs.W + "bookmarkEnd"]), partName, story, "Bookmarks are preserved but are not modeled.");
        AddWarningIfAny(diagnostics, "W1006", "content-control", Count(document, OoxmlNs.W + "sdt"), partName, story, "Content controls are preserved but are not modeled.");
        AddWarningIfAny(diagnostics, "W1007", "floating-image", Count(document, OoxmlNs.Wp + "anchor"), partName, story, "Floating images are preserved but are not listed as editable images.");

        int externalImages = package
            .GetRelationships(partName, cancellationToken)
            .Count(relationship => relationship.IsExternal && relationship.Type == OoxmlRelTypes.Image);
        AddWarningIfAny(diagnostics, "W1008", "external-image", externalImages, partName, story, "External images are not fetched and are not listed as editable images.");
    }

    private static int Count(XDocument document, XName name)
    {
        return document.Descendants(name).Count();
    }

    private static int CountAny(XDocument document, IReadOnlyCollection<XName> names)
    {
        return document.Descendants().Count(element => names.Contains(element.Name));
    }

    private static void AddWarningIfAny(
        List<DocxDiagnostic> diagnostics,
        string code,
        string feature,
        int count,
        string partName,
        string story,
        string message)
    {
        if (count == 0)
        {
            return;
        }

        diagnostics.Add(new DocxDiagnostic(
            DocxSeverity.Warning,
            code,
            $"{message} Count={count}.",
            PartName: partName,
            Story: story,
            Feature: feature));
    }
}
