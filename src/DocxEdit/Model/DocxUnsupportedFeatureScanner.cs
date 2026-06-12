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
        AddWarningIfAny(diagnostics, "W1001", "tracked-changes", "selected-text-view", CountAny(document, RevisionElements), partName, story, "Tracked-change markup is present; read text can use final, original, or lightweight markup views. Use the changes command for markup metadata.");
        AddWarningIfAny(diagnostics, "W1002", "hyperlink", "modeled-metadata", Count(document, OoxmlNs.W + "hyperlink"), partName, story, "Hyperlinks are surfaced as metadata and preserved; hyperlink editing is limited.");
        AddWarningIfAny(diagnostics, "W1003", "field", "modeled-metadata", CountAny(document, [OoxmlNs.W + "fldSimple", OoxmlNs.W + "fldChar", OoxmlNs.W + "instrText"]), partName, story, "Fields are surfaced as metadata and preserved; DocxEdit does not evaluate field results.");
        AddWarningIfAny(diagnostics, "W1004", "comment", "changes-metadata", CountAny(document, [OoxmlNs.W + "commentRangeStart", OoxmlNs.W + "commentRangeEnd", OoxmlNs.W + "commentReference"]), partName, story, "Comment anchors are surfaced by changes metadata; comment text is not exposed by the read model.");
        AddWarningIfAny(diagnostics, "W1005", "bookmark", "modeled-metadata", CountAny(document, [OoxmlNs.W + "bookmarkStart", OoxmlNs.W + "bookmarkEnd"]), partName, story, "Bookmarks are surfaced as metadata and preserved; bookmark range editing is limited.");
        AddWarningIfAny(diagnostics, "W1006", "content-control", "modeled-metadata", Count(document, OoxmlNs.W + "sdt"), partName, story, "Content controls are surfaced as metadata and preserved; content-control editing is limited.");
        AddWarningIfAny(diagnostics, "W1007", "floating-image", "modeled-metadata", Count(document, OoxmlNs.Wp + "anchor"), partName, story, "Floating images are surfaced as image metadata and preserved; full layout editing is limited.");
        AddWarningIfAny(diagnostics, "W1009", "chart", "preserve-only", CountWhere(document, element => element.Name.LocalName == "chart" && element.Name.NamespaceName.Contains("/chart", StringComparison.OrdinalIgnoreCase)), partName, story, "Charts are preserved but are not modeled.");
        AddWarningIfAny(diagnostics, "W1010", "smart-art", "preserve-only", CountWhere(document, element => element.Name.NamespaceName.Contains("/diagram", StringComparison.OrdinalIgnoreCase)), partName, story, "SmartArt and diagram content are preserved but are not modeled.");
        AddWarningIfAny(diagnostics, "W1011", "equation", "preserve-only", CountWhere(document, element => element.Name.NamespaceName == "http://schemas.openxmlformats.org/officeDocument/2006/math"), partName, story, "Equations are preserved but are not modeled.");
        AddWarningIfAny(diagnostics, "W1012", "shape", "preserve-only", CountWhere(document, element =>
            element.Name == OoxmlNs.W + "pict" ||
            element.Name.LocalName == "shape" ||
            element.Name.NamespaceName.Contains("vml", StringComparison.OrdinalIgnoreCase) ||
            element.Name.NamespaceName.Contains("wordprocessingShape", StringComparison.OrdinalIgnoreCase)), partName, story, "Shapes are preserved but are not modeled.");
        AddWarningIfAny(diagnostics, "W1013", "alt-chunk", "preserve-only", Count(document, OoxmlNs.W + "altChunk"), partName, story, "altChunk content is preserved but is not imported or modeled.");
        AddWarningIfAny(diagnostics, "W1014", "section-flow", "basic-section-model", CountComplexSectionFlows(document), partName, story, "Complex section flow is present; section read/edit support does not model full section inheritance.");

        int externalImages = package
            .GetRelationships(partName, cancellationToken)
            .Count(relationship => relationship.IsExternal && relationship.Type == OoxmlRelTypes.Image);
        AddWarningIfAny(diagnostics, "W1008", "external-image", "omit-from-editable-images", externalImages, partName, story, "External images are not fetched and are not listed as editable images.");

        IReadOnlyList<OoxmlRelationship> relationships = package.GetRelationships(partName, cancellationToken);
        int brokenHyperlinks = CountBrokenHyperlinks(document, relationships);
        AddWarningIfAny(diagnostics, "W1015", "hyperlink", "broken-relationship", brokenHyperlinks, partName, story, "Hyperlink relationship IDs are missing from the part relationship table.");
        AddWarningIfAny(diagnostics, "W1016", "hyperlink", "invalid-uri", CountInvalidHyperlinkUris(document, relationships), partName, story, "Hyperlink targets include unsupported or malformed external URIs.");
        AddWarningIfAny(diagnostics, "W1017", "hyperlink", "missing-anchor", CountMissingHyperlinkAnchors(document), partName, story, "Internal hyperlink anchors do not match any bookmark in the same story.");
        AddWarningIfAny(diagnostics, "W1018", "hyperlink", "duplicate-anchor", CountDuplicateHyperlinkAnchors(document), partName, story, "Internal hyperlink anchors match multiple bookmarks in the same story.");
    }

    private static int Count(XDocument document, XName name)
    {
        return document.Descendants(name).Count();
    }

    private static int CountAny(XDocument document, IReadOnlyCollection<XName> names)
    {
        return document.Descendants().Count(element => names.Contains(element.Name));
    }

    private static int CountWhere(XDocument document, Func<XElement, bool> predicate)
    {
        return document.Descendants().Count(predicate);
    }

    private static int CountComplexSectionFlows(XDocument document)
    {
        XElement? root = document.Root;
        if (root is null)
        {
            return 0;
        }

        int sectionProperties = root.Descendants(OoxmlNs.W + "sectPr").Count();
        int paragraphSectionBreaks = root
            .Descendants(OoxmlNs.W + "pPr")
            .Elements(OoxmlNs.W + "sectPr")
            .Count();
        return sectionProperties > 1
            ? sectionProperties
            : paragraphSectionBreaks;
    }

    private static int CountBrokenHyperlinks(XDocument document, IReadOnlyList<OoxmlRelationship> relationships)
    {
        var relationshipIds = relationships.Select(relationship => relationship.Id).ToHashSet(StringComparer.Ordinal);
        return document
            .Descendants(OoxmlNs.W + "hyperlink")
            .Select(hyperlink => (string?)hyperlink.Attribute(OoxmlNs.R + "id"))
            .Count(id => !string.IsNullOrWhiteSpace(id) && !relationshipIds.Contains(id));
    }

    private static int CountInvalidHyperlinkUris(XDocument document, IReadOnlyList<OoxmlRelationship> relationships)
    {
        IReadOnlyDictionary<string, OoxmlRelationship> relationshipsById = relationships.ToDictionary(relationship => relationship.Id, StringComparer.Ordinal);
        return document
            .Descendants(OoxmlNs.W + "hyperlink")
            .Select(hyperlink => (string?)hyperlink.Attribute(OoxmlNs.R + "id"))
            .Where(id => !string.IsNullOrWhiteSpace(id) &&
                relationshipsById.TryGetValue(id!, out OoxmlRelationship? relationship) &&
                relationship.IsExternal &&
                relationship.Type == OoxmlRelTypes.Hyperlink)
            .Count(id => !IsSupportedHyperlinkUri(relationshipsById[id!].Target));
    }

    private static bool IsSupportedHyperlinkUri(string uri)
    {
        return Uri.TryCreate(uri, UriKind.Absolute, out Uri? parsed) &&
            parsed.Scheme is "http" or "https" or "mailto";
    }

    private static int CountMissingHyperlinkAnchors(XDocument document)
    {
        IReadOnlyDictionary<string, int> bookmarkCounts = CountBookmarkNames(document);
        return document
            .Descendants(OoxmlNs.W + "hyperlink")
            .Select(hyperlink => (string?)hyperlink.Attribute(OoxmlNs.W + "anchor"))
            .Where(anchor => !string.IsNullOrWhiteSpace(anchor))
            .Count(anchor => !bookmarkCounts.ContainsKey(anchor!));
    }

    private static int CountDuplicateHyperlinkAnchors(XDocument document)
    {
        IReadOnlyDictionary<string, int> bookmarkCounts = CountBookmarkNames(document);
        return document
            .Descendants(OoxmlNs.W + "hyperlink")
            .Select(hyperlink => (string?)hyperlink.Attribute(OoxmlNs.W + "anchor"))
            .Where(anchor => !string.IsNullOrWhiteSpace(anchor))
            .Count(anchor => bookmarkCounts.TryGetValue(anchor!, out int count) && count > 1);
    }

    private static IReadOnlyDictionary<string, int> CountBookmarkNames(XDocument document)
    {
        return document
            .Descendants(OoxmlNs.W + "bookmarkStart")
            .Select(bookmark => (string?)bookmark.Attribute(OoxmlNs.W + "name"))
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .GroupBy(name => name!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
    }

    private static void AddWarningIfAny(
        List<DocxDiagnostic> diagnostics,
        string code,
        string feature,
        string fallback,
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
            Feature: feature,
            Fallback: fallback));
    }
}
