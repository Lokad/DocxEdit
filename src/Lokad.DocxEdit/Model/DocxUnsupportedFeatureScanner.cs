using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

internal static class DocxUnsupportedFeatureScanner
{
    private static readonly HashSet<string> SupportedNumberingFormats = new(StringComparer.Ordinal)
    {
        "decimal",
        "decimalZero",
        "upperLetter",
        "lowerLetter",
        "upperRoman",
        "lowerRoman",
        "bullet"
    };

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
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var diagnostics = new List<DocxDiagnostic>();
        ScanStory(package, package.MainDocumentPartName, "main", diagnostics, cancellationToken);
        if (includeHeadersFooters)
        {
            int headerIndex = 1;
            foreach (ResolvedOoxmlRelationship relationship in package
                .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
                .Where(relationship => relationship.Type == OoxmlRelTypes.Header)
                .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
            {
                ScanStory(package, relationship.ResolvedTarget, $"header[{headerIndex++}]", diagnostics, cancellationToken);
            }

            int footerIndex = 1;
            foreach (ResolvedOoxmlRelationship relationship in package
                .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
                .Where(relationship => relationship.Type == OoxmlRelTypes.Footer)
                .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
            {
                ScanStory(package, relationship.ResolvedTarget, $"footer[{footerIndex++}]", diagnostics, cancellationToken);
            }
        }

        ScanNumbering(package, diagnostics, cancellationToken);
        return diagnostics;
    }

    private static void ScanNumbering(
        OoxmlPackage package,
        List<DocxDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        OoxmlRelationship? relationship = package
            .GetRelationships(package.MainDocumentPartName, cancellationToken)
            .FirstOrDefault(relationship =>
                !relationship.IsExternal &&
                relationship.Type == OoxmlRelTypes.Numbering &&
                relationship.ResolvedTarget is not null);
        if (relationship?.ResolvedTarget is null)
        {
            return;
        }

        OoxmlPart? part = package.GetPart(relationship.ResolvedTarget);
        if (part is null)
        {
            return;
        }

        using Stream stream = part.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        AddWarningIfAny(diagnostics, "W1024", "numbering", "unsupported-picture-bullet", CountAny(document, [OoxmlNs.W + "numPicBullet", OoxmlNs.W + "lvlPicBulletId"]), relationship.ResolvedTarget, "numbering", "Picture bullets are preserved but cannot be expanded into deterministic labels.");
        AddWarningIfAny(diagnostics, "W1025", "numbering", "unsupported-numbering-format", CountUnsupportedNumberingFormats(document), relationship.ResolvedTarget, "numbering", "Numbering definitions contain numFmt values that cannot be expanded into deterministic labels.");
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
        AddWarningIfAny(diagnostics, "W1026", "numbering", "tracked-numbering-property-revision", CountNumberingPropertyRevisions(document), partName, story, "Tracked paragraph property revisions include previous numbering state; final-view labels use current numbering, but original-view label reconstruction for previous numbering is not modeled.");
        AddWarningIfAny(diagnostics, "W1009", "chart", "preserve-only", CountWhere(document, element => element.Name.LocalName == "chart" && element.Name.NamespaceName.Contains("/chart", StringComparison.OrdinalIgnoreCase)), partName, story, "Charts are preserved but are not modeled.");
        AddWarningIfAny(diagnostics, "W1010", "smart-art", "preserve-only", CountWhere(document, element => element.Name.NamespaceName.Contains("/diagram", StringComparison.OrdinalIgnoreCase)), partName, story, "SmartArt and diagram content are preserved but are not modeled.");
        AddWarningIfAny(diagnostics, "W1011", "equation", "whole-equation", Count(document, OoxmlNs.M + "oMath"), partName, story, "Equations are modeled as whole native objects; editing their internal runs and tracked equation revisions is not supported.");
        AddWarningIfAny(diagnostics, "W1019", "linked-image", "omit-from-editable-images", CountLinkedImages(document), partName, story, "Linked images are not fetched and are not listed as editable images.");
        AddWarningIfAny(diagnostics, "W1020", "vml", "preserve-only", CountWhere(document, IsVmlDrawing), partName, story, "VML drawings are preserved but are not modeled.");
        AddWarningIfAny(diagnostics, "W1021", "grouped-drawing", "preserve-only", CountWhere(document, IsGroupedDrawing), partName, story, "Grouped drawings are preserved but are not modeled.");
        AddWarningIfAny(diagnostics, "W1022", "ole-object", "preserve-only", CountWhere(document, IsOleObject), partName, story, "OLE objects are preserved but are not modeled or executed.");
        AddWarningIfAny(diagnostics, "W1012", "shape", "preserve-only", CountWhere(document, element =>
            !IsVmlDrawing(element) &&
            !IsGroupedDrawing(element) &&
            !IsOleObject(element) &&
            (element.Name.LocalName == "shape" ||
                element.Name.NamespaceName.Contains("wordprocessingShape", StringComparison.OrdinalIgnoreCase))), partName, story, "Shapes are preserved but are not modeled.");
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
        AddWarningIfAny(diagnostics, "W1023", "hyperlink", "unsupported-internal-part-link", CountUnsupportedInternalPartHyperlinks(document, relationships), partName, story, "Internal part hyperlinks are preserved and modeled as target parts, but hyperlink patch edits support external URIs or anchors only.");
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

    private static int CountLinkedImages(XDocument document)
    {
        return document
            .Descendants(OoxmlNs.A + "blip")
            .Count(blip => !string.IsNullOrWhiteSpace((string?)blip.Attribute(OoxmlNs.R + "link")));
    }

    private static bool IsVmlDrawing(XElement element)
    {
        return element.Name == OoxmlNs.W + "pict" ||
            element.Name.NamespaceName.Contains(":vml", StringComparison.OrdinalIgnoreCase) ||
            element.Name.NamespaceName.Contains("/vml", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(element.Name.NamespaceName, "urn:schemas-microsoft-com:vml", StringComparison.Ordinal);
    }

    private static bool IsGroupedDrawing(XElement element)
    {
        return string.Equals(element.Name.LocalName, "grpSp", StringComparison.Ordinal) ||
            string.Equals(element.Name.LocalName, "wgp", StringComparison.Ordinal) ||
            element.Name.NamespaceName.Contains("wordprocessingGroup", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsOleObject(XElement element)
    {
        return element.Name == OoxmlNs.W + "object" ||
            string.Equals(element.Name.LocalName, "OLEObject", StringComparison.Ordinal);
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

    private static int CountUnsupportedNumberingFormats(XDocument document)
    {
        return document
            .Descendants(OoxmlNs.W + "numFmt")
            .Select(element => (string?)element.Attribute(OoxmlNs.W + "val"))
            .Count(format => !string.IsNullOrWhiteSpace(format) && !SupportedNumberingFormats.Contains(format));
    }

    private static int CountNumberingPropertyRevisions(XDocument document)
    {
        return document
            .Descendants(OoxmlNs.W + "pPrChange")
            .Count(change => change
                .Element(OoxmlNs.W + "pPr")
                ?.Element(OoxmlNs.W + "numPr") is not null);
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
                relationshipsById.TryGetValue(id, out OoxmlRelationship? relationship) &&
                relationship.IsExternal &&
                relationship.Type == OoxmlRelTypes.Hyperlink)
            .OfType<string>()
            .Count(id => !IsSupportedHyperlinkUri(relationshipsById[id].Target));
    }

    private static bool IsSupportedHyperlinkUri(string uri)
    {
        return Uri.TryCreate(uri, UriKind.Absolute, out Uri? parsed) &&
            parsed.Scheme is "http" or "https" or "mailto";
    }

    private static int CountUnsupportedInternalPartHyperlinks(XDocument document, IReadOnlyList<OoxmlRelationship> relationships)
    {
        IReadOnlyDictionary<string, OoxmlRelationship> relationshipsById = relationships.ToDictionary(relationship => relationship.Id, StringComparer.Ordinal);
        return document
            .Descendants(OoxmlNs.W + "hyperlink")
            .Select(hyperlink => (string?)hyperlink.Attribute(OoxmlNs.R + "id"))
            .Count(id => !string.IsNullOrWhiteSpace(id) &&
                relationshipsById.TryGetValue(id, out OoxmlRelationship? relationship) &&
                !relationship.IsExternal &&
                relationship.Type == OoxmlRelTypes.Hyperlink);
    }

    private static int CountMissingHyperlinkAnchors(XDocument document)
    {
        IReadOnlyDictionary<string, int> bookmarkCounts = CountBookmarkNames(document);
        return document
            .Descendants(OoxmlNs.W + "hyperlink")
            .Select(hyperlink => (string?)hyperlink.Attribute(OoxmlNs.W + "anchor"))
            .Where(anchor => !string.IsNullOrWhiteSpace(anchor))
            .OfType<string>()
            .Count(anchor => !bookmarkCounts.ContainsKey(anchor));
    }

    private static int CountDuplicateHyperlinkAnchors(XDocument document)
    {
        IReadOnlyDictionary<string, int> bookmarkCounts = CountBookmarkNames(document);
        return document
            .Descendants(OoxmlNs.W + "hyperlink")
            .Select(hyperlink => (string?)hyperlink.Attribute(OoxmlNs.W + "anchor"))
            .Where(anchor => !string.IsNullOrWhiteSpace(anchor))
            .OfType<string>()
            .Count(anchor => bookmarkCounts.TryGetValue(anchor, out int count) && count > 1);
    }

    private static IReadOnlyDictionary<string, int> CountBookmarkNames(XDocument document)
    {
        return document
            .Descendants(OoxmlNs.W + "bookmarkStart")
            .Select(bookmark => (string?)bookmark.Attribute(OoxmlNs.W + "name"))
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .OfType<string>()
            .GroupBy(name => name, StringComparer.Ordinal)
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

        diagnostics.Add(new DocxDiagnostic(DocxSeverity.Warning, code, $"{message} Count={count}.") with
        {
            PartName = partName,
            Story = story,
            Feature = feature,
            Fallback = fallback
        });
    }
}
