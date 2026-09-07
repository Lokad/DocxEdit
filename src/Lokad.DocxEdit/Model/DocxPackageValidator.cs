using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

// Checks the package and reports content problems as diagnostics with part names
// and codes. Malformed bytes still surface as document exceptions at the loading
// boundary, which the editor converts into failed results; validator logic itself
// never throws for the content it can read.
internal static partial class DocxPackageValidator
{
    public static IReadOnlyList<DocxDiagnostic> Validate(
        OoxmlPackage package,
        DocxValidationProfile profile,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        IReadOnlyDictionary<string, string> storyPrefixes = profile == DocxValidationProfile.Structural
            ? BuildStoryPrefixes(package, cancellationToken)
            : new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (OoxmlPart part in package.Parts.Values
            .Where(part => part.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) ||
                part.Name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            .OrderBy(part => part.Name, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            ValidateRoot(package, part.Name, document.Root, diagnostics, cancellationToken);
            bool isXmlPart = part.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);
            if (profile == DocxValidationProfile.Structural &&
                isXmlPart &&
                DocxPartRoles.IsValidatedWordPart(package, part.Name, cancellationToken))
            {
                storyPrefixes.TryGetValue(part.Name, out string? storyPrefix);
                ValidateWordPart(package, part.Name, storyPrefix, document, diagnostics, cancellationToken);
            }
        }

        if (profile == DocxValidationProfile.Structural)
        {
            ValidateCommentConsistency(package, diagnostics, cancellationToken);
        }

        return diagnostics;
    }

    private static IReadOnlyDictionary<string, string> BuildStoryPrefixes(
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        return DocxPartRoles.GetStoryPrefixes(package, cancellationToken);
    }

    private static void ValidateRoot(
        OoxmlPackage package,
        string partName,
        XElement? root,
        List<DocxDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        if (root is null)
        {
            diagnostics.Add(Error("E9101", "XML part has no root element.", partName));
            return;
        }

        XName? expected = partName switch
        {
            _ when partName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) => OoxmlNs.Rel + "Relationships",
            _ => DocxPartRoles.GetExpectedRoot(package, partName, cancellationToken)
        };
        if (expected is not null && root.Name != expected)
        {
            diagnostics.Add(Error("E9102", $"Expected root '{expected.LocalName}', found '{root.Name.LocalName}'.", partName));
        }
    }

    private static void ValidateWordPart(
        OoxmlPackage package,
        string partName,
        string? storyPrefix,
        XDocument document,
        List<DocxDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        ValidatePairedIds(document, OoxmlNs.W + "bookmarkStart", OoxmlNs.W + "bookmarkEnd", "bookmark", partName, diagnostics);
        ValidatePairedIds(document, OoxmlNs.W + "commentRangeStart", OoxmlNs.W + "commentRangeEnd", "comment range", partName, diagnostics);
        if (storyPrefix is not null)
        {
            ValidateDuplicateSemanticSelectors(document, partName, storyPrefix, diagnostics);
        }

        ValidateFieldBalance(document, partName, diagnostics);
        ValidateFieldFlags(document, partName, diagnostics);
        ValidateRevisionMarkup(document, partName, diagnostics);
        ValidateContentControls(document, partName, diagnostics);
        ValidateParagraphStyleReferences(package, partName, document, diagnostics, cancellationToken);
        ValidateNumberingReferences(package, partName, document, diagnostics, cancellationToken);
        ValidateNumberingDefinitions(package, partName, document, diagnostics, cancellationToken);
        ValidateHeaderFooterReferences(package, partName, document, diagnostics, cancellationToken);
        ValidateSectionProperties(document, partName, diagnostics);
        ValidateDrawingRelationships(package, partName, document, diagnostics, cancellationToken);
        ValidateDrawingProperties(document, partName, diagnostics);
        ValidateDrawingGeometry(document, partName, diagnostics);
        if (string.Equals(partName, DocxPartRoles.FindCommentsExtendedPartName(package, cancellationToken), StringComparison.OrdinalIgnoreCase))
        {
            ValidateCommentsExtended(package, partName, document, diagnostics, cancellationToken);
        }
        else if (string.Equals(partName, DocxPartRoles.FindCommentsIdsPartName(package, cancellationToken), StringComparison.OrdinalIgnoreCase))
        {
            ValidateCommentsIds(package, partName, document, diagnostics, cancellationToken);
        }
        else if (string.Equals(partName, DocxPartRoles.FindSettingsPartName(package, cancellationToken), StringComparison.OrdinalIgnoreCase))
        {
            ValidateSettings(document, partName, diagnostics);
        }

        ValidateTables(document, partName, diagnostics);
    }

    internal static void ValidateRevisionMarkup(
        XDocument document,
        string partName,
        List<DocxDiagnostic> diagnostics)
    {
        foreach (XElement element in document.Descendants())
        {
            if (element.Name == OoxmlNs.W + "ins" || element.Name == OoxmlNs.W + "del")
            {
                ValidateRevisionMetadata(element, element.Name.LocalName, partName, diagnostics);
            }
            else if (element.Name == OoxmlNs.W + "pPrChange")
            {
                ValidateRevisionMetadata(element, "pPrChange", partName, diagnostics);
                if (element.Element(OoxmlNs.W + "pPr") is null)
                {
                    diagnostics.Add(Error("E9121", "Paragraph property revision w:pPrChange is missing child w:pPr.", partName));
                }
            }
            else if (element.Name == OoxmlNs.W + "tblPrChange")
            {
                ValidateRevisionMetadata(element, "tblPrChange", partName, diagnostics);
                if (element.Element(OoxmlNs.W + "tblPr") is null)
                {
                    diagnostics.Add(Error("E9121", "Table property revision w:tblPrChange is missing child w:tblPr.", partName));
                }
            }
            else if (element.Name == OoxmlNs.W + "trPrChange")
            {
                ValidateRevisionMetadata(element, "trPrChange", partName, diagnostics);
                if (element.Element(OoxmlNs.W + "trPr") is null)
                {
                    diagnostics.Add(Error("E9121", "Row property revision w:trPrChange is missing child w:trPr.", partName));
                }
            }
            else if (element.Name == OoxmlNs.W + "tcPrChange")
            {
                ValidateRevisionMetadata(element, "tcPrChange", partName, diagnostics);
                if (element.Element(OoxmlNs.W + "tcPr") is null)
                {
                    diagnostics.Add(Error("E9121", "Cell property revision w:tcPrChange is missing child w:tcPr.", partName));
                }
            }
            else if (element.Name == OoxmlNs.W + "sectPrChange")
            {
                ValidateRevisionMetadata(element, "sectPrChange", partName, diagnostics);
                if (element.Element(OoxmlNs.W + "sectPr") is null)
                {
                    diagnostics.Add(Error("E9121", "Section property revision w:sectPrChange is missing child w:sectPr.", partName));
                }
            }
            else if (element.Name == OoxmlNs.W + "delText" &&
                !element.Ancestors(OoxmlNs.W + "del").Any())
            {
                diagnostics.Add(Error("E9121", "Deleted text w:delText appears outside w:del revision markup.", partName));
            }
        }
    }

    private static void ValidateRevisionMetadata(
        XElement element,
        string label,
        string partName,
        List<DocxDiagnostic> diagnostics)
    {
        string? id = (string?)element.Attribute(OoxmlNs.W + "id");
        if (!int.TryParse(id, out int parsedId) || parsedId < 0)
        {
            diagnostics.Add(Error("E9121", $"{label} revision has missing or invalid non-negative w:id.", partName));
        }

        string? author = (string?)element.Attribute(OoxmlNs.W + "author");
        if (string.IsNullOrWhiteSpace(author))
        {
            diagnostics.Add(Error("E9121", $"{label} revision has missing or empty w:author.", partName));
        }

        string? date = (string?)element.Attribute(OoxmlNs.W + "date");
        if (!DateTimeOffset.TryParse(date, null, System.Globalization.DateTimeStyles.RoundtripKind, out _))
        {
            diagnostics.Add(Error("E9121", $"{label} revision has missing or invalid w:date.", partName));
        }
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
            .OfType<string>()
            .GroupBy(id => id, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        var ends = document.Descendants(endName)
            .Select(element => (string?)element.Attribute(OoxmlNs.W + "id"))
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .OfType<string>()
            .GroupBy(id => id, StringComparer.Ordinal)
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

    private static void ValidateDuplicateSemanticSelectors(
        XDocument document,
        string partName,
        string storyPrefix,
        List<DocxDiagnostic> diagnostics)
    {
        var bookmarks = document
            .Descendants(OoxmlNs.W + "bookmarkStart")
            .Select((element, index) => (
                Id: $"{storyPrefix}.B{index + 1:0000}",
                Name: (string?)element.Attribute(OoxmlNs.W + "name")))
            .Where(item => !string.IsNullOrWhiteSpace(item.Name))
            .ToArray();
        AddDuplicateSelectorWarnings(
            bookmarks.Select(item => (item.Name, item.Id)),
            "bookmark name",
            "bookmark",
            partName,
            diagnostics);

        var contentControls = document
            .Descendants(OoxmlNs.W + "sdt")
            .Select((element, index) =>
            {
                XElement? properties = element.Element(OoxmlNs.W + "sdtPr");
                return (
                    Id: $"{storyPrefix}.CC{index + 1:0000}",
                    Tag: (string?)properties?.Element(OoxmlNs.W + "tag")?.Attribute(OoxmlNs.W + "val"),
                    Alias: (string?)properties?.Element(OoxmlNs.W + "alias")?.Attribute(OoxmlNs.W + "val"));
            })
            .ToArray();
        AddDuplicateSelectorWarnings(
            contentControls
                .Select(item => (item.Tag, item.Id)),
            "content-control tag",
            "content-control",
            partName,
            diagnostics);
        AddDuplicateSelectorWarnings(
            contentControls
                .Select(item => (item.Alias, item.Id)),
            "content-control alias",
            "content-control",
            partName,
            diagnostics);
    }

    private static void AddDuplicateSelectorWarnings(
        IEnumerable<(string? Value, string Id)> candidates,
        string label,
        string feature,
        string partName,
        List<DocxDiagnostic> diagnostics)
    {
        foreach (IGrouping<string, string> group in candidates
            .WithNonBlankKey(candidate => candidate.Value)
            .GroupBy(pair => pair.Key, pair => pair.Item.Id, StringComparer.Ordinal)
            .Where(group => group.Count() > 1)
            .OrderBy(group => group.Key, StringComparer.Ordinal))
        {
            string ids = string.Join(", ", group);
            diagnostics.Add(Warning(
                "W9109",
                $"Duplicate {label} '{group.Key}' appears {group.Count()} times; candidate IDs: {ids}.",
                partName,
                feature,
                "ambiguous-selector"));
        }
    }

    private static HashSet<string>? ReadNumberingIds(
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        string? numberingPartName = DocxPartRoles.FindNumberingPartName(package, cancellationToken);
        OoxmlPart? numberingPart = numberingPartName is null ? null : package.GetPart(numberingPartName);
        if (numberingPart is null)
        {
            return null;
        }

        using Stream stream = numberingPart.OpenRead();
        XDocument numbering = SafeXml.Load(stream, cancellationToken);
        return numbering
            .Descendants(OoxmlNs.W + "num")
            .Select(num => (string?)num.Attribute(OoxmlNs.W + "numId"))
            .Where(numberingId => !string.IsNullOrWhiteSpace(numberingId))
            .OfType<string>()
            .ToHashSet(StringComparer.Ordinal);
    }

    private static void ValidateHeaderFooterReferences(
        OoxmlPackage package,
        string partName,
        XDocument document,
        List<DocxDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        if (!string.Equals(partName, package.MainDocumentPartName, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        IReadOnlyDictionary<string, OoxmlRelationship> relationships = package
            .GetRelationships(partName, cancellationToken)
            .ToDictionary(relationship => relationship.Id, StringComparer.Ordinal);
        ValidateStoryReferences(document, relationships, OoxmlNs.W + "headerReference", OoxmlRelTypes.Header, "header", partName, diagnostics);
        ValidateStoryReferences(document, relationships, OoxmlNs.W + "footerReference", OoxmlRelTypes.Footer, "footer", partName, diagnostics);
    }

    private static void ValidateStoryReferences(
        XDocument document,
        IReadOnlyDictionary<string, OoxmlRelationship> relationships,
        XName elementName,
        string expectedRelationshipType,
        string label,
        string partName,
        List<DocxDiagnostic> diagnostics)
    {
        foreach (XElement reference in document.Descendants(elementName))
        {
            string? relationshipId = (string?)reference.Attribute(OoxmlNs.R + "id");
            if (string.IsNullOrWhiteSpace(relationshipId))
            {
                diagnostics.Add(Error("E9119", $"{label}Reference is missing r:id.", partName));
                continue;
            }

            if (!relationships.TryGetValue(relationshipId, out OoxmlRelationship? relationship))
            {
                diagnostics.Add(Error("E9119", $"{label}Reference targets missing relationship '{relationshipId}'.", partName));
                continue;
            }

            if (relationship.Type != expectedRelationshipType)
            {
                diagnostics.Add(Error("E9119", $"{label}Reference relationship '{relationshipId}' has type '{relationship.Type}', expected {label} relationship.", partName));
            }
        }
    }

    private static void ValidateSectionProperties(
        XDocument document,
        string partName,
        List<DocxDiagnostic> diagnostics)
    {
        foreach (XElement sectionProperties in document.Descendants(OoxmlNs.W + "sectPr"))
        {
            string? columnCount = (string?)sectionProperties
                .Element(OoxmlNs.W + "cols")
                ?.Attribute(OoxmlNs.W + "num");
            if (!string.IsNullOrWhiteSpace(columnCount) &&
                (!int.TryParse(columnCount, out int parsedColumns) || parsedColumns <= 0))
            {
                diagnostics.Add(Error("E9120", $"Section columns w:num has invalid positive integer value '{columnCount}'.", partName));
            }

            string? orientation = (string?)sectionProperties
                .Element(OoxmlNs.W + "pgSz")
                ?.Attribute(OoxmlNs.W + "orient");
            if (!string.IsNullOrWhiteSpace(orientation) && !DocxOrientationExtensions.TryParseWireValue(orientation, out _))
            {
                diagnostics.Add(Error("E9120", $"Section page size w:orient has invalid value '{orientation}'.", partName));
            }
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
            if (string.IsNullOrWhiteSpace(relationshipId))
            {
                continue;
            }

            if (!relationships.TryGetValue(relationshipId, out OoxmlRelationship? relationship))
            {
                diagnostics.Add(Error("E9105", $"Drawing references missing relationship '{relationshipId}'.", partName));
                continue;
            }

            if (relationship.Type != OoxmlRelTypes.Image)
            {
                diagnostics.Add(Error("E9113", $"Drawing relationship '{relationshipId}' has type '{relationship.Type}', expected image relationship.", partName));
                continue;
            }

            if (relationship.IsExternal)
            {
                continue;
            }

            if (relationship.ResolvedTarget is null)
            {
                diagnostics.Add(Error("E9113", $"Drawing image relationship '{relationshipId}' has no resolved internal target.", partName));
                continue;
            }

            OoxmlPart? imagePart = package.GetPart(relationship.ResolvedTarget);
            if (imagePart is null)
            {
                diagnostics.Add(Error("E9113", $"Drawing image relationship '{relationshipId}' targets missing part '{relationship.ResolvedTarget}'.", partName));
                continue;
            }

            if (!IsImageContentType(imagePart.ContentType))
            {
                diagnostics.Add(Error("E9113", $"Drawing image relationship '{relationshipId}' targets part '{imagePart.Name}' with non-image content type '{imagePart.ContentType ?? "unknown"}'.", partName));
            }
        }
    }

    private static bool IsImageContentType(string? contentType)
    {
        return contentType?.StartsWith("image/", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static void ValidateDrawingProperties(XDocument document, string partName, List<DocxDiagnostic> diagnostics)
    {
        var docPrIds = document
            .Descendants(OoxmlNs.Wp + "docPr")
            .Select(element => (string?)element.Attribute("id"))
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .OfType<string>()
            .GroupBy(id => id, StringComparer.Ordinal)
            .Where(group => group.Count() > 1)
            .OrderBy(group => group.Key, StringComparer.Ordinal);
        foreach (IGrouping<string, string> group in docPrIds)
        {
            diagnostics.Add(Error("E9107", $"Duplicate drawing docPr id '{group.Key}' appears {group.Count()} times.", partName));
        }
    }

    private static void ValidateDrawingGeometry(XDocument document, string partName, List<DocxDiagnostic> diagnostics)
    {
        foreach (XElement extent in document.Descendants(OoxmlNs.Wp + "extent"))
        {
            if (!TryReadPositiveLongAttribute(extent, "cx", out _) ||
                !TryReadPositiveLongAttribute(extent, "cy", out _))
            {
                diagnostics.Add(Error("E9109", "Drawing wp:extent must have positive integer cx and cy attributes.", partName));
            }
        }

        foreach (XElement crop in document.Descendants(OoxmlNs.A + "srcRect"))
        {
            if (!TryReadCropPerThousandPercent(crop, "l", out int left) ||
                !TryReadCropPerThousandPercent(crop, "t", out int top) ||
                !TryReadCropPerThousandPercent(crop, "r", out int right) ||
                !TryReadCropPerThousandPercent(crop, "b", out int bottom))
            {
                diagnostics.Add(Error("E9110", "Drawing a:srcRect crop values must be percentages between 0 and 100.", partName));
                continue;
            }

            if (left + right >= 100_000 || top + bottom >= 100_000)
            {
                diagnostics.Add(Error("E9110", "Drawing a:srcRect opposing crop sides must sum to less than 100 percent.", partName));
            }
        }
    }

    private static bool TryReadPositiveLongAttribute(XElement element, string localName, out long value)
    {
        return long.TryParse((string?)element.Attribute(localName), System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out value) &&
            value > 0;
    }

    private static bool TryReadCropPerThousandPercent(XElement element, string localName, out int value)
    {
        value = 0;
        string? text = (string?)element.Attribute(localName);
        if (string.IsNullOrWhiteSpace(text))
        {
            return true;
        }

        if (text.EndsWith("%", StringComparison.Ordinal))
        {
            if (!decimal.TryParse(text[..^1], System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out decimal percent))
            {
                return false;
            }

            if (percent < 0 || percent > 100)
            {
                return false;
            }

            value = (int)Math.Round(percent * 1000m, MidpointRounding.AwayFromZero);
            return value is >= 0 and <= 100_000;
        }

        return int.TryParse(text, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out value) &&
            value is >= 0 and <= 100_000;
    }

    private static IReadOnlyList<string> GetCommentsPartNames(OoxmlPackage package, CancellationToken cancellationToken)
    {
        var partNames = new SortedSet<string>(StringComparer.Ordinal);
        foreach (ResolvedOoxmlRelationship relationship in package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => relationship.Type == OoxmlRelTypes.Comments))
        {
            partNames.Add(relationship.ResolvedTarget);
        }

        string? commentsPartName = DocxPartRoles.FindCommentsPartName(package, cancellationToken);
        if (commentsPartName is not null)
        {
            partNames.Add(commentsPartName);
        }

        return partNames.ToArray();
    }

    private static DocxDiagnostic Error(string code, string message, string? partName)
    {
        return new DocxDiagnostic(DocxSeverity.Error, code, message) with { PartName = partName };
    }

    private static DocxDiagnostic Warning(string code, string message, string partName, string feature, string fallback)
    {
        return new DocxDiagnostic(DocxSeverity.Warning, code, message) with
        {
            PartName = partName,
            Feature = feature,
            Fallback = fallback
        };
    }

    private sealed class ComplexFieldValidationState
    {
        public bool HasSeparate { get; set; }
    }

    private sealed record VerticalMergeValidationState(int StartColumn, int ColumnSpan);
}
