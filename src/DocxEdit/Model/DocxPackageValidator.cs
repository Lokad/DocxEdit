using System.Xml.Linq;
using DocxEdit.Ooxml;

namespace DocxEdit.Model;

internal static class DocxPackageValidator
{
    public static IReadOnlyList<DocxDiagnostic> Validate(
        OoxmlPackage package,
        DocxValidationProfile profile,
        CancellationToken cancellationToken = default)
    {
        var diagnostics = new List<DocxDiagnostic>();
        IReadOnlyDictionary<string, string> storyPrefixes = profile == DocxValidationProfile.Structural
            ? BuildStoryPrefixes(package, cancellationToken)
            : new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (OoxmlPart part in package.Parts.Values
            .Where(part => part.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(part => part.Name, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            ValidateRoot(part.Name, document.Root, diagnostics);
            if (profile == DocxValidationProfile.Structural &&
                part.Name.StartsWith("/word/", StringComparison.OrdinalIgnoreCase))
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
        var prefixes = new Dictionary<string, string>(StringComparer.Ordinal);
        if (package.MainDocumentPartName is null)
        {
            return prefixes;
        }

        prefixes[package.MainDocumentPartName] = "M";
        IReadOnlyList<OoxmlRelationship> relationships = package.GetRelationships(package.MainDocumentPartName, cancellationToken);
        int headerIndex = 1;
        foreach (OoxmlRelationship relationship in relationships
            .Where(relationship => !relationship.IsExternal && relationship.Type == OoxmlRelTypes.Header && relationship.ResolvedTarget is not null)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
        {
            if (package.GetPart(relationship.ResolvedTarget!) is not null)
            {
                prefixes[relationship.ResolvedTarget!] = $"H{headerIndex++:000}";
            }
        }

        int footerIndex = 1;
        foreach (OoxmlRelationship relationship in relationships
            .Where(relationship => !relationship.IsExternal && relationship.Type == OoxmlRelTypes.Footer && relationship.ResolvedTarget is not null)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
        {
            if (package.GetPart(relationship.ResolvedTarget!) is not null)
            {
                prefixes[relationship.ResolvedTarget!] = $"F{footerIndex++:000}";
            }
        }

        return prefixes;
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
        ValidateContentControls(document, partName, diagnostics);
        ValidateParagraphStyleReferences(package, partName, document, diagnostics, cancellationToken);
        ValidateNumberingReferences(package, partName, document, diagnostics, cancellationToken);
        ValidateNumberingDefinitions(partName, document, diagnostics);
        ValidateHeaderFooterReferences(package, partName, document, diagnostics, cancellationToken);
        ValidateDrawingRelationships(package, partName, document, diagnostics, cancellationToken);
        ValidateDrawingProperties(document, partName, diagnostics);
        ValidateDrawingGeometry(document, partName, diagnostics);
        if (string.Equals(partName, "/word/commentsExtended.xml", StringComparison.OrdinalIgnoreCase))
        {
            ValidateCommentsExtended(package, partName, document, diagnostics, cancellationToken);
        }
        else if (string.Equals(partName, "/word/settings.xml", StringComparison.OrdinalIgnoreCase))
        {
            ValidateSettings(document, partName, diagnostics);
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
            bookmarks.Select(item => (item.Name!, item.Id)),
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
                .Where(item => !string.IsNullOrWhiteSpace(item.Tag))
                .Select(item => (item.Tag!, item.Id)),
            "content-control tag",
            "content-control",
            partName,
            diagnostics);
        AddDuplicateSelectorWarnings(
            contentControls
                .Where(item => !string.IsNullOrWhiteSpace(item.Alias))
                .Select(item => (item.Alias!, item.Id)),
            "content-control alias",
            "content-control",
            partName,
            diagnostics);
    }

    private static void AddDuplicateSelectorWarnings(
        IEnumerable<(string Value, string Id)> candidates,
        string label,
        string feature,
        string partName,
        List<DocxDiagnostic> diagnostics)
    {
        foreach (IGrouping<string, (string Value, string Id)> group in candidates
            .GroupBy(candidate => candidate.Value, StringComparer.Ordinal)
            .Where(group => group.Count() > 1)
            .OrderBy(group => group.Key, StringComparer.Ordinal))
        {
            string ids = string.Join(", ", group.Select(candidate => candidate.Id));
            diagnostics.Add(Warning(
                "W9109",
                $"Duplicate {label} '{group.Key}' appears {group.Count()} times; candidate IDs: {ids}.",
                partName,
                feature,
                "ambiguous-selector"));
        }
    }

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

    private static void ValidateParagraphStyleReferences(
        OoxmlPackage package,
        string partName,
        XDocument document,
        List<DocxDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        HashSet<string>? paragraphStyleIds = ReadParagraphStyleIds(package, cancellationToken);
        if (paragraphStyleIds is null)
        {
            return;
        }

        foreach (string styleId in document
            .Descendants(OoxmlNs.W + "pStyle")
            .Select(style => (string?)style.Attribute(OoxmlNs.W + "val"))
            .Where(styleId => !string.IsNullOrWhiteSpace(styleId))
            .Select(styleId => styleId!)
            .Distinct(StringComparer.Ordinal)
            .Order(StringComparer.Ordinal))
        {
            if (!paragraphStyleIds.Contains(styleId))
            {
                diagnostics.Add(Warning(
                    "W9116",
                    $"Paragraph style reference '{styleId}' is not defined in /word/styles.xml.",
                    partName,
                    "style",
                    "missing-style-definition"));
            }
        }
    }

    private static HashSet<string>? ReadParagraphStyleIds(
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        OoxmlPart? stylesPart = package.GetPart("/word/styles.xml");
        if (stylesPart is null)
        {
            return null;
        }

        using Stream stream = stylesPart.OpenRead();
        XDocument styles = SafeXml.Load(stream, cancellationToken);
        return styles
            .Descendants(OoxmlNs.W + "style")
            .Where(style => string.Equals((string?)style.Attribute(OoxmlNs.W + "type"), "paragraph", StringComparison.Ordinal))
            .Select(style => (string?)style.Attribute(OoxmlNs.W + "styleId"))
            .Where(styleId => !string.IsNullOrWhiteSpace(styleId))
            .Select(styleId => styleId!)
            .ToHashSet(StringComparer.Ordinal);
    }

    private static void ValidateNumberingReferences(
        OoxmlPackage package,
        string partName,
        XDocument document,
        List<DocxDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        if (string.Equals(partName, "/word/numbering.xml", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        HashSet<string>? numberingIds = ReadNumberingIds(package, cancellationToken);
        if (numberingIds is null)
        {
            return;
        }

        foreach (string numberingId in document
            .Descendants(OoxmlNs.W + "numPr")
            .Elements(OoxmlNs.W + "numId")
            .Select(numId => (string?)numId.Attribute(OoxmlNs.W + "val"))
            .Where(numberingId => !string.IsNullOrWhiteSpace(numberingId) && numberingId != "0")
            .Select(numberingId => numberingId!)
            .Distinct(StringComparer.Ordinal)
            .Order(StringComparer.Ordinal))
        {
            if (!numberingIds.Contains(numberingId))
            {
                diagnostics.Add(Warning(
                    "W9117",
                    $"Numbering reference '{numberingId}' is not defined in /word/numbering.xml.",
                    partName,
                    "numbering",
                    "missing-numbering-definition"));
            }
        }
    }

    private static void ValidateNumberingDefinitions(
        string partName,
        XDocument document,
        List<DocxDiagnostic> diagnostics)
    {
        if (!string.Equals(partName, "/word/numbering.xml", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var abstractIds = document
            .Descendants(OoxmlNs.W + "abstractNum")
            .Select(abstractNum => (string?)abstractNum.Attribute(OoxmlNs.W + "abstractNumId"))
            .Where(abstractId => !string.IsNullOrWhiteSpace(abstractId))
            .Select(abstractId => abstractId!)
            .ToHashSet(StringComparer.Ordinal);

        foreach (string abstractId in document
            .Descendants(OoxmlNs.W + "num")
            .Elements(OoxmlNs.W + "abstractNumId")
            .Select(abstractNumId => (string?)abstractNumId.Attribute(OoxmlNs.W + "val"))
            .Where(abstractId => !string.IsNullOrWhiteSpace(abstractId))
            .Select(abstractId => abstractId!)
            .Distinct(StringComparer.Ordinal)
            .Order(StringComparer.Ordinal))
        {
            if (!abstractIds.Contains(abstractId))
            {
                diagnostics.Add(Warning(
                    "W9117",
                    $"Numbering definition references missing abstractNumId '{abstractId}'.",
                    partName,
                    "numbering",
                    "missing-abstract-numbering-definition"));
            }
        }
    }

    private static HashSet<string>? ReadNumberingIds(
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        OoxmlPart? numberingPart = package.GetPart("/word/numbering.xml");
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
            .Select(numberingId => numberingId!)
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
            .Select(id => id!)
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

            value = (int)Math.Round(percent * 1000m, MidpointRounding.AwayFromZero);
            return value is >= 0 and <= 100_000;
        }

        return int.TryParse(text, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out value) &&
            value is >= 0 and <= 100_000;
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

    private static void ValidateSettings(
        XDocument document,
        string partName,
        List<DocxDiagnostic> diagnostics)
    {
        foreach (XElement updateFields in document.Descendants(OoxmlNs.W + "updateFields"))
        {
            string? value = (string?)updateFields.Attribute(OoxmlNs.W + "val");
            if (value is null || value is "0" or "1" or "true" or "false" or "on" or "off")
            {
                continue;
            }

            diagnostics.Add(Error("E9118", $"Settings w:updateFields has invalid OnOff value '{value}'.", partName));
        }
    }

    private static HashSet<string> ReadCommentParaIds(OoxmlPackage package, CancellationToken cancellationToken)
    {
        var paraIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (string partName in GetCommentsPartNames(package, cancellationToken))
        {
            OoxmlPart? commentsPart = package.GetPart(partName);
            if (commentsPart is null)
            {
                continue;
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
        }

        return paraIds;
    }

    private static void ValidateCommentConsistency(
        OoxmlPackage package,
        List<DocxDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        var bodyIds = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        foreach (string partName in GetCommentsPartNames(package, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            foreach (XElement comment in document.Descendants(OoxmlNs.W + "comment"))
            {
                string? id = (string?)comment.Attribute(OoxmlNs.W + "id");
                if (string.IsNullOrWhiteSpace(id))
                {
                    diagnostics.Add(Error("E9111", "Comment body is missing w:id.", partName));
                    continue;
                }

                if (!bodyIds.TryGetValue(id, out List<string>? partNames))
                {
                    partNames = [];
                    bodyIds[id] = partNames;
                }

                partNames.Add(partName);
            }
        }

        foreach (KeyValuePair<string, List<string>> item in bodyIds
            .Where(item => item.Value.Count > 1)
            .OrderBy(item => item.Key, StringComparer.Ordinal))
        {
            string partNames = string.Join(", ", item.Value.Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal));
            diagnostics.Add(Error("E9111", $"Duplicate comment body id '{item.Key}' appears {item.Value.Count} times in comments parts: {partNames}.", item.Value[0]));
        }

        var rangeIds = new HashSet<string>(StringComparer.Ordinal);
        var referenceIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (OoxmlPart part in package.Parts.Values
            .Where(part => part.Name.StartsWith("/word/", StringComparison.OrdinalIgnoreCase) &&
                part.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                !part.Name.Contains("/_rels/", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(part.Name, "/word/comments.xml", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(part.Name, "/word/commentsExtended.xml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(part => part.Name, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            foreach (XElement marker in document
                .Descendants()
                .Where(element => element.Name == OoxmlNs.W + "commentRangeStart" ||
                    element.Name == OoxmlNs.W + "commentRangeEnd" ||
                    element.Name == OoxmlNs.W + "commentReference"))
            {
                string? id = (string?)marker.Attribute(OoxmlNs.W + "id");
                if (string.IsNullOrWhiteSpace(id))
                {
                    diagnostics.Add(Error("E9111", $"{marker.Name.LocalName} is missing w:id.", part.Name));
                    continue;
                }

                if (marker.Name == OoxmlNs.W + "commentReference")
                {
                    referenceIds.Add(id);
                }
                else
                {
                    rangeIds.Add(id);
                }
            }
        }

        foreach (string id in rangeIds.Concat(referenceIds).Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal))
        {
            if (!bodyIds.ContainsKey(id))
            {
                diagnostics.Add(Error("E9111", $"Comment markup id '{id}' has no matching comment body.", null));
            }
        }

        foreach (string id in rangeIds.Except(referenceIds, StringComparer.Ordinal).Order(StringComparer.Ordinal))
        {
            diagnostics.Add(Error("E9111", $"Comment range id '{id}' has no matching commentReference.", null));
        }
    }

    private static IReadOnlyList<string> GetCommentsPartNames(OoxmlPackage package, CancellationToken cancellationToken)
    {
        var partNames = new SortedSet<string>(StringComparer.Ordinal);
        if (package.MainDocumentPartName is not null)
        {
            foreach (OoxmlRelationship relationship in package
                .GetRelationships(package.MainDocumentPartName, cancellationToken)
                .Where(relationship => !relationship.IsExternal &&
                    relationship.Type == OoxmlRelTypes.Comments &&
                    relationship.ResolvedTarget is not null))
            {
                partNames.Add(relationship.ResolvedTarget!);
            }
        }

        if (package.GetPart("/word/comments.xml") is not null)
        {
            partNames.Add("/word/comments.xml");
        }

        return partNames.ToArray();
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

            ValidateTableVisualGrid(table, partName, diagnostics);
        }
    }

    private static void ValidateTableVisualGrid(XElement table, string partName, List<DocxDiagnostic> diagnostics)
    {
        int declaredGridColumns = table
            .Element(OoxmlNs.W + "tblGrid")
            ?.Elements(OoxmlNs.W + "gridCol")
            .Count() ?? 0;
        var activeVerticalMerges = new Dictionary<int, VerticalMergeValidationState>();
        int rowIndex = 1;
        foreach (XElement row in table.Elements(OoxmlNs.W + "tr"))
        {
            int gridBefore = ReadTableGridOffset(row, "gridBefore", rowIndex, partName, diagnostics);
            int gridAfter = ReadTableGridOffset(row, "gridAfter", rowIndex, partName, diagnostics);
            for (int column = 1; column <= gridBefore; column++)
            {
                activeVerticalMerges.Remove(column);
            }

            int columnIndex = 1 + gridBefore;
            foreach (XElement cell in row.Elements(OoxmlNs.W + "tc"))
            {
                int columnSpan = ReadTableCellColumnSpan(cell, rowIndex, columnIndex, partName, diagnostics);
                string? verticalMerge = ReadTableCellVerticalMerge(cell);
                if (string.Equals(verticalMerge, "restart", StringComparison.Ordinal))
                {
                    var state = new VerticalMergeValidationState(columnIndex, columnSpan);
                    for (int column = columnIndex; column < columnIndex + columnSpan; column++)
                    {
                        activeVerticalMerges[column] = state;
                    }
                }
                else if (string.Equals(verticalMerge, "continue", StringComparison.Ordinal))
                {
                    if (!activeVerticalMerges.TryGetValue(columnIndex, out VerticalMergeValidationState? state))
                    {
                        diagnostics.Add(Error("E9114", $"Table vertical merge continuation at row {rowIndex}, column {columnIndex} has no active restart.", partName));
                    }
                    else if (state.StartColumn != columnIndex || state.ColumnSpan != columnSpan)
                    {
                        diagnostics.Add(Error("E9114", $"Table vertical merge continuation at row {rowIndex}, column {columnIndex} span {columnSpan} does not match active restart span {state.ColumnSpan} at column {state.StartColumn}.", partName));
                    }
                }
                else
                {
                    for (int column = columnIndex; column < columnIndex + columnSpan; column++)
                    {
                        activeVerticalMerges.Remove(column);
                    }
                }

                columnIndex += columnSpan;
            }

            int visualColumnCount = columnIndex - 1 + gridAfter;
            if (declaredGridColumns > 0 && visualColumnCount > declaredGridColumns)
            {
                diagnostics.Add(Error("E9114", $"Table row {rowIndex} spans {visualColumnCount} visual column(s), exceeding declared tblGrid column count {declaredGridColumns}.", partName));
            }

            rowIndex++;
        }
    }

    private static int ReadTableGridOffset(
        XElement row,
        string localName,
        int rowIndex,
        string partName,
        List<DocxDiagnostic> diagnostics)
    {
        XElement? element = row
            .Element(OoxmlNs.W + "trPr")
            ?.Element(OoxmlNs.W + localName);
        if (element is null)
        {
            return 0;
        }

        string? value = (string?)element.Attribute(OoxmlNs.W + "val");
        if (int.TryParse(value, out int parsed) && parsed >= 0)
        {
            return parsed;
        }

        diagnostics.Add(Error("E9114", $"Table row {rowIndex} has invalid {localName} value '{value ?? "unknown"}'.", partName));
        return 0;
    }

    private static int ReadTableCellColumnSpan(
        XElement cell,
        int rowIndex,
        int columnIndex,
        string partName,
        List<DocxDiagnostic> diagnostics)
    {
        XElement? gridSpan = cell
            .Element(OoxmlNs.W + "tcPr")
            ?.Element(OoxmlNs.W + "gridSpan");
        if (gridSpan is null)
        {
            return 1;
        }

        string? value = (string?)gridSpan.Attribute(OoxmlNs.W + "val");
        if (int.TryParse(value, out int parsed) && parsed > 0)
        {
            return parsed;
        }

        diagnostics.Add(Error("E9114", $"Table cell at row {rowIndex}, column {columnIndex} has invalid gridSpan value '{value ?? "unknown"}'.", partName));
        return 1;
    }

    private static string? ReadTableCellVerticalMerge(XElement cell)
    {
        XElement? verticalMerge = cell
            .Element(OoxmlNs.W + "tcPr")
            ?.Element(OoxmlNs.W + "vMerge");
        if (verticalMerge is null)
        {
            return null;
        }

        return (string?)verticalMerge.Attribute(OoxmlNs.W + "val") ?? "continue";
    }

    private static DocxDiagnostic Error(string code, string message, string? partName)
    {
        return new DocxDiagnostic(DocxSeverity.Error, code, message, PartName: partName);
    }

    private static DocxDiagnostic Warning(string code, string message, string partName, string feature, string fallback)
    {
        return new DocxDiagnostic(
            DocxSeverity.Warning,
            code,
            message,
            PartName: partName,
            Feature: feature,
            Fallback: fallback);
    }

    private sealed class ComplexFieldValidationState
    {
        public bool HasSeparate { get; set; }
    }

    private sealed record VerticalMergeValidationState(int StartColumn, int ColumnSpan);
}
