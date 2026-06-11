using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocxEdit.Model;
using DocxEdit.Ooxml;

namespace DocxEdit;

internal static class DocxPatchEngine
{
    private const string SettingsContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";

    private static readonly IReadOnlyDictionary<XName, string> ProtectedTextEditElements = new Dictionary<XName, string>
    {
        [OoxmlNs.W + "hyperlink"] = "hyperlink",
        [OoxmlNs.W + "fldSimple"] = "field",
        [OoxmlNs.W + "fldChar"] = "field",
        [OoxmlNs.W + "instrText"] = "field",
        [OoxmlNs.W + "commentRangeStart"] = "comment",
        [OoxmlNs.W + "commentRangeEnd"] = "comment",
        [OoxmlNs.W + "commentReference"] = "comment",
        [OoxmlNs.W + "bookmarkStart"] = "bookmark",
        [OoxmlNs.W + "bookmarkEnd"] = "bookmark",
        [OoxmlNs.W + "sdt"] = "content-control",
        [OoxmlNs.W + "ins"] = "tracked-insertion",
        [OoxmlNs.W + "del"] = "tracked-deletion",
        [OoxmlNs.W + "moveFrom"] = "tracked-move-from",
        [OoxmlNs.W + "moveTo"] = "tracked-move-to",
        [OoxmlNs.W + "moveFromRangeStart"] = "tracked-move-from-range",
        [OoxmlNs.W + "moveFromRangeEnd"] = "tracked-move-from-range",
        [OoxmlNs.W + "moveToRangeStart"] = "tracked-move-to-range",
        [OoxmlNs.W + "moveToRangeEnd"] = "tracked-move-to-range",
        [OoxmlNs.W + "customXmlInsRangeStart"] = "tracked-custom-xml-insertion",
        [OoxmlNs.W + "customXmlInsRangeEnd"] = "tracked-custom-xml-insertion",
        [OoxmlNs.W + "customXmlDelRangeStart"] = "tracked-custom-xml-deletion",
        [OoxmlNs.W + "customXmlDelRangeEnd"] = "tracked-custom-xml-deletion",
        [OoxmlNs.W + "customXmlMoveFromRangeStart"] = "tracked-custom-xml-move-from",
        [OoxmlNs.W + "customXmlMoveFromRangeEnd"] = "tracked-custom-xml-move-from",
        [OoxmlNs.W + "customXmlMoveToRangeStart"] = "tracked-custom-xml-move-to",
        [OoxmlNs.W + "customXmlMoveToRangeEnd"] = "tracked-custom-xml-move-to",
        [OoxmlNs.W + "drawing"] = "drawing",
        [OoxmlNs.W + "pict"] = "picture",
        [OoxmlNs.W + "object"] = "object"
    };

    public static PatchExecutionResult Check(
        OoxmlPackage package,
        DocxPatch patch,
        DocxEditOptions options,
        CancellationToken cancellationToken = default)
    {
        return Execute(package, patch, options, apply: false, cancellationToken);
    }

    public static PatchExecutionResult Apply(
        OoxmlPackage package,
        DocxPatch patch,
        DocxEditOptions options,
        CancellationToken cancellationToken = default)
    {
        return Execute(package, patch, options, apply: true, cancellationToken);
    }

    private static PatchExecutionResult Execute(
        OoxmlPackage package,
        DocxPatch patch,
        DocxEditOptions options,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        var reports = new List<DocxPatchOperationReport>();
        foreach (DocxPatchOperation operation in patch.Operations)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var operationDiagnostics = new List<DocxDiagnostic>();
            if (options.TrackChanges == TrackChangesMode.Require)
            {
                operationDiagnostics.Add(Diagnostic(DocxSeverity.Error, "E6001", $"TrackChangesMode.Require is not supported for operation '{operation.OperationName}'.", operation, operation.Fields.GetValueOrDefault("target")));
            }
            else
            {
                if (options.TrackChanges == TrackChangesMode.Suggest)
                {
                    operationDiagnostics.Add(Diagnostic(DocxSeverity.Warning, "W4001", $"TrackChangesMode.Suggest is not supported for operation '{operation.OperationName}'; applying the edit directly.", operation, operation.Fields.GetValueOrDefault("target")));
                }

                operationDiagnostics.AddRange(operation.OperationName switch
                {
                    "replace-text" => ExecuteReplaceText(package, operation, apply, cancellationToken),
                    "replace-paragraph" => ExecuteReplaceParagraph(package, operation, apply, cancellationToken),
                    "insert-before" => ExecuteInsertBlock(package, operation, insertAfter: false, apply, cancellationToken),
                    "insert-after" => ExecuteInsertBlock(package, operation, insertAfter: true, apply, cancellationToken),
                    "delete-block" => ExecuteDeleteBlock(package, operation, apply, cancellationToken),
                    "set-style" => ExecuteSetStyle(package, operation, apply, cancellationToken),
                    "set-cell" => ExecuteSetCell(package, operation, apply, cancellationToken),
                    "append-row" => ExecuteAppendRow(package, operation, apply, cancellationToken),
                    "insert-row-before" => ExecuteInsertRow(package, operation, insertAfter: false, apply, cancellationToken),
                    "insert-row-after" => ExecuteInsertRow(package, operation, insertAfter: true, apply, cancellationToken),
                    "delete-row" => ExecuteDeleteRow(package, operation, apply, cancellationToken),
                    "replace-image" => ExecuteReplaceImage(package, operation, options, apply, cancellationToken),
                    "insert-image-after" => ExecuteInsertImageAfter(package, operation, options, apply, cancellationToken),
                    "set-image-alt" => ExecuteSetImageAlt(package, operation, apply, cancellationToken),
                    "delete-image" => ExecuteDeleteImage(package, operation, apply, cancellationToken),
                    "set-section-columns" => ExecuteSetSectionColumns(package, operation, apply, cancellationToken),
                    "set-section-orientation" => ExecuteSetSectionOrientation(package, operation, apply, cancellationToken),
                    _ => [Diagnostic(DocxSeverity.Error, "E4201", $"Unsupported operation '{operation.OperationName}'.", operation)]
                });
            }
            bool operationSuccess = operationDiagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error);
            diagnostics.AddRange(operationDiagnostics);
            reports.Add(new DocxPatchOperationReport(
                operation.Index,
                operation.OperationName,
                operation.Fields.GetValueOrDefault("target"),
                operationSuccess,
                operationDiagnostics));
        }

        if (apply &&
            options.MarkFieldsDirtyWhenEditing &&
            patch.Operations.Count != 0 &&
            diagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error))
        {
            diagnostics.AddRange(MarkFieldsDirty(package, cancellationToken));
        }

        if (apply && diagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error))
        {
            diagnostics.AddRange(ValidateEditedPackage(package, cancellationToken));
        }

        return new PatchExecutionResult(diagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error), diagnostics, reports);
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteReplaceText(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? find = ReadRequiredField(operation, "find", diagnostics);
        string? replacement = ReadRequiredField(operation, "with", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        bool? preserveRuns = ReadBooleanField(operation, "preserve-runs", diagnostics);
        int? occurrence = ReadPositiveOccurrence(operation, diagnostics);
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        bool shouldPreserveRuns = preserveRuns ?? true;
        ParagraphTarget? paragraphTarget = ResolveParagraphTarget(package, operation, target!, cancellationToken, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
        if (selectorDiagnostics.Count != 0)
        {
            return selectorDiagnostics;
        }

        if (paragraphTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (find!.Length == 0)
        {
            return [Diagnostic(DocxSeverity.Error, "E4205", "Field 'find' must not be empty.", operation, target)];
        }

        string current = ReadVisibleText(paragraphTarget.Paragraph);
        if (expected is not null && !string.Equals(current, expected, StringComparison.Ordinal))
        {
            return
            [
                Diagnostic(
                    DocxSeverity.Error,
                    "E3201",
                    $"Guard failed for {target}. Expected text does not match current text.",
                    operation,
                    target)
            ];
        }

        if (TryGetProtectedTextEditFeature(paragraphTarget.Paragraph, out string protectedFeature))
        {
            return [Diagnostic(DocxSeverity.Error, "E4305", $"Text edit for {target} crosses protected OOXML boundary '{protectedFeature}'.", operation, target)];
        }

        IReadOnlyList<TextRange> matches = FindTextMatches(current, find!, occurrence);
        if (matches.Count == 0)
        {
            return [Diagnostic(DocxSeverity.Error, "E4203", $"Find text was not found in {target}.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        if (shouldPreserveRuns)
        {
            if (!TryReplaceParagraphTextPreservingRuns(paragraphTarget.Paragraph, matches, replacement!, out string? unsupportedReason))
            {
                return [Diagnostic(DocxSeverity.Error, "E4306", $"Run-preserving replacement is not supported for {target}: {unsupportedReason}. Use preserve-runs false to allow paragraph-level rewriting.", operation, target)];
            }
        }
        else
        {
            string edited = ApplyTextReplacement(current, matches, replacement!);
            ReplaceParagraphText(paragraphTarget.Paragraph, edited);
        }

        SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteReplaceParagraph(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        string? style = operation.Fields.GetValueOrDefault("style");
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ParagraphTarget? paragraphTarget = ResolveParagraphTarget(package, operation, target!, cancellationToken, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
        if (selectorDiagnostics.Count != 0)
        {
            return selectorDiagnostics;
        }

        if (paragraphTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        string current = ReadVisibleText(paragraphTarget.Paragraph);
        if (expected is not null && !string.Equals(current, expected, StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected text does not match current text.", operation, target)];
        }

        if (TryGetProtectedTextEditFeature(paragraphTarget.Paragraph, out string protectedFeature))
        {
            return [Diagnostic(DocxSeverity.Error, "E4305", $"Paragraph replacement for {target} would remove protected OOXML boundary '{protectedFeature}'.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        ReplaceParagraphText(paragraphTarget.Paragraph, text!);
        if (style is not null)
        {
            SetParagraphStyle(paragraphTarget.Paragraph, style);
        }

        SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteInsertBlock(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool insertAfter,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        string? style = operation.Fields.GetValueOrDefault("style");
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
        XElement? targetBlock = ResolveMainBlock(body, operation, target!, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
        if (selectorDiagnostics.Count != 0)
        {
            return selectorDiagnostics;
        }

        if (targetBlock is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        XElement paragraph = CreateSimpleParagraph(text!, style);
        if (insertAfter)
        {
            targetBlock.AddAfterSelf(paragraph);
        }
        else
        {
            targetBlock.AddBeforeSelf(paragraph);
        }

        SaveMainDocument(package, document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteDeleteBlock(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
        XElement? targetBlock = ResolveMainBlock(body, operation, target!, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
        if (selectorDiagnostics.Count != 0)
        {
            return selectorDiagnostics;
        }

        if (targetBlock is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (expected is not null)
        {
            string current = ReadVisibleText(targetBlock);
            if (!string.Equals(current, expected, StringComparison.Ordinal))
            {
                return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected text does not match current text.", operation, target)];
            }
        }

        if (!apply)
        {
            return [];
        }

        targetBlock.Remove();
        SaveMainDocument(package, document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetStyle(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? style = ReadRequiredField(operation, "style", diagnostics);
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ParagraphTarget? paragraphTarget = ResolveParagraphTarget(package, operation, target!, cancellationToken, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
        if (selectorDiagnostics.Count != 0)
        {
            return selectorDiagnostics;
        }

        if (paragraphTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!TryResolveStyleId(package, style!, "paragraph", cancellationToken, out string? styleId, out DocxDiagnostic? styleDiagnostic, operation, target))
        {
            return [styleDiagnostic!];
        }

        if (!apply)
        {
            return [];
        }

        SetParagraphStyle(paragraphTarget.Paragraph, styleId!);
        SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteReplaceImage(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? asset = ReadRequiredField(operation, "asset", diagnostics);
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!TryParseMainImageTarget(target!, out int imageOrdinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported replace-image target '{target}'. Expected an image ID such as M.I0001.", operation, target)];
        }

        ImageTarget? imageTarget = FindMainImageTarget(package, imageOrdinal, cancellationToken);
        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!TryReadAsset(options.AssetProvider, asset!, cancellationToken, out byte[] bytes, out string? contentType, out DocxDiagnostic? assetDiagnostic, operation, target))
        {
            return [assetDiagnostic!];
        }

        if (imageTarget.Part.ContentType is not null &&
            !string.Equals(imageTarget.Part.ContentType, contentType, StringComparison.OrdinalIgnoreCase))
        {
            return [Diagnostic(DocxSeverity.Error, "E5204", $"Replacing image content type '{imageTarget.Part.ContentType}' with '{contentType}' is not supported for existing media part {imageTarget.Part.Name}.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        package.ReplacePartBytes(imageTarget.Part.Name, bytes);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteInsertImageAfter(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? asset = ReadRequiredField(operation, "asset", diagnostics);
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ParagraphTarget? paragraphTarget = ResolveParagraphTarget(package, operation, target!, cancellationToken, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
        if (selectorDiagnostics.Count != 0)
        {
            return selectorDiagnostics;
        }

        if (paragraphTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!TryReadAsset(options.AssetProvider, asset!, cancellationToken, out byte[] bytes, out string? contentType, out DocxDiagnostic? assetDiagnostic, operation, target))
        {
            return [assetDiagnostic!];
        }

        if (!TryReadImageExtent(operation, out long widthEmus, out long heightEmus, out DocxDiagnostic? dimensionDiagnostic))
        {
            return [dimensionDiagnostic!];
        }

        if (!apply)
        {
            return [];
        }

        string imagePartName = OoxmlMediaParts.AllocateImagePartName(package.Parts.Keys, contentType!);
        string relationshipId = OoxmlIds.AllocateRelationshipId(package.GetRelationships(paragraphTarget.PartName, cancellationToken).Select(relationship => relationship.Id));
        package.AddPart(imagePartName, contentType!, bytes);
        package.AddRelationship(paragraphTarget.PartName, relationshipId, OoxmlRelTypes.Image, GetRelativeRelationshipTarget(paragraphTarget.PartName, imagePartName));
        XElement imageParagraph = CreateInlineImageParagraph(relationshipId, widthEmus, heightEmus, operation.Fields.GetValueOrDefault("alt") ?? string.Empty);
        paragraphTarget.Paragraph.AddAfterSelf(imageParagraph);
        SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
        return [];
    }

    private static bool TryReadImageExtent(
        DocxPatchOperation operation,
        out long widthEmus,
        out long heightEmus,
        out DocxDiagnostic? diagnostic)
    {
        widthEmus = OoxmlUnits.InchesToEmu(1);
        heightEmus = widthEmus;
        diagnostic = null;
        bool hasWidth = operation.Fields.TryGetValue("width", out string? width);
        bool hasHeight = operation.Fields.TryGetValue("height", out string? height);
        if (hasWidth && !OoxmlUnits.TryParseDimension(width!, out widthEmus))
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E5206", $"Invalid image width '{width}'.", operation, operation.Fields.GetValueOrDefault("target"));
            return false;
        }

        if (hasHeight && !OoxmlUnits.TryParseDimension(height!, out heightEmus))
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E5206", $"Invalid image height '{height}'.", operation, operation.Fields.GetValueOrDefault("target"));
            return false;
        }

        if (hasWidth && !hasHeight)
        {
            heightEmus = widthEmus;
        }
        else if (!hasWidth && hasHeight)
        {
            widthEmus = heightEmus;
        }

        return true;
    }

    private static string GetRelativeRelationshipTarget(string sourcePartName, string targetPartName)
    {
        string normalizedSource = OoxmlPath.NormalizePartName(sourcePartName);
        string normalizedTarget = OoxmlPath.NormalizePartName(targetPartName);
        string sourceDirectory = normalizedSource[..(normalizedSource.LastIndexOf('/') + 1)];
        return normalizedTarget.StartsWith(sourceDirectory, StringComparison.Ordinal)
            ? normalizedTarget[sourceDirectory.Length..]
            : normalizedTarget.TrimStart('/');
    }

    private static XElement CreateInlineImageParagraph(string relationshipId, long widthEmus, long heightEmus, string alt)
    {
        return new XElement(
            OoxmlNs.W + "p",
            new XElement(
                OoxmlNs.W + "r",
                new XElement(
                    OoxmlNs.W + "drawing",
                    new XElement(
                        OoxmlNs.Wp + "inline",
                        new XElement(OoxmlNs.Wp + "extent", new XAttribute("cx", widthEmus), new XAttribute("cy", heightEmus)),
                        new XElement(OoxmlNs.Wp + "docPr", new XAttribute("id", "1"), new XAttribute("name", "Picture"), new XAttribute("descr", alt)),
                        new XElement(
                            OoxmlNs.A + "graphic",
                            new XElement(
                                OoxmlNs.A + "graphicData",
                                new XAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture"),
                                new XElement(
                                    OoxmlNs.Pic + "pic",
                                    new XElement(
                                        OoxmlNs.Pic + "nvPicPr",
                                        new XElement(OoxmlNs.Pic + "cNvPr", new XAttribute("id", "0"), new XAttribute("name", "Picture")),
                                        new XElement(OoxmlNs.Pic + "cNvPicPr")),
                                    new XElement(
                                        OoxmlNs.Pic + "blipFill",
                                        new XElement(OoxmlNs.A + "blip", new XAttribute(OoxmlNs.R + "embed", relationshipId)),
                                        new XElement(OoxmlNs.A + "stretch", new XElement(OoxmlNs.A + "fillRect"))),
                                    new XElement(
                                        OoxmlNs.Pic + "spPr",
                                        new XElement(
                                            OoxmlNs.A + "xfrm",
                                            new XElement(OoxmlNs.A + "off", new XAttribute("x", "0"), new XAttribute("y", "0")),
                                            new XElement(OoxmlNs.A + "ext", new XAttribute("cx", widthEmus), new XAttribute("cy", heightEmus))),
                                        new XElement(OoxmlNs.A + "prstGeom", new XAttribute("prst", "rect"), new XElement(OoxmlNs.A + "avLst"))))))))));
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetImageAlt(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? alt = ReadRequiredField(operation, "alt", diagnostics);
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!TryParseMainImageTarget(target!, out int imageOrdinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-image-alt target '{target}'. Expected an image ID such as M.I0001.", operation, target)];
        }

        ImageBlipTarget? imageTarget = FindMainImageBlipTarget(package, imageOrdinal, cancellationToken);
        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        XElement? drawing = imageTarget.Blip.Ancestors(OoxmlNs.W + "drawing").FirstOrDefault();
        XElement? container = drawing?.Descendants(OoxmlNs.Wp + "inline").FirstOrDefault()
            ?? drawing?.Descendants(OoxmlNs.Wp + "anchor").FirstOrDefault();
        if (container is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E5205", $"Image '{target}' does not have editable DrawingML properties.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        XElement? docPr = container.Element(OoxmlNs.Wp + "docPr");
        if (docPr is null)
        {
            docPr = new XElement(OoxmlNs.Wp + "docPr");
            docPr.SetAttributeValue("id", "1");
            docPr.SetAttributeValue("name", target);
            container.AddFirst(docPr);
        }

        docPr.SetAttributeValue("descr", alt);
        SaveMainDocument(package, imageTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteDeleteImage(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!TryParseMainImageTarget(target!, out int imageOrdinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported delete-image target '{target}'. Expected an image ID such as M.I0001.", operation, target)];
        }

        ImageBlipTarget? imageTarget = FindMainImageBlipTarget(package, imageOrdinal, cancellationToken);
        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        XElement? drawing = imageTarget.Blip.Ancestors(OoxmlNs.W + "drawing").FirstOrDefault();
        if (drawing is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E5205", $"Image '{target}' does not have an editable DrawingML object.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        drawing.Remove();
        SaveMainDocument(package, imageTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetSectionColumns(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? countText = ReadRequiredField(operation, "count", diagnostics);
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!int.TryParse(countText, out int count) || count is < 1 or > 4)
        {
            return [Diagnostic(DocxSeverity.Error, "E6201", "Section column count must be between 1 and 4.", operation, target)];
        }

        SectionTarget? sectionTarget = ResolveMainSectionTarget(package, target!, cancellationToken);
        if (sectionTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        XElement? columns = sectionTarget.SectionProperties.Element(OoxmlNs.W + "cols");
        if (columns is null)
        {
            columns = new XElement(OoxmlNs.W + "cols");
            sectionTarget.SectionProperties.Add(columns);
        }

        columns.SetAttributeValue(OoxmlNs.W + "num", count.ToString());
        SaveMainDocument(package, sectionTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetSectionOrientation(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? orientation = ReadRequiredField(operation, "orientation", diagnostics);
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (orientation is not ("portrait" or "landscape"))
        {
            return [Diagnostic(DocxSeverity.Error, "E6202", "Section orientation must be portrait or landscape.", operation, target)];
        }

        SectionTarget? sectionTarget = ResolveMainSectionTarget(package, target!, cancellationToken);
        if (sectionTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        XElement? pageSize = sectionTarget.SectionProperties.Element(OoxmlNs.W + "pgSz");
        if (pageSize is null)
        {
            pageSize = new XElement(OoxmlNs.W + "pgSz");
            sectionTarget.SectionProperties.AddFirst(pageSize);
        }

        string? currentOrientation = (string?)pageSize.Attribute(OoxmlNs.W + "orient") ?? "portrait";
        if (!string.Equals(currentOrientation, orientation, StringComparison.Ordinal) &&
            pageSize.Attribute(OoxmlNs.W + "w") is XAttribute width &&
            pageSize.Attribute(OoxmlNs.W + "h") is XAttribute height)
        {
            (width.Value, height.Value) = (height.Value, width.Value);
        }

        pageSize.SetAttributeValue(OoxmlNs.W + "orient", orientation);
        SaveMainDocument(package, sectionTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetCell(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        bool force = ReadBooleanField(operation, "force", diagnostics) ?? false;
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!TryParseMainCellTarget(target!, out int tableOrdinal, out int rowOrdinal, out int cellOrdinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-cell target '{target}'. Expected a table cell ID such as M.T0001.R02.C03.", operation, target)];
        }

        XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
        XElement? table = FindTable(body, tableOrdinal);
        XElement? row = table?.Elements(OoxmlNs.W + "tr").ElementAtOrDefault(rowOrdinal - 1);
        XElement? cell = row?.Elements(OoxmlNs.W + "tc").ElementAtOrDefault(cellOrdinal - 1);
        if (cell is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        string current = ReadVisibleText(cell);
        if (expected is not null && !string.Equals(current, expected, StringComparison.Ordinal))
        {
            return
            [
                Diagnostic(
                    DocxSeverity.Error,
                    "E3201",
                    $"Guard failed for {target}. Expected text does not match current text.",
                    operation,
                    target)
            ];
        }

        if (IsVerticalMergeContinuation(cell))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Unsupported merged-cell target '{target}'.", operation, target)];
        }

        if (!force && !IsSimpleEditableCell(cell))
        {
            return [Diagnostic(DocxSeverity.Error, "E4302", $"Cell '{target}' contains unsupported content. Use force true only when replacing all cell content is intended.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        ReplaceCellText(cell, text!);
        SaveMainDocument(package, document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteAppendRow(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!TryParseMainTableTarget(target!, out int tableOrdinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported append-row target '{target}'. Expected a table ID such as M.T0001.", operation, target)];
        }

        string[] cellTexts = operation.FieldValues
            .Where(field => field.Name == "cell")
            .Select(field => field.Value)
            .ToArray();
        if (cellTexts.Length == 0)
        {
            return [Diagnostic(DocxSeverity.Error, "E4202", "Operation 'append-row' is missing required field 'cell'.", operation, target)];
        }

        XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
        XElement? table = FindTable(body, tableOrdinal);
        if (table is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        XElement[] rows = table.Elements(OoxmlNs.W + "tr").ToArray();
        if (rows.Length == 0)
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' has no rows to clone.", operation, target)];
        }

        if (ContainsVerticalMerges(table))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' contains vertical merges that are not supported by append-row.", operation, target)];
        }

        if (!IsRectangular(table, out int columnCount))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' is not rectangular and cannot be appended safely.", operation, target)];
        }

        if (cellTexts.Length != columnCount)
        {
            return [Diagnostic(DocxSeverity.Error, "E4303", $"append-row expected {columnCount} cell field(s), but received {cellTexts.Length}.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        XElement newRow = CreateRowFromTemplate(rows[^1], cellTexts);
        rows[^1].AddAfterSelf(newRow);
        SaveMainDocument(package, document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteInsertRow(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool insertAfter,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        bool force = ReadBooleanField(operation, "force", diagnostics) ?? false;
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!TryParseMainRowTarget(target!, out int tableOrdinal, out int rowOrdinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported {operation.OperationName} target '{target}'. Expected a table row ID such as M.T0001.R02.", operation, target)];
        }

        string[] cellTexts = operation.FieldValues
            .Where(field => field.Name == "cell")
            .Select(field => field.Value)
            .ToArray();
        if (cellTexts.Length == 0)
        {
            return [Diagnostic(DocxSeverity.Error, "E4202", $"Operation '{operation.OperationName}' is missing required field 'cell'.", operation, target)];
        }

        XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
        XElement? table = FindTable(body, tableOrdinal);
        XElement[] rows = table?.Elements(OoxmlNs.W + "tr").ToArray() ?? [];
        XElement? templateRow = rowOrdinal < 1 || rowOrdinal > rows.Length ? null : rows[rowOrdinal - 1];
        if (templateRow is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (ContainsVerticalMerges(table!))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' contains vertical merges that are not supported by {operation.OperationName}.", operation, target)];
        }

        int expectedCellCount;
        if (IsRectangular(table!, out int columnCount))
        {
            expectedCellCount = columnCount;
        }
        else
        {
            if (!force)
            {
                return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' is not rectangular and cannot be edited safely without force true.", operation, target)];
            }

            expectedCellCount = templateRow.Elements(OoxmlNs.W + "tc").Count();
        }

        if (expectedCellCount == 0)
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Row '{target}' has no cells to clone.", operation, target)];
        }

        if (cellTexts.Length != expectedCellCount)
        {
            return [Diagnostic(DocxSeverity.Error, "E4303", $"{operation.OperationName} expected {expectedCellCount} cell field(s), but received {cellTexts.Length}.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        XElement newRow = CreateRowFromTemplate(templateRow, cellTexts);
        if (insertAfter)
        {
            templateRow.AddAfterSelf(newRow);
        }
        else
        {
            templateRow.AddBeforeSelf(newRow);
        }

        SaveMainDocument(package, document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteDeleteRow(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        bool force = ReadBooleanField(operation, "force", diagnostics) ?? false;
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!TryParseMainRowTarget(target!, out int tableOrdinal, out int rowOrdinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported delete-row target '{target}'. Expected a table row ID such as M.T0001.R02.", operation, target)];
        }

        XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
        XElement? table = FindTable(body, tableOrdinal);
        XElement[] rows = table?.Elements(OoxmlNs.W + "tr").ToArray() ?? [];
        XElement? row = rowOrdinal < 1 || rowOrdinal > rows.Length ? null : rows[rowOrdinal - 1];
        if (row is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (rows.Length == 1)
        {
            return [Diagnostic(DocxSeverity.Error, "E4304", $"Cannot delete the last row of table '{target}'.", operation, target)];
        }

        if (ContainsVerticalMerges(table!))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' contains vertical merges that are not supported by delete-row.", operation, target)];
        }

        if (!force && !IsRectangular(table!, out _))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' is not rectangular and cannot be edited safely without force true.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        row.Remove();
        SaveMainDocument(package, document);
        return [];
    }

    private static string? ReadRequiredField(
        DocxPatchOperation operation,
        string fieldName,
        List<DocxDiagnostic> diagnostics)
    {
        if (operation.Fields.TryGetValue(fieldName, out string? value) && value.Length != 0)
        {
            return value;
        }

        diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4202", $"Operation '{operation.OperationName}' is missing required field '{fieldName}'.", operation));
        return null;
    }

    private static bool TryParseMainParagraphTarget(string target, out int paragraphOrdinal)
    {
        paragraphOrdinal = 0;
        return target.Length == 7 &&
            target.StartsWith("M.P", StringComparison.Ordinal) &&
            int.TryParse(target[3..], out paragraphOrdinal);
    }

    private static bool TryParseMainTableTarget(string target, out int tableOrdinal)
    {
        tableOrdinal = 0;
        return target.Length == 7 &&
            target.StartsWith("M.T", StringComparison.Ordinal) &&
            int.TryParse(target[3..], out tableOrdinal);
    }

    private static bool TryParseMainCellTarget(string target, out int tableOrdinal, out int rowOrdinal, out int cellOrdinal)
    {
        tableOrdinal = 0;
        rowOrdinal = 0;
        cellOrdinal = 0;
        if (target.Length != 15 ||
            !target.StartsWith("M.T", StringComparison.Ordinal) ||
            target[7..9] != ".R" ||
            target[11..13] != ".C")
        {
            return false;
        }

        return int.TryParse(target[3..7], out tableOrdinal) &&
            int.TryParse(target[9..11], out rowOrdinal) &&
            int.TryParse(target[13..15], out cellOrdinal);
    }

    private static bool TryParseMainRowTarget(string target, out int tableOrdinal, out int rowOrdinal)
    {
        tableOrdinal = 0;
        rowOrdinal = 0;
        if (target.Length != 11 ||
            !target.StartsWith("M.T", StringComparison.Ordinal) ||
            target[7..9] != ".R")
        {
            return false;
        }

        return int.TryParse(target[3..7], out tableOrdinal) &&
            int.TryParse(target[9..11], out rowOrdinal);
    }

    private static bool TryParseMainImageTarget(string target, out int imageOrdinal)
    {
        imageOrdinal = 0;
        return target.Length == 7 &&
            target.StartsWith("M.I", StringComparison.Ordinal) &&
            int.TryParse(target[3..], out imageOrdinal);
    }

    private static bool TryParseMainSectionTarget(string target, out int sectionOrdinal)
    {
        sectionOrdinal = 0;
        return target.Length == 7 &&
            target.StartsWith("M.S", StringComparison.Ordinal) &&
            int.TryParse(target[3..], out sectionOrdinal);
    }

    private static XElement? FindTable(XElement body, int tableOrdinal)
    {
        return tableOrdinal < 1
            ? null
            : body.Elements(OoxmlNs.W + "tbl").ElementAtOrDefault(tableOrdinal - 1);
    }

    private static XElement? FindParagraph(XElement body, int paragraphOrdinal)
    {
        return paragraphOrdinal < 1
            ? null
            : body.Elements(OoxmlNs.W + "p").ElementAtOrDefault(paragraphOrdinal - 1);
    }

    private static XElement? ResolveMainBlock(
        XElement body,
        DocxPatchOperation operation,
        string target,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        diagnostics = [];
        if (!TryParseTargetSelector(target, operation, out TargetSelector? selector, out DocxDiagnostic? diagnostic))
        {
            diagnostics = [diagnostic!];
            return null;
        }

        if (selector is ExplicitIdTargetSelector)
        {
            if (TryParseMainParagraphTarget(target, out int paragraphOrdinal))
            {
                return FindParagraph(body, paragraphOrdinal);
            }

            if (TryParseMainTableTarget(target, out int tableOrdinal))
            {
                return FindTable(body, tableOrdinal);
            }

            return null;
        }

        return ResolveMainParagraphElementBySelector(body, selector!, operation, out diagnostics);
    }

    private static ParagraphTarget? ResolveParagraphTarget(
        OoxmlPackage package,
        DocxPatchOperation operation,
        string target,
        CancellationToken cancellationToken,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        diagnostics = [];
        if (!TryParseTargetSelector(target, operation, out TargetSelector? selector, out DocxDiagnostic? diagnostic))
        {
            diagnostics = [diagnostic!];
            return null;
        }

        if (selector is ExplicitIdTargetSelector)
        {
            if (TryParseMainParagraphTarget(target, out int mainParagraphOrdinal))
            {
                XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
                XElement? paragraph = FindParagraph(body, mainParagraphOrdinal);
                return paragraph is null || package.MainDocumentPartName is null
                    ? null
                    : new ParagraphTarget(package.MainDocumentPartName, document, paragraph);
            }

            if (TryParseStoryParagraphTarget(target, 'H', out int headerOrdinal, out int headerParagraphOrdinal))
            {
                return ResolveRelatedStoryParagraphTarget(package, OoxmlRelTypes.Header, headerOrdinal, headerParagraphOrdinal, cancellationToken);
            }

            if (TryParseStoryParagraphTarget(target, 'F', out int footerOrdinal, out int footerParagraphOrdinal))
            {
                return ResolveRelatedStoryParagraphTarget(package, OoxmlRelTypes.Footer, footerOrdinal, footerParagraphOrdinal, cancellationToken);
            }

            return null;
        }

        if (package.MainDocumentPartName is null)
        {
            return null;
        }

        XDocument mainDocument = LoadMainDocument(package, cancellationToken, out XElement mainBody);
        XElement? selectedParagraph = ResolveMainParagraphElementBySelector(mainBody, selector!, operation, out diagnostics);
        return selectedParagraph is null
            ? null
            : new ParagraphTarget(package.MainDocumentPartName, mainDocument, selectedParagraph);
    }

    private static bool TryParseTargetSelector(
        string target,
        DocxPatchOperation operation,
        out TargetSelector? selector,
        out DocxDiagnostic? diagnostic)
    {
        selector = null;
        diagnostic = null;
        if (target.StartsWith("heading:", StringComparison.Ordinal))
        {
            string value = target["heading:".Length..].Trim();
            int? level = null;
            int separator = value.IndexOf(':');
            if (separator > 0 && value[..separator].All(char.IsDigit))
            {
                if (!int.TryParse(value[..separator], out int parsedLevel) || parsedLevel is < 1 or > 9)
                {
                    diagnostic = Diagnostic(DocxSeverity.Error, "E1203", $"Invalid heading selector '{target}'. Heading level must be between 1 and 9.", operation, target);
                    return false;
                }

                level = parsedLevel;
                value = value[(separator + 1)..].Trim();
            }

            if (!TryReadSelectorText(value, out string selectorText))
            {
                diagnostic = Diagnostic(DocxSeverity.Error, "E1203", $"Invalid heading selector '{target}'. Expected heading:\"Text\" or heading:2:\"Text\".", operation, target);
                return false;
            }

            selector = new HeadingTargetSelector(target, level, selectorText);
            return true;
        }

        if (target.StartsWith("text:", StringComparison.Ordinal))
        {
            string value = target["text:".Length..].Trim();
            if (!TryReadSelectorText(value, out string selectorText))
            {
                diagnostic = Diagnostic(DocxSeverity.Error, "E1203", $"Invalid text selector '{target}'. Expected text:\"Text\".", operation, target);
                return false;
            }

            selector = new ParagraphTextTargetSelector(target, selectorText);
            return true;
        }

        selector = new ExplicitIdTargetSelector(target);
        return true;
    }

    private static bool TryReadSelectorText(string value, out string selectorText)
    {
        selectorText = string.Empty;
        if (value.Length >= 2 && value[0] == '"' && value[^1] == '"')
        {
            selectorText = value[1..^1];
            return selectorText.Length != 0;
        }

        if (value.StartsWith('"') || value.EndsWith('"'))
        {
            return false;
        }

        selectorText = value;
        return selectorText.Length != 0;
    }

    private static XElement? ResolveMainParagraphElementBySelector(
        XElement body,
        TargetSelector selector,
        DocxPatchOperation operation,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        diagnostics = [];
        if (selector is HeadingTargetSelector headingSelector)
        {
            return ResolveMainParagraphElementByPredicate(
                body,
                paragraph =>
                {
                    int? headingLevel = ReadHeadingLevel(paragraph);
                    return headingLevel is not null &&
                        (headingSelector.Level is null || headingSelector.Level == headingLevel) &&
                        string.Equals(ReadVisibleText(paragraph), headingSelector.Text, StringComparison.Ordinal);
                },
                selector.Raw,
                operation,
                out diagnostics);
        }

        if (selector is ParagraphTextTargetSelector paragraphTextSelector)
        {
            return ResolveMainParagraphElementByPredicate(
                body,
                paragraph => ReadVisibleText(paragraph).Contains(paragraphTextSelector.Text, StringComparison.Ordinal),
                selector.Raw,
                operation,
                out diagnostics);
        }

        return null;
    }

    private static XElement? ResolveMainParagraphElementByPredicate(
        XElement body,
        Func<XElement, bool> predicate,
        string rawSelector,
        DocxPatchOperation operation,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        diagnostics = [];
        var matches = new List<ParagraphSelectorMatch>();
        int paragraphOrdinal = 0;
        foreach (XElement paragraph in body.Elements(OoxmlNs.W + "p"))
        {
            paragraphOrdinal++;
            if (predicate(paragraph))
            {
                matches.Add(new ParagraphSelectorMatch($"M.P{paragraphOrdinal:0000}", paragraph));
            }
        }

        if (matches.Count > 1)
        {
            diagnostics =
            [
                Diagnostic(
                    DocxSeverity.Error,
                    "E1202",
                    $"Selector matched {matches.Count} targets: {string.Join(", ", matches.Select(match => match.Id))}. Use a more specific selector or an explicit ID.",
                    operation,
                    rawSelector)
            ];
            return null;
        }

        return matches.Count == 0 ? null : matches[0].Paragraph;
    }

    private static int? ReadHeadingLevel(XElement paragraph)
    {
        string? styleId = (string?)paragraph
            .Element(OoxmlNs.W + "pPr")
            ?.Element(OoxmlNs.W + "pStyle")
            ?.Attribute(OoxmlNs.W + "val");
        if (styleId is null)
        {
            return null;
        }

        string digits = new(styleId.Where(char.IsDigit).ToArray());
        return int.TryParse(digits, out int level) && level is >= 1 and <= 9
            ? level
            : null;
    }

    private static SectionTarget? ResolveMainSectionTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (!TryParseMainSectionTarget(target, out int sectionOrdinal) || sectionOrdinal < 1)
        {
            return null;
        }

        XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
        int currentOrdinal = 0;
        foreach (XElement element in body.Descendants(OoxmlNs.W + "sectPr"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            currentOrdinal++;
            if (currentOrdinal == sectionOrdinal)
            {
                return new SectionTarget(document, element);
            }
        }

        return null;
    }

    private static ParagraphTarget? ResolveRelatedStoryParagraphTarget(
        OoxmlPackage package,
        string relationshipType,
        int storyOrdinal,
        int paragraphOrdinal,
        CancellationToken cancellationToken)
    {
        if (package.MainDocumentPartName is null || storyOrdinal < 1)
        {
            return null;
        }

        OoxmlRelationship? relationship = package
            .GetRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => !relationship.IsExternal && relationship.Type == relationshipType && relationship.ResolvedTarget is not null)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .ElementAtOrDefault(storyOrdinal - 1);
        if (relationship?.ResolvedTarget is null)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, relationship.ResolvedTarget, cancellationToken, out XElement root);
        XElement? paragraph = FindParagraph(root, paragraphOrdinal);
        return paragraph is null ? null : new ParagraphTarget(relationship.ResolvedTarget, document, paragraph);
    }

    private static bool TryParseStoryParagraphTarget(
        string target,
        char storyPrefix,
        out int storyOrdinal,
        out int paragraphOrdinal)
    {
        storyOrdinal = 0;
        paragraphOrdinal = 0;
        if (target.Length != 10 ||
            target[0] != storyPrefix ||
            target[4..6] != ".P")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[6..], out paragraphOrdinal);
    }

    private static ImageTarget? FindMainImageTarget(
        OoxmlPackage package,
        int imageOrdinal,
        CancellationToken cancellationToken)
    {
        ImageBlipTarget? target = FindMainImageBlipTarget(package, imageOrdinal, cancellationToken);
        return target is null ? null : new ImageTarget(target.RelationshipId, target.Part);
    }

    private static ImageBlipTarget? FindMainImageBlipTarget(
        OoxmlPackage package,
        int imageOrdinal,
        CancellationToken cancellationToken)
    {
        if (imageOrdinal < 1 || package.MainDocumentPartName is null)
        {
            return null;
        }

        XDocument document = LoadMainDocument(package, cancellationToken, out _);
        IReadOnlyDictionary<string, OoxmlRelationship> relationships = package
            .GetRelationships(package.MainDocumentPartName, cancellationToken)
            .ToDictionary(relationship => relationship.Id, StringComparer.Ordinal);
        var seenParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int currentOrdinal = 0;
        foreach (XElement blip in document.Descendants(OoxmlNs.A + "blip"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? relationshipId = (string?)blip.Attribute(OoxmlNs.R + "embed");
            if (relationshipId is null ||
                !relationships.TryGetValue(relationshipId, out OoxmlRelationship? relationship) ||
                relationship.IsExternal ||
                relationship.ResolvedTarget is null)
            {
                continue;
            }

            OoxmlPart? part = package.GetPart(relationship.ResolvedTarget);
            if (part is null || !seenParts.Add(part.Name))
            {
                continue;
            }

            currentOrdinal++;
            if (currentOrdinal == imageOrdinal)
            {
                return new ImageBlipTarget(document, blip, relationshipId, part);
            }
        }

        return null;
    }

    private static bool TryReadAsset(
        IDocxAssetProvider? assetProvider,
        string asset,
        CancellationToken cancellationToken,
        out byte[] bytes,
        out string? contentType,
        out DocxDiagnostic? diagnostic,
        DocxPatchOperation operation,
        string? target)
    {
        bytes = [];
        contentType = null;
        diagnostic = null;
        if (assetProvider is null)
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E5201", "No asset provider is configured for image operation assets.", operation, target);
            return false;
        }

        if (!assetProvider.TryOpen(asset, out Stream stream, out string? contentTypeHint, out string? fileNameHint))
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E5202", $"Asset '{asset}' could not be resolved.", operation, target);
            return false;
        }

        using (stream)
        using (var memory = new MemoryStream())
        {
            CopyTo(stream, memory, cancellationToken);
            bytes = memory.ToArray();
        }

        contentType = DetectImageContentType(bytes, contentTypeHint, fileNameHint ?? asset);
        if (contentType is null)
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E5203", $"Asset '{asset}' is not a supported PNG or JPEG image.", operation, target);
            return false;
        }

        return true;
    }

    private static bool TryResolveStyleId(
        OoxmlPackage package,
        string requestedStyle,
        string styleType,
        CancellationToken cancellationToken,
        out string? styleId,
        out DocxDiagnostic? diagnostic,
        DocxPatchOperation operation,
        string? target)
    {
        styleId = null;
        diagnostic = null;
        IReadOnlyList<DocxStyleInfo> styles = DocxStyleScanner.Scan(package, cancellationToken)
            .Where(style => style.Type == styleType)
            .ToArray();
        DocxStyleInfo? byId = styles.FirstOrDefault(style => string.Equals(style.StyleId, requestedStyle, StringComparison.Ordinal));
        if (byId is not null)
        {
            styleId = byId.StyleId;
            return true;
        }

        DocxStyleInfo[] byName = styles
            .Where(style => string.Equals(style.Name, requestedStyle, StringComparison.Ordinal))
            .ToArray();
        if (byName.Length == 1)
        {
            styleId = byName[0].StyleId;
            return true;
        }

        diagnostic = byName.Length > 1
            ? Diagnostic(DocxSeverity.Error, "E7102", $"Style name '{requestedStyle}' is ambiguous.", operation, target)
            : Diagnostic(DocxSeverity.Error, "E7101", $"Style '{requestedStyle}' was not found.", operation, target);
        return false;
    }

    private static string? DetectImageContentType(byte[] bytes, string? contentTypeHint, string? fileNameHint)
    {
        string? normalizedHint = NormalizeImageContentType(contentTypeHint);
        if (normalizedHint is not null)
        {
            return normalizedHint;
        }

        string extension = Path.GetExtension(fileNameHint ?? string.Empty).ToLowerInvariant();
        if (extension is ".png")
        {
            return "image/png";
        }

        if (extension is ".jpg" or ".jpeg")
        {
            return "image/jpeg";
        }

        if (bytes.Length >= 8 &&
            bytes[0] == 0x89 &&
            bytes[1] == 0x50 &&
            bytes[2] == 0x4E &&
            bytes[3] == 0x47 &&
            bytes[4] == 0x0D &&
            bytes[5] == 0x0A &&
            bytes[6] == 0x1A &&
            bytes[7] == 0x0A)
        {
            return "image/png";
        }

        if (bytes.Length >= 3 &&
            bytes[0] == 0xFF &&
            bytes[1] == 0xD8 &&
            bytes[2] == 0xFF)
        {
            return "image/jpeg";
        }

        return null;
    }

    private static string? NormalizeImageContentType(string? contentType)
    {
        return contentType?.ToLowerInvariant() switch
        {
            "image/png" => "image/png",
            "image/jpeg" => "image/jpeg",
            "image/jpg" => "image/jpeg",
            _ => null
        };
    }

    private static void CopyTo(Stream source, Stream destination, CancellationToken cancellationToken)
    {
        byte[] buffer = new byte[81920];
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int read = source.Read(buffer, 0, buffer.Length);
            if (read == 0)
            {
                return;
            }

            destination.Write(buffer, 0, read);
        }
    }

    private static string ReadVisibleText(XElement container)
    {
        var builder = new StringBuilder();
        foreach (XElement element in container.Descendants())
        {
            if (element.Ancestors(OoxmlNs.W + "del").Any() ||
                element.Ancestors(OoxmlNs.W + "moveFrom").Any())
            {
                continue;
            }

            if (element.Name == OoxmlNs.W + "t")
            {
                builder.Append(element.Value);
            }
            else if (element.Name == OoxmlNs.W + "tab")
            {
                builder.Append('\t');
            }
            else if (element.Name == OoxmlNs.W + "br")
            {
                builder.Append('\n');
            }
        }

        return builder.ToString();
    }

    private static bool TryGetProtectedTextEditFeature(XElement paragraph, out string feature)
    {
        foreach (XElement element in paragraph.Descendants())
        {
            if (ProtectedTextEditElements.TryGetValue(element.Name, out feature!))
            {
                return true;
            }
        }

        feature = string.Empty;
        return false;
    }

    private static int? ReadPositiveOccurrence(DocxPatchOperation operation, List<DocxDiagnostic> diagnostics)
    {
        if (!operation.Fields.TryGetValue("occurrence", out string? value))
        {
            return null;
        }

        if (!int.TryParse(value, out int occurrence) || occurrence <= 0)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Field 'occurrence' must be greater than 0.", operation, operation.Fields.GetValueOrDefault("target")));
            return null;
        }

        return occurrence;
    }

    private static IReadOnlyList<TextRange> FindTextMatches(string text, string find, int? occurrence)
    {
        var matches = new List<TextRange>();
        int index = 0;
        int seen = 0;
        while (index <= text.Length)
        {
            int found = text.IndexOf(find, index, StringComparison.Ordinal);
            if (found < 0)
            {
                break;
            }

            seen++;
            if (occurrence is null || seen == occurrence)
            {
                matches.Add(new TextRange(found, find.Length));
                if (occurrence is not null)
                {
                    break;
                }
            }

            index = found + find.Length;
        }

        return matches;
    }

    private static string ApplyTextReplacement(string text, IReadOnlyList<TextRange> matches, string replacement)
    {
        var builder = new StringBuilder(text);
        for (int i = matches.Count - 1; i >= 0; i--)
        {
            TextRange match = matches[i];
            builder.Remove(match.Start, match.Length);
            builder.Insert(match.Start, replacement);
        }

        return builder.ToString();
    }

    private static bool TryReplaceParagraphTextPreservingRuns(
        XElement paragraph,
        IReadOnlyList<TextRange> matches,
        string replacement,
        out string? unsupportedReason)
    {
        unsupportedReason = null;
        if (replacement.Contains('\t') || replacement.Contains('\n'))
        {
            unsupportedReason = "replacement contains tabs or line breaks";
            return false;
        }

        TextPosition?[] positions = BuildTextPositions(paragraph);
        foreach (TextRange match in matches)
        {
            for (int i = match.Start; i < match.Start + match.Length; i++)
            {
                if (i < 0 || i >= positions.Length || positions[i] is null)
                {
                    unsupportedReason = "match includes tabs, line breaks, or non-text run content";
                    return false;
                }
            }
        }

        for (int i = matches.Count - 1; i >= 0; i--)
        {
            TextRange match = matches[i];
            TextPosition start = positions[match.Start]!;
            TextPosition end = positions[match.Start + match.Length - 1]!;
            if (ReferenceEquals(start.TextElement, end.TextElement))
            {
                string value = start.TextElement.Value;
                string edited = value.Remove(start.Offset, match.Length).Insert(start.Offset, replacement);
                SetTextElementValue(start.TextElement, edited);
                continue;
            }

            string startValue = start.TextElement.Value;
            string endValue = end.TextElement.Value;
            SetTextElementValue(start.TextElement, startValue[..start.Offset] + replacement);
            SetTextElementValue(end.TextElement, endValue[(end.Offset + 1)..]);

            bool insideRange = false;
            foreach (XElement textElement in paragraph.Elements(OoxmlNs.W + "r").Elements(OoxmlNs.W + "t"))
            {
                if (ReferenceEquals(textElement, start.TextElement))
                {
                    insideRange = true;
                    continue;
                }

                if (ReferenceEquals(textElement, end.TextElement))
                {
                    break;
                }

                if (insideRange)
                {
                    SetTextElementValue(textElement, string.Empty);
                }
            }
        }

        return true;
    }

    private static TextPosition?[] BuildTextPositions(XElement paragraph)
    {
        var positions = new List<TextPosition?>();
        foreach (XElement run in paragraph.Elements(OoxmlNs.W + "r"))
        {
            foreach (XElement child in run.Elements())
            {
                if (child.Name == OoxmlNs.W + "t")
                {
                    for (int i = 0; i < child.Value.Length; i++)
                    {
                        positions.Add(new TextPosition(child, i));
                    }
                }
                else if (child.Name == OoxmlNs.W + "tab" || child.Name == OoxmlNs.W + "br")
                {
                    positions.Add(null);
                }
            }
        }

        return positions.ToArray();
    }

    private static void SetTextElementValue(XElement textElement, string text)
    {
        textElement.Value = text;
        if (RequiresPreserveSpace(text))
        {
            textElement.SetAttributeValue(OoxmlNs.Xml + "space", "preserve");
        }
        else
        {
            textElement.SetAttributeValue(OoxmlNs.Xml + "space", null);
        }
    }

    private static void ReplaceParagraphText(XElement paragraph, string text)
    {
        XElement? paragraphProperties = paragraph.Element(OoxmlNs.W + "pPr");
        XElement? firstRunProperties = paragraph
            .Elements(OoxmlNs.W + "r")
            .Elements(OoxmlNs.W + "rPr")
            .FirstOrDefault();

        paragraph.RemoveNodes();
        if (paragraphProperties is not null)
        {
            paragraph.Add(new XElement(paragraphProperties));
        }

        var run = new XElement(OoxmlNs.W + "r");
        if (firstRunProperties is not null)
        {
            run.Add(new XElement(firstRunProperties));
        }

        var textElement = new XElement(OoxmlNs.W + "t", text);
        if (RequiresPreserveSpace(text))
        {
            textElement.SetAttributeValue(OoxmlNs.Xml + "space", "preserve");
        }

        run.Add(textElement);
        paragraph.Add(run);
    }

    private static void ReplaceCellText(XElement cell, string text)
    {
        XElement? cellProperties = cell.Element(OoxmlNs.W + "tcPr");
        cell.RemoveNodes();
        if (cellProperties is not null)
        {
            cell.Add(new XElement(cellProperties));
        }

        cell.Add(CreateSimpleParagraph(text));
    }

    private static XElement CreateRowFromTemplate(XElement templateRow, IReadOnlyList<string> cellTexts)
    {
        var row = new XElement(OoxmlNs.W + "tr");
        XElement? rowProperties = templateRow.Element(OoxmlNs.W + "trPr");
        if (rowProperties is not null)
        {
            row.Add(new XElement(rowProperties));
        }

        XElement[] templateCells = templateRow.Elements(OoxmlNs.W + "tc").ToArray();
        for (int i = 0; i < cellTexts.Count; i++)
        {
            var cell = new XElement(OoxmlNs.W + "tc");
            XElement? cellProperties = templateCells[i].Element(OoxmlNs.W + "tcPr");
            if (cellProperties is not null)
            {
                cell.Add(new XElement(cellProperties));
            }

            cell.Add(CreateSimpleParagraph(cellTexts[i]));
            row.Add(cell);
        }

        return row;
    }

    private static XElement CreateSimpleParagraph(string text, string? style = null)
    {
        var paragraph = new XElement(OoxmlNs.W + "p");
        if (style is not null)
        {
            SetParagraphStyle(paragraph, style);
        }

        var run = new XElement(OoxmlNs.W + "r");
        foreach (XNode node in CreateTextNodes(text))
        {
            run.Add(node);
        }

        paragraph.Add(run);
        return paragraph;
    }

    private static void SetParagraphStyle(XElement paragraph, string style)
    {
        XElement? paragraphProperties = paragraph.Element(OoxmlNs.W + "pPr");
        if (paragraphProperties is null)
        {
            paragraphProperties = new XElement(OoxmlNs.W + "pPr");
            paragraph.AddFirst(paragraphProperties);
        }

        XElement? paragraphStyle = paragraphProperties.Element(OoxmlNs.W + "pStyle");
        if (paragraphStyle is null)
        {
            paragraphStyle = new XElement(OoxmlNs.W + "pStyle");
            paragraphProperties.AddFirst(paragraphStyle);
        }

        paragraphStyle.SetAttributeValue(OoxmlNs.W + "val", style);
    }

    private static IEnumerable<XNode> CreateTextNodes(string text)
    {
        var buffer = new StringBuilder();
        foreach (char ch in text)
        {
            if (ch == '\t' || ch == '\n')
            {
                if (buffer.Length != 0)
                {
                    yield return CreateTextElement(buffer.ToString());
                    buffer.Clear();
                }

                yield return ch == '\t'
                    ? new XElement(OoxmlNs.W + "tab")
                    : new XElement(OoxmlNs.W + "br");
                continue;
            }

            buffer.Append(ch);
        }

        yield return CreateTextElement(buffer.ToString());
    }

    private static XElement CreateTextElement(string text)
    {
        var textElement = new XElement(OoxmlNs.W + "t", text);
        if (RequiresPreserveSpace(text))
        {
            textElement.SetAttributeValue(OoxmlNs.Xml + "space", "preserve");
        }

        return textElement;
    }

    private static bool IsSimpleEditableCell(XElement cell)
    {
        XElement[] paragraphs = cell.Elements(OoxmlNs.W + "p").ToArray();
        return paragraphs.Length == 1 &&
            !cell.Elements(OoxmlNs.W + "tbl").Any() &&
            !cell.Descendants(OoxmlNs.W + "drawing").Any() &&
            !cell.Descendants(OoxmlNs.W + "fldSimple").Any() &&
            !cell.Descendants(OoxmlNs.W + "fldChar").Any() &&
            !cell.Descendants(OoxmlNs.W + "instrText").Any();
    }

    private static bool ContainsVerticalMerges(XElement table)
    {
        return table.Descendants(OoxmlNs.W + "vMerge").Any();
    }

    private static bool IsRectangular(XElement table, out int columnCount)
    {
        XElement[] rows = table.Elements(OoxmlNs.W + "tr").ToArray();
        columnCount = rows.Length == 0 ? 0 : rows[0].Elements(OoxmlNs.W + "tc").Count();
        int expectedColumnCount = columnCount;
        return expectedColumnCount != 0 && rows.All(row => row.Elements(OoxmlNs.W + "tc").Count() == expectedColumnCount);
    }

    private static bool IsVerticalMergeContinuation(XElement cell)
    {
        XElement? verticalMerge = cell.Element(OoxmlNs.W + "tcPr")?.Element(OoxmlNs.W + "vMerge");
        if (verticalMerge is null)
        {
            return false;
        }

        string? value = (string?)verticalMerge.Attribute(OoxmlNs.W + "val");
        return !string.Equals(value, "restart", StringComparison.OrdinalIgnoreCase);
    }

    private static bool? ReadBooleanField(DocxPatchOperation operation, string fieldName, List<DocxDiagnostic> diagnostics)
    {
        if (!operation.Fields.TryGetValue(fieldName, out string? value))
        {
            return null;
        }

        if (string.Equals(value, "true", StringComparison.Ordinal))
        {
            return true;
        }

        if (string.Equals(value, "false", StringComparison.Ordinal))
        {
            return false;
        }

        diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4204", $"Field '{fieldName}' must be true or false.", operation));
        return null;
    }

    private static IReadOnlyList<DocxDiagnostic> MarkFieldsDirty(OoxmlPackage package, CancellationToken cancellationToken)
    {
        if (package.MainDocumentPartName is null)
        {
            return [new DocxDiagnostic(DocxSeverity.Error, "E9001", "Post-edit validation failed: main document part is missing.")];
        }

        string settingsPartName = ResolveOrCreateSettingsPart(package, cancellationToken);
        OoxmlPart settingsPart = package.GetPart(settingsPartName)
            ?? throw new InvalidDataException($"Settings part '{settingsPartName}' does not exist.");
        using Stream stream = settingsPart.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        XElement settings = document.Root
            ?? throw new InvalidDataException($"Settings part '{settingsPartName}' has no XML root.");
        if (settings.Name != OoxmlNs.W + "settings")
        {
            return [new DocxDiagnostic(DocxSeverity.Error, "E9001", $"Post-edit validation failed for {settingsPartName}: Expected root element 'settings', found '{settings.Name.LocalName}'.", PartName: settingsPartName)];
        }

        XElement? updateFields = settings.Element(OoxmlNs.W + "updateFields");
        if (updateFields is null)
        {
            updateFields = new XElement(OoxmlNs.W + "updateFields");
            settings.Add(updateFields);
        }

        updateFields.SetAttributeValue(OoxmlNs.W + "val", "true");
        SaveDocumentPart(package, settingsPartName, document);
        return [];
    }

    private static string ResolveOrCreateSettingsPart(OoxmlPackage package, CancellationToken cancellationToken)
    {
        OoxmlRelationship? relationship = package
            .GetRelationships(package.MainDocumentPartName!, cancellationToken)
            .FirstOrDefault(relationship =>
                !relationship.IsExternal &&
                relationship.Type == OoxmlRelTypes.Settings &&
                relationship.ResolvedTarget is not null);
        if (relationship?.ResolvedTarget is not null)
        {
            return relationship.ResolvedTarget;
        }

        const string settingsPartName = "/word/settings.xml";
        if (package.GetPart(settingsPartName) is null)
        {
            byte[] bytes = Encoding.UTF8.GetBytes("""
                <?xml version="1.0" encoding="utf-8"?>
                <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />
                """);
            package.AddPart(settingsPartName, SettingsContentType, bytes);
        }

        string relationshipId = OoxmlIds.AllocateRelationshipId(package
            .GetRelationships(package.MainDocumentPartName!, cancellationToken)
            .Select(relationship => relationship.Id));
        package.AddRelationship(package.MainDocumentPartName!, relationshipId, OoxmlRelTypes.Settings, GetRelativeRelationshipTarget(package.MainDocumentPartName!, settingsPartName));
        return settingsPartName;
    }

    private static IReadOnlyList<DocxDiagnostic> ValidateEditedPackage(OoxmlPackage package, CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        foreach (string partName in package.TouchedPartNames.Order(StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                diagnostics.Add(PostEditValidationDiagnostic(partName, $"Touched part '{partName}' is missing."));
                continue;
            }

            if (!ShouldValidateTouchedPart(part))
            {
                continue;
            }

            try
            {
                using Stream stream = part.OpenRead();
                XDocument document = SafeXml.Load(stream, cancellationToken);
                ValidateTouchedPartRoot(package, part, document, diagnostics, cancellationToken);
            }
            catch (Exception ex) when (ex is InvalidDataException or IOException or XmlException)
            {
                diagnostics.Add(PostEditValidationDiagnostic(part.Name, ex.Message));
            }
        }

        return diagnostics;
    }

    private static bool ShouldValidateTouchedPart(OoxmlPart part)
    {
        return string.Equals(part.Name, "/[Content_Types].xml", StringComparison.OrdinalIgnoreCase) ||
            part.Name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) ||
            part.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(part.ContentType, OoxmlContentTypeNames.Xml, StringComparison.OrdinalIgnoreCase) ||
            part.ContentType?.EndsWith("+xml", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static void ValidateTouchedPartRoot(
        OoxmlPackage package,
        OoxmlPart part,
        XDocument document,
        List<DocxDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        XElement root = document.Root
            ?? throw new InvalidDataException("XML part has no root element.");

        if (string.Equals(part.Name, "/[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
        {
            RequireRoot(part, root, OoxmlNs.Ct + "Types", diagnostics);
            return;
        }

        if (part.Name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
        {
            RequireRoot(part, root, OoxmlNs.Rel + "Relationships", diagnostics);
            ValidateRelationshipTargets(package, part, diagnostics, cancellationToken);
            return;
        }

        if (string.Equals(part.Name, package.MainDocumentPartName, StringComparison.OrdinalIgnoreCase))
        {
            RequireRoot(part, root, OoxmlNs.W + "document", diagnostics);
            if (root.Element(OoxmlNs.W + "body") is null)
            {
                diagnostics.Add(PostEditValidationDiagnostic(part.Name, "Main document part is missing w:body."));
            }

            return;
        }

        if (string.Equals(part.ContentType, "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml", StringComparison.OrdinalIgnoreCase))
        {
            RequireRoot(part, root, OoxmlNs.W + "hdr", diagnostics);
            return;
        }

        if (string.Equals(part.ContentType, "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml", StringComparison.OrdinalIgnoreCase))
        {
            RequireRoot(part, root, OoxmlNs.W + "ftr", diagnostics);
            return;
        }

        if (string.Equals(part.ContentType, SettingsContentType, StringComparison.OrdinalIgnoreCase))
        {
            RequireRoot(part, root, OoxmlNs.W + "settings", diagnostics);
        }
    }

    private static void RequireRoot(OoxmlPart part, XElement root, XName expectedRoot, List<DocxDiagnostic> diagnostics)
    {
        if (root.Name != expectedRoot)
        {
            diagnostics.Add(PostEditValidationDiagnostic(part.Name, $"Expected root element '{expectedRoot.LocalName}', found '{root.Name.LocalName}'."));
        }
    }

    private static void ValidateRelationshipTargets(
        OoxmlPackage package,
        OoxmlPart relationshipPart,
        List<DocxDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        string sourcePartName = GetSourcePartNameFromRelationshipPartName(relationshipPart.Name);
        using Stream stream = relationshipPart.OpenRead();
        foreach (OoxmlRelationship relationship in OoxmlPackage.ParseRelationships(stream, sourcePartName, cancellationToken))
        {
            if (!relationship.IsExternal &&
                relationship.ResolvedTarget is not null &&
                package.GetPart(relationship.ResolvedTarget) is null)
            {
                diagnostics.Add(PostEditValidationDiagnostic(relationshipPart.Name, $"Relationship '{relationship.Id}' targets missing part '{relationship.ResolvedTarget}'."));
            }
        }
    }

    private static string GetSourcePartNameFromRelationshipPartName(string relationshipPartName)
    {
        string normalized = OoxmlPath.NormalizePartName(relationshipPartName);
        if (normalized == "/_rels/.rels")
        {
            return "/";
        }

        const string relationshipMarker = "/_rels/";
        int markerIndex = normalized.LastIndexOf(relationshipMarker, StringComparison.Ordinal);
        if (markerIndex < 0 || !normalized.EndsWith(".rels", StringComparison.Ordinal))
        {
            throw new InvalidDataException($"Invalid relationship part name '{relationshipPartName}'.");
        }

        string directory = normalized[..markerIndex];
        string fileName = normalized[(markerIndex + relationshipMarker.Length)..^".rels".Length];
        return OoxmlPath.NormalizePartName($"{directory}/{fileName}");
    }

    private static DocxDiagnostic PostEditValidationDiagnostic(string partName, string message)
    {
        return new DocxDiagnostic(DocxSeverity.Error, "E9001", $"Post-edit validation failed for {partName}: {message}", PartName: partName);
    }

    private static XDocument LoadMainDocument(
        OoxmlPackage package,
        CancellationToken cancellationToken,
        out XElement body)
    {
        XDocument document = LoadDocumentPart(package, package.MainDocumentPartName!, cancellationToken, out XElement root);
        body = root.Element(OoxmlNs.W + "body")
            ?? throw new InvalidDataException("Main document part is missing w:body.");
        return document;
    }

    private static XDocument LoadDocumentPart(
        OoxmlPackage package,
        string partName,
        CancellationToken cancellationToken,
        out XElement root)
    {
        OoxmlPart documentPart = package.GetPart(partName)
            ?? throw new InvalidDataException($"Document part '{partName}' does not exist.");
        using Stream stream = documentPart.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        root = document.Root
            ?? throw new InvalidDataException($"Document part '{partName}' has no XML root.");
        return document;
    }

    private static void SaveMainDocument(OoxmlPackage package, XDocument document)
    {
        SaveDocumentPart(package, package.MainDocumentPartName!, document);
    }

    private static void SaveDocumentPart(OoxmlPackage package, string partName, XDocument document)
    {
        using var output = new MemoryStream();
        document.Save(output, SaveOptions.DisableFormatting);
        package.ReplacePartBytes(partName, output.ToArray());
    }

    private static bool RequiresPreserveSpace(string text)
    {
        return text.Length != 0 &&
            (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[^1]) || text.Contains("  ", StringComparison.Ordinal));
    }

    private static DocxDiagnostic Diagnostic(
        DocxSeverity severity,
        string code,
        string message,
        DocxPatchOperation operation,
        string? targetId = null)
    {
        return new DocxDiagnostic(severity, code, message, TargetId: targetId, OperationIndex: operation.Index);
    }
}

internal sealed record PatchExecutionResult(
    bool Success,
    IReadOnlyList<DocxDiagnostic> Diagnostics,
    IReadOnlyList<DocxPatchOperationReport> Reports);

internal sealed record ImageTarget(string RelationshipId, OoxmlPart Part);

internal sealed record ImageBlipTarget(XDocument Document, XElement Blip, string RelationshipId, OoxmlPart Part);

internal sealed record ParagraphTarget(string PartName, XDocument Document, XElement Paragraph);

internal sealed record SectionTarget(XDocument Document, XElement SectionProperties);

internal readonly record struct TextRange(int Start, int Length);

internal sealed record TextPosition(XElement TextElement, int Offset);

internal abstract record TargetSelector(string Raw);

internal sealed record ExplicitIdTargetSelector(string Raw) : TargetSelector(Raw);

internal sealed record HeadingTargetSelector(string Raw, int? Level, string Text) : TargetSelector(Raw);

internal sealed record ParagraphTextTargetSelector(string Raw, string Text) : TargetSelector(Raw);

internal sealed record ParagraphSelectorMatch(string Id, XElement Paragraph);
