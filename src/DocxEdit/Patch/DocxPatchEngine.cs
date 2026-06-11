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
            bool supportsTrackedChanges = SupportsTrackedChangeOutput(operation.OperationName);
            if (options.TrackChanges == TrackChangesMode.Require && !supportsTrackedChanges)
            {
                operationDiagnostics.Add(Diagnostic(DocxSeverity.Error, "E6001", $"TrackChangesMode.Require is not supported for operation '{operation.OperationName}'.", operation, operation.Fields.GetValueOrDefault("target")));
            }
            else
            {
                if (options.TrackChanges == TrackChangesMode.Suggest && !supportsTrackedChanges)
                {
                    operationDiagnostics.Add(Diagnostic(DocxSeverity.Warning, "W4001", $"TrackChangesMode.Suggest is not supported for operation '{operation.OperationName}'; applying the edit directly.", operation, operation.Fields.GetValueOrDefault("target")));
                }

                operationDiagnostics.AddRange(operation.OperationName switch
                {
                    "replace-text" => ExecuteReplaceText(package, operation, options, apply, cancellationToken),
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

    private static bool SupportsTrackedChangeOutput(string operationName)
    {
        return string.Equals(operationName, "replace-text", StringComparison.Ordinal);
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteReplaceText(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
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

        bool useTrackedChanges = options.TrackChanges is TrackChangesMode.Require or TrackChangesMode.Suggest;
        bool canUseTrackedChanges = true;
        string? trackedUnsupportedReason = null;
        if (useTrackedChanges)
        {
            canUseTrackedChanges = TryValidateTrackedTextReplacement(paragraphTarget.Paragraph, current, matches, replacement!, out trackedUnsupportedReason);
            if (!canUseTrackedChanges && options.TrackChanges == TrackChangesMode.Require)
            {
                return [Diagnostic(DocxSeverity.Error, "E6002", $"Tracked-change replacement is not supported for {target}: {trackedUnsupportedReason}.", operation, target)];
            }

            if (!canUseTrackedChanges)
            {
                diagnostics.Add(Diagnostic(DocxSeverity.Warning, "W4002", $"Tracked-change replacement is not supported for {target}: {trackedUnsupportedReason}; applying the edit directly.", operation, target));
            }
        }

        if (!apply)
        {
            return diagnostics;
        }

        if (useTrackedChanges && canUseTrackedChanges)
        {
            ReplaceParagraphTextWithTrackedChanges(package, paragraphTarget.Paragraph, current, matches, replacement!, options, cancellationToken);
            SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
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
        return diagnostics;
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

        BlockTarget? blockTarget = ResolveBlockTarget(package, operation, target!, cancellationToken, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
        if (selectorDiagnostics.Count != 0)
        {
            return selectorDiagnostics;
        }

        if (blockTarget is null)
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
            blockTarget.Block.AddAfterSelf(paragraph);
        }
        else
        {
            blockTarget.Block.AddBeforeSelf(paragraph);
        }

        SaveDocumentPart(package, blockTarget.PartName, blockTarget.Document);
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

        BlockTarget? blockTarget = ResolveBlockTarget(package, operation, target!, cancellationToken, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
        if (selectorDiagnostics.Count != 0)
        {
            return selectorDiagnostics;
        }

        if (blockTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (expected is not null)
        {
            string current = ReadVisibleText(blockTarget.Block);
            if (!string.Equals(current, expected, StringComparison.Ordinal))
            {
                return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected text does not match current text.", operation, target)];
            }
        }

        if (!apply)
        {
            return [];
        }

        blockTarget.Block.Remove();
        SaveDocumentPart(package, blockTarget.PartName, blockTarget.Document);
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
        bool hasAlt = operation.Fields.ContainsKey("alt");
        string? alt = operation.Fields.GetValueOrDefault("alt");
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target!, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported replace-image target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
        }

        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateImageContentTypeGuard(operation, target!, imageTarget.Part.ContentType, diagnostics))
        {
            return diagnostics;
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

        XElement? imageContainer = null;
        if (hasAlt &&
            !TryGetImageDrawingContainer(imageTarget, target!, operation, out imageContainer, out DocxDiagnostic? altDiagnostic))
        {
            return [altDiagnostic!];
        }

        if (!apply)
        {
            return [];
        }

        package.ReplacePartBytes(imageTarget.Part.Name, bytes);
        if (hasAlt)
        {
            SetImageAlt(imageContainer!, alt!, target!);
            SaveDocumentPart(package, imageTarget.PartName, imageTarget.Document);
        }

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

        if (!ValidateImageContentTypeGuard(operation, target!, contentType!, diagnostics))
        {
            return diagnostics;
        }

        if (!TryReadImageExtent(operation, bytes, contentType!, out long widthEmus, out long heightEmus, out DocxDiagnostic? dimensionDiagnostic))
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
        int docPrId = AllocateDrawingDocPrId(paragraphTarget.Document);
        XElement imageParagraph = CreateInlineImageParagraph(relationshipId, docPrId, widthEmus, heightEmus, operation.Fields.GetValueOrDefault("alt") ?? string.Empty);
        paragraphTarget.Paragraph.AddAfterSelf(imageParagraph);
        SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
        return [];
    }

    private static bool TryReadImageExtent(
        DocxPatchOperation operation,
        byte[] bytes,
        string contentType,
        out long widthEmus,
        out long heightEmus,
        out DocxDiagnostic? diagnostic)
    {
        widthEmus = OoxmlUnits.InchesToEmu(1);
        heightEmus = widthEmus;
        diagnostic = null;
        bool hasWidth = operation.Fields.TryGetValue("width", out string? width);
        bool hasHeight = operation.Fields.TryGetValue("height", out string? height);
        bool hasPixelSize = TryReadImagePixelSize(bytes, contentType, out int pixelWidth, out int pixelHeight);
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
            heightEmus = hasPixelSize && pixelWidth > 0
                ? checked((long)Math.Round(widthEmus * (pixelHeight / (double)pixelWidth), MidpointRounding.AwayFromZero))
                : widthEmus;
        }
        else if (!hasWidth && hasHeight)
        {
            widthEmus = hasPixelSize && pixelHeight > 0
                ? checked((long)Math.Round(heightEmus * (pixelWidth / (double)pixelHeight), MidpointRounding.AwayFromZero))
                : heightEmus;
        }
        else if (!hasWidth && !hasHeight && hasPixelSize)
        {
            widthEmus = OoxmlUnits.PixelsToEmu(pixelWidth);
            heightEmus = OoxmlUnits.PixelsToEmu(pixelHeight);
        }

        return true;
    }

    private static bool TryValidateTrackedTextReplacement(
        XElement paragraph,
        string current,
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

        foreach (TextRange match in matches)
        {
            string deletedText = current.Substring(match.Start, match.Length);
            if (deletedText.Contains('\t') || deletedText.Contains('\n'))
            {
                unsupportedReason = "matched text contains tabs or line breaks";
                return false;
            }
        }

        if (HasMixedDirectTextRunProperties(paragraph))
        {
            unsupportedReason = "paragraph contains mixed direct run formatting";
            return false;
        }

        return true;
    }

    private static bool HasMixedDirectTextRunProperties(XElement paragraph)
    {
        string? firstSignature = null;
        foreach (XElement run in paragraph.Elements(OoxmlNs.W + "r"))
        {
            if (!RunHasVisibleText(run))
            {
                continue;
            }

            string signature = run.Element(OoxmlNs.W + "rPr")?.ToString(SaveOptions.DisableFormatting) ?? string.Empty;
            if (firstSignature is null)
            {
                firstSignature = signature;
                continue;
            }

            if (!string.Equals(firstSignature, signature, StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    private static bool RunHasVisibleText(XElement run)
    {
        return run.Elements().Any(element =>
            element.Name == OoxmlNs.W + "t" ||
            element.Name == OoxmlNs.W + "tab" ||
            element.Name == OoxmlNs.W + "br");
    }

    private static void ReplaceParagraphTextWithTrackedChanges(
        OoxmlPackage package,
        XElement paragraph,
        string current,
        IReadOnlyList<TextRange> matches,
        string replacement,
        DocxEditOptions options,
        CancellationToken cancellationToken)
    {
        string[] revisionIds = AllocateRevisionIds(package, matches.Count * 2, cancellationToken);
        XElement? paragraphProperties = paragraph.Element(OoxmlNs.W + "pPr");
        XElement? firstRunProperties = paragraph
            .Elements(OoxmlNs.W + "r")
            .Elements(OoxmlNs.W + "rPr")
            .FirstOrDefault();
        string author = options.Author;
        string timestamp = options.TimestampUtc.ToUniversalTime().ToString("O");

        var nodes = new List<XNode>();
        if (paragraphProperties is not null)
        {
            nodes.Add(new XElement(paragraphProperties));
        }

        int cursor = 0;
        int revisionIndex = 0;
        foreach (TextRange match in matches)
        {
            if (match.Start > cursor)
            {
                AddTextRun(nodes, current[cursor..match.Start], firstRunProperties);
            }

            string deletedText = current.Substring(match.Start, match.Length);
            nodes.Add(CreateDeletedRun(deletedText, firstRunProperties, revisionIds[revisionIndex++], author, timestamp));
            nodes.Add(CreateInsertedRun(replacement, firstRunProperties, revisionIds[revisionIndex++], author, timestamp));
            cursor = match.Start + match.Length;
        }

        if (cursor < current.Length)
        {
            AddTextRun(nodes, current[cursor..], firstRunProperties);
        }

        paragraph.RemoveNodes();
        paragraph.Add(nodes);
    }

    private static void AddTextRun(List<XNode> nodes, string text, XElement? runProperties)
    {
        if (text.Length == 0)
        {
            return;
        }

        var run = new XElement(OoxmlNs.W + "r");
        if (runProperties is not null)
        {
            run.Add(new XElement(runProperties));
        }

        foreach (XNode node in CreateTextNodes(text))
        {
            run.Add(node);
        }

        nodes.Add(run);
    }

    private static XElement CreateDeletedRun(
        string text,
        XElement? runProperties,
        string revisionId,
        string author,
        string timestamp)
    {
        var run = new XElement(OoxmlNs.W + "r");
        if (runProperties is not null)
        {
            run.Add(new XElement(runProperties));
        }

        run.Add(CreateDeletedTextElement(text));
        return new XElement(
            OoxmlNs.W + "del",
            new XAttribute(OoxmlNs.W + "id", revisionId),
            new XAttribute(OoxmlNs.W + "author", author),
            new XAttribute(OoxmlNs.W + "date", timestamp),
            run);
    }

    private static XElement CreateInsertedRun(
        string text,
        XElement? runProperties,
        string revisionId,
        string author,
        string timestamp)
    {
        var run = new XElement(OoxmlNs.W + "r");
        if (runProperties is not null)
        {
            run.Add(new XElement(runProperties));
        }

        run.Add(CreateTextElement(text));
        return new XElement(
            OoxmlNs.W + "ins",
            new XAttribute(OoxmlNs.W + "id", revisionId),
            new XAttribute(OoxmlNs.W + "author", author),
            new XAttribute(OoxmlNs.W + "date", timestamp),
            run);
    }

    private static bool TryReadImagePixelSize(byte[] bytes, string contentType, out int width, out int height)
    {
        width = 0;
        height = 0;
        return contentType switch
        {
            "image/png" => TryReadPngPixelSize(bytes, out width, out height),
            "image/jpeg" => TryReadJpegPixelSize(bytes, out width, out height),
            _ => false
        };
    }

    private static bool TryReadPngPixelSize(byte[] bytes, out int width, out int height)
    {
        width = 0;
        height = 0;
        if (bytes.Length < 24 ||
            bytes[0] != 0x89 ||
            bytes[1] != 0x50 ||
            bytes[2] != 0x4E ||
            bytes[3] != 0x47 ||
            bytes[4] != 0x0D ||
            bytes[5] != 0x0A ||
            bytes[6] != 0x1A ||
            bytes[7] != 0x0A ||
            bytes[12] != 0x49 ||
            bytes[13] != 0x48 ||
            bytes[14] != 0x44 ||
            bytes[15] != 0x52)
        {
            return false;
        }

        width = ReadBigEndianInt32(bytes, 16);
        height = ReadBigEndianInt32(bytes, 20);
        return width > 0 && height > 0;
    }

    private static bool TryReadJpegPixelSize(byte[] bytes, out int width, out int height)
    {
        width = 0;
        height = 0;
        if (bytes.Length < 4 || bytes[0] != 0xFF || bytes[1] != 0xD8)
        {
            return false;
        }

        int index = 2;
        while (index + 3 < bytes.Length)
        {
            if (bytes[index] != 0xFF)
            {
                index++;
                continue;
            }

            while (index < bytes.Length && bytes[index] == 0xFF)
            {
                index++;
            }

            if (index >= bytes.Length)
            {
                return false;
            }

            byte marker = bytes[index++];
            if (marker is 0xD9 or 0xDA)
            {
                return false;
            }

            if (marker is 0x01 || marker is >= 0xD0 and <= 0xD7)
            {
                continue;
            }

            if (index + 1 >= bytes.Length)
            {
                return false;
            }

            int segmentLength = ReadBigEndianUInt16(bytes, index);
            if (segmentLength < 2 || index + segmentLength > bytes.Length)
            {
                return false;
            }

            if (IsJpegStartOfFrame(marker) && segmentLength >= 7)
            {
                height = ReadBigEndianUInt16(bytes, index + 3);
                width = ReadBigEndianUInt16(bytes, index + 5);
                return width > 0 && height > 0;
            }

            index += segmentLength;
        }

        return false;
    }

    private static bool IsJpegStartOfFrame(byte marker)
    {
        return marker is 0xC0 or 0xC1 or 0xC2 or 0xC3 or 0xC5 or 0xC6 or 0xC7 or 0xC9 or 0xCA or 0xCB or 0xCD or 0xCE or 0xCF;
    }

    private static int ReadBigEndianInt32(byte[] bytes, int offset)
    {
        return (bytes[offset] << 24) |
            (bytes[offset + 1] << 16) |
            (bytes[offset + 2] << 8) |
            bytes[offset + 3];
    }

    private static int ReadBigEndianUInt16(byte[] bytes, int offset)
    {
        return (bytes[offset] << 8) | bytes[offset + 1];
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

    private static int AllocateDrawingDocPrId(XDocument document)
    {
        var existing = new HashSet<int>();
        foreach (XElement docPr in document.Descendants(OoxmlNs.Wp + "docPr"))
        {
            if (int.TryParse((string?)docPr.Attribute("id"), out int id) && id > 0)
            {
                existing.Add(id);
            }
        }

        for (int id = 1; ; id++)
        {
            if (!existing.Contains(id))
            {
                return id;
            }
        }
    }

    private static XElement CreateInlineImageParagraph(string relationshipId, int docPrId, long widthEmus, long heightEmus, string alt)
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
                        new XElement(OoxmlNs.Wp + "docPr", new XAttribute("id", docPrId), new XAttribute("name", $"Picture {docPrId}"), new XAttribute("descr", alt)),
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

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target!, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-image-alt target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
        }

        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateImageContentTypeGuard(operation, target!, imageTarget.Part.ContentType, diagnostics))
        {
            return diagnostics;
        }

        if (!TryGetImageDrawingContainer(imageTarget, target!, operation, out XElement? imageContainer, out DocxDiagnostic? diagnostic))
        {
            return [diagnostic!];
        }

        if (!apply)
        {
            return [];
        }

        SetImageAlt(imageContainer!, alt!, target!);
        SaveDocumentPart(package, imageTarget.PartName, imageTarget.Document);
        return [];
    }

    private static bool TryGetImageDrawingContainer(
        ImageBlipTarget imageTarget,
        string target,
        DocxPatchOperation operation,
        out XElement? container,
        out DocxDiagnostic? diagnostic)
    {
        XElement? drawing = imageTarget.Blip.Ancestors(OoxmlNs.W + "drawing").FirstOrDefault();
        container = drawing?.Descendants(OoxmlNs.Wp + "inline").FirstOrDefault()
            ?? drawing?.Descendants(OoxmlNs.Wp + "anchor").FirstOrDefault();
        if (container is null)
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E5205", $"Image '{target}' does not have editable DrawingML properties.", operation, target);
            return false;
        }

        diagnostic = null;
        return true;
    }

    private static void SetImageAlt(XElement container, string alt, string target)
    {
        XElement? docPr = container.Element(OoxmlNs.Wp + "docPr");
        if (docPr is null)
        {
            docPr = new XElement(OoxmlNs.Wp + "docPr");
            docPr.SetAttributeValue("id", "1");
            docPr.SetAttributeValue("name", target);
            container.AddFirst(docPr);
        }

        docPr.SetAttributeValue("descr", alt);
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

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target!, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported delete-image target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
        }

        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateImageContentTypeGuard(operation, target!, imageTarget.Part.ContentType, diagnostics))
        {
            return diagnostics;
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
        if (!UsesRelationship(imageTarget.Document, imageTarget.RelationshipId))
        {
            package.RemoveRelationship(imageTarget.PartName, imageTarget.RelationshipId);
            if (!AnyRelationshipTargetsPart(package, imageTarget.Part.Name, cancellationToken))
            {
                package.RemovePart(imageTarget.Part.Name);
            }
        }

        SaveDocumentPart(package, imageTarget.PartName, imageTarget.Document);
        return [];
    }

    private static bool UsesRelationship(XDocument document, string relationshipId)
    {
        return document
            .Descendants(OoxmlNs.A + "blip")
            .Any(blip => string.Equals((string?)blip.Attribute(OoxmlNs.R + "embed"), relationshipId, StringComparison.Ordinal));
    }

    private static bool AnyRelationshipTargetsPart(OoxmlPackage package, string partName, CancellationToken cancellationToken)
    {
        foreach (OoxmlPart relationshipPart in package.Parts.Values.Where(part => part.Name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)))
        {
            string sourcePartName = GetSourcePartNameFromRelationshipPartName(relationshipPart.Name);
            foreach (OoxmlRelationship relationship in package.GetRelationships(sourcePartName, cancellationToken))
            {
                if (!relationship.IsExternal &&
                    relationship.ResolvedTarget is not null &&
                    string.Equals(relationship.ResolvedTarget, partName, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
        }

        return false;
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

        if (!ValidateSectionGuards(operation, target!, sectionTarget.SectionProperties, diagnostics))
        {
            return diagnostics;
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

        if (!ValidateSectionGuards(operation, target!, sectionTarget.SectionProperties, diagnostics))
        {
            return diagnostics;
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

        CellTarget? cellTarget = ResolveCellTarget(package, target!, cancellationToken);
        if (cellTarget is null && !IsSupportedCellTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-cell target '{target}'. Expected a table cell ID such as M.T0001.R02.C03 or H001.T0001.R02.C03.", operation, target)];
        }

        if (cellTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateTableGuards(operation, target!, cellTarget.Table, cellTarget.Row, diagnostics))
        {
            return diagnostics;
        }

        string current = ReadVisibleText(cellTarget.Cell);
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

        if (IsVerticalMergeContinuation(cellTarget.Cell))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Unsupported merged-cell target '{target}'.", operation, target)];
        }

        if (!force && !IsSimpleEditableCell(cellTarget.Cell))
        {
            return [Diagnostic(DocxSeverity.Error, "E4302", $"Cell '{target}' contains unsupported content. Use force true only when replacing all cell content is intended.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        ReplaceCellText(cellTarget.Cell, text!);
        SaveDocumentPart(package, cellTarget.PartName, cellTarget.Document);
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

        TableTarget? tableTarget = ResolveTableTarget(package, target!, cancellationToken);
        if (tableTarget is null && !IsSupportedTableTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported append-row target '{target}'. Expected a table ID such as M.T0001 or H001.T0001.", operation, target)];
        }

        string[] cellTexts = operation.FieldValues
            .Where(field => field.Name == "cell")
            .Select(field => field.Value)
            .ToArray();
        if (cellTexts.Length == 0)
        {
            return [Diagnostic(DocxSeverity.Error, "E4202", "Operation 'append-row' is missing required field 'cell'.", operation, target)];
        }

        if (tableTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateTableGuards(operation, target!, tableTarget.Table, row: null, diagnostics))
        {
            return diagnostics;
        }

        XElement[] rows = tableTarget.Table.Elements(OoxmlNs.W + "tr").ToArray();
        if (rows.Length == 0)
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' has no rows to clone.", operation, target)];
        }

        if (ContainsVerticalMerges(tableTarget.Table))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' contains vertical merges that are not supported by append-row.", operation, target)];
        }

        if (!IsRectangular(tableTarget.Table, out int columnCount))
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
        SaveDocumentPart(package, tableTarget.PartName, tableTarget.Document);
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

        RowTarget? rowTarget = ResolveRowTarget(package, target!, cancellationToken);
        if (rowTarget is null && !IsSupportedRowTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported {operation.OperationName} target '{target}'. Expected a table row ID such as M.T0001.R02 or H001.T0001.R02.", operation, target)];
        }

        string[] cellTexts = operation.FieldValues
            .Where(field => field.Name == "cell")
            .Select(field => field.Value)
            .ToArray();
        if (cellTexts.Length == 0)
        {
            return [Diagnostic(DocxSeverity.Error, "E4202", $"Operation '{operation.OperationName}' is missing required field 'cell'.", operation, target)];
        }

        if (rowTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateTableGuards(operation, target!, rowTarget.Table, rowTarget.Row, diagnostics))
        {
            return diagnostics;
        }

        if (ContainsVerticalMerges(rowTarget.Table))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' contains vertical merges that are not supported by {operation.OperationName}.", operation, target)];
        }

        int expectedCellCount;
        if (IsRectangular(rowTarget.Table, out int columnCount))
        {
            expectedCellCount = columnCount;
        }
        else
        {
            if (!force)
            {
                return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' is not rectangular and cannot be edited safely without force true.", operation, target)];
            }

            expectedCellCount = rowTarget.Row.Elements(OoxmlNs.W + "tc").Count();
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

        XElement newRow = CreateRowFromTemplate(rowTarget.Row, cellTexts);
        if (insertAfter)
        {
            rowTarget.Row.AddAfterSelf(newRow);
        }
        else
        {
            rowTarget.Row.AddBeforeSelf(newRow);
        }

        SaveDocumentPart(package, rowTarget.PartName, rowTarget.Document);
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
        string? expectedContains = operation.Fields.GetValueOrDefault("expect-contains");
        bool force = ReadBooleanField(operation, "force", diagnostics) ?? false;
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        RowTarget? rowTarget = ResolveRowTarget(package, target!, cancellationToken);
        if (rowTarget is null && !IsSupportedRowTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported delete-row target '{target}'. Expected a table row ID such as M.T0001.R02 or H001.T0001.R02.", operation, target)];
        }

        if (rowTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateTableGuards(operation, target!, rowTarget.Table, rowTarget.Row, diagnostics))
        {
            return diagnostics;
        }

        string rowText = ReadVisibleText(rowTarget.Row);
        if (expectedContains is not null && !rowText.Contains(expectedContains, StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected row text to contain '{expectedContains}'.", operation, target)];
        }

        XElement[] rows = rowTarget.Table.Elements(OoxmlNs.W + "tr").ToArray();
        if (rows.Length == 1)
        {
            return [Diagnostic(DocxSeverity.Error, "E4304", $"Cannot delete the last row of table '{target}'.", operation, target)];
        }

        if (ContainsVerticalMerges(rowTarget.Table))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' contains vertical merges that are not supported by delete-row.", operation, target)];
        }

        if (!force && !IsRectangular(rowTarget.Table, out _))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' is not rectangular and cannot be edited safely without force true.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        rowTarget.Row.Remove();
        SaveDocumentPart(package, rowTarget.PartName, rowTarget.Document);
        return [];
    }

    private static bool ValidateTableGuards(
        DocxPatchOperation operation,
        string target,
        XElement table,
        XElement? row,
        List<DocxDiagnostic> diagnostics)
    {
        if (!TryReadPositiveIntegerGuard(operation, "expect-row-count", target, diagnostics, out int? expectedRowCount) ||
            !TryReadPositiveIntegerGuard(operation, "expect-column-count", target, diagnostics, out int? expectedColumnCount) ||
            !TryReadPositiveIntegerGuard(operation, "expect-cell-count", target, diagnostics, out int? expectedCellCount))
        {
            return false;
        }

        int actualRowCount = table.Elements(OoxmlNs.W + "tr").Count();
        if (expectedRowCount is not null && actualRowCount != expectedRowCount)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected {expectedRowCount} row(s), found {actualRowCount}.", operation, target));
        }

        if (expectedColumnCount is not null)
        {
            if (!IsRectangular(table, out int actualColumnCount))
            {
                diagnostics.Add(Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected {expectedColumnCount} column(s), but table is not rectangular.", operation, target));
            }
            else if (actualColumnCount != expectedColumnCount)
            {
                diagnostics.Add(Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected {expectedColumnCount} column(s), found {actualColumnCount}.", operation, target));
            }
        }

        if (expectedCellCount is not null)
        {
            if (row is null)
            {
                diagnostics.Add(Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Field 'expect-cell-count' requires a row or cell target.", operation, target));
            }
            else
            {
                int actualCellCount = row.Elements(OoxmlNs.W + "tc").Count();
                if (actualCellCount != expectedCellCount)
                {
                    diagnostics.Add(Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected {expectedCellCount} cell(s), found {actualCellCount}.", operation, target));
                }
            }
        }

        return diagnostics.Count == 0;
    }

    private static bool ValidateImageContentTypeGuard(
        DocxPatchOperation operation,
        string target,
        string? actualContentType,
        List<DocxDiagnostic> diagnostics)
    {
        if (!operation.Fields.TryGetValue("expect-content-type", out string? expectedContentType))
        {
            return true;
        }

        string? normalizedExpected = NormalizeImageContentType(expectedContentType);
        if (normalizedExpected is null)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Field 'expect-content-type' must be image/png or image/jpeg.", operation, target));
            return false;
        }

        if (!string.Equals(actualContentType, normalizedExpected, StringComparison.Ordinal))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected image content type '{normalizedExpected}', found '{actualContentType ?? "unknown"}'.", operation, target));
        }

        return diagnostics.Count == 0;
    }

    private static bool ValidateSectionGuards(
        DocxPatchOperation operation,
        string target,
        XElement sectionProperties,
        List<DocxDiagnostic> diagnostics)
    {
        if (!TryReadPositiveIntegerGuard(operation, "expect-columns", target, diagnostics, out int? expectedColumns))
        {
            return false;
        }

        if (expectedColumns is not null)
        {
            int actualColumns = ReadSectionColumnCount(sectionProperties);
            if (actualColumns != expectedColumns)
            {
                diagnostics.Add(Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected {expectedColumns} section column(s), found {actualColumns}.", operation, target));
            }
        }

        if (operation.Fields.TryGetValue("expect-orientation", out string? expectedOrientation))
        {
            if (expectedOrientation is not ("portrait" or "landscape"))
            {
                diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Field 'expect-orientation' must be portrait or landscape.", operation, target));
            }
            else
            {
                string actualOrientation = ReadSectionOrientation(sectionProperties);
                if (!string.Equals(actualOrientation, expectedOrientation, StringComparison.Ordinal))
                {
                    diagnostics.Add(Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected section orientation '{expectedOrientation}', found '{actualOrientation}'.", operation, target));
                }
            }
        }

        return diagnostics.Count == 0;
    }

    private static int ReadSectionColumnCount(XElement sectionProperties)
    {
        string? countText = (string?)sectionProperties
            .Element(OoxmlNs.W + "cols")
            ?.Attribute(OoxmlNs.W + "num");
        return int.TryParse(countText, out int count) && count > 0 ? count : 1;
    }

    private static string ReadSectionOrientation(XElement sectionProperties)
    {
        return (string?)sectionProperties
            .Element(OoxmlNs.W + "pgSz")
            ?.Attribute(OoxmlNs.W + "orient")
            ?? "portrait";
    }

    private static bool TryReadPositiveIntegerGuard(
        DocxPatchOperation operation,
        string fieldName,
        string target,
        List<DocxDiagnostic> diagnostics,
        out int? value)
    {
        value = null;
        if (!operation.Fields.TryGetValue(fieldName, out string? text))
        {
            return true;
        }

        if (!int.TryParse(text, out int parsed) || parsed <= 0)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", $"Field '{fieldName}' must be greater than 0.", operation, target));
            return false;
        }

        value = parsed;
        return true;
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

    private static bool TryParseStoryImageTarget(
        string target,
        char storyPrefix,
        out int storyOrdinal,
        out int imageOrdinal)
    {
        storyOrdinal = 0;
        imageOrdinal = 0;
        if (target.Length != 10 ||
            target[0] != storyPrefix ||
            target[4..6] != ".I")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[6..], out imageOrdinal);
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

    private static BlockTarget? ResolveBlockTarget(
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

        if (selector is not ExplicitIdTargetSelector)
        {
            if (package.MainDocumentPartName is null)
            {
                return null;
            }

            XDocument mainDocument = LoadMainDocument(package, cancellationToken, out XElement mainBody);
            XElement? selectedBlock = ResolveMainBlock(mainBody, operation, target, out diagnostics);
            return selectedBlock is null
                ? null
                : new BlockTarget(package.MainDocumentPartName, mainDocument, selectedBlock);
        }

        if (TryParseMainParagraphTarget(target, out int mainParagraphOrdinal))
        {
            XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
            XElement? paragraph = FindParagraph(body, mainParagraphOrdinal);
            return paragraph is null || package.MainDocumentPartName is null
                ? null
                : new BlockTarget(package.MainDocumentPartName, document, paragraph);
        }

        if (TryParseMainTableTarget(target, out int mainTableOrdinal))
        {
            XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
            XElement? table = FindTable(body, mainTableOrdinal);
            return table is null || package.MainDocumentPartName is null
                ? null
                : new BlockTarget(package.MainDocumentPartName, document, table);
        }

        if (TryParseStoryParagraphTarget(target, 'H', out int headerOrdinal, out int headerParagraphOrdinal))
        {
            return ResolveRelatedStoryBlockTarget(package, OoxmlRelTypes.Header, headerOrdinal, OoxmlNs.W + "p", headerParagraphOrdinal, cancellationToken);
        }

        if (TryParseStoryParagraphTarget(target, 'F', out int footerOrdinal, out int footerParagraphOrdinal))
        {
            return ResolveRelatedStoryBlockTarget(package, OoxmlRelTypes.Footer, footerOrdinal, OoxmlNs.W + "p", footerParagraphOrdinal, cancellationToken);
        }

        if (TryParseStoryTableTarget(target, 'H', out headerOrdinal, out int headerTableOrdinal))
        {
            return ResolveRelatedStoryBlockTarget(package, OoxmlRelTypes.Header, headerOrdinal, OoxmlNs.W + "tbl", headerTableOrdinal, cancellationToken);
        }

        if (TryParseStoryTableTarget(target, 'F', out footerOrdinal, out int footerTableOrdinal))
        {
            return ResolveRelatedStoryBlockTarget(package, OoxmlRelTypes.Footer, footerOrdinal, OoxmlNs.W + "tbl", footerTableOrdinal, cancellationToken);
        }

        return null;
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

        if (target.StartsWith("bookmark:", StringComparison.Ordinal))
        {
            string value = target["bookmark:".Length..].Trim();
            if (!TryReadSelectorText(value, out string selectorText))
            {
                diagnostic = Diagnostic(DocxSeverity.Error, "E1203", $"Invalid bookmark selector '{target}'. Expected bookmark:\"Name\".", operation, target);
                return false;
            }

            selector = new BookmarkTargetSelector(target, selectorText);
            return true;
        }

        if (target.StartsWith("content-control:", StringComparison.Ordinal))
        {
            string value = target["content-control:".Length..].Trim();
            if (!TryReadSelectorText(value, out string selectorText))
            {
                diagnostic = Diagnostic(DocxSeverity.Error, "E1203", $"Invalid content-control selector '{target}'. Expected content-control:\"TagOrAlias\".", operation, target);
                return false;
            }

            selector = new ContentControlTargetSelector(target, selectorText);
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

        if (selector is BookmarkTargetSelector bookmarkSelector)
        {
            return ResolveMainParagraphElementByPredicate(
                body,
                paragraph => ParagraphHasBookmark(paragraph, bookmarkSelector.Name),
                selector.Raw,
                operation,
                out diagnostics);
        }

        if (selector is ContentControlTargetSelector contentControlSelector)
        {
            return ResolveMainParagraphElementByPredicate(
                body,
                paragraph => ParagraphHasContentControl(paragraph, contentControlSelector.Name),
                selector.Raw,
                operation,
                out diagnostics);
        }

        return null;
    }

    private static bool ParagraphHasBookmark(XElement paragraph, string name)
    {
        return paragraph
            .Descendants(OoxmlNs.W + "bookmarkStart")
            .Any(bookmark => string.Equals((string?)bookmark.Attribute(OoxmlNs.W + "name"), name, StringComparison.Ordinal));
    }

    private static bool ParagraphHasContentControl(XElement paragraph, string name)
    {
        return paragraph
            .Descendants(OoxmlNs.W + "sdt")
            .Any(contentControl =>
            {
                XElement? properties = contentControl.Element(OoxmlNs.W + "sdtPr");
                string? tag = (string?)properties
                    ?.Element(OoxmlNs.W + "tag")
                    ?.Attribute(OoxmlNs.W + "val");
                string? alias = (string?)properties
                    ?.Element(OoxmlNs.W + "alias")
                    ?.Attribute(OoxmlNs.W + "val");
                return string.Equals(tag, name, StringComparison.Ordinal) ||
                    string.Equals(alias, name, StringComparison.Ordinal);
            });
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

        if (matches.Count == 0)
        {
            diagnostics =
            [
                Diagnostic(
                    DocxSeverity.Error,
                    "E1201",
                    $"Selector matched 0 targets: {rawSelector}.{BuildNoMatchSuggestion(body, rawSelector)}",
                    operation,
                    rawSelector)
            ];
            return null;
        }

        return matches[0].Paragraph;
    }

    private static string BuildNoMatchSuggestion(XElement body, string rawSelector)
    {
        string[] suggestions = rawSelector.StartsWith("heading:", StringComparison.Ordinal)
            ? EnumerateMainParagraphs(body)
                .Select(match => (match.Id, HeadingLevel: ReadHeadingLevel(match.Paragraph)))
                .Where(match => match.HeadingLevel is not null)
                .Take(3)
                .Select(match => $"{match.Id} heading level={match.HeadingLevel}")
                .ToArray()
            : EnumerateMainParagraphs(body)
                .Take(3)
                .Select(match => match.Id)
                .ToArray();
        if (suggestions.Length == 0)
        {
            return " No nearby paragraph targets are available.";
        }

        string label = rawSelector.StartsWith("heading:", StringComparison.Ordinal)
            ? "Nearby headings"
            : "Nearby paragraphs";
        return $" {label}: {string.Join(", ", suggestions)}.";
    }

    private static IEnumerable<ParagraphSelectorMatch> EnumerateMainParagraphs(XElement body)
    {
        int paragraphOrdinal = 0;
        foreach (XElement paragraph in body.Elements(OoxmlNs.W + "p"))
        {
            paragraphOrdinal++;
            yield return new ParagraphSelectorMatch($"M.P{paragraphOrdinal:0000}", paragraph);
        }
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

    private static bool TryParseStoryTableTarget(
        string target,
        char storyPrefix,
        out int storyOrdinal,
        out int tableOrdinal)
    {
        storyOrdinal = 0;
        tableOrdinal = 0;
        if (target.Length != 10 ||
            target[0] != storyPrefix ||
            target[4..6] != ".T")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[6..], out tableOrdinal);
    }

    private static bool TryParseStoryRowTarget(
        string target,
        char storyPrefix,
        out int storyOrdinal,
        out int tableOrdinal,
        out int rowOrdinal)
    {
        storyOrdinal = 0;
        tableOrdinal = 0;
        rowOrdinal = 0;
        if (target.Length != 14 ||
            target[0] != storyPrefix ||
            target[4..6] != ".T" ||
            target[10..12] != ".R")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[6..10], out tableOrdinal) &&
            int.TryParse(target[12..], out rowOrdinal);
    }

    private static bool TryParseStoryCellTarget(
        string target,
        char storyPrefix,
        out int storyOrdinal,
        out int tableOrdinal,
        out int rowOrdinal,
        out int cellOrdinal)
    {
        storyOrdinal = 0;
        tableOrdinal = 0;
        rowOrdinal = 0;
        cellOrdinal = 0;
        if (target.Length != 18 ||
            target[0] != storyPrefix ||
            target[4..6] != ".T" ||
            target[10..12] != ".R" ||
            target[14..16] != ".C")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[6..10], out tableOrdinal) &&
            int.TryParse(target[12..14], out rowOrdinal) &&
            int.TryParse(target[16..], out cellOrdinal);
    }

    private static BlockTarget? ResolveRelatedStoryBlockTarget(
        OoxmlPackage package,
        string relationshipType,
        int storyOrdinal,
        XName blockName,
        int blockOrdinal,
        CancellationToken cancellationToken)
    {
        if (package.MainDocumentPartName is null || storyOrdinal < 1 || blockOrdinal < 1)
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
        XElement? block = root.Elements(blockName).ElementAtOrDefault(blockOrdinal - 1);
        return block is null ? null : new BlockTarget(relationship.ResolvedTarget, document, block);
    }

    private static bool IsSupportedTableTargetShape(string target)
    {
        return TryParseMainTableTarget(target, out _) ||
            TryParseStoryTableTarget(target, 'H', out _, out _) ||
            TryParseStoryTableTarget(target, 'F', out _, out _);
    }

    private static bool IsSupportedRowTargetShape(string target)
    {
        return TryParseMainRowTarget(target, out _, out _) ||
            TryParseStoryRowTarget(target, 'H', out _, out _, out _) ||
            TryParseStoryRowTarget(target, 'F', out _, out _, out _);
    }

    private static bool IsSupportedCellTargetShape(string target)
    {
        return TryParseMainCellTarget(target, out _, out _, out _) ||
            TryParseStoryCellTarget(target, 'H', out _, out _, out _, out _) ||
            TryParseStoryCellTarget(target, 'F', out _, out _, out _, out _);
    }

    private static TableTarget? ResolveTableTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (TryParseMainTableTarget(target, out int mainTableOrdinal))
        {
            XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
            XElement? table = FindTable(body, mainTableOrdinal);
            return table is null || package.MainDocumentPartName is null
                ? null
                : new TableTarget(package.MainDocumentPartName, document, table);
        }

        if (TryParseStoryTableTarget(target, 'H', out int headerOrdinal, out int headerTableOrdinal))
        {
            return ResolveRelatedStoryTableTarget(package, OoxmlRelTypes.Header, headerOrdinal, headerTableOrdinal, cancellationToken);
        }

        if (TryParseStoryTableTarget(target, 'F', out int footerOrdinal, out int footerTableOrdinal))
        {
            return ResolveRelatedStoryTableTarget(package, OoxmlRelTypes.Footer, footerOrdinal, footerTableOrdinal, cancellationToken);
        }

        return null;
    }

    private static RowTarget? ResolveRowTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        TableTarget? tableTarget;
        int rowOrdinal;
        if (TryParseMainRowTarget(target, out int mainTableOrdinal, out rowOrdinal))
        {
            tableTarget = ResolveTableTarget(package, $"M.T{mainTableOrdinal:0000}", cancellationToken);
        }
        else if (TryParseStoryRowTarget(target, 'H', out int headerOrdinal, out int headerTableOrdinal, out rowOrdinal))
        {
            tableTarget = ResolveRelatedStoryTableTarget(package, OoxmlRelTypes.Header, headerOrdinal, headerTableOrdinal, cancellationToken);
        }
        else if (TryParseStoryRowTarget(target, 'F', out int footerOrdinal, out int footerTableOrdinal, out rowOrdinal))
        {
            tableTarget = ResolveRelatedStoryTableTarget(package, OoxmlRelTypes.Footer, footerOrdinal, footerTableOrdinal, cancellationToken);
        }
        else
        {
            return null;
        }

        XElement? row = tableTarget?.Table.Elements(OoxmlNs.W + "tr").ElementAtOrDefault(rowOrdinal - 1);
        return tableTarget is null || row is null
            ? null
            : new RowTarget(tableTarget.PartName, tableTarget.Document, tableTarget.Table, row);
    }

    private static CellTarget? ResolveCellTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        RowTarget? rowTarget;
        int cellOrdinal;
        if (TryParseMainCellTarget(target, out int mainTableOrdinal, out int rowOrdinal, out cellOrdinal))
        {
            rowTarget = ResolveRowTarget(package, $"M.T{mainTableOrdinal:0000}.R{rowOrdinal:00}", cancellationToken);
        }
        else if (TryParseStoryCellTarget(target, 'H', out int headerOrdinal, out int headerTableOrdinal, out rowOrdinal, out cellOrdinal))
        {
            rowTarget = ResolveRowTarget(package, $"H{headerOrdinal:000}.T{headerTableOrdinal:0000}.R{rowOrdinal:00}", cancellationToken);
        }
        else if (TryParseStoryCellTarget(target, 'F', out int footerOrdinal, out int footerTableOrdinal, out rowOrdinal, out cellOrdinal))
        {
            rowTarget = ResolveRowTarget(package, $"F{footerOrdinal:000}.T{footerTableOrdinal:0000}.R{rowOrdinal:00}", cancellationToken);
        }
        else
        {
            return null;
        }

        XElement? cell = rowTarget?.Row.Elements(OoxmlNs.W + "tc").ElementAtOrDefault(cellOrdinal - 1);
        return rowTarget is null || cell is null
            ? null
            : new CellTarget(rowTarget.PartName, rowTarget.Document, rowTarget.Table, rowTarget.Row, cell);
    }

    private static TableTarget? ResolveRelatedStoryTableTarget(
        OoxmlPackage package,
        string relationshipType,
        int storyOrdinal,
        int tableOrdinal,
        CancellationToken cancellationToken)
    {
        BlockTarget? blockTarget = ResolveRelatedStoryBlockTarget(package, relationshipType, storyOrdinal, OoxmlNs.W + "tbl", tableOrdinal, cancellationToken);
        return blockTarget is null ? null : new TableTarget(blockTarget.PartName, blockTarget.Document, blockTarget.Block);
    }

    private static bool IsSupportedImageTargetShape(string target)
    {
        return TryParseMainImageTarget(target, out _) ||
            TryParseStoryImageTarget(target, 'H', out _, out _) ||
            TryParseStoryImageTarget(target, 'F', out _, out _);
    }

    private static ImageBlipTarget? ResolveImageBlipTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (TryParseMainImageTarget(target, out int mainImageOrdinal))
        {
            return package.MainDocumentPartName is null
                ? null
                : FindImageBlipTarget(package, package.MainDocumentPartName, mainImageOrdinal, cancellationToken);
        }

        if (TryParseStoryImageTarget(target, 'H', out int headerOrdinal, out int headerImageOrdinal))
        {
            string? partName = ResolveRelatedStoryPartName(package, OoxmlRelTypes.Header, headerOrdinal, cancellationToken);
            return partName is null ? null : FindImageBlipTarget(package, partName, headerImageOrdinal, cancellationToken);
        }

        if (TryParseStoryImageTarget(target, 'F', out int footerOrdinal, out int footerImageOrdinal))
        {
            string? partName = ResolveRelatedStoryPartName(package, OoxmlRelTypes.Footer, footerOrdinal, cancellationToken);
            return partName is null ? null : FindImageBlipTarget(package, partName, footerImageOrdinal, cancellationToken);
        }

        return null;
    }

    private static ImageBlipTarget? FindImageBlipTarget(
        OoxmlPackage package,
        string partName,
        int imageOrdinal,
        CancellationToken cancellationToken)
    {
        if (imageOrdinal < 1)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
        IReadOnlyDictionary<string, OoxmlRelationship> relationships = package
            .GetRelationships(partName, cancellationToken)
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
                return new ImageBlipTarget(partName, document, blip, relationshipId, part);
            }
        }

        return null;
    }

    private static string? ResolveRelatedStoryPartName(
        OoxmlPackage package,
        string relationshipType,
        int storyOrdinal,
        CancellationToken cancellationToken)
    {
        if (package.MainDocumentPartName is null || storyOrdinal < 1)
        {
            return null;
        }

        return package
            .GetRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => !relationship.IsExternal && relationship.Type == relationshipType && relationship.ResolvedTarget is not null)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .ElementAtOrDefault(storyOrdinal - 1)
            ?.ResolvedTarget;
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
        XElement? paragraphProperties = cell
            .Elements(OoxmlNs.W + "p")
            .Elements(OoxmlNs.W + "pPr")
            .FirstOrDefault();
        cell.RemoveNodes();
        if (cellProperties is not null)
        {
            cell.Add(new XElement(cellProperties));
        }

        cell.Add(CreateSimpleParagraph(text, paragraphProperties: paragraphProperties));
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

            XElement? paragraphProperties = templateCells[i]
                .Elements(OoxmlNs.W + "p")
                .Elements(OoxmlNs.W + "pPr")
                .FirstOrDefault();
            cell.Add(CreateSimpleParagraph(cellTexts[i], paragraphProperties: paragraphProperties));
            row.Add(cell);
        }

        return row;
    }

    private static XElement CreateSimpleParagraph(string text, string? style = null, XElement? paragraphProperties = null)
    {
        var paragraph = new XElement(OoxmlNs.W + "p");
        if (paragraphProperties is not null)
        {
            paragraph.Add(new XElement(paragraphProperties));
        }

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

    private static XElement CreateDeletedTextElement(string text)
    {
        var textElement = new XElement(OoxmlNs.W + "delText", text);
        if (RequiresPreserveSpace(text))
        {
            textElement.SetAttributeValue(OoxmlNs.Xml + "space", "preserve");
        }

        return textElement;
    }

    private static string[] AllocateRevisionIds(
        OoxmlPackage package,
        int count,
        CancellationToken cancellationToken)
    {
        int nextId = 1;
        foreach (OoxmlPart part in package.Parts.Values
            .Where(part => part.Name.StartsWith("/word/", StringComparison.OrdinalIgnoreCase) && part.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(part => part.Name, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            foreach (XAttribute idAttribute in document.Descendants().Attributes(OoxmlNs.W + "id"))
            {
                if (int.TryParse(idAttribute.Value, out int id) && id >= nextId)
                {
                    nextId = id + 1;
                }
            }
        }

        string[] ids = new string[count];
        for (int i = 0; i < ids.Length; i++)
        {
            ids[i] = (nextId++).ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        return ids;
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

internal sealed record ImageBlipTarget(string PartName, XDocument Document, XElement Blip, string RelationshipId, OoxmlPart Part);

internal sealed record ParagraphTarget(string PartName, XDocument Document, XElement Paragraph);

internal sealed record BlockTarget(string PartName, XDocument Document, XElement Block);

internal sealed record TableTarget(string PartName, XDocument Document, XElement Table);

internal sealed record RowTarget(string PartName, XDocument Document, XElement Table, XElement Row);

internal sealed record CellTarget(string PartName, XDocument Document, XElement Table, XElement Row, XElement Cell);

internal sealed record SectionTarget(XDocument Document, XElement SectionProperties);

internal readonly record struct TextRange(int Start, int Length);

internal sealed record TextPosition(XElement TextElement, int Offset);

internal abstract record TargetSelector(string Raw);

internal sealed record ExplicitIdTargetSelector(string Raw) : TargetSelector(Raw);

internal sealed record HeadingTargetSelector(string Raw, int? Level, string Text) : TargetSelector(Raw);

internal sealed record ParagraphTextTargetSelector(string Raw, string Text) : TargetSelector(Raw);

internal sealed record BookmarkTargetSelector(string Raw, string Name) : TargetSelector(Raw);

internal sealed record ContentControlTargetSelector(string Raw, string Name) : TargetSelector(Raw);

internal sealed record ParagraphSelectorMatch(string Id, XElement Paragraph);
