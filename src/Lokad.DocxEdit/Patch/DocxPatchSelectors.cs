using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

internal static partial class DocxPatchEngine
{
    private static XElement? ResolveMainBlock(
        OoxmlPackage package,
        XElement body,
        DocxPatchOperation operation,
        string target,
        CancellationToken cancellationToken,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        diagnostics = [];
        if (!TryParseTargetSelector(target, operation, out TargetSelector? selector, out DocxDiagnostic? diagnostic))
        {
            diagnostics = [diagnostic];
            return null;
        }

        if (selector is not ExplicitIdTargetSelector explicitId)
        {
            return ResolveMainParagraphElementBySelector(package, body, selector, operation, cancellationToken, out diagnostics);
        }

        if (explicitId.TargetId is not { } blockId)
        {
            return null;
        }

        if (blockId is { Story: 'M', Kind: DocxTargetKind.Paragraph })
        {
            return FindSnapshotElementInContainer(body, OoxmlNs.W + "p", blockId.ToWireValue());
        }

        if (blockId is { Story: 'M', Kind: DocxTargetKind.Table })
        {
            return FindSnapshotElementInContainer(body, OoxmlNs.W + "tbl", blockId.ToWireValue());
        }

        return null;
    }

    private static BlockTarget? ResolveBlockTarget(
        OoxmlPackage package,
        DocxPatchOperation operation,
        string target,
        CancellationToken cancellationToken,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        diagnostics = [];
        if (IsAliasReference(target))
        {
            return ResolveAliasBlockTarget(package, operation, target, cancellationToken, out diagnostics);
        }

        if (!TryParseTargetSelector(target, operation, out TargetSelector? selector, out DocxDiagnostic? diagnostic))
        {
            diagnostics = [diagnostic];
            return null;
        }

        if (selector is not ExplicitIdTargetSelector)
        {

            XDocument mainDocument = LoadMainDocument(package, cancellationToken, out XElement mainBody);
            XElement? selectedBlock = ResolveMainBlock(package, mainBody, operation, target, cancellationToken, out diagnostics);
            return selectedBlock is null
                ? null
                : new BlockTarget(package.MainDocumentPartName, mainDocument, selectedBlock);
        }

        if (selector is ExplicitIdTargetSelector { TargetId: { } blockId })
        {
            if (blockId is { Story: 'M', Kind: DocxTargetKind.Paragraph })
            {
                XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
                XElement? paragraph = FindSnapshotElementInContainer(body, OoxmlNs.W + "p", blockId.ToWireValue());
                return paragraph is null
                    ? null
                    : new BlockTarget(package.MainDocumentPartName, document, paragraph);
            }

            if (blockId is { Story: 'M', Kind: DocxTargetKind.Table })
            {
                XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
                XElement? table = FindSnapshotElementInContainer(body, OoxmlNs.W + "tbl", blockId.ToWireValue());
                return table is null
                    ? null
                    : new BlockTarget(package.MainDocumentPartName, document, table);
            }

            bool isStoryBlock = blockId is { Kind: DocxTargetKind.Paragraph, Story: 'H' or 'F' }
                or { Kind: DocxTargetKind.Table, Story: 'H' or 'F' };
            if (isStoryBlock)
            {
                string relationshipType = blockId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
                return ResolveRelatedStoryBlockTarget(package, relationshipType, blockId, cancellationToken);
            }
        diagnostics = [WrongKindDiagnostic(operation, target, blockId, "a paragraph ID such as M.P0001 or a table ID such as M.T0001")];
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
        if (IsAliasReference(target))
        {
            return ResolveAliasParagraphTarget(package, operation, target, cancellationToken, out diagnostics);
        }

        if (!TryParseTargetSelector(target, operation, out TargetSelector? selector, out DocxDiagnostic? diagnostic))
        {
            diagnostics = [diagnostic];
            return null;
        }

        if (selector is ExplicitIdTargetSelector { TargetId: { } paragraphId })
        {
            if (paragraphId is { Story: 'M', Kind: DocxTargetKind.Paragraph })
            {
                XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
                XElement? paragraph = FindSnapshotElementInContainer(body, OoxmlNs.W + "p", paragraphId.ToWireValue());
                return paragraph is null
                    ? null
                    : new ParagraphTarget(package.MainDocumentPartName, document, paragraph);
            }

            bool isStoryParagraph = paragraphId is { Kind: DocxTargetKind.Paragraph, Story: 'H' or 'F' };
            if (isStoryParagraph)
            {
                string relationshipType = paragraphId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
                return ResolveRelatedStoryParagraphTarget(package, relationshipType, paragraphId, cancellationToken);
            }

            diagnostics = [WrongKindDiagnostic(operation, target, paragraphId, DescribeAcceptedKindsForOperation(operation))];
            return null;
        }


        XDocument mainDocument = LoadMainDocument(package, cancellationToken, out XElement mainBody);
        XElement? selectedParagraph = ResolveMainParagraphElementBySelector(package, mainBody, selector, operation, cancellationToken, out diagnostics);
        return selectedParagraph is null
            ? null
            : new ParagraphTarget(package.MainDocumentPartName, mainDocument, selectedParagraph);
    }

    // D14: capture the resolved target for paragraph and block text ops so
    // reports identify the resolved paragraph (or table anchor) instead of
    // echoing only the selector. Snapshot marks supply input-snapshot IDs;
    // a missing mark or failed resolution means no reporting, never a failure.
    private static DocxTargetId? CaptureResolvedTargetSnapshot(
        OoxmlPackage package,
        DocxPatchOperation operation,
        CancellationToken cancellationToken,
        out string? createdMark)
    {
        createdMark = null;
        string? target = operation.Fields.GetValueOrDefault("target");
        if (string.IsNullOrWhiteSpace(target))
        {
            return null;
        }

        if (IsAliasReference(target))
        {
            ParagraphTarget? aliasCapture = ResolveAliasParagraphTarget(package, operation, target, cancellationToken, out _);
            createdMark = (string?)aliasCapture?.Paragraph.Attribute(SnapshotCreatedName);
            return TryResolveAliasParagraphId(package, target, cancellationToken);
        }

        if (DocxTargetId.TryParse(target, out DocxTargetId explicitId))
        {
            if (explicitId.Kind == DocxTargetKind.Table
                && operation.OperationName is ("insert-before" or "insert-after" or "delete-block"))
            {
                return explicitId;
            }

            if (explicitId.Kind is DocxTargetKind.Paragraph)
            {
                if (operation.OperationName is not ("replace-text" or "replace-paragraph" or "set-style" or "insert-before" or "insert-after" or "delete-block"))
                {
                    return null;
                }

                ParagraphTarget? paragraphTarget = ResolveParagraphTarget(package, operation, target, cancellationToken, out _);
                if (paragraphTarget is null)
                {
                    return null;
                }

                string? snapshotId = (string?)paragraphTarget.Paragraph.Attribute(SnapshotIdName);
                return snapshotId is not null && DocxTargetId.TryParse(snapshotId, out DocxTargetId resolved)
                    ? resolved
                    : null;
            }

            return CaptureExplicitTargetSnapshotId(package, operation, explicitId, cancellationToken);
        }

        if (operation.OperationName is not ("replace-text" or "replace-paragraph" or "set-style" or "insert-before" or "insert-after" or "delete-block"))
        {
            return null;
        }

        ParagraphTarget? semanticTarget = ResolveParagraphTarget(package, operation, target, cancellationToken, out _);
        if (semanticTarget is null)
        {
            return null;
        }

        string? semanticSnapshotId = (string?)semanticTarget.Paragraph.Attribute(SnapshotIdName);
        createdMark = (string?)semanticTarget.Paragraph.Attribute(SnapshotCreatedName);
        if (semanticSnapshotId is not null && DocxTargetId.TryParse(semanticSnapshotId, out DocxTargetId semanticResolved))
        {
            return semanticResolved;
        }
        if (createdMark is not null)
        {
            int? fallbackOrdinal = PhysicalParagraphOrdinal(semanticTarget.Document, semanticTarget.Paragraph);
            if (fallbackOrdinal is not null && DocxPartRoles.GetStoryPrefixes(package, cancellationToken).TryGetValue(semanticTarget.PartName, out string? fallbackPrefix) && TryParseStoryPrefix(fallbackPrefix, out char fallbackStory, out int fallbackPart))
            {
                return new DocxTargetId(fallbackStory, fallbackPart, DocxTargetKind.Paragraph, fallbackOrdinal.Value, 0, 0);
            }
        }
        return null;
    }

    private static DocxTargetId? CaptureExplicitTargetSnapshotId(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxTargetId explicitId,
        CancellationToken cancellationToken)
    {
        string target = explicitId.ToWireValue();
        XElement? element = explicitId.Kind switch
        {
            DocxTargetKind.Table => ResolveTableTarget(package, target, cancellationToken)?.Table,
            DocxTargetKind.Row => ResolveRowTarget(package, target, cancellationToken)?.Row,
            DocxTargetKind.Cell or DocxTargetKind.MergeGroup => ResolveCellTarget(package, target, cancellationToken)?.Cell,
            DocxTargetKind.Hyperlink => ResolveHyperlinkTarget(package, target, cancellationToken)?.Hyperlink,
            DocxTargetKind.Bookmark => ResolveBookmarkTarget(package, operation, target, cancellationToken, out _)?.Start,
            DocxTargetKind.ContentControl => ResolveContentControlTarget(package, operation, target, cancellationToken, out _)?.ContentControl,
            DocxTargetKind.Field => ResolveFieldTarget(package, target, cancellationToken)?.Element,
            DocxTargetKind.Image => ResolveImageBlipTarget(package, target, cancellationToken)?.Blip,
            DocxTargetKind.Section => ResolveMainSectionTarget(package, target, cancellationToken)?.SectionProperties,
            _ => null
        };

        if (element is null)
        {
            return null;
        }

        XName markName = explicitId.Kind == DocxTargetKind.MergeGroup ? SnapshotMergeGroupName : SnapshotIdName;
        string? snapshotId = (string?)element.Attribute(markName);
        return snapshotId is not null && DocxTargetId.TryParse(snapshotId, out DocxTargetId resolved)
            ? resolved
            : null;
    }

    // D14: opt-in bounded before/after preview state for text and property edits.
    // Reads run against the live package: before-values are captured pre-operation
    // and after-values are re-read post-operation. A null record means previews are
    // disabled or inapplicable; a record with a null value means the value itself
    // was absent (for example no paragraph style). Missing resolution never fails.
    private static string? ResolvePreviewLocator(
        OoxmlPackage package,
        DocxPatchOperation operation,
        CancellationToken cancellationToken)
    {
        string? target = operation.Fields.GetValueOrDefault("target");
        if (string.IsNullOrWhiteSpace(target))
        {
            return null;
        }

        XElement? element = ResolvePreviewElement(package, operation, target, cancellationToken);
        if (element is null)
        {
            return null;
        }

        string? mark = (string?)element.Attribute(SnapshotIdName)
            ?? (string?)element.Attribute(SnapshotMergeGroupName)
            ?? (string?)element.Attribute(SnapshotAliasName);
        if (mark is not null && DocxTargetId.TryParse(mark, out _))
        {
            return mark;
        }

        string? commentWire = (string?)element.Attribute(SnapshotIdName);
        if (commentWire is not null && TryParseCommentBodyTarget(commentWire, out _, out _))
        {
            return commentWire;
        }

        string? aliasName = (string?)element.Attribute(SnapshotAliasName);
        if (aliasName is not null)
        {
            return "@" + aliasName;
        }
        string? createdLoc = (string?)element.Attribute(SnapshotCreatedName);
        if (createdLoc is not null)
        {
            return createdLoc;
        }
        return null;
    }

    private static XElement? ResolvePreviewElement(
        OoxmlPackage package,
        DocxPatchOperation operation,
        string target,
        CancellationToken cancellationToken)
    {
        var created = FindMarkedStoryElement(package, SnapshotCreatedName, target, OoxmlNs.W + "p", cancellationToken);
        if (created is not null)
        {
            return created.Value.Element;
        }
        return operation.OperationName switch
        {
            "set-cell" or "set-cell-shading" => ResolveCellTarget(package, target, cancellationToken)?.Cell,
            "set-hyperlink-text" or "set-hyperlink-target" or "remove-hyperlink" => ResolveHyperlinkTarget(package, target, cancellationToken)?.Hyperlink,
            "set-table-style" => ResolveTableTarget(package, target, cancellationToken)?.Table,
            "set-row-header" => ResolveRowTarget(package, target, cancellationToken)?.Row,
            "set-content-control-text" => ResolveContentControlTarget(package, operation, target, cancellationToken, out _)?.ContentControl,
            "set-image-alt" => ResolveImageBlipTarget(package, target, cancellationToken)?.Blip,
            "set-field-result" or "set-field-code" or "set-field-dirty" or "set-field-lock" => ResolveFieldTarget(package, target, cancellationToken)?.Element,
            "set-comment-text" => ResolveCommentTarget(package, target, operation, cancellationToken, out _)?.Comment,
            "replace-bookmark-text" => ResolveBookmarkTarget(package, operation, target, cancellationToken, out _)?.Start,
            "set-table-metadata" => ResolveTableTarget(package, target, cancellationToken)?.Table,
            "delete-block" => ResolveBlockTarget(package, operation, target, cancellationToken, out _)?.Block,
            "set-section-columns" or "set-section-orientation" => ResolveMainSectionTarget(package, target, cancellationToken)?.SectionProperties,
            "replace-text" => ResolveCellTarget(package, target, cancellationToken)?.Cell
                ?? ResolveParagraphTarget(package, operation, target, cancellationToken, out _)?.Paragraph,
            _ => ResolveParagraphTarget(package, operation, target, cancellationToken, out _)?.Paragraph
        };
    }

    private sealed record PreviewSnapshot(string? Before, string? Locator);

    private static PreviewSnapshot? CapturePreviewBefore(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        CancellationToken cancellationToken)
    {
        if (options.MaxPreviewChars <= 0)
        {
            return null;
        }

        string? before = ReadPreviewValue(package, operation, cancellationToken);
        if (before is null && !PreviewValueMayBeAbsent(operation.OperationName))
        {
            return null;
        }

        return new PreviewSnapshot(before, ResolvePreviewLocator(package, operation, cancellationToken));
    }

    private static bool PreviewValueMayBeAbsent(string operationName)
    {
        return operationName is "set-style" or "set-cell-shading" or "set-table-style" or "set-image-alt" or "delete-block" or "set-hyperlink-target" or "set-field-dirty" or "set-field-lock" or "set-field-code";
    }

    private static (string? Before, string? After, bool Truncated) FinalizePreview(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        PreviewSnapshot? before,
        bool success,
        CancellationToken cancellationToken)
    {
        if (before is null || !success || options.MaxPreviewChars <= 0)
        {
            return (null, null, false);
        }

        DocxPatchOperation readOperation = before.Locator is null
            ? operation
            : operation with { Fields = new Dictionary<string, string>(operation.Fields) { ["target"] = before.Locator } };
        string? after = ReadPreviewValue(package, readOperation, cancellationToken);

        int budget = options.MaxPreviewChars;
        bool truncated = false;
        string? boundedBefore = before.Before;
        if (boundedBefore is not null && boundedBefore.Length > budget)
        {
            boundedBefore = boundedBefore.Substring(0, budget);
            truncated = true;
        }

        string? boundedAfter = after;
        if (boundedAfter is not null && boundedAfter.Length > budget)
        {
            boundedAfter = boundedAfter.Substring(0, budget);
            truncated = true;
        }

        return (boundedBefore, boundedAfter, truncated);
    }

    private static string? ReadPreviewValue(
        OoxmlPackage package,
        DocxPatchOperation operation,
        CancellationToken cancellationToken)
    {
        if (operation.OperationName is not ("replace-text" or "replace-paragraph" or "set-cell" or "set-style" or "set-hyperlink-text" or "set-cell-shading" or "set-table-style" or "set-row-header" or "set-content-control-text" or "set-image-alt" or "set-field-result" or "set-field-code" or "set-comment-text" or "replace-bookmark-text" or "set-table-metadata" or "delete-block" or "set-section-columns" or "set-hyperlink-target" or "set-field-dirty" or "set-field-lock" or "set-section-orientation"))
        {
            return null;
        }

        string? target = operation.Fields.GetValueOrDefault("target");
        if (string.IsNullOrWhiteSpace(target))
        {
            return null;
        }

        if (operation.OperationName == "set-cell")
        {
            CellTarget? cellTarget = ResolveCellTarget(package, target, cancellationToken);
            return cellTarget is null ? null : ReadVisibleText(cellTarget.Cell);
        }

        if (operation.OperationName == "set-hyperlink-text")
        {
            HyperlinkTarget? hyperlinkTarget = ResolveHyperlinkTarget(package, target, cancellationToken);
            return hyperlinkTarget is null ? null : ReadVisibleText(hyperlinkTarget.Hyperlink);
        }

        if (operation.OperationName == "set-cell-shading")
        {
            CellTarget? cellTarget = ResolveCellTarget(package, target, cancellationToken);
            return cellTarget is null ? null : ReadCellShadingFill(cellTarget.Cell);
        }

        if (operation.OperationName == "set-table-style")
        {
            TableTarget? tableTarget = ResolveTableTarget(package, target, cancellationToken);
            return tableTarget is null ? null : ReadTableStyleId(tableTarget.Table);
        }

        if (operation.OperationName == "set-row-header")
        {
            RowTarget? rowTarget = ResolveRowTarget(package, target, cancellationToken);
            return rowTarget is null ? null : ReadTableRowHeader(rowTarget.Row).ToString().ToLowerInvariant();
        }

        if (operation.OperationName == "set-content-control-text")
        {
            ContentControlTarget? controlTarget = ResolveContentControlTarget(package, operation, target, cancellationToken, out _);
            return controlTarget is null ? null : ReadVisibleText(controlTarget.ContentControl);
        }

        if (operation.OperationName == "set-image-alt")
        {
            ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target, cancellationToken);
            if (imageTarget is null)
            {
                return null;
            }

            if (!TryGetImageDrawingContainer(imageTarget, target, operation, out XElement? imageContainer, out _))
            {
                return null;
            }

            return (string?)imageContainer.Element(OoxmlNs.Wp + "docPr")?.Attribute("descr");
        }

        if (operation.OperationName == "set-field-result")
        {
            FieldTarget? fieldTarget = ResolveFieldTarget(package, target, cancellationToken);
            return fieldTarget is null ? null : ReadVisibleText(fieldTarget.Element);
        }

        if (operation.OperationName == "set-field-code")
        {
            FieldTarget? codeTarget = ResolveFieldTarget(package, target, cancellationToken);
            return codeTarget is null ? null : (string?)codeTarget.Element.Attribute(OoxmlNs.W + "instr");
        }

        if (operation.OperationName == "set-comment-text")
        {
            CommentTarget? commentTarget = ResolveCommentTarget(package, target, operation, cancellationToken, out _);
            return commentTarget is null ? null : ReadVisibleText(commentTarget.Comment);
        }

        if (operation.OperationName == "replace-bookmark-text")
        {
            BookmarkTarget? bookmarkTarget = ResolveBookmarkTarget(package, operation, target, cancellationToken, out _);
            if (bookmarkTarget?.End is null)
            {
                return null;
            }

            XElement? startParagraph = bookmarkTarget.Start.Parent;
            XElement? endParagraph = bookmarkTarget.End.Parent;
            if (startParagraph is null || !ReferenceEquals(startParagraph, endParagraph))
            {
                return null;
            }

            XNode[] bookmarkNodes = bookmarkTarget.Start.NodesAfterSelf().TakeWhile(node => node != bookmarkTarget.End).ToArray();
            return ReadVisibleText(new XElement(OoxmlNs.W + "p", bookmarkNodes));
        }

        if (operation.OperationName == "set-table-metadata")
        {
            TableTarget? metadataTarget = ResolveTableTarget(package, target, cancellationToken);
            if (metadataTarget is null)
            {
                return null;
            }

            string? caption = ReadTableTextProperty(metadataTarget.Table, "tblCaption");
            string? description = ReadTableTextProperty(metadataTarget.Table, "tblDescription");
            return "caption=" + (caption ?? string.Empty) + "; description=" + (description ?? string.Empty);
        }

        if (operation.OperationName == "delete-block")
        {
            BlockTarget? blockTarget = ResolveBlockTarget(package, operation, target, cancellationToken, out _);
            return blockTarget is null ? null : ReadVisibleText(blockTarget.Block);
        }

        if (operation.OperationName == "set-section-columns")
        {
            SectionTarget? sectionTarget = ResolveMainSectionTarget(package, target, cancellationToken);
            return sectionTarget is null ? null : ReadSectionColumnCount(sectionTarget.SectionProperties).ToString();
        }

        if (operation.OperationName == "set-hyperlink-target")
        {
            HyperlinkTarget? hyperlinkTarget = ResolveHyperlinkTarget(package, target, cancellationToken);
            return hyperlinkTarget is null ? null : ReadHyperlinkTargetValue(package, hyperlinkTarget, cancellationToken);
        }

        if (operation.OperationName == "set-field-dirty")
        {
            FieldTarget? fieldTarget = ResolveFieldTarget(package, target, cancellationToken);
            return fieldTarget is null ? null : (string?)fieldTarget.Element.Attribute(OoxmlNs.W + "dirty");
        }

        if (operation.OperationName == "set-field-lock")
        {
            FieldTarget? lockTarget = ResolveFieldTarget(package, target, cancellationToken);
            return lockTarget is null ? null : (string?)lockTarget.Element.Attribute(OoxmlNs.W + "fldLock");
        }

        if (operation.OperationName == "set-section-orientation")
        {
            SectionTarget? orientationTarget = ResolveMainSectionTarget(package, target, cancellationToken);
            return orientationTarget is null ? null : ReadSectionOrientation(orientationTarget.SectionProperties).ToWireValue();
        }

        if (operation.OperationName == "replace-text"
            && DocxTargetId.TryParse(target, out DocxTargetId cellId)
            && cellId.Kind is DocxTargetKind.Cell or DocxTargetKind.MergeGroup)
        {
            CellTarget? cellTarget = ResolveCellTarget(package, target, cancellationToken);
            return cellTarget is null ? null : ReadVisibleText(cellTarget.Cell);
        }

        ParagraphTarget? paragraphTarget = ResolveParagraphTarget(package, operation, target, cancellationToken, out _);
        if (paragraphTarget is null)
        {
            return null;
        }

        if (operation.OperationName == "set-style")
        {
            return (string?)paragraphTarget.Paragraph.Element(OoxmlNs.W + "pPr")?.Element(OoxmlNs.W + "pStyle")?.Attribute(OoxmlNs.W + "val");
        }

        return ReadVisibleText(paragraphTarget.Paragraph);
    }

    private static string? ReadHyperlinkTargetValue(
        OoxmlPackage package,
        HyperlinkTarget hyperlinkTarget,
        CancellationToken cancellationToken)
    {
        string? relationshipId = (string?)hyperlinkTarget.Hyperlink.Attribute(OoxmlNs.R + "id");
        if (!string.IsNullOrWhiteSpace(relationshipId))
        {
            OoxmlRelationship? relationship = package
                .GetRelationships(hyperlinkTarget.PartName, cancellationToken)
                .FirstOrDefault(candidate => string.Equals(candidate.Id, relationshipId, StringComparison.Ordinal));
            return relationship?.Target;
        }

        return (string?)hyperlinkTarget.Hyperlink.Attribute(OoxmlNs.W + "anchor");
    }

    private static bool TryParseTargetSelector(
        string target,
        DocxPatchOperation operation,
        [NotNullWhen(true)] out TargetSelector? selector,
        [NotNullWhen(false)] out DocxDiagnostic? diagnostic)
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

        DocxTargetId? parsedTargetId = DocxTargetId.TryParse(target, out DocxTargetId parsedTargetIdValue) ? parsedTargetIdValue : null;
        selector = new ExplicitIdTargetSelector(target, parsedTargetId);
        return true;
    }

    // SPEC 10.6: semantic text selectors (text:, heading:) match with whitespace
    // normalized (every whitespace run collapses to one space, ends trimmed) while
    // staying case-sensitive. Guards such as expect-text keep exact matching.
    private static string NormalizeSelectorText(string value)
    {
        StringBuilder normalized = new(value.Length);
        bool pendingSpace = false;
        foreach (char ch in value)
        {
            if (char.IsWhiteSpace(ch))
            {
                pendingSpace = normalized.Length != 0;
                continue;
            }

            if (pendingSpace)
            {
                normalized.Append((char)32);
                pendingSpace = false;
            }

            normalized.Append(ch);
        }

        return normalized.ToString();
    }

    private static bool TryReadSelectorText(string value, out string selectorText)
    {
        selectorText = string.Empty;
        if (value.Length >= 2 && value[0] == '"' && value[^1] == '"')
        {
            selectorText = DocxPatchParser.UnescapePatchValue(value[1..^1]);
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
        OoxmlPackage package,
        XElement body,
        TargetSelector selector,
        DocxPatchOperation operation,
        CancellationToken cancellationToken,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        diagnostics = [];
        if (selector is HeadingTargetSelector headingSelector)
        {
            string normalizedHeading = NormalizeSelectorText(headingSelector.Text);
            return ResolveMainParagraphElementByPredicate(
                package,
                body,
                paragraph =>
                {
                    int? headingLevel = DocxHeadingLevels.GetHeadingLevel(package, paragraph, cancellationToken);
                    return headingLevel is not null &&
                        (headingSelector.Level is null || headingSelector.Level == headingLevel) &&
                        string.Equals(NormalizeSelectorText(ReadVisibleText(paragraph)), normalizedHeading, StringComparison.Ordinal);
                },
                selector.Raw,
                operation,
                cancellationToken,
                out diagnostics);
        }

        if (selector is ParagraphTextTargetSelector paragraphTextSelector)
        {
            string normalizedSearch = NormalizeSelectorText(paragraphTextSelector.Text);
            return ResolveMainParagraphElementByPredicate(
                package,
                body,
                paragraph => NormalizeSelectorText(ReadVisibleText(paragraph)).Contains(normalizedSearch, StringComparison.Ordinal),
                selector.Raw,
                operation,
                cancellationToken,
                out diagnostics);
        }

        if (selector is BookmarkTargetSelector bookmarkSelector)
        {
            return ResolveMainParagraphElementByPredicate(
                package,
                body,
                paragraph => ParagraphHasBookmark(paragraph, bookmarkSelector.Name),
                selector.Raw,
                operation,
                cancellationToken,
                out diagnostics);
        }

        if (selector is ContentControlTargetSelector contentControlSelector)
        {
            return ResolveMainParagraphElementByPredicate(
                package,
                body,
                paragraph => ParagraphHasContentControl(paragraph, contentControlSelector.Name),
                selector.Raw,
                operation,
                cancellationToken,
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
        OoxmlPackage package,
        XElement body,
        Func<XElement, bool> predicate,
        string rawSelector,
        DocxPatchOperation operation,
        CancellationToken cancellationToken,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        diagnostics = [];
        var matches = new List<ParagraphSelectorMatch>();
        int paragraphOrdinal = 0;
        foreach (DocxStoryBlocks.StoryBlock entry in DocxStoryBlocks.EnumeratePhysicalBlocks(body))
        {
            if (entry.Block.Name != OoxmlNs.W + "p")
            {
                continue;
            }

            paragraphOrdinal++;
            if (predicate(entry.Block))
            {
                matches.Add(new ParagraphSelectorMatch($"M.P{paragraphOrdinal:0000}", entry.Block));
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
                    rawSelector) with { MatchCount = matches.Count, CandidateIds = matches.Select(match => match.Id).ToArray() }
            ];
            return null;
        }

        if (matches.Count == 0)
        {
            (string suggestion, string[] candidates) = BuildNoMatchSuggestion(package, body, rawSelector, cancellationToken);
            diagnostics =
            [
                Diagnostic(
                    DocxSeverity.Error,
                    "E1201",
                    $"Selector matched 0 targets: {rawSelector}.{suggestion} Semantic selectors search only the main story; use explicit header/footer IDs for header and footer text.",
                    operation,
                    rawSelector) with { MatchCount = 0, CandidateIds = candidates }
            ];
            return null;
        }

        return matches[0].Paragraph;
    }

    private static (string Text, string[] CandidateIds) BuildNoMatchSuggestion(OoxmlPackage package, XElement body, string rawSelector, CancellationToken cancellationToken)
    {
        string query = ExtractSelectorQuery(rawSelector);
        if (rawSelector.StartsWith("heading:", StringComparison.Ordinal))
        {
            var headings = EnumerateMainParagraphs(body)
                .Select((match, order) => (match.Id, match.Paragraph, Order: order, HeadingLevel: DocxHeadingLevels.GetHeadingLevel(package, match.Paragraph, cancellationToken)))
                .Where(match => match.HeadingLevel is not null)
                .Select(match => (Label: match.Id + " heading level=" + match.HeadingLevel, match.Id, Score: ScoreParagraphMatch(ReadVisibleText(match.Paragraph), query), match.Order))
                .OrderByDescending(match => match.Score.Contains)
                .ThenByDescending(match => match.Score.Overlap)
                .ThenBy(match => match.Order)
                .Take(3)
                .ToArray();
            return headings.Length == 0
                ? (" No nearby paragraph targets are available.", [])
                : ("Nearby headings: " + string.Join(", ", headings.Select(match => match.Label)) + ".", headings.Select(match => match.Id).ToArray());
        }

        var ranked = EnumerateMainParagraphs(body)
            .Select((match, order) => (match.Id, Score: ScoreParagraphMatch(ReadVisibleText(match.Paragraph), query), Order: order))
            .OrderByDescending(match => match.Score.Contains)
            .ThenByDescending(match => match.Score.Overlap)
            .ThenBy(match => match.Order)
            .Take(3)
            .ToArray();
        string[] paragraphs = ranked.Select(match => match.Id).ToArray();
        string suggestion = paragraphs.Length == 0
            ? " No nearby paragraph targets are available."
            : "Nearby paragraphs: " + string.Join(", ", paragraphs) + ".";
        if (!string.IsNullOrEmpty(query) &&
            !ranked.Any(match => match.Score.Contains) &&
            body.Descendants(OoxmlNs.W + "tc").Any(cell => ScoreParagraphMatch(ReadVisibleText(cell), query).Contains))
        {
            suggestion += " Matching text is in a table cell; use find to locate its cell ID.";
        }

        return (suggestion, paragraphs);
    }

    private static string ExtractSelectorQuery(string rawSelector)
    {
        int first = rawSelector.IndexOf(":");
        if (first < 0)
        {
            return rawSelector.Trim();
        }

        string rest = rawSelector.Substring(first + 1).Trim();
        int second = rest.IndexOf(":");
        if (second > 0 && int.TryParse(rest.Substring(0, second).Trim(), out _))
        {
            return UnquoteSelectorText(rest.Substring(second + 1));
        }

        return UnquoteSelectorText(rest);
    }

    private static string UnquoteSelectorText(string value)
    {
        string trimmed = value.Trim();
        return trimmed.Length >= 2 && trimmed.StartsWith((char)34) && trimmed.EndsWith((char)34)
            ? trimmed.Substring(1, trimmed.Length - 2)
            : trimmed;
    }

    private static (bool Contains, int Overlap) ScoreParagraphMatch(string text, string query)
    {
        if (query.Length == 0)
        {
            return (false, 0);
        }

        bool contains = text.Contains(query, StringComparison.OrdinalIgnoreCase);
        HashSet<string> queryWords = new(
            query.Split([" ", "\t", "\n", "\r"], StringSplitOptions.RemoveEmptyEntries),
            StringComparer.OrdinalIgnoreCase);
        int overlap = 0;
        foreach (string word in text.Split([" ", "\t", "\n", "\r"], StringSplitOptions.RemoveEmptyEntries))
        {
            if (queryWords.Contains(word))
            {
                overlap++;
            }
        }

        return (contains, overlap);
    }

    private static IEnumerable<ParagraphSelectorMatch> EnumerateMainParagraphs(XElement body)
    {
        int paragraphOrdinal = 0;
        foreach (DocxStoryBlocks.StoryBlock entry in DocxStoryBlocks.EnumeratePhysicalBlocks(body))
        {
            if (entry.Block.Name != OoxmlNs.W + "p")
            {
                continue;
            }

            paragraphOrdinal++;
            yield return new ParagraphSelectorMatch($"M.P{paragraphOrdinal:0000}", entry.Block);
        }
    }

    private static string DescribeTargetKind(DocxTargetKind kind)
    {
        return kind switch
        {
            DocxTargetKind.Paragraph => "paragraph",
            DocxTargetKind.Table => "table",
            DocxTargetKind.Row => "table row",
            DocxTargetKind.Cell => "table cell",
            DocxTargetKind.MergeGroup => "merge group",
            DocxTargetKind.Image => "image",
            DocxTargetKind.Hyperlink => "hyperlink",
            DocxTargetKind.Field => "field",
            DocxTargetKind.Bookmark => "bookmark",
            DocxTargetKind.ContentControl => "content control",
            DocxTargetKind.Section => "section",
            _ => "unknown",
        };
    }

    private static string? WrongKindAlternative(DocxTargetKind kind)
    {
        return kind switch
        {
            DocxTargetKind.Cell => "Use set-cell with this cell ID to edit its text.",
            DocxTargetKind.MergeGroup => "Use set-cell with this merge-group ID to edit its text.",
            DocxTargetKind.Row => "Use a table-row operation with this row ID.",
            DocxTargetKind.Table => "Use a table operation with this table ID.",
            DocxTargetKind.Image => "Use an image operation with this image ID.",
            DocxTargetKind.Hyperlink => "Use a hyperlink operation with this hyperlink ID.",
            DocxTargetKind.Field => "Use a field operation with this field ID.",
            DocxTargetKind.Bookmark => "Use a bookmark operation with this bookmark ID.",
            DocxTargetKind.ContentControl => "Use a content-control operation with this content-control ID.",
            DocxTargetKind.Section => "Use a section operation with this section ID.",
            _ => null,
        };
    }

    private static string DescribeAcceptedKindsForOperation(DocxPatchOperation operation)
    {
        if (OperationsByName.TryGetValue(operation.OperationName, out OperationRegistration? registration) &&
            registration.AcceptedKinds.Length != 0)
        {
            return DescribeAcceptedKinds(registration.AcceptedKinds);
        }

        return "a paragraph ID such as M.P0001";
    }

    private static DocxDiagnostic WrongKindDiagnostic(DocxPatchOperation operation, string target, DocxTargetId parsed, string acceptedForms)
    {
        string message = "Target " + target + " is a " + DescribeTargetKind(parsed.Kind) + " ID. Operation " + operation.OperationName + " requires " + acceptedForms + ".";
        string? alternative = WrongKindAlternative(parsed.Kind);
        if (alternative is not null)
        {
            message += " " + alternative;
        }

        return Diagnostic(DocxSeverity.Error, "E1201", message, operation, target);
    }

    private static SectionTarget? ResolveMainSectionTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (!DocxTargetId.TryParse(target, out DocxTargetId sectionId) ||
            sectionId is not { Story: 'M', Kind: DocxTargetKind.Section } ||
            sectionId.Primary < 1)
        {
            return null;
        }

        XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
        // D01: sections bind to the input snapshot, not live ordinals.
        XElement? element = FindSnapshotElement(document, OoxmlNs.W + "sectPr", sectionId.ToWireValue());
        return element is null ? null : new SectionTarget(document, element);
    }

    private static ParagraphTarget? ResolveRelatedStoryParagraphTarget(
        OoxmlPackage package,
        string relationshipType,
        DocxTargetId paragraphId,
        CancellationToken cancellationToken)
    {
        if (paragraphId.StoryPart < 1 || paragraphId.Primary < 1)
        {
            return null;
        }

        ResolvedOoxmlRelationship? relationship = package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => relationship.Type == relationshipType)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .ElementAtOrDefault(paragraphId.StoryPart - 1);
        if (relationship is null)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, relationship.ResolvedTarget, cancellationToken, out XElement root);
        // D01: explicit IDs bind to the input snapshot.
        XElement? paragraph = FindSnapshotElement(document, OoxmlNs.W + "p", paragraphId.ToWireValue());
        return paragraph is null ? null : new ParagraphTarget(relationship.ResolvedTarget, document, paragraph);
    }

    private static BlockTarget? ResolveRelatedStoryBlockTarget(
        OoxmlPackage package,
        string relationshipType,
        DocxTargetId blockId,
        CancellationToken cancellationToken)
    {
        if (blockId.StoryPart < 1 || blockId.Primary < 1)
        {
            return null;
        }

        XName blockName = blockId.Kind == DocxTargetKind.Paragraph ? OoxmlNs.W + "p" : OoxmlNs.W + "tbl";
        ResolvedOoxmlRelationship? relationship = package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => relationship.Type == relationshipType)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .ElementAtOrDefault(blockId.StoryPart - 1);
        if (relationship is null)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, relationship.ResolvedTarget, cancellationToken, out XElement root);
        // D01: explicit IDs bind to the input snapshot.
        XElement? block = FindSnapshotElement(document, blockName, blockId.ToWireValue());
        return block is null ? null : new BlockTarget(relationship.ResolvedTarget, document, block);
    }

    private static bool IsSupportedTableTargetShape(string target)
    {
        return DocxTargetId.TryParse(target, out DocxTargetId tableId) && tableId.Kind == DocxTargetKind.Table;
    }

    private static bool IsSupportedRowTargetShape(string target)
    {
        return DocxTargetId.TryParse(target, out DocxTargetId rowId) && rowId.Kind == DocxTargetKind.Row;
    }

    private static bool IsSupportedCellTargetShape(string target)
    {
        return DocxTargetId.TryParse(target, out DocxTargetId cellId)
            && cellId.Kind is (DocxTargetKind.Cell or DocxTargetKind.MergeGroup);
    }

    private static TableTarget? ResolveTableTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (!DocxTargetId.TryParse(target, out DocxTargetId tableId) || tableId.Kind != DocxTargetKind.Table)
        {
            return null;
        }

        return ResolveTableTarget(package, tableId, cancellationToken);
    }

    private static TableTarget? ResolveTableTarget(
        OoxmlPackage package,
        DocxTargetId tableId,
        CancellationToken cancellationToken)
    {
        if (tableId.Story == 'M')
        {
            XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
            XElement? table = FindSnapshotElementInContainer(body, OoxmlNs.W + "tbl", tableId.ToWireValue());
            return table is null
                ? null
                : new TableTarget(package.MainDocumentPartName, document, table);
        }

        if (tableId.Story is 'H' or 'F')
        {
            string relationshipType = tableId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
            return ResolveRelatedStoryTableTarget(package, relationshipType, tableId, cancellationToken);
        }

        return null;
    }

    private static RowTarget? ResolveRowTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (IsAliasReference(target))
        {
            return ResolveAliasRowTarget(package, target, cancellationToken);
        }

        if (!DocxTargetId.TryParse(target, out DocxTargetId rowId) || rowId.Kind != DocxTargetKind.Row)
        {
            return null;
        }

        return ResolveRowTarget(package, rowId, cancellationToken);
    }

    private static RowTarget? ResolveRowTarget(
        OoxmlPackage package,
        DocxTargetId rowId,
        CancellationToken cancellationToken)
    {
        // D01: rows bind to the input snapshot.
        string wireId = rowId.ToWireValue();
        if (rowId.Story == 'M')
        {
            XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
            XElement? row = FindSnapshotElement(document, OoxmlNs.W + "tr", wireId);
            if (row is null)
            {
                return null;
            }

            XElement? table = row.Parent?.Name == OoxmlNs.W + "tbl" ? row.Parent : row.Ancestors(OoxmlNs.W + "tbl").FirstOrDefault();
            return table is null ? null : new RowTarget(package.MainDocumentPartName, document, table, row);
        }

        if (rowId.Story is 'H' or 'F')
        {
            string relationshipType = rowId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
            string? partName = ResolveRelatedStoryPartName(package, relationshipType, rowId.StoryPart, cancellationToken);
            if (partName is null)
            {
                return null;
            }

            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out XElement root);
            XElement? row = FindSnapshotElement(document, OoxmlNs.W + "tr", wireId);
            if (row is null)
            {
                return null;
            }

            XElement? table = row.Parent?.Name == OoxmlNs.W + "tbl" ? row.Parent : row.Ancestors(OoxmlNs.W + "tbl").FirstOrDefault();
            return table is null ? null : new RowTarget(partName, document, table, row);
        }

        return null;
    }

    private static CellTarget? ResolveCellTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (!DocxTargetId.TryParse(target, out DocxTargetId cellId) ||
            cellId.Kind is not (DocxTargetKind.Cell or DocxTargetKind.MergeGroup))
        {
            return null;
        }

        if (cellId.Kind == DocxTargetKind.MergeGroup)
        {
            TableTarget? mergeTableTarget = ResolveTableTarget(package, cellId.TableId, cancellationToken);
            return mergeTableTarget is null ? null : ResolveMergeGroupCellTarget(mergeTableTarget, cellId.Secondary);
        }

        // D01: cells bind to the input snapshot.
        string cellWireId = cellId.ToWireValue();
        if (cellId.Story == 'M')
        {
            XDocument cellDocument = LoadMainDocument(package, cancellationToken, out XElement cellBody);
            XElement? cell = FindSnapshotCellByVisualColumn(cellDocument, cellId);
            if (cell is null)
            {
                return null;
            }

            XElement? row = cell.Parent?.Name == OoxmlNs.W + "tr" ? cell.Parent : cell.Ancestors(OoxmlNs.W + "tr").FirstOrDefault();
            XElement? table = row?.Parent?.Name == OoxmlNs.W + "tbl" ? row.Parent : row?.Ancestors(OoxmlNs.W + "tbl").FirstOrDefault();
            return row is null || table is null ? null : new CellTarget(package.MainDocumentPartName, cellDocument, table, row, cell, cellId.Tertiary);
        }

        if (cellId.Story is 'H' or 'F')
        {
            string relationshipType = cellId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
            string? partName = ResolveRelatedStoryPartName(package, relationshipType, cellId.StoryPart, cancellationToken);
            if (partName is null)
            {
                return null;
            }

            XDocument cellDocument = LoadDocumentPart(package, partName, cancellationToken, out XElement cellRoot);
            XElement? cell = FindSnapshotCellByVisualColumn(cellDocument, cellId);
            if (cell is null)
            {
                return null;
            }

            XElement? row = cell.Parent?.Name == OoxmlNs.W + "tr" ? cell.Parent : cell.Ancestors(OoxmlNs.W + "tr").FirstOrDefault();
            XElement? table = row?.Parent?.Name == OoxmlNs.W + "tbl" ? row.Parent : row?.Ancestors(OoxmlNs.W + "tbl").FirstOrDefault();
            return row is null || table is null ? null : new CellTarget(partName, cellDocument, table, row, cell, cellId.Tertiary);
        }

        return null;
    }

    private static XElement? FindCellByVisualColumn(XElement row, int visualColumnIndex)
    {
        if (visualColumnIndex < 1)
        {
            return null;
        }

        int columnIndex = 1 + ReadTableRowGridOffset(row, "gridBefore");
        foreach (XElement cell in row.Elements(OoxmlNs.W + "tc"))
        {
            int columnSpan = ReadTableCellColumnSpan(cell);
            if (visualColumnIndex >= columnIndex && visualColumnIndex < columnIndex + columnSpan)
            {
                return cell;
            }

            columnIndex += columnSpan;
        }

        return null;
    }

    private static CellTarget? ResolveMergeGroupCellTarget(TableTarget tableTarget, int mergeGroupOrdinal)
    {
        if (mergeGroupOrdinal < 1)
        {
            return null;
        }

        string? tableSnapshotId = (string?)tableTarget.Table.Attribute(SnapshotIdName);
        if (tableSnapshotId is null ||
            !DocxTargetId.TryParse(tableSnapshotId, out DocxTargetId tableId))
        {
            return null;
        }

        var groupId = new DocxTargetId(tableId.Story, tableId.StoryPart, DocxTargetKind.MergeGroup, tableId.Primary, mergeGroupOrdinal, 0);
        string groupWireId = groupId.ToWireValue();
        XElement? root = tableTarget.Table
            .Descendants(OoxmlNs.W + "tc")
            .FirstOrDefault(cell => string.Equals((string?)cell.Attribute(SnapshotMergeGroupName), groupWireId, StringComparison.Ordinal));
        if (root is null)
        {
            return null;
        }

        XElement? row = root.Parent?.Name == OoxmlNs.W + "tr" ? root.Parent : root.Ancestors(OoxmlNs.W + "tr").FirstOrDefault();
        if (row is null)
        {
            return null;
        }

        int columnIndex = 1 + ReadTableRowGridOffset(row, "gridBefore");
        foreach (XElement cell in row.Elements(OoxmlNs.W + "tc"))
        {
            if (ReferenceEquals(cell, root))
            {
                break;
            }

            columnIndex += ReadTableCellColumnSpan(cell);
        }

        return new CellTarget(tableTarget.PartName, tableTarget.Document, tableTarget.Table, row, root, columnIndex);
    }

    private static MergeGroupRootState? FindActiveMergeGroup(
        IReadOnlyDictionary<int, MergeGroupRootState> activeVerticalMerges,
        int columnIndex,
        int columnSpan)
    {
        MergeGroupRootState? root = null;
        for (int column = columnIndex; column < columnIndex + columnSpan; column++)
        {
            if (!activeVerticalMerges.TryGetValue(column, out MergeGroupRootState? current))
            {
                return null;
            }

            root ??= current;
            if (!ReferenceEquals(root.Cell, current.Cell))
            {
                return null;
            }
        }

        return root;
    }

    private static void SetActiveMergeGroup(
        Dictionary<int, MergeGroupRootState> activeVerticalMerges,
        int columnIndex,
        int columnSpan,
        MergeGroupRootState root)
    {
        for (int column = columnIndex; column < columnIndex + columnSpan; column++)
        {
            activeVerticalMerges[column] = root;
        }
    }

    private static void RemoveActiveMergeGroups(
        Dictionary<int, MergeGroupRootState> activeVerticalMerges,
        int columnIndex,
        int columnSpan)
    {
        for (int column = columnIndex; column < columnIndex + columnSpan; column++)
        {
            activeVerticalMerges.Remove(column);
        }
    }

    private static TableTarget? ResolveRelatedStoryTableTarget(
        OoxmlPackage package,
        string relationshipType,
        DocxTargetId tableId,
        CancellationToken cancellationToken)
    {
        BlockTarget? blockTarget = ResolveRelatedStoryBlockTarget(package, relationshipType, tableId, cancellationToken);
        return blockTarget is null ? null : new TableTarget(blockTarget.PartName, blockTarget.Document, blockTarget.Block);
    }

    private static bool IsSupportedImageTargetShape(string target)
    {
        return DocxTargetId.TryParse(target, out DocxTargetId imageId) && imageId.Kind == DocxTargetKind.Image;
    }

    private static bool IsSupportedHyperlinkTargetShape(string target)
    {
        return DocxTargetId.TryParse(target, out DocxTargetId hyperlinkId) && hyperlinkId.Kind == DocxTargetKind.Hyperlink;
    }

    private static bool IsSupportedContentControlTargetShape(string target)
    {
        return (DocxTargetId.TryParse(target, out DocxTargetId contentcontrolId) && contentcontrolId.Kind == DocxTargetKind.ContentControl)
            || target.StartsWith("content-control:", StringComparison.Ordinal);
    }

    private static bool IsSupportedFieldTargetShape(string target)
    {
        return DocxTargetId.TryParse(target, out DocxTargetId fieldId) && fieldId.Kind == DocxTargetKind.Field;
    }

    private static bool IsSupportedBookmarkTargetShape(string target)
    {
        return (DocxTargetId.TryParse(target, out DocxTargetId bookmarkId) && bookmarkId.Kind == DocxTargetKind.Bookmark)
            || target.StartsWith("bookmark:", StringComparison.Ordinal);
    }

    private static ContentControlTarget? ResolveContentControlTarget(
        OoxmlPackage package,
        DocxPatchOperation operation,
        string target,
        CancellationToken cancellationToken,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        diagnostics = [];
        if (TryParseTargetSelector(target, operation, out TargetSelector? selector, out DocxDiagnostic? selectorDiagnostic)
            && selector is ContentControlTargetSelector contentControlSelector)
        {
            return ResolveContentControlTargetByName(package, operation, target, contentControlSelector.Name, cancellationToken, out diagnostics);
        }

        if (selectorDiagnostic is not null)
        {
            diagnostics = [selectorDiagnostic];
            return null;
        }

        if (!DocxTargetId.TryParse(target, out DocxTargetId controlId) || controlId.Kind != DocxTargetKind.ContentControl)
        {
            return null;
        }

        if (controlId.Story == 'M')
        {
            return FindContentControlTarget(package, package.MainDocumentPartName, controlId.Story, controlId.StoryPart, controlId.Primary, cancellationToken);
        }

        if (controlId.Story is 'H' or 'F')
        {
            string relationshipType = controlId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
            string? partName = ResolveRelatedStoryPartName(package, relationshipType, controlId.StoryPart, cancellationToken);
            return partName is null ? null : FindContentControlTarget(package, partName, controlId.Story, controlId.StoryPart, controlId.Primary, cancellationToken);
        }

        return null;
    }

    private static FieldTarget? ResolveFieldTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (!DocxTargetId.TryParse(target, out DocxTargetId fieldId) || fieldId.Kind != DocxTargetKind.Field)
        {
            return null;
        }

        if (fieldId.Story == 'M')
        {
            XDocument document = LoadMainDocument(package, cancellationToken, out _);
            XElement? field = FindSnapshotFieldElement(document, fieldId);
            return field is null
                ? null
                : new FieldTarget(package.MainDocumentPartName, document, field);
        }

        if (fieldId.Story is 'H' or 'F')
        {
            string relationshipType = fieldId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
            string? partName = ResolveRelatedStoryPartName(package, relationshipType, fieldId.StoryPart, cancellationToken);
            return partName is null ? null : FindFieldTarget(package, partName, fieldId.Story, fieldId.StoryPart, fieldId.Primary, cancellationToken);
        }

        return null;
    }

    private static bool IsAllFieldsTarget(string target)
    {
        return target is "all" or "all-fields";
    }

    private static IReadOnlyList<FieldTarget> ResolveAllFieldTargets(
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        var targets = new List<FieldTarget>();
        foreach (string partName in GetEditableStoryPartNames(package, cancellationToken))
        {
            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out XElement root);
            targets.AddRange(FindFields(root).Select(field => new FieldTarget(partName, document, field)));
        }

        return targets;
    }

    private static FieldTarget? FindFieldTarget(
        OoxmlPackage package,
        string partName,
        char story,
        int storyPart,
        int fieldOrdinal,
        CancellationToken cancellationToken,
        bool allowLiveFallback = false)
    {
        if (fieldOrdinal < 1)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, partName, cancellationToken, out XElement root);
        var snapshotId = new DocxTargetId(story, storyPart, DocxTargetKind.Field, fieldOrdinal, 0, 0);
        XElement? field = FindSnapshotFieldElement(document, snapshotId);
        if (field is null && allowLiveFallback)
        {
            field = FindField(root, fieldOrdinal);
        }

        return field is null ? null : new FieldTarget(partName, document, field);
    }

    private static XElement? FindSnapshotFieldElement(XDocument document, DocxTargetId fieldId)
    {
        string wireId = fieldId.ToWireValue();
        return FindSnapshotElement(document, OoxmlNs.W + "fldSimple", wireId)
            ?? FindSnapshotElement(document, OoxmlNs.W + "fldChar", wireId);
    }

    private static XElement? FindField(XElement root, int fieldOrdinal)
    {
        if (fieldOrdinal < 1)
        {
            return null;
        }

        return FindFields(root).ElementAtOrDefault(fieldOrdinal - 1);
    }

    private static IReadOnlyList<XElement> FindFields(XElement root)
    {
        var fields = new List<XElement>();
        var stack = new Stack<XElement>();
        foreach (XElement element in root.Descendants())
        {
            if (element.Name == OoxmlNs.W + "fldSimple")
            {
                fields.Add(element);
                continue;
            }

            if (element.Name != OoxmlNs.W + "fldChar")
            {
                continue;
            }

            string? fieldCharType = (string?)element.Attribute(OoxmlNs.W + "fldCharType");
            if (string.Equals(fieldCharType, "begin", StringComparison.Ordinal))
            {
                stack.Push(element);
            }
            else if (string.Equals(fieldCharType, "end", StringComparison.Ordinal) && stack.Count > 0)
            {
                XElement begin = stack.Pop();
                fields.Add(begin);
            }
        }

        foreach (XElement incompleteBegin in stack)
        {
            fields.Add(incompleteBegin);
        }

        return fields;
    }

    private static ContentControlTarget? FindContentControlTarget(
        OoxmlPackage package,
        string partName,
        char story,
        int storyPart,
        int contentControlOrdinal,
        CancellationToken cancellationToken)
    {
        if (contentControlOrdinal < 1)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
        var snapshotId = new DocxTargetId(story, storyPart, DocxTargetKind.ContentControl, contentControlOrdinal, 0, 0);
        XElement? contentControl = FindSnapshotElement(document, OoxmlNs.W + "sdt", snapshotId.ToWireValue());
        return contentControl is null ? null : new ContentControlTarget(partName, document, contentControl);
    }

    private static BookmarkTarget? ResolveBookmarkTarget(
        OoxmlPackage package,
        DocxPatchOperation operation,
        string target,
        CancellationToken cancellationToken,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        diagnostics = [];
        if (IsAliasReference(target))
        {
            return ResolveAliasBookmarkTarget(package, operation, target, cancellationToken, out diagnostics);
        }

        if (TryParseTargetSelector(target, operation, out TargetSelector? selector, out DocxDiagnostic? selectorDiagnostic)
            && selector is BookmarkTargetSelector bookmarkSelector)
        {
            return ResolveBookmarkTargetByName(package, operation, target, bookmarkSelector.Name, cancellationToken, out diagnostics);
        }

        if (selectorDiagnostic is not null)
        {
            diagnostics = [selectorDiagnostic];
            return null;
        }

        if (!DocxTargetId.TryParse(target, out DocxTargetId bookmarkId) || bookmarkId.Kind != DocxTargetKind.Bookmark)
        {
            return null;
        }

        if (bookmarkId.Story == 'M')
        {
            return FindBookmarkTarget(package, package.MainDocumentPartName, bookmarkId.Story, bookmarkId.StoryPart, bookmarkId.Primary, cancellationToken);
        }

        if (bookmarkId.Story is 'H' or 'F')
        {
            string relationshipType = bookmarkId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
            string? partName = ResolveRelatedStoryPartName(package, relationshipType, bookmarkId.StoryPart, cancellationToken);
            return partName is null ? null : FindBookmarkTarget(package, partName, bookmarkId.Story, bookmarkId.StoryPart, bookmarkId.Primary, cancellationToken);
        }

        return null;
    }

    private static BookmarkTarget? FindBookmarkTarget(
        OoxmlPackage package,
        string partName,
        char story,
        int storyPart,
        int bookmarkOrdinal,
        CancellationToken cancellationToken,
        bool allowLiveFallback = false)
    {
        if (bookmarkOrdinal < 1)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
        var snapshotId = new DocxTargetId(story, storyPart, DocxTargetKind.Bookmark, bookmarkOrdinal, 0, 0);
        XElement? start = FindSnapshotElement(document, OoxmlNs.W + "bookmarkStart", snapshotId.ToWireValue());
        if (start is null && allowLiveFallback && bookmarkOrdinal >= 1)
        {
            start = document
                .Descendants(OoxmlNs.W + "bookmarkStart")
                .ElementAtOrDefault(bookmarkOrdinal - 1);
        }
        string? ooxmlId = (string?)start?.Attribute(OoxmlNs.W + "id");
        XElement? end = ooxmlId is null
            ? null
            : document
                .Descendants(OoxmlNs.W + "bookmarkEnd")
                .FirstOrDefault(element => string.Equals((string?)element.Attribute(OoxmlNs.W + "id"), ooxmlId, StringComparison.Ordinal));
        return start is null ? null : new BookmarkTarget(partName, document, start, end);
    }

    private static ContentControlTarget? ResolveContentControlTargetByName(
        OoxmlPackage package,
        DocxPatchOperation operation,
        string target,
        string name,
        CancellationToken cancellationToken,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        diagnostics = [];
        var matches = new List<ContentControlTarget>();
        var matchIds = new List<string>();
        IReadOnlyDictionary<string, string> prefixes = DocxPartRoles.GetStoryPrefixes(package, cancellationToken);
        foreach (string partName in GetEditableStoryPartNames(package, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!prefixes.TryGetValue(partName, out string? prefix))
            {
                continue;
            }

            (char story, int storyPart) = DocxTargetId.ParseStoryPrefix(prefix);
            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
            int ordinal = 0;
            foreach (XElement contentControl in document.Descendants(OoxmlNs.W + "sdt"))
            {
                ordinal++;
                XElement? properties = contentControl.Element(OoxmlNs.W + "sdtPr");
                string? tag = (string?)properties
                    ?.Element(OoxmlNs.W + "tag")
                    ?.Attribute(OoxmlNs.W + "val");
                string? alias = (string?)properties
                    ?.Element(OoxmlNs.W + "alias")
                    ?.Attribute(OoxmlNs.W + "val");
                if (!string.Equals(tag, name, StringComparison.Ordinal) && !string.Equals(alias, name, StringComparison.Ordinal))
                {
                    continue;
                }

                var id = new DocxTargetId(story, storyPart, DocxTargetKind.ContentControl, ordinal, 0, 0);
                matches.Add(new ContentControlTarget(partName, document, contentControl));
                matchIds.Add(id.ToWireValue());
            }
        }

        if (matches.Count > 1)
        {
            diagnostics = [Diagnostic(DocxSeverity.Error, "E1202", "Content-control name matched " + matches.Count + " controls: " + string.Join(", ", matchIds) + ". Use an explicit content-control ID.", operation, target) with { MatchCount = matches.Count, CandidateIds = matchIds.ToArray() }];
            return null;
        }

        return matches.Count == 0 ? null : matches[0];
    }

    private static BookmarkTarget? ResolveBookmarkTargetByName(
        OoxmlPackage package,
        DocxPatchOperation operation,
        string target,
        string name,
        CancellationToken cancellationToken,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        diagnostics = [];
        var matches = new List<BookmarkTarget>();
        var matchIds = new List<string>();
        IReadOnlyDictionary<string, string> prefixes = DocxPartRoles.GetStoryPrefixes(package, cancellationToken);
        foreach (string partName in GetEditableStoryPartNames(package, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!prefixes.TryGetValue(partName, out string? prefix))
            {
                continue;
            }

            (char story, int storyPart) = DocxTargetId.ParseStoryPrefix(prefix);
            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
            int ordinal = 0;
            foreach (XElement start in document.Descendants(OoxmlNs.W + "bookmarkStart"))
            {
                ordinal++;
                if (!string.Equals((string?)start.Attribute(OoxmlNs.W + "name"), name, StringComparison.Ordinal))
                {
                    continue;
                }

                string? ooxmlId = (string?)start.Attribute(OoxmlNs.W + "id");
                XElement? end = ooxmlId is null
                    ? null
                    : document.Descendants(OoxmlNs.W + "bookmarkEnd").FirstOrDefault(element => string.Equals((string?)element.Attribute(OoxmlNs.W + "id"), ooxmlId, StringComparison.Ordinal));
                var id = new DocxTargetId(story, storyPart, DocxTargetKind.Bookmark, ordinal, 0, 0);
                matches.Add(new BookmarkTarget(partName, document, start, end));
                matchIds.Add(id.ToWireValue());
            }
        }

        if (matches.Count > 1)
        {
            diagnostics = [Diagnostic(DocxSeverity.Error, "E1202", "Bookmark name matched " + matches.Count + " bookmarks: " + string.Join(", ", matchIds) + ". Use an explicit bookmark ID.", operation, target) with { MatchCount = matches.Count, CandidateIds = matchIds.ToArray() }];
            return null;
        }

        return matches.Count == 0 ? null : matches[0];
    }

    private static HyperlinkTarget? ResolveHyperlinkTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (IsAliasReference(target))
        {
            return ResolveAliasHyperlinkTarget(package, target, cancellationToken);
        }

        if (!DocxTargetId.TryParse(target, out DocxTargetId hyperlinkId) || hyperlinkId.Kind != DocxTargetKind.Hyperlink)
        {
            return null;
        }

        if (hyperlinkId.Story == 'M')
        {
            return FindHyperlinkTarget(package, package.MainDocumentPartName, hyperlinkId.Story, hyperlinkId.StoryPart, hyperlinkId.Primary, cancellationToken);
        }

        if (hyperlinkId.Story is 'H' or 'F')
        {
            string relationshipType = hyperlinkId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
            string? partName = ResolveRelatedStoryPartName(package, relationshipType, hyperlinkId.StoryPart, cancellationToken);
            return partName is null ? null : FindHyperlinkTarget(package, partName, hyperlinkId.Story, hyperlinkId.StoryPart, hyperlinkId.Primary, cancellationToken);
        }

        return null;
    }

    private static HyperlinkTarget? FindHyperlinkTarget(
        OoxmlPackage package,
        string partName,
        char story,
        int storyPart,
        int hyperlinkOrdinal,
        CancellationToken cancellationToken,
        bool allowLiveFallback = false)
    {
        if (hyperlinkOrdinal < 1)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
        var snapshotId = new DocxTargetId(story, storyPart, DocxTargetKind.Hyperlink, hyperlinkOrdinal, 0, 0);
        XElement? hyperlink = FindSnapshotElement(document, OoxmlNs.W + "hyperlink", snapshotId.ToWireValue());
        if (hyperlink is null && allowLiveFallback && hyperlinkOrdinal >= 1)
        {
            hyperlink = document
                .Descendants(OoxmlNs.W + "hyperlink")
                .ElementAtOrDefault(hyperlinkOrdinal - 1);
        }

        return hyperlink is null ? null : new HyperlinkTarget(partName, document, hyperlink);
    }

    private static ImageBlipTarget? ResolveImageBlipTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (IsAliasReference(target))
        {
            return ResolveAliasImageBlipTarget(package, target, cancellationToken);
        }

        if (!DocxTargetId.TryParse(target, out DocxTargetId imageId) || imageId.Kind != DocxTargetKind.Image)
        {
            return null;
        }

        if (imageId.Story == 'M')
        {
            return FindImageBlipTarget(package, package.MainDocumentPartName, imageId.Story, imageId.StoryPart, imageId.Primary, cancellationToken);
        }

        if (imageId.Story is 'H' or 'F')
        {
            string relationshipType = imageId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
            string? partName = ResolveRelatedStoryPartName(package, relationshipType, imageId.StoryPart, cancellationToken);
            return partName is null ? null : FindImageBlipTarget(package, partName, imageId.Story, imageId.StoryPart, imageId.Primary, cancellationToken);
        }

        return null;
    }

    private sealed record ImageBlipEntry(XElement Blip, string RelationshipId, OoxmlPart Part);

    // C02: image identity is placement based. Each drawing placement in a
    // story gets its own public ID in document order. The media part is
    // a property of the placement, not the identity. Alt text, position,
    // and size belong to the placement. Repeated use of one media part
    // yields one ID per placement. Discovery, snapshot binding, mutation,
    // and reporting must use this same enumeration so an ID from discovery
    // always addresses the drawing it describes.
    // Public IDs for repeated media change under this contract.
    private static List<ImageBlipEntry> FindImageBlipEntries(
        OoxmlPackage package,
        string partName,
        XDocument document,
        CancellationToken cancellationToken)
    {
        IReadOnlyDictionary<string, OoxmlRelationship> relationships = package
            .GetRelationships(partName, cancellationToken)
            .ToDictionary(relationship => relationship.Id, StringComparer.Ordinal);
        var entries = new List<ImageBlipEntry>();
        XElement? storyRoot = document.Root;
        if (storyRoot is null)
        {
            return entries;
        }

        XElement storyContainer = storyRoot.Element(OoxmlNs.W + "body") ?? storyRoot;
        foreach (XElement blip in document.Descendants(OoxmlNs.A + "blip"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsBlipHiddenInFinalView(blip, storyContainer))
            {
                continue;
            }

            string? relationshipId = (string?)blip.Attribute(OoxmlNs.R + "embed");
            if (relationshipId is null ||
                !relationships.TryGetValue(relationshipId, out OoxmlRelationship? relationship) ||
                relationship.IsExternal ||
                relationship.ResolvedTarget is null)
            {
                continue;
            }

            OoxmlPart? part = package.GetPart(relationship.ResolvedTarget);
            if (part is null)
            {
                continue;
            }

            entries.Add(new ImageBlipEntry(blip, relationshipId, part));
        }

        return entries;
    }

    // C02: revision-hidden placements stay out of the placement enumeration.
    // Read and dump report images from Final-view-visible blocks only, so a
    // drawing inside a del/moveFrom-wrapped block owns no public ID. The patch
    // enumeration agrees: hidden blips are skipped, while blips in revision
    // runs inside a visible block still count on both sides.
    private static bool IsBlipHiddenInFinalView(XElement blip, XElement container)
    {
        XElement top = blip;
        while (top.Parent is not null && !ReferenceEquals(top.Parent, container))
        {
            top = top.Parent;
        }

        if (ReferenceEquals(top, container) || DocxStoryBlocks.IsStoryBlock(top))
        {
            return false;
        }

        return !DocxStoryBlocks.IsVisibleInView(top, DocxTextView.Final);
    }

    private static ImageBlipTarget? FindImageBlipTarget(
        OoxmlPackage package,
        string partName,
        char story,
        int storyPart,
        int imageOrdinal,
        CancellationToken cancellationToken,
        bool allowLiveFallback = false)
    {
        if (imageOrdinal < 1)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
        List<ImageBlipEntry> entries = FindImageBlipEntries(package, partName, document, cancellationToken);
        var snapshotId = new DocxTargetId(story, storyPart, DocxTargetKind.Image, imageOrdinal, 0, 0);
        string wireId = snapshotId.ToWireValue();
        ImageBlipEntry? entry = entries.FirstOrDefault(candidate => string.Equals((string?)candidate.Blip.Attribute(SnapshotIdName), wireId, StringComparison.Ordinal));
        entry ??= allowLiveFallback ? entries.ElementAtOrDefault(imageOrdinal - 1) : null;
        return entry is null ? null : new ImageBlipTarget(partName, document, entry.Blip, entry.RelationshipId, entry.Part);
    }

    private static string? ResolveRelatedStoryPartName(
        OoxmlPackage package,
        string relationshipType,
        int storyOrdinal,
        CancellationToken cancellationToken)
    {
        if (storyOrdinal < 1)
        {
            return null;
        }

        return package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => relationship.Type == relationshipType)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .ElementAtOrDefault(storyOrdinal - 1)
            ?.ResolvedTarget;
    }
}
