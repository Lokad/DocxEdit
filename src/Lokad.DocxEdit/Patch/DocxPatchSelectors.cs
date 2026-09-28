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

            diagnostics = [WrongKindDiagnostic(operation, target, paragraphId, "a paragraph ID such as M.P0001")];
            return null;
        }


        XDocument mainDocument = LoadMainDocument(package, cancellationToken, out XElement mainBody);
        XElement? selectedParagraph = ResolveMainParagraphElementBySelector(package, mainBody, selector, operation, cancellationToken, out diagnostics);
        return selectedParagraph is null
            ? null
            : new ParagraphTarget(package.MainDocumentPartName, mainDocument, selectedParagraph);
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
            return ResolveMainParagraphElementByPredicate(
                package,
                body,
                paragraph =>
                {
                    int? headingLevel = DocxHeadingLevels.GetHeadingLevel(package, paragraph, cancellationToken);
                    return headingLevel is not null &&
                        (headingSelector.Level is null || headingSelector.Level == headingLevel) &&
                        string.Equals(ReadVisibleText(paragraph), headingSelector.Text, StringComparison.Ordinal);
                },
                selector.Raw,
                operation,
                cancellationToken,
                out diagnostics);
        }

        if (selector is ParagraphTextTargetSelector paragraphTextSelector)
        {
            return ResolveMainParagraphElementByPredicate(
                package,
                body,
                paragraph => ReadVisibleText(paragraph).Contains(paragraphTextSelector.Text, StringComparison.Ordinal),
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
                    $"Selector matched 0 targets: {rawSelector}.{BuildNoMatchSuggestion(package, body, rawSelector, cancellationToken)}",
                    operation,
                    rawSelector)
            ];
            return null;
        }

        return matches[0].Paragraph;
    }

    private static string BuildNoMatchSuggestion(OoxmlPackage package, XElement body, string rawSelector, CancellationToken cancellationToken)
    {
        string[] suggestions = rawSelector.StartsWith("heading:", StringComparison.Ordinal)
            ? EnumerateMainParagraphs(body)
                .Select(match => (match.Id, HeadingLevel: DocxHeadingLevels.GetHeadingLevel(package, match.Paragraph, cancellationToken)))
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

        int mergeGroupIndex = 1;
        var activeVerticalMerges = new Dictionary<int, MergeGroupRootState>();
        foreach (XElement row in tableTarget.Table.Elements(OoxmlNs.W + "tr"))
        {
            int gridBefore = ReadTableRowGridOffset(row, "gridBefore");
            RemoveActiveMergeGroups(activeVerticalMerges, 1, gridBefore);
            int columnIndex = 1 + gridBefore;
            foreach (XElement cell in row.Elements(OoxmlNs.W + "tc"))
            {
                int columnSpan = ReadTableCellColumnSpan(cell);
                DocxVerticalMerge? verticalMerge = ReadTableCellVerticalMerge(cell);
                if (verticalMerge == DocxVerticalMerge.Restart)
                {
                    int currentMergeGroup = mergeGroupIndex++;
                    SetActiveMergeGroup(activeVerticalMerges, columnIndex, columnSpan, new MergeGroupRootState(row, cell, columnIndex));
                    if (currentMergeGroup == mergeGroupOrdinal)
                    {
                        return new CellTarget(tableTarget.PartName, tableTarget.Document, tableTarget.Table, row, cell, columnIndex);
                    }
                }
                else if (verticalMerge is not null)
                {
                    MergeGroupRootState? root = FindActiveMergeGroup(activeVerticalMerges, columnIndex, columnSpan);
                    if (root is null)
                    {
                        int currentMergeGroup = mergeGroupIndex++;
                        if (currentMergeGroup == mergeGroupOrdinal)
                        {
                            return new CellTarget(tableTarget.PartName, tableTarget.Document, tableTarget.Table, row, cell, columnIndex);
                        }

                        SetActiveMergeGroup(activeVerticalMerges, columnIndex, columnSpan, new MergeGroupRootState(row, cell, columnIndex));
                    }
                }
                else
                {
                    RemoveActiveMergeGroups(activeVerticalMerges, columnIndex, columnSpan);
                    if (columnSpan > 1)
                    {
                        int currentMergeGroup = mergeGroupIndex++;
                        if (currentMergeGroup == mergeGroupOrdinal)
                        {
                            return new CellTarget(tableTarget.PartName, tableTarget.Document, tableTarget.Table, row, cell, columnIndex);
                        }
                    }
                }

                columnIndex += columnSpan;
            }

            int gridAfter = ReadTableRowGridOffset(row, "gridAfter");
            RemoveActiveMergeGroups(activeVerticalMerges, columnIndex, gridAfter);
        }

        return null;
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
        return DocxTargetId.TryParse(target, out DocxTargetId contentcontrolId) && contentcontrolId.Kind == DocxTargetKind.ContentControl;
    }

    private static bool IsSupportedFieldTargetShape(string target)
    {
        return DocxTargetId.TryParse(target, out DocxTargetId fieldId) && fieldId.Kind == DocxTargetKind.Field;
    }

    private static bool IsSupportedBookmarkTargetShape(string target)
    {
        return DocxTargetId.TryParse(target, out DocxTargetId bookmarkId) && bookmarkId.Kind == DocxTargetKind.Bookmark;
    }

    private static ContentControlTarget? ResolveContentControlTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (!DocxTargetId.TryParse(target, out DocxTargetId controlId) || controlId.Kind != DocxTargetKind.ContentControl)
        {
            return null;
        }

        if (controlId.Story == 'M')
        {
            return FindContentControlTarget(package, package.MainDocumentPartName, controlId.Primary, cancellationToken);
        }

        if (controlId.Story is 'H' or 'F')
        {
            string relationshipType = controlId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
            string? partName = ResolveRelatedStoryPartName(package, relationshipType, controlId.StoryPart, cancellationToken);
            return partName is null ? null : FindContentControlTarget(package, partName, controlId.Primary, cancellationToken);
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
            XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
            XElement? field = FindField(body, fieldId.Primary);
            return field is null
                ? null
                : new FieldTarget(package.MainDocumentPartName, document, field);
        }

        if (fieldId.Story is 'H' or 'F')
        {
            string relationshipType = fieldId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
            string? partName = ResolveRelatedStoryPartName(package, relationshipType, fieldId.StoryPart, cancellationToken);
            return partName is null ? null : FindFieldTarget(package, partName, fieldId.Primary, cancellationToken);
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
        int fieldOrdinal,
        CancellationToken cancellationToken)
    {
        if (fieldOrdinal < 1)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, partName, cancellationToken, out XElement root);
        XElement? field = FindField(root, fieldOrdinal);
        return field is null ? null : new FieldTarget(partName, document, field);
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
        int contentControlOrdinal,
        CancellationToken cancellationToken)
    {
        if (contentControlOrdinal < 1)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
        XElement? contentControl = document
            .Descendants(OoxmlNs.W + "sdt")
            .ElementAtOrDefault(contentControlOrdinal - 1);
        return contentControl is null ? null : new ContentControlTarget(partName, document, contentControl);
    }

    private static BookmarkTarget? ResolveBookmarkTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (!DocxTargetId.TryParse(target, out DocxTargetId bookmarkId) || bookmarkId.Kind != DocxTargetKind.Bookmark)
        {
            return null;
        }

        if (bookmarkId.Story == 'M')
        {
            return FindBookmarkTarget(package, package.MainDocumentPartName, bookmarkId.Primary, cancellationToken);
        }

        if (bookmarkId.Story is 'H' or 'F')
        {
            string relationshipType = bookmarkId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
            string? partName = ResolveRelatedStoryPartName(package, relationshipType, bookmarkId.StoryPart, cancellationToken);
            return partName is null ? null : FindBookmarkTarget(package, partName, bookmarkId.Primary, cancellationToken);
        }

        return null;
    }

    private static BookmarkTarget? FindBookmarkTarget(
        OoxmlPackage package,
        string partName,
        int bookmarkOrdinal,
        CancellationToken cancellationToken)
    {
        if (bookmarkOrdinal < 1)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
        XElement? start = document
            .Descendants(OoxmlNs.W + "bookmarkStart")
            .ElementAtOrDefault(bookmarkOrdinal - 1);
        string? ooxmlId = (string?)start?.Attribute(OoxmlNs.W + "id");
        XElement? end = ooxmlId is null
            ? null
            : document
                .Descendants(OoxmlNs.W + "bookmarkEnd")
                .FirstOrDefault(element => string.Equals((string?)element.Attribute(OoxmlNs.W + "id"), ooxmlId, StringComparison.Ordinal));
        return start is null ? null : new BookmarkTarget(partName, document, start, end);
    }

    private static HyperlinkTarget? ResolveHyperlinkTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (!DocxTargetId.TryParse(target, out DocxTargetId hyperlinkId) || hyperlinkId.Kind != DocxTargetKind.Hyperlink)
        {
            return null;
        }

        if (hyperlinkId.Story == 'M')
        {
            return FindHyperlinkTarget(package, package.MainDocumentPartName, hyperlinkId.Primary, cancellationToken);
        }

        if (hyperlinkId.Story is 'H' or 'F')
        {
            string relationshipType = hyperlinkId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
            string? partName = ResolveRelatedStoryPartName(package, relationshipType, hyperlinkId.StoryPart, cancellationToken);
            return partName is null ? null : FindHyperlinkTarget(package, partName, hyperlinkId.Primary, cancellationToken);
        }

        return null;
    }

    private static HyperlinkTarget? FindHyperlinkTarget(
        OoxmlPackage package,
        string partName,
        int hyperlinkOrdinal,
        CancellationToken cancellationToken)
    {
        if (hyperlinkOrdinal < 1)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
        XElement? hyperlink = document
            .Descendants(OoxmlNs.W + "hyperlink")
            .ElementAtOrDefault(hyperlinkOrdinal - 1);
        return hyperlink is null ? null : new HyperlinkTarget(partName, document, hyperlink);
    }

    private static ImageBlipTarget? ResolveImageBlipTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (!DocxTargetId.TryParse(target, out DocxTargetId imageId) || imageId.Kind != DocxTargetKind.Image)
        {
            return null;
        }

        if (imageId.Story == 'M')
        {
            return FindImageBlipTarget(package, package.MainDocumentPartName, imageId.Primary, cancellationToken);
        }

        if (imageId.Story is 'H' or 'F')
        {
            string relationshipType = imageId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
            string? partName = ResolveRelatedStoryPartName(package, relationshipType, imageId.StoryPart, cancellationToken);
            return partName is null ? null : FindImageBlipTarget(package, partName, imageId.Primary, cancellationToken);
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
