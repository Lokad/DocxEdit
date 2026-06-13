using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocxEdit.Model;
using DocxEdit.Ooxml;

namespace DocxEdit;

internal static class DocxPatchEngine
{
    private const string SettingsContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
    private const string CommentsContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
    private const string CommentsExtendedContentType = "application/vnd.ms-word.commentsExtended+xml";

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
            TableOperationSnapshot? tableBefore = CaptureTableOperationSnapshot(package, operation, cancellationToken);
            bool supportsTrackedChanges = SupportsTrackedChangeOutput(operation.OperationName);
            if (options.TrackChanges == TrackChangesMode.Require && !supportsTrackedChanges)
            {
                operationDiagnostics.Add(Diagnostic(DocxSeverity.Error, "E6001", BuildUnsupportedTrackedOperationMessage(options.TrackChanges, operation), operation, operation.Fields.GetValueOrDefault("target")));
            }
            else
            {
                if (options.TrackChanges == TrackChangesMode.Suggest && !supportsTrackedChanges)
                {
                    operationDiagnostics.Add(Diagnostic(DocxSeverity.Warning, "W4001", BuildUnsupportedTrackedOperationMessage(options.TrackChanges, operation), operation, operation.Fields.GetValueOrDefault("target")));
                }

                operationDiagnostics.AddRange(operation.OperationName switch
                {
                    "replace-text" => ExecuteReplaceText(package, operation, options, apply, cancellationToken),
                    "replace-paragraph" => ExecuteReplaceParagraph(package, operation, options, apply, cancellationToken),
                    "insert-before" => ExecuteInsertBlock(package, operation, options, insertAfter: false, apply, cancellationToken),
                    "insert-after" => ExecuteInsertBlock(package, operation, options, insertAfter: true, apply, cancellationToken),
                    "delete-block" => ExecuteDeleteBlock(package, operation, options, apply, cancellationToken),
                    "set-style" => ExecuteSetStyle(package, operation, options, apply, cancellationToken),
                    "set-content-control-text" => ExecuteSetContentControlText(package, operation, apply, cancellationToken),
                    "set-content-control-checkbox" => ExecuteSetContentControlCheckbox(package, operation, apply, cancellationToken),
                    "set-content-control-choice" => ExecuteSetContentControlChoice(package, operation, apply, cancellationToken),
                    "set-content-control-date" => ExecuteSetContentControlDate(package, operation, apply, cancellationToken),
                    "replace-bookmark-text" => ExecuteReplaceBookmarkText(package, operation, apply, cancellationToken),
                    "rename-bookmark" => ExecuteRenameBookmark(package, operation, apply, cancellationToken),
                    "delete-bookmark" => ExecuteDeleteBookmark(package, operation, apply, cancellationToken),
                    "add-comment" => ExecuteAddComment(package, operation, options, apply, cancellationToken),
                    "set-comment-text" => ExecuteSetCommentText(package, operation, apply, cancellationToken),
                    "resolve-comment" => ExecuteSetCommentResolved(package, operation, resolved: true, apply, cancellationToken),
                    "reopen-comment" => ExecuteSetCommentResolved(package, operation, resolved: false, apply, cancellationToken),
                    "delete-comment" => ExecuteDeleteComment(package, operation, apply, cancellationToken),
                    "add-comment-reply" or "delete-comment-reply" => ExecuteUnsupportedCommentThreadOperation(operation),
                    "set-field-dirty" => ExecuteSetFieldFlag(package, operation, "dirty", "dirty", apply, cancellationToken),
                    "set-field-lock" => ExecuteSetFieldFlag(package, operation, "locked", "fldLock", apply, cancellationToken),
                    "set-field-code" => ExecuteSetFieldCode(package, operation, apply, cancellationToken),
                    "set-field-result" => ExecuteSetFieldResult(package, operation, apply, cancellationToken),
                    "set-hyperlink-target" => ExecuteSetHyperlinkTarget(package, operation, apply, cancellationToken),
                    "set-hyperlink-text" => ExecuteSetHyperlinkText(package, operation, apply, cancellationToken),
                    "insert-hyperlink-after" => ExecuteInsertHyperlinkAfter(package, operation, apply, cancellationToken),
                    "remove-hyperlink" => ExecuteRemoveHyperlink(package, operation, apply, cancellationToken),
                    "set-cell" => ExecuteSetCell(package, operation, options, apply, cancellationToken),
                    "set-table-style" => ExecuteSetTableStyle(package, operation, apply, cancellationToken),
                    "set-table-metadata" => ExecuteSetTableMetadata(package, operation, apply, cancellationToken),
                    "set-row-header" => ExecuteSetRowHeader(package, operation, apply, cancellationToken),
                    "append-row" => ExecuteAppendRow(package, operation, apply, cancellationToken),
                    "insert-row-before" => ExecuteInsertRow(package, operation, insertAfter: false, apply, cancellationToken),
                    "insert-row-after" => ExecuteInsertRow(package, operation, insertAfter: true, apply, cancellationToken),
                    "delete-row" => ExecuteDeleteRow(package, operation, apply, cancellationToken),
                    "replace-image" => ExecuteReplaceImage(package, operation, options, apply, cancellationToken),
                    "insert-image-after" => ExecuteInsertImageAfter(package, operation, options, apply, cancellationToken),
                    "set-image-alt" => ExecuteSetImageAlt(package, operation, apply, cancellationToken),
                    "set-image-metadata" => ExecuteSetImageMetadata(package, operation, apply, cancellationToken),
                    "set-image-size" => ExecuteSetImageSize(package, operation, apply, cancellationToken),
                    "set-image-wrap" => ExecuteSetImageWrap(package, operation, apply, cancellationToken),
                    "set-image-position" => ExecuteSetImagePosition(package, operation, apply, cancellationToken),
                    "set-image-crop" => ExecuteSetImageCrop(package, operation, apply, cancellationToken),
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
                operationDiagnostics)
            {
                AffectedTargets = operationSuccess ? BuildAffectedTargets(operation, tableBefore) : []
            });
        }

        bool shouldMarkFieldsDirty = apply &&
            options.MarkFieldsDirtyWhenEditing &&
            patch.Operations.Count != 0 &&
            patch.Operations.Any(operation => ShouldMarkFieldsDirtyAfterOperation(operation.OperationName)) &&
            diagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error);
        bool containsFieldsBeforeRefresh = shouldMarkFieldsDirty && PackageContainsFieldMarkup(package, cancellationToken);
        if (shouldMarkFieldsDirty)
        {
            diagnostics.AddRange(MarkFieldsDirty(package, cancellationToken));
            if (containsFieldsBeforeRefresh && diagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error))
            {
                diagnostics.Add(new DocxDiagnostic(
                    DocxSeverity.Warning,
                    "W5103",
                    "Document contains fields and was marked for Word-side field refresh; DocxEdit does not recalculate field results.",
                    Feature: "field",
                    Fallback: "word-refresh-required"));
            }
        }

        if (apply && diagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error))
        {
            diagnostics.AddRange(ValidateEditedPackage(package, cancellationToken));
        }

        return new PatchExecutionResult(diagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error), diagnostics, reports);
    }

    private static bool SupportsTrackedChangeOutput(string operationName)
    {
        return DocxHelp.TryGetPatchOperation(operationName, out DocxPatchOperationInfo operation) &&
            operation.TrackChangesSupport.StartsWith("tracked-", StringComparison.Ordinal);
    }

    private static string BuildUnsupportedTrackedOperationMessage(TrackChangesMode mode, DocxPatchOperation operation)
    {
        string operationName = operation.OperationName;
        string support = DocxHelp.TryGetPatchOperation(operationName, out DocxPatchOperationInfo operationInfo)
            ? operationInfo.TrackChangesSupport
            : "unclassified";

        return mode switch
        {
            TrackChangesMode.Require =>
                $"TrackChangesMode.Require cannot apply operation '{operationName}' as tracked output because its catalog support is '{support}'. Existing tracked-change markup is preserved, but this operation does not generate new revision markup.",
            TrackChangesMode.Suggest =>
                $"TrackChangesMode.Suggest will apply operation '{operationName}' directly because its catalog support is '{support}'. Existing tracked-change markup is preserved, but this operation does not generate new revision markup.",
            _ =>
                $"TrackChangesMode.{mode} does not generate tracked output for operation '{operationName}' because its catalog support is '{support}'."
        };
    }

    private static bool ShouldMarkFieldsDirtyAfterOperation(string operationName)
    {
        return operationName is not "set-field-result";
    }

    private static bool PackageContainsFieldMarkup(OoxmlPackage package, CancellationToken cancellationToken)
    {
        foreach (OoxmlPart part in package.Parts.Values
            .Where(part => part.Name.StartsWith("/word/", StringComparison.OrdinalIgnoreCase) &&
                part.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                !part.Name.Contains("/_rels/", StringComparison.OrdinalIgnoreCase))
            .OrderBy(part => part.Name, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            if (document.Descendants().Any(element =>
                element.Name == OoxmlNs.W + "fldSimple" ||
                element.Name == OoxmlNs.W + "fldChar" ||
                element.Name == OoxmlNs.W + "instrText"))
            {
                return true;
            }
        }

        return false;
    }

    private static TableOperationSnapshot? CaptureTableOperationSnapshot(
        OoxmlPackage package,
        DocxPatchOperation operation,
        CancellationToken cancellationToken)
    {
        string? target = operation.Fields.GetValueOrDefault("target");
        if (string.IsNullOrWhiteSpace(target))
        {
            return null;
        }

        return operation.OperationName switch
        {
            "set-cell" => CaptureCellSnapshot(package, target, cancellationToken),
            "append-row" => CaptureTableSnapshot(package, target, cancellationToken),
            "insert-row-before" or "insert-row-after" or "delete-row" => CaptureRowSnapshot(package, target, cancellationToken),
            _ => null
        };
    }

    private static TableOperationSnapshot? CaptureCellSnapshot(OoxmlPackage package, string target, CancellationToken cancellationToken)
    {
        CellTarget? cellTarget = ResolveCellTarget(package, target, cancellationToken);
        if (cellTarget is null)
        {
            return null;
        }

        XElement[] rows = cellTarget.Table.Elements(OoxmlNs.W + "tr").ToArray();
        XElement[] cells = cellTarget.Row.Elements(OoxmlNs.W + "tc").ToArray();
        int rowIndex = Array.IndexOf(rows, cellTarget.Row) + 1;
        int columnIndex = Array.IndexOf(cells, cellTarget.Cell) + 1;
        return CreateTableOperationSnapshot(target, cellTarget.Table, rowIndex, columnIndex, cells.Length);
    }

    private static TableOperationSnapshot? CaptureRowSnapshot(OoxmlPackage package, string target, CancellationToken cancellationToken)
    {
        RowTarget? rowTarget = ResolveRowTarget(package, target, cancellationToken);
        if (rowTarget is null)
        {
            return null;
        }

        XElement[] rows = rowTarget.Table.Elements(OoxmlNs.W + "tr").ToArray();
        int rowIndex = Array.IndexOf(rows, rowTarget.Row) + 1;
        int cellCount = rowTarget.Row.Elements(OoxmlNs.W + "tc").Count();
        return CreateTableOperationSnapshot(target, rowTarget.Table, rowIndex, columnIndex: null, cellCount);
    }

    private static TableOperationSnapshot? CaptureTableSnapshot(OoxmlPackage package, string target, CancellationToken cancellationToken)
    {
        TableTarget? tableTarget = ResolveTableTarget(package, target, cancellationToken);
        return tableTarget is null
            ? null
            : CreateTableOperationSnapshot(target, tableTarget.Table, rowIndex: null, columnIndex: null, cellCount: null);
    }

    private static TableOperationSnapshot CreateTableOperationSnapshot(
        string target,
        XElement table,
        int? rowIndex,
        int? columnIndex,
        int? cellCount)
    {
        int rowCount = table.Elements(OoxmlNs.W + "tr").Count();
        int columnCount = IsRectangular(table, out int rectangularColumnCount)
            ? rectangularColumnCount
            : table.Elements(OoxmlNs.W + "tr").Select(row => row.Elements(OoxmlNs.W + "tc").Count()).DefaultIfEmpty(0).Max();
        string? tableId = ExtractTableId(target);
        return new TableOperationSnapshot(target, tableId, rowIndex, columnIndex, rowCount, columnCount, cellCount);
    }

    private static IReadOnlyList<DocxPatchAffectedTarget> BuildAffectedTargets(DocxPatchOperation operation, TableOperationSnapshot? before)
    {
        if (before is null)
        {
            return [];
        }

        return operation.OperationName switch
        {
            "set-cell" => BuildSetCellAffectedTargets(before),
            "append-row" => BuildInsertedRowAffectedTargets(before, before.RowCountBefore + 1, operation.FieldValues.Count(field => field.Name == "cell"), "append"),
            "insert-row-before" => BuildInsertedRowAffectedTargets(before, before.RowIndex ?? 1, operation.FieldValues.Count(field => field.Name == "cell"), "insert"),
            "insert-row-after" => BuildInsertedRowAffectedTargets(before, (before.RowIndex ?? before.RowCountBefore) + 1, operation.FieldValues.Count(field => field.Name == "cell"), "insert"),
            "delete-row" => BuildDeletedRowAffectedTargets(before),
            _ => []
        };
    }

    private static IReadOnlyList<DocxPatchAffectedTarget> BuildSetCellAffectedTargets(TableOperationSnapshot before)
    {
        return
        [
            new(before.TargetId, "cell", "update")
            {
                ParentId = before.TableId,
                RowIndex = before.RowIndex,
                ColumnIndex = before.ColumnIndex,
                RowCountBefore = before.RowCountBefore,
                RowCountAfter = before.RowCountBefore,
                ColumnCount = before.ColumnCount
            }
        ];
    }

    private static IReadOnlyList<DocxPatchAffectedTarget> BuildInsertedRowAffectedTargets(
        TableOperationSnapshot before,
        int insertedRowIndex,
        int requestedCellCount,
        string action)
    {
        string tableId = before.TableId ?? before.TargetId;
        int cellCount = requestedCellCount == 0 ? before.ColumnCount : requestedCellCount;
        string rowId = $"{tableId}.R{insertedRowIndex:00}";
        var affected = new List<DocxPatchAffectedTarget>
        {
            new(rowId, "row", action)
            {
                ParentId = tableId,
                RowIndex = insertedRowIndex,
                RowCountBefore = before.RowCountBefore,
                RowCountAfter = before.RowCountBefore + 1,
                ColumnCount = before.ColumnCount,
                CellCount = cellCount
            }
        };
        for (int column = 1; column <= cellCount; column++)
        {
            affected.Add(new DocxPatchAffectedTarget($"{rowId}.C{column:00}", "cell", action)
            {
                ParentId = rowId,
                RowIndex = insertedRowIndex,
                ColumnIndex = column,
                RowCountBefore = before.RowCountBefore,
                RowCountAfter = before.RowCountBefore + 1,
                ColumnCount = before.ColumnCount
            });
        }

        return affected;
    }

    private static IReadOnlyList<DocxPatchAffectedTarget> BuildDeletedRowAffectedTargets(TableOperationSnapshot before)
    {
        string tableId = before.TableId ?? ExtractTableId(before.TargetId) ?? before.TargetId;
        int rowIndex = before.RowIndex ?? 1;
        int cellCount = before.CellCount ?? before.ColumnCount;
        var affected = new List<DocxPatchAffectedTarget>
        {
            new(before.TargetId, "row", "delete")
            {
                ParentId = tableId,
                RowIndex = rowIndex,
                RowCountBefore = before.RowCountBefore,
                RowCountAfter = before.RowCountBefore - 1,
                ColumnCount = before.ColumnCount,
                CellCount = cellCount
            }
        };
        for (int column = 1; column <= cellCount; column++)
        {
            affected.Add(new DocxPatchAffectedTarget($"{before.TargetId}.C{column:00}", "cell", "delete")
            {
                ParentId = before.TargetId,
                RowIndex = rowIndex,
                ColumnIndex = column,
                RowCountBefore = before.RowCountBefore,
                RowCountAfter = before.RowCountBefore - 1,
                ColumnCount = before.ColumnCount
            });
        }

        return affected;
    }

    private static string? ExtractTableId(string target)
    {
        int rowMarker = target.IndexOf(".R", StringComparison.Ordinal);
        return rowMarker < 0 ? target : target[..rowMarker];
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
        DocxEditOptions options,
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

        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges &&
            !TryValidateTrackedWholeParagraphReplacement(paragraphTarget.Paragraph, current, text!, style, out string? trackedUnsupportedReason))
        {
            if (!TrackUnsupportedShape(options, operation, target!, trackedUnsupportedReason!, diagnostics))
            {
                return diagnostics;
            }

            useTrackedChanges = false;
        }

        if (!apply)
        {
            return diagnostics;
        }

        if (useTrackedChanges)
        {
            ReplaceWholeParagraphTextWithTrackedChanges(package, paragraphTarget.Paragraph, current, text!, options, cancellationToken);
            SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
            return diagnostics;
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
        DocxEditOptions options,
        bool insertAfter,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        string? style = operation.Fields.GetValueOrDefault("style");
        bool copyParagraphProperties = ReadBooleanField(operation, "copy-paragraph-properties", diagnostics) ?? false;
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

        if (copyParagraphProperties && blockTarget.Block.Name != OoxmlNs.W + "p")
        {
            return [Diagnostic(DocxSeverity.Error, "E4307", $"Field 'copy-paragraph-properties' requires paragraph target '{target}'.", operation, target)];
        }

        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges && TextContainsTrackedUnsupportedCharacters(text!))
        {
            if (!TrackUnsupportedShape(options, operation, target!, "inserted paragraph text contains tabs or line breaks", diagnostics))
            {
                return diagnostics;
            }

            useTrackedChanges = false;
        }

        if (!apply)
        {
            return diagnostics;
        }

        XElement? paragraphProperties = copyParagraphProperties
            ? blockTarget.Block.Element(OoxmlNs.W + "pPr")
            : null;
        XElement paragraph = useTrackedChanges
            ? CreateTrackedInsertedParagraph(package, text!, style, paragraphProperties, options, cancellationToken)
            : CreateSimpleParagraph(text!, style, paragraphProperties);
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
        DocxEditOptions options,
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

        string current = ReadVisibleText(blockTarget.Block);
        if (expected is not null)
        {
            if (!string.Equals(current, expected, StringComparison.Ordinal))
            {
                return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected text does not match current text.", operation, target)];
            }
        }

        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges)
        {
            if (blockTarget.Block.Name != OoxmlNs.W + "p")
            {
                if (!TrackUnsupportedShape(options, operation, target!, "tracked block deletion is supported only for paragraph targets", diagnostics))
                {
                    return diagnostics;
                }

                useTrackedChanges = false;
            }
            else if (TryGetProtectedTextEditFeature(blockTarget.Block, out string protectedFeature))
            {
                if (!TrackUnsupportedShape(options, operation, target!, $"paragraph contains protected OOXML boundary '{protectedFeature}'", diagnostics))
                {
                    return diagnostics;
                }

                useTrackedChanges = false;
            }
            else if (TextContainsTrackedUnsupportedCharacters(current))
            {
                if (!TrackUnsupportedShape(options, operation, target!, "deleted paragraph text contains tabs or line breaks", diagnostics))
                {
                    return diagnostics;
                }

                useTrackedChanges = false;
            }
        }

        if (!apply)
        {
            return diagnostics;
        }

        if (useTrackedChanges)
        {
            ReplaceWholeParagraphTextWithTrackedChanges(package, blockTarget.Block, current, string.Empty, options, cancellationToken);
            SaveDocumentPart(package, blockTarget.PartName, blockTarget.Document);
            return diagnostics;
        }

        blockTarget.Block.Remove();
        SaveDocumentPart(package, blockTarget.PartName, blockTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetStyle(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
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

        if (IsTrackedMode(options))
        {
            SetParagraphStyleWithTrackedChange(package, paragraphTarget.Paragraph, styleId!, options, cancellationToken);
        }
        else
        {
            SetParagraphStyle(paragraphTarget.Paragraph, styleId!);
        }

        SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetContentControlText(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ContentControlTarget? controlTarget = ResolveContentControlTarget(package, target!, cancellationToken);
        if (controlTarget is null && !IsSupportedContentControlTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported content-control target '{target}'. Expected a content control ID such as M.CC0001 or H001.CC0001.", operation, target)];
        }

        if (controlTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        DocxDiagnostic? lockDiagnostic = ValidateContentControlUnlocked(controlTarget.ContentControl, operation, target!);
        if (lockDiagnostic is not null)
        {
            return [lockDiagnostic];
        }

        XElement? content = controlTarget.ContentControl.Element(OoxmlNs.W + "sdtContent");
        if (content is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E4310", $"Content control '{target}' has no editable content container.", operation, target)];
        }

        string current = ReadVisibleText(content);
        if (expected is not null && !string.Equals(current, expected, StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected content-control text does not match current text.", operation, target)];
        }

        bool isPlainText = IsPlainTextContentControl(controlTarget.ContentControl);
        if (!isPlainText)
        {
            if (!IsRichTextContentControl(controlTarget.ContentControl))
            {
                return [Diagnostic(DocxSeverity.Error, "E4310", $"Content control '{target}' is not a plain-text or rich-text content control.", operation, target)];
            }

            if (expected is null)
            {
                return [Diagnostic(DocxSeverity.Error, "E4205", $"Rich-text content control '{target}' requires expect-text before replacement.", operation, target)];
            }

            if (content.Elements().Any(element => element.Name != OoxmlNs.W + "p"))
            {
                return [Diagnostic(DocxSeverity.Error, "E4310", $"Rich-text content control '{target}' contains non-paragraph content.", operation, target)];
            }

            if (TryGetProtectedTextEditFeature(content, out string protectedFeature))
            {
                return [Diagnostic(DocxSeverity.Error, "E4310", $"Rich-text content control '{target}' contains protected OOXML boundary '{protectedFeature}'.", operation, target)];
            }
        }

        if (!apply)
        {
            return [];
        }

        ReplaceContentControlText(content, text!);
        SaveDocumentPart(package, controlTarget.PartName, controlTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetContentControlCheckbox(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        _ = ReadRequiredField(operation, "checked", diagnostics);
        bool? checkedValue = ReadBooleanField(operation, "checked", diagnostics);
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ContentControlTarget? controlTarget = ResolveContentControlTarget(package, target!, cancellationToken);
        if (controlTarget is null && !IsSupportedContentControlTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported content-control target '{target}'. Expected a content control ID such as M.CC0001 or H001.CC0001.", operation, target)];
        }

        if (controlTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        XElement? checkBox = controlTarget.ContentControl
            .Element(OoxmlNs.W + "sdtPr")
            ?.Element(OoxmlNs.W + "checkBox");
        if (checkBox is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E4310", $"Content control '{target}' is not a checkbox content control.", operation, target)];
        }

        DocxDiagnostic? lockDiagnostic = ValidateContentControlUnlocked(controlTarget.ContentControl, operation, target!);
        if (lockDiagnostic is not null)
        {
            return [lockDiagnostic];
        }

        XElement? content = controlTarget.ContentControl.Element(OoxmlNs.W + "sdtContent");
        if (content is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E4310", $"Content control '{target}' has no editable content container.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        SetCheckboxChecked(checkBox, checkedValue!.Value);
        ReplaceContentControlText(content, GetCheckboxDisplaySymbol(checkBox, checkedValue.Value));
        SaveDocumentPart(package, controlTarget.PartName, controlTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetContentControlChoice(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? value = operation.Fields.GetValueOrDefault("value");
        string? displayText = operation.Fields.GetValueOrDefault("display-text");
        if ((value is null) == (displayText is null))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Exactly one of 'value' or 'display-text' is required for set-content-control-choice.", operation, target));
        }

        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ContentControlTarget? controlTarget = ResolveContentControlTarget(package, target!, cancellationToken);
        if (controlTarget is null && !IsSupportedContentControlTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported content-control target '{target}'. Expected a content control ID such as M.CC0001 or H001.CC0001.", operation, target)];
        }

        if (controlTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        XElement? list = controlTarget.ContentControl
            .Element(OoxmlNs.W + "sdtPr")
            ?.Elements()
            .FirstOrDefault(element => element.Name == OoxmlNs.W + "dropDownList" || element.Name == OoxmlNs.W + "comboBox");
        if (list is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E4310", $"Content control '{target}' is not a dropdown or combo box content control.", operation, target)];
        }

        DocxDiagnostic? lockDiagnostic = ValidateContentControlUnlocked(controlTarget.ContentControl, operation, target!);
        if (lockDiagnostic is not null)
        {
            return [lockDiagnostic];
        }

        ContentControlChoice? choice = ResolveContentControlChoice(list, value, displayText);
        if (choice is null)
        {
            string selector = value is not null ? $"value '{value}'" : $"display-text '{displayText}'";
            return [Diagnostic(DocxSeverity.Error, "E4205", $"Content control '{target}' has no list item with {selector}.", operation, target)];
        }

        XElement? content = controlTarget.ContentControl.Element(OoxmlNs.W + "sdtContent");
        if (content is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E4310", $"Content control '{target}' has no editable content container.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        ReplaceContentControlText(content, choice.DisplayText);
        SaveDocumentPart(package, controlTarget.PartName, controlTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetContentControlDate(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? value = ReadRequiredField(operation, "value", diagnostics);
        string? displayText = operation.Fields.GetValueOrDefault("display-text");
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ContentControlTarget? controlTarget = ResolveContentControlTarget(package, target!, cancellationToken);
        if (controlTarget is null && !IsSupportedContentControlTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported content-control target '{target}'. Expected a content control ID such as M.CC0001 or H001.CC0001.", operation, target)];
        }

        if (controlTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        XElement? date = controlTarget.ContentControl
            .Element(OoxmlNs.W + "sdtPr")
            ?.Element(OoxmlNs.W + "date");
        if (date is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E4310", $"Content control '{target}' is not a date content control.", operation, target)];
        }

        DocxDiagnostic? lockDiagnostic = ValidateContentControlUnlocked(controlTarget.ContentControl, operation, target!);
        if (lockDiagnostic is not null)
        {
            return [lockDiagnostic];
        }

        XElement? content = controlTarget.ContentControl.Element(OoxmlNs.W + "sdtContent");
        if (content is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E4310", $"Content control '{target}' has no editable content container.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        SetContentControlDateValue(date, value!);
        ReplaceContentControlText(content, displayText ?? value!);
        SaveDocumentPart(package, controlTarget.PartName, controlTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteReplaceBookmarkText(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        BookmarkTarget? bookmarkTarget = ResolveBookmarkTarget(package, target!, cancellationToken);
        if (bookmarkTarget is null && !IsSupportedBookmarkTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported bookmark target '{target}'. Expected a bookmark ID such as M.B0001 or H001.B0001.", operation, target)];
        }

        if (bookmarkTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (bookmarkTarget.End is null ||
            bookmarkTarget.Start.Parent is null ||
            bookmarkTarget.Start.Parent != bookmarkTarget.End.Parent ||
            bookmarkTarget.Start.Parent.Name != OoxmlNs.W + "p")
        {
            return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark '{target}' is not a simple same-paragraph range.", operation, target)];
        }

        XNode[] nodes = bookmarkTarget.Start.NodesAfterSelf()
            .TakeWhile(node => node != bookmarkTarget.End)
            .ToArray();
        if (ContainsProtectedBookmarkReplacementNode(nodes, out string? protectedFeature))
        {
            return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark '{target}' replacement would remove protected OOXML boundary '{protectedFeature}'.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        foreach (XNode node in nodes)
        {
            node.Remove();
        }

        bookmarkTarget.Start.AddAfterSelf(CreateSimpleRun(text!));
        SaveDocumentPart(package, bookmarkTarget.PartName, bookmarkTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteRenameBookmark(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? name = ReadRequiredField(operation, "name", diagnostics);
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!IsValidBookmarkName(name!))
        {
            return [Diagnostic(DocxSeverity.Error, "E4205", "Field 'name' must be a non-empty bookmark name without whitespace.", operation, target)];
        }

        BookmarkTarget? bookmarkTarget = ResolveBookmarkTarget(package, target!, cancellationToken);
        if (bookmarkTarget is null && !IsSupportedBookmarkTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported bookmark target '{target}'. Expected a bookmark ID such as M.B0001 or H001.B0001.", operation, target)];
        }

        if (bookmarkTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        string? oldName = (string?)bookmarkTarget.Start.Attribute(OoxmlNs.W + "name");
        if (string.IsNullOrWhiteSpace(oldName))
        {
            return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark '{target}' has no current name.", operation, target)];
        }

        if (string.Equals(oldName, name, StringComparison.Ordinal))
        {
            return [];
        }

        if (BookmarkNameExists(bookmarkTarget.Document, bookmarkTarget.Start, name!))
        {
            return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark name '{name}' already exists in part '{bookmarkTarget.PartName}'.", operation, target)];
        }

        if (CountBookmarkName(bookmarkTarget.Document, oldName) > 1 &&
            HasInternalHyperlinkAnchor(bookmarkTarget.Document, oldName))
        {
            return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark '{target}' has duplicate name '{oldName}' and same-part hyperlink anchors; rename would be ambiguous.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        bookmarkTarget.Start.SetAttributeValue(OoxmlNs.W + "name", name);
        UpdateInternalHyperlinkAnchors(bookmarkTarget.Document, oldName, name!);
        SaveDocumentPart(package, bookmarkTarget.PartName, bookmarkTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteDeleteBookmark(
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

        BookmarkTarget? bookmarkTarget = ResolveBookmarkTarget(package, target!, cancellationToken);
        if (bookmarkTarget is null && !IsSupportedBookmarkTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported bookmark target '{target}'. Expected a bookmark ID such as M.B0001 or H001.B0001.", operation, target)];
        }

        if (bookmarkTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (bookmarkTarget.End is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark '{target}' is incomplete and cannot be deleted safely.", operation, target)];
        }

        string? name = (string?)bookmarkTarget.Start.Attribute(OoxmlNs.W + "name");
        if (!string.IsNullOrWhiteSpace(name) &&
            HasInternalHyperlinkAnchor(bookmarkTarget.Document, name!))
        {
            return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark '{target}' is referenced by same-part hyperlink anchors; update or remove those hyperlinks before deleting the bookmark.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        bookmarkTarget.Start.Remove();
        bookmarkTarget.End.Remove();
        SaveDocumentPart(package, bookmarkTarget.PartName, bookmarkTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteAddComment(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        string author = operation.Fields.GetValueOrDefault("author") ?? options.Author;
        string? initials = operation.Fields.GetValueOrDefault("initials");
        DateTimeOffset timestampUtc = options.TimestampUtc.ToUniversalTime();
        if (operation.Fields.TryGetValue("date", out string? date))
        {
            if (!DateTimeOffset.TryParse(date, null, System.Globalization.DateTimeStyles.RoundtripKind, out DateTimeOffset parsedDate))
            {
                diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", $"Field 'date' must be an ISO-8601 timestamp: {date}.", operation, target));
            }
            else
            {
                timestampUtc = parsedDate.ToUniversalTime();
            }
        }

        if (string.IsNullOrWhiteSpace(author))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Field 'author' must not be empty.", operation, target));
        }

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
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 paragraph targets: {target}.", operation, target)];
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

        DocxDiagnostic? commentsPartDiagnostic = ValidateExistingCommentsPart(package, cancellationToken);
        if (commentsPartDiagnostic is not null)
        {
            return [commentsPartDiagnostic];
        }

        if (!apply)
        {
            return [];
        }

        CommentsPartTarget commentsPart = ResolveOrCreateCommentsPart(package, cancellationToken);
        string commentId = AllocateCommentId(package, cancellationToken);
        commentsPart.Root.Add(CreateComment(commentId, text!, author, initials, timestampUtc));
        AddCommentAnchor(paragraphTarget.Paragraph, commentId);
        SaveDocumentPart(package, commentsPart.PartName, commentsPart.Document);
        SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetCommentText(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        CommentTarget? commentTarget = ResolveCommentTarget(package, target!, operation, cancellationToken, out DocxDiagnostic? diagnostic);
        if (diagnostic is not null)
        {
            return [diagnostic];
        }

        if (commentTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 comments: {target}.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        ReplaceCommentText(commentTarget.Comment, text!);
        SaveDocumentPart(package, commentTarget.PartName, commentTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetCommentResolved(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool resolved,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        CommentTarget? commentTarget = ResolveCommentTarget(package, target!, operation, cancellationToken, out DocxDiagnostic? diagnostic);
        if (diagnostic is not null)
        {
            return [diagnostic];
        }

        if (commentTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 comments: {target}.", operation, target)];
        }

        DocxDiagnostic? commentsExtendedDiagnostic = ValidateExistingCommentsExtendedPart(package, cancellationToken);
        if (commentsExtendedDiagnostic is not null)
        {
            return [commentsExtendedDiagnostic];
        }

        if (!apply)
        {
            return [];
        }

        CommentExtensionTarget? extensionTarget = ResolveOrCreateCommentExtensionTarget(
            package,
            commentTarget,
            operation,
            target!,
            cancellationToken,
            out diagnostic,
            out bool commentDocumentChanged);
        if (diagnostic is not null)
        {
            return [diagnostic];
        }

        if (extensionTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E4312", $"Comment '{target}' has no commentsExtended resolution metadata.", operation, target)];
        }

        extensionTarget.CommentExtension.SetAttributeValue(OoxmlNs.W15 + "done", resolved ? "1" : "0");
        if (commentDocumentChanged)
        {
            SaveDocumentPart(package, commentTarget.PartName, commentTarget.Document);
        }

        SaveDocumentPart(package, extensionTarget.PartName, extensionTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteDeleteComment(
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

        CommentTarget? commentTarget = ResolveCommentTarget(package, target!, operation, cancellationToken, out DocxDiagnostic? diagnostic);
        if (diagnostic is not null)
        {
            return [diagnostic];
        }

        if (commentTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 comments: {target}.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        string commentId = (string?)commentTarget.Comment.Attribute(OoxmlNs.W + "id") ?? string.Empty;
        RemoveCommentExtensionRecords(package, commentTarget.Comment, cancellationToken);
        commentTarget.Comment.Remove();
        SaveDocumentPart(package, commentTarget.PartName, commentTarget.Document);
        RemoveCommentAnchors(package, commentId, cancellationToken);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetFieldFlag(
        OoxmlPackage package,
        DocxPatchOperation operation,
        string fieldName,
        string attributeName,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        _ = ReadRequiredField(operation, fieldName, diagnostics);
        bool? value = ReadBooleanField(operation, fieldName, diagnostics);
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (IsAllFieldsTarget(target!))
        {
            IReadOnlyList<FieldTarget> fieldTargets = ResolveAllFieldTargets(package, cancellationToken);
            if (fieldTargets.Count == 0)
            {
                return [Diagnostic(DocxSeverity.Error, "E1201", "Selector matched 0 fields: all.", operation, target)];
            }

            if (!apply)
            {
                return [];
            }

            foreach (FieldTarget targetField in fieldTargets)
            {
                targetField.Element.SetAttributeValue(OoxmlNs.W + attributeName, value!.Value ? "true" : "false");
            }

            foreach (IGrouping<string, FieldTarget> partGroup in fieldTargets.GroupBy(fieldTarget => fieldTarget.PartName, StringComparer.Ordinal))
            {
                SaveDocumentPart(package, partGroup.Key, partGroup.First().Document);
            }

            return [];
        }

        FieldTarget? fieldTarget = ResolveFieldTarget(package, target!, cancellationToken);
        if (fieldTarget is null && !IsSupportedFieldTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported field target '{target}'. Expected 'all' or a field ID such as M.F0001 or H001.F0001.", operation, target)];
        }

        if (fieldTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 fields: {target}.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        fieldTarget.Element.SetAttributeValue(OoxmlNs.W + attributeName, value!.Value ? "true" : "false");
        SaveDocumentPart(package, fieldTarget.PartName, fieldTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetFieldCode(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? code = ReadRequiredField(operation, "code", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-code");
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        FieldTarget? fieldTarget = ResolveFieldTarget(package, target!, cancellationToken);
        if (fieldTarget is null && !IsSupportedFieldTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported field target '{target}'. Expected a field ID such as M.F0001 or H001.F0001.", operation, target)];
        }

        if (fieldTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 fields: {target}.", operation, target)];
        }

        if (fieldTarget.Element.Name != OoxmlNs.W + "fldSimple")
        {
            return [Diagnostic(DocxSeverity.Error, "E4313", $"Field code replacement for {target} currently supports only simple w:fldSimple fields.", operation, target)];
        }

        string current = NormalizeFieldCodeForGuard((string?)fieldTarget.Element.Attribute(OoxmlNs.W + "instr") ?? string.Empty);
        if (expected is not null && !string.Equals(current, NormalizeFieldCodeForGuard(expected), StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected field code does not match current code.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        fieldTarget.Element.SetAttributeValue(OoxmlNs.W + "instr", code!);
        fieldTarget.Element.SetAttributeValue(OoxmlNs.W + "dirty", "true");
        SaveDocumentPart(package, fieldTarget.PartName, fieldTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetFieldResult(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-result");
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        FieldTarget? fieldTarget = ResolveFieldTarget(package, target!, cancellationToken);
        if (fieldTarget is null && !IsSupportedFieldTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported field target '{target}'. Expected a field ID such as M.F0001 or H001.F0001.", operation, target)];
        }

        if (fieldTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 fields: {target}.", operation, target)];
        }

        if (fieldTarget.Element.Name != OoxmlNs.W + "fldSimple")
        {
            return [Diagnostic(DocxSeverity.Error, "E4313", $"Cached-result replacement for {target} currently supports only simple w:fldSimple fields.", operation, target)];
        }

        string current = ReadVisibleText(fieldTarget.Element);
        if (expected is not null && !string.Equals(current, expected, StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected field result does not match current result.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        ReplaceSimpleFieldResult(fieldTarget.Element, text!);
        SaveDocumentPart(package, fieldTarget.PartName, fieldTarget.Document);
        return [];
    }

    private static string NormalizeFieldCodeForGuard(string code)
    {
        return string.Join(
            " ",
            code.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));
    }

    private static void ReplaceSimpleFieldResult(XElement field, string text)
    {
        XElement? firstRunProperties = field
            .Elements(OoxmlNs.W + "r")
            .Elements(OoxmlNs.W + "rPr")
            .FirstOrDefault();
        field.RemoveNodes();
        var run = new XElement(OoxmlNs.W + "r");
        if (firstRunProperties is not null)
        {
            run.Add(new XElement(firstRunProperties));
        }

        foreach (XNode node in CreateTextNodes(text))
        {
            run.Add(node);
        }

        field.Add(run);
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetHyperlinkTarget(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        if (!TryReadHyperlinkDestination(operation, diagnostics, out string? uri, out string? anchor))
        {
            return diagnostics;
        }

        bool? history = ReadBooleanField(operation, "history", diagnostics);

        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        HyperlinkTarget? hyperlinkTarget = ResolveHyperlinkTarget(package, target!, cancellationToken);
        if (hyperlinkTarget is null && !IsSupportedHyperlinkTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported hyperlink target '{target}'. Expected a hyperlink ID such as M.L0001 or H001.L0001.", operation, target)];
        }

        if (hyperlinkTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        if (uri is not null)
        {
            SetExternalHyperlinkTarget(package, hyperlinkTarget, uri, cancellationToken);
        }
        else
        {
            SetInternalHyperlinkAnchor(package, hyperlinkTarget, anchor!);
        }

        if (operation.Fields.TryGetValue("tooltip", out string? tooltip))
        {
            hyperlinkTarget.Hyperlink.SetAttributeValue(OoxmlNs.W + "tooltip", tooltip);
        }

        if (operation.Fields.TryGetValue("target-frame", out string? targetFrame))
        {
            hyperlinkTarget.Hyperlink.SetAttributeValue(OoxmlNs.W + "tgtFrame", targetFrame);
        }

        if (history is not null)
        {
            hyperlinkTarget.Hyperlink.SetAttributeValue(OoxmlNs.W + "history", history.Value ? "true" : "false");
        }

        SaveDocumentPart(package, hyperlinkTarget.PartName, hyperlinkTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetHyperlinkText(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        HyperlinkTarget? hyperlinkTarget = ResolveHyperlinkTarget(package, target!, cancellationToken);
        if (hyperlinkTarget is null && !IsSupportedHyperlinkTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported hyperlink target '{target}'. Expected a hyperlink ID such as M.L0001 or H001.L0001.", operation, target)];
        }

        if (hyperlinkTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        ReplaceHyperlinkText(hyperlinkTarget.Hyperlink, text!);
        SaveDocumentPart(package, hyperlinkTarget.PartName, hyperlinkTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteInsertHyperlinkAfter(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        if (!TryReadHyperlinkDestination(operation, diagnostics, out string? uri, out string? anchor))
        {
            return diagnostics;
        }

        bool? history = ReadBooleanField(operation, "history", diagnostics);

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

        string? relationshipId = null;
        if (uri is not null)
        {
            relationshipId = OoxmlIds.AllocateRelationshipId(package.GetRelationships(blockTarget.PartName, cancellationToken).Select(relationship => relationship.Id));
            package.AddRelationship(blockTarget.PartName, relationshipId, OoxmlRelTypes.Hyperlink, uri, "External");
        }

        XElement paragraph = CreateHyperlinkParagraph(
            text!,
            relationshipId,
            anchor,
            operation.Fields.GetValueOrDefault("tooltip"),
            operation.Fields.GetValueOrDefault("target-frame"),
            history);
        blockTarget.Block.AddAfterSelf(paragraph);
        SaveDocumentPart(package, blockTarget.PartName, blockTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteRemoveHyperlink(
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

        HyperlinkTarget? hyperlinkTarget = ResolveHyperlinkTarget(package, target!, cancellationToken);
        if (hyperlinkTarget is null && !IsSupportedHyperlinkTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported hyperlink target '{target}'. Expected a hyperlink ID such as M.L0001 or H001.L0001.", operation, target)];
        }

        if (hyperlinkTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        string? relationshipId = (string?)hyperlinkTarget.Hyperlink.Attribute(OoxmlNs.R + "id");
        hyperlinkTarget.Hyperlink.ReplaceWith(hyperlinkTarget.Hyperlink.Nodes().ToArray());
        if (!string.IsNullOrWhiteSpace(relationshipId) &&
            !UsesHyperlinkRelationship(hyperlinkTarget.Document, relationshipId!))
        {
            package.RemoveRelationship(hyperlinkTarget.PartName, relationshipId!);
        }

        SaveDocumentPart(package, hyperlinkTarget.PartName, hyperlinkTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteUnsupportedCommentThreadOperation(DocxPatchOperation operation)
    {
        string? target = operation.Fields.GetValueOrDefault("target");
        return
        [
            Diagnostic(
                DocxSeverity.Error,
                "E4314",
                $"Operation '{operation.OperationName}' is not supported because threaded comment replies require commentsIds/threaded-comments metadata that DocxEdit does not safely model yet.",
                operation,
                target)
        ];
    }

    private static bool TryReadHyperlinkDestination(
        DocxPatchOperation operation,
        List<DocxDiagnostic> diagnostics,
        out string? uri,
        out string? anchor)
    {
        uri = operation.Fields.GetValueOrDefault("uri");
        anchor = operation.Fields.GetValueOrDefault("anchor");
        string? target = operation.Fields.GetValueOrDefault("target");
        if (string.IsNullOrWhiteSpace(uri) == string.IsNullOrWhiteSpace(anchor))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Exactly one of 'uri' or 'anchor' is required for hyperlink destination operations.", operation, target));
            return false;
        }

        if (uri is not null &&
            !TryValidateExternalHyperlinkUri(uri, out string? uriDiagnostic))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", uriDiagnostic!, operation, target));
            return false;
        }

        if (anchor is not null && (anchor.Length == 0 || anchor.Any(char.IsWhiteSpace)))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Field 'anchor' must be a non-empty bookmark anchor without whitespace.", operation, target));
            return false;
        }

        return true;
    }

    private static bool TryValidateExternalHyperlinkUri(string uri, out string? diagnostic)
    {
        diagnostic = null;
        if (!Uri.TryCreate(uri, UriKind.RelativeOrAbsolute, out Uri? parsed))
        {
            diagnostic = $"Field 'uri' must be a well-formed absolute http, https, or mailto URI; received malformed URI '{uri}'.";
            return false;
        }

        if (!parsed.IsAbsoluteUri)
        {
            diagnostic = $"Field 'uri' must be an absolute http, https, or mailto URI; relative hyperlink targets are not supported: {uri}.";
            return false;
        }

        if (parsed.Scheme is not ("http" or "https" or "mailto"))
        {
            diagnostic = $"Unsupported hyperlink URI scheme '{parsed.Scheme}'. Allowed schemes are http, https, and mailto.";
            return false;
        }

        return true;
    }

    private static void SetExternalHyperlinkTarget(
        OoxmlPackage package,
        HyperlinkTarget hyperlinkTarget,
        string uri,
        CancellationToken cancellationToken)
    {
        string? oldRelationshipId = (string?)hyperlinkTarget.Hyperlink.Attribute(OoxmlNs.R + "id");
        bool canReuseRelationship = !string.IsNullOrWhiteSpace(oldRelationshipId) &&
            CountHyperlinkRelationshipUses(hyperlinkTarget.Document, oldRelationshipId!) == 1;
        string relationshipId = canReuseRelationship
            ? oldRelationshipId!
            : OoxmlIds.AllocateRelationshipId(package.GetRelationships(hyperlinkTarget.PartName, cancellationToken).Select(relationship => relationship.Id));
        if (canReuseRelationship)
        {
            package.RemoveRelationship(hyperlinkTarget.PartName, relationshipId);
        }

        package.AddRelationship(hyperlinkTarget.PartName, relationshipId, OoxmlRelTypes.Hyperlink, uri, "External");
        hyperlinkTarget.Hyperlink.SetAttributeValue(OoxmlNs.R + "id", relationshipId);
        hyperlinkTarget.Hyperlink.SetAttributeValue(OoxmlNs.W + "anchor", null);
    }

    private static void SetInternalHyperlinkAnchor(
        OoxmlPackage package,
        HyperlinkTarget hyperlinkTarget,
        string anchor)
    {
        string? oldRelationshipId = (string?)hyperlinkTarget.Hyperlink.Attribute(OoxmlNs.R + "id");
        hyperlinkTarget.Hyperlink.SetAttributeValue(OoxmlNs.R + "id", null);
        hyperlinkTarget.Hyperlink.SetAttributeValue(OoxmlNs.W + "anchor", anchor);
        if (!string.IsNullOrWhiteSpace(oldRelationshipId) &&
            !UsesHyperlinkRelationship(hyperlinkTarget.Document, oldRelationshipId!))
        {
            package.RemoveRelationship(hyperlinkTarget.PartName, oldRelationshipId!);
        }
    }

    private static void ReplaceHyperlinkText(XElement hyperlink, string text)
    {
        var run = new XElement(OoxmlNs.W + "r");
        foreach (XNode node in CreateTextNodes(text))
        {
            run.Add(node);
        }

        hyperlink.RemoveNodes();
        hyperlink.Add(run);
    }

    private static XElement CreateHyperlinkParagraph(
        string text,
        string? relationshipId,
        string? anchor,
        string? tooltip,
        string? targetFrame,
        bool? history)
    {
        var hyperlink = new XElement(OoxmlNs.W + "hyperlink");
        if (relationshipId is not null)
        {
            hyperlink.SetAttributeValue(OoxmlNs.R + "id", relationshipId);
        }

        if (anchor is not null)
        {
            hyperlink.SetAttributeValue(OoxmlNs.W + "anchor", anchor);
        }

        if (!string.IsNullOrWhiteSpace(tooltip))
        {
            hyperlink.SetAttributeValue(OoxmlNs.W + "tooltip", tooltip);
        }

        if (!string.IsNullOrWhiteSpace(targetFrame))
        {
            hyperlink.SetAttributeValue(OoxmlNs.W + "tgtFrame", targetFrame);
        }

        if (history is not null)
        {
            hyperlink.SetAttributeValue(OoxmlNs.W + "history", history.Value ? "true" : "false");
        }

        ReplaceHyperlinkText(hyperlink, text);
        return new XElement(OoxmlNs.W + "p", hyperlink);
    }

    private static bool UsesHyperlinkRelationship(XDocument document, string relationshipId)
    {
        return CountHyperlinkRelationshipUses(document, relationshipId) > 0;
    }

    private static int CountHyperlinkRelationshipUses(XDocument document, string relationshipId)
    {
        return document
            .Descendants(OoxmlNs.W + "hyperlink")
            .Count(hyperlink => string.Equals((string?)hyperlink.Attribute(OoxmlNs.R + "id"), relationshipId, StringComparison.Ordinal));
    }

    private static bool IsPlainTextContentControl(XElement contentControl)
    {
        return contentControl
            .Element(OoxmlNs.W + "sdtPr")
            ?.Element(OoxmlNs.W + "text") is not null;
    }

    private static bool IsRichTextContentControl(XElement contentControl)
    {
        XElement? properties = contentControl.Element(OoxmlNs.W + "sdtPr");
        if (properties is null)
        {
            return true;
        }

        if (properties.Element(OoxmlNs.W + "richText") is not null)
        {
            return true;
        }

        return !properties.Elements().Any(element => element.Name.LocalName is
            "text" or
            "checkBox" or
            "dropDownList" or
            "comboBox" or
            "date" or
            "picture" or
            "group" or
            "repeatingSection" or
            "repeatingSectionItem");
    }

    private static DocxDiagnostic? ValidateContentControlUnlocked(
        XElement contentControl,
        DocxPatchOperation operation,
        string target)
    {
        XElement? lockElement = contentControl.Element(OoxmlNs.W + "sdtPr")?.Element(OoxmlNs.W + "lock");
        if (lockElement is null)
        {
            return null;
        }

        string lockValue = (string?)lockElement.Attribute(OoxmlNs.W + "val") ?? "locked";
        if (string.Equals(lockValue, "unlocked", StringComparison.Ordinal))
        {
            return null;
        }

        return Diagnostic(
            DocxSeverity.Error,
            "E4310",
            $"Content control '{target}' is locked by w:lock='{lockValue}'.",
            operation,
            target);
    }

    private static void SetCheckboxChecked(XElement checkBox, bool checkedValue)
    {
        XElement? checkedElement = checkBox.Element(OoxmlNs.W + "checked");
        if (checkedElement is null)
        {
            checkedElement = new XElement(OoxmlNs.W + "checked");
            checkBox.Add(checkedElement);
        }

        checkedElement.SetAttributeValue(OoxmlNs.W + "val", checkedValue ? "1" : "0");
    }

    private static string GetCheckboxDisplaySymbol(XElement checkBox, bool checkedValue)
    {
        string stateElementName = checkedValue ? "checkedState" : "uncheckedState";
        string? stateValue = (string?)checkBox
            .Element(OoxmlNs.W + stateElementName)
            ?.Attribute(OoxmlNs.W + "val");
        return TryDecodeStateSymbol(stateValue, out string? symbol)
            ? symbol!
            : char.ConvertFromUtf32(checkedValue ? 0x2612 : 0x2610);
    }

    private static bool TryDecodeStateSymbol(string? value, out string? symbol)
    {
        symbol = null;
        if (string.IsNullOrEmpty(value))
        {
            return false;
        }

        if (value.Length == 1)
        {
            symbol = value;
            return true;
        }

        if (value.All(Uri.IsHexDigit) &&
            int.TryParse(value, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out int codePoint) &&
            codePoint > 0)
        {
            symbol = char.ConvertFromUtf32(codePoint);
            return true;
        }

        return false;
    }

    private static ContentControlChoice? ResolveContentControlChoice(XElement list, string? value, string? displayText)
    {
        foreach (XElement item in list.Elements(OoxmlNs.W + "listItem"))
        {
            string? itemValue = (string?)item.Attribute(OoxmlNs.W + "value");
            string? itemDisplayText = (string?)item.Attribute(OoxmlNs.W + "displayText") ?? itemValue;
            if (value is not null && string.Equals(itemValue, value, StringComparison.Ordinal))
            {
                return new ContentControlChoice(itemDisplayText ?? string.Empty);
            }

            if (displayText is not null && string.Equals(itemDisplayText, displayText, StringComparison.Ordinal))
            {
                return new ContentControlChoice(itemDisplayText ?? string.Empty);
            }
        }

        return null;
    }

    private static void SetContentControlDateValue(XElement date, string value)
    {
        XElement? fullDate = date.Element(OoxmlNs.W + "fullDate");
        if (fullDate is null)
        {
            fullDate = new XElement(OoxmlNs.W + "fullDate");
            date.Add(fullDate);
        }

        fullDate.SetAttributeValue(OoxmlNs.W + "val", value);
    }

    private static void ReplaceContentControlText(XElement content, string text)
    {
        bool blockLevel = content.Elements(OoxmlNs.W + "p").Any();
        content.RemoveNodes();
        if (blockLevel)
        {
            content.Add(CreateSimpleParagraph(text));
        }
        else
        {
            content.Add(CreateSimpleRun(text));
        }
    }

    private static bool ContainsProtectedBookmarkReplacementNode(IEnumerable<XNode> nodes, out string? feature)
    {
        foreach (XElement element in nodes.OfType<XElement>().SelectMany(ElementAndDescendants))
        {
            if (ProtectedTextEditElements.TryGetValue(element.Name, out string? protectedFeature) &&
                element.Name != OoxmlNs.W + "bookmarkStart" &&
                element.Name != OoxmlNs.W + "bookmarkEnd")
            {
                feature = protectedFeature;
                return true;
            }
        }

        feature = null;
        return false;
    }

    private static bool IsValidBookmarkName(string name)
    {
        return !string.IsNullOrWhiteSpace(name) && !name.Any(char.IsWhiteSpace);
    }

    private static int CountBookmarkName(XDocument document, string name)
    {
        return document
            .Descendants(OoxmlNs.W + "bookmarkStart")
            .Count(bookmark => string.Equals((string?)bookmark.Attribute(OoxmlNs.W + "name"), name, StringComparison.Ordinal));
    }

    private static bool BookmarkNameExists(XDocument document, XElement excludedStart, string name)
    {
        return document
            .Descendants(OoxmlNs.W + "bookmarkStart")
            .Any(bookmark => bookmark != excludedStart &&
                string.Equals((string?)bookmark.Attribute(OoxmlNs.W + "name"), name, StringComparison.Ordinal));
    }

    private static bool HasInternalHyperlinkAnchor(XDocument document, string anchor)
    {
        return document
            .Descendants(OoxmlNs.W + "hyperlink")
            .Any(hyperlink => string.Equals((string?)hyperlink.Attribute(OoxmlNs.W + "anchor"), anchor, StringComparison.Ordinal));
    }

    private static void UpdateInternalHyperlinkAnchors(XDocument document, string oldName, string newName)
    {
        foreach (XElement hyperlink in document.Descendants(OoxmlNs.W + "hyperlink"))
        {
            if (string.Equals((string?)hyperlink.Attribute(OoxmlNs.W + "anchor"), oldName, StringComparison.Ordinal))
            {
                hyperlink.SetAttributeValue(OoxmlNs.W + "anchor", newName);
            }
        }
    }

    private static CommentTarget? ResolveCommentTarget(
        OoxmlPackage package,
        string target,
        DocxPatchOperation operation,
        CancellationToken cancellationToken,
        out DocxDiagnostic? diagnostic)
    {
        diagnostic = null;
        if (target.StartsWith("comment:", StringComparison.Ordinal))
        {
            string commentId = target["comment:".Length..].Trim();
            if (commentId.Length == 0)
            {
                diagnostic = Diagnostic(DocxSeverity.Error, "E1203", "Comment target must be comment:<id> or C001.C0001.", operation, target);
                return null;
            }

            return FindCommentTargetById(package, commentId, cancellationToken);
        }

        if (TryParseCommentBodyTarget(target, out int storyOrdinal, out int commentOrdinal))
        {
            string? partName = GetCommentsPartNames(package, cancellationToken).ElementAtOrDefault(storyOrdinal - 1);
            return partName is null ? null : FindCommentTargetByOrdinal(package, partName, commentOrdinal, cancellationToken);
        }

        diagnostic = Diagnostic(DocxSeverity.Error, "E1203", "Comment target must be comment:<id> or C001.C0001.", operation, target);
        return null;
    }

    private static bool TryParseCommentBodyTarget(string target, out int storyOrdinal, out int commentOrdinal)
    {
        storyOrdinal = 0;
        commentOrdinal = 0;
        if (target.Length != 10 ||
            target[0] != 'C' ||
            target[4..6] != ".C")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[6..], out commentOrdinal) &&
            storyOrdinal > 0 &&
            commentOrdinal > 0;
    }

    private static CommentTarget? FindCommentTargetById(
        OoxmlPackage package,
        string commentId,
        CancellationToken cancellationToken)
    {
        foreach (string partName in GetCommentsPartNames(package, cancellationToken))
        {
            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
            XElement? comment = document
                .Descendants(OoxmlNs.W + "comment")
                .FirstOrDefault(comment => string.Equals((string?)comment.Attribute(OoxmlNs.W + "id"), commentId, StringComparison.Ordinal));
            if (comment is not null)
            {
                return new CommentTarget(partName, document, comment);
            }
        }

        return null;
    }

    private static CommentTarget? FindCommentTargetByOrdinal(
        OoxmlPackage package,
        string partName,
        int commentOrdinal,
        CancellationToken cancellationToken)
    {
        if (commentOrdinal < 1)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
        XElement? comment = document
            .Descendants(OoxmlNs.W + "comment")
            .ElementAtOrDefault(commentOrdinal - 1);
        return comment is null ? null : new CommentTarget(partName, document, comment);
    }

    private static IReadOnlyList<string> GetCommentsPartNames(OoxmlPackage package, CancellationToken cancellationToken)
    {
        if (package.MainDocumentPartName is null)
        {
            return [];
        }

        return package
            .GetRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => !relationship.IsExternal && relationship.Type == OoxmlRelTypes.Comments && relationship.ResolvedTarget is not null)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .Select(relationship => relationship.ResolvedTarget!)
            .ToArray();
    }

    private static IReadOnlyList<string> GetCommentsExtendedPartNames(OoxmlPackage package, CancellationToken cancellationToken)
    {
        if (package.MainDocumentPartName is null)
        {
            return [];
        }

        var partNames = package
            .GetRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => !relationship.IsExternal && relationship.Type == OoxmlRelTypes.CommentsExtended && relationship.ResolvedTarget is not null)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .Select(relationship => relationship.ResolvedTarget!)
            .ToList();
        if (package.GetPart("/word/commentsExtended.xml") is not null &&
            !partNames.Contains("/word/commentsExtended.xml", StringComparer.Ordinal))
        {
            partNames.Add("/word/commentsExtended.xml");
        }

        return partNames;
    }

    private static DocxDiagnostic? ValidateExistingCommentsPart(OoxmlPackage package, CancellationToken cancellationToken)
    {
        foreach (string partName in GetCommentsPartNames(package, cancellationToken))
        {
            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out XElement root);
            if (root.Name != OoxmlNs.W + "comments")
            {
                return new DocxDiagnostic(DocxSeverity.Error, "E9001", $"Comments part '{partName}' has root '{root.Name.LocalName}', expected 'comments'.", PartName: partName);
            }
        }

        return null;
    }

    private static DocxDiagnostic? ValidateExistingCommentsExtendedPart(OoxmlPackage package, CancellationToken cancellationToken)
    {
        foreach (string partName in GetCommentsExtendedPartNames(package, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out XElement root);
            if (root.Name != OoxmlNs.W15 + "commentsEx")
            {
                return new DocxDiagnostic(DocxSeverity.Error, "E9001", $"commentsExtended part '{partName}' has root '{root.Name.LocalName}', expected 'commentsEx'.", PartName: partName);
            }
        }

        return null;
    }

    private static CommentsPartTarget ResolveOrCreateCommentsPart(OoxmlPackage package, CancellationToken cancellationToken)
    {
        string? commentsPartName = GetCommentsPartNames(package, cancellationToken).FirstOrDefault();
        if (commentsPartName is null)
        {
            commentsPartName = "/word/comments.xml";
            if (package.GetPart(commentsPartName) is null)
            {
                byte[] bytes = Encoding.UTF8.GetBytes("""
                    <?xml version="1.0" encoding="utf-8"?>
                    <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />
                    """);
                package.AddPart(commentsPartName, CommentsContentType, bytes);
            }

            string relationshipId = OoxmlIds.AllocateRelationshipId(package
                .GetRelationships(package.MainDocumentPartName!, cancellationToken)
                .Select(relationship => relationship.Id));
            package.AddRelationship(package.MainDocumentPartName!, relationshipId, OoxmlRelTypes.Comments, GetRelativeRelationshipTarget(package.MainDocumentPartName!, commentsPartName));
        }
        else if (package.GetPart(commentsPartName) is null)
        {
            byte[] bytes = Encoding.UTF8.GetBytes("""
                <?xml version="1.0" encoding="utf-8"?>
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />
                """);
            package.AddPart(commentsPartName, CommentsContentType, bytes);
        }

        XDocument document = LoadDocumentPart(package, commentsPartName, cancellationToken, out XElement root);
        if (root.Name != OoxmlNs.W + "comments")
        {
            throw new InvalidDataException($"Comments part '{commentsPartName}' has root '{root.Name.LocalName}', expected 'comments'.");
        }

        return new CommentsPartTarget(commentsPartName, document, root);
    }

    private static string AllocateCommentId(OoxmlPackage package, CancellationToken cancellationToken)
    {
        int maxId = -1;
        foreach (string partName in GetCommentsPartNames(package, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            foreach (string id in document.Descendants(OoxmlNs.W + "comment").Select(comment => (string?)comment.Attribute(OoxmlNs.W + "id")).OfType<string>())
            {
                maxId = Math.Max(maxId, ParseNonNegativeIdOrDefault(id, -1));
            }
        }

        foreach (OoxmlPart part in package.Parts.Values
            .Where(part => part.Name.StartsWith("/word/", StringComparison.OrdinalIgnoreCase) &&
                part.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(part => part.Name, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            foreach (string id in document
                .Descendants()
                .Where(element => element.Name == OoxmlNs.W + "commentRangeStart" ||
                    element.Name == OoxmlNs.W + "commentRangeEnd" ||
                    element.Name == OoxmlNs.W + "commentReference")
                .Select(element => (string?)element.Attribute(OoxmlNs.W + "id"))
                .OfType<string>())
            {
                maxId = Math.Max(maxId, ParseNonNegativeIdOrDefault(id, -1));
            }
        }

        return (maxId + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    private static int ParseNonNegativeIdOrDefault(string value, int fallback)
    {
        return int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int id) && id >= 0
            ? id
            : fallback;
    }

    private static CommentExtensionTarget? ResolveCommentExtensionTarget(
        OoxmlPackage package,
        XElement comment,
        DocxPatchOperation operation,
        string target,
        CancellationToken cancellationToken,
        out DocxDiagnostic? diagnostic)
    {
        diagnostic = null;
        string? paraId = ReadCommentParaId(comment);
        if (string.IsNullOrWhiteSpace(paraId))
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E4312", $"Comment '{target}' has no w15:paraId, so resolution state cannot be edited safely.", operation, target);
            return null;
        }

        IReadOnlyList<string> partNames = GetCommentsExtendedPartNames(package, cancellationToken);
        if (partNames.Count == 0)
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E4312", $"Comment '{target}' has no commentsExtended part, so resolution state cannot be edited safely.", operation, target);
            return null;
        }

        foreach (string partName in partNames)
        {
            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
            XElement? commentExtension = document
                .Descendants(OoxmlNs.W15 + "commentEx")
                .FirstOrDefault(element => string.Equals((string?)element.Attribute(OoxmlNs.W15 + "paraId"), paraId, StringComparison.Ordinal));
            if (commentExtension is not null)
            {
                return new CommentExtensionTarget(partName, document, commentExtension);
            }
        }

        return null;
    }

    private static CommentExtensionTarget? ResolveOrCreateCommentExtensionTarget(
        OoxmlPackage package,
        CommentTarget commentTarget,
        DocxPatchOperation operation,
        string target,
        CancellationToken cancellationToken,
        out DocxDiagnostic? diagnostic,
        out bool commentDocumentChanged)
    {
        diagnostic = null;
        commentDocumentChanged = false;
        string? paraId = ReadCommentParaId(commentTarget.Comment);
        if (string.IsNullOrWhiteSpace(paraId))
        {
            XElement? firstParagraph = commentTarget.Comment.Elements(OoxmlNs.W + "p").FirstOrDefault();
            if (firstParagraph is null)
            {
                diagnostic = Diagnostic(DocxSeverity.Error, "E4312", $"Comment '{target}' has no body paragraph for resolution metadata.", operation, target);
                return null;
            }

            paraId = AllocateCommentParaId(package, cancellationToken);
            EnsureNamespaceDeclaration(commentTarget.Document.Root, "w15", OoxmlNs.W15);
            firstParagraph.SetAttributeValue(OoxmlNs.W15 + "paraId", paraId);
            commentDocumentChanged = true;
        }

        foreach (string partName in GetCommentsExtendedPartNames(package, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
            XElement? commentExtension = document
                .Descendants(OoxmlNs.W15 + "commentEx")
                .FirstOrDefault(element => string.Equals((string?)element.Attribute(OoxmlNs.W15 + "paraId"), paraId, StringComparison.Ordinal));
            if (commentExtension is not null)
            {
                return new CommentExtensionTarget(partName, document, commentExtension);
            }
        }

        CommentsExtendedPartTarget extensionPart = ResolveOrCreateCommentsExtendedPart(package, cancellationToken);
        var extension = new XElement(OoxmlNs.W15 + "commentEx", new XAttribute(OoxmlNs.W15 + "paraId", paraId));
        extensionPart.Root.Add(extension);
        return new CommentExtensionTarget(extensionPart.PartName, extensionPart.Document, extension);
    }

    private static CommentsExtendedPartTarget ResolveOrCreateCommentsExtendedPart(OoxmlPackage package, CancellationToken cancellationToken)
    {
        string? commentsExtendedPartName = GetCommentsExtendedPartNames(package, cancellationToken).FirstOrDefault();
        if (commentsExtendedPartName is null)
        {
            commentsExtendedPartName = "/word/commentsExtended.xml";
            if (package.GetPart(commentsExtendedPartName) is null)
            {
                byte[] bytes = Encoding.UTF8.GetBytes("""
                    <?xml version="1.0" encoding="utf-8"?>
                    <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" />
                    """);
                package.AddPart(commentsExtendedPartName, CommentsExtendedContentType, bytes);
            }

            string relationshipId = OoxmlIds.AllocateRelationshipId(package
                .GetRelationships(package.MainDocumentPartName!, cancellationToken)
                .Select(relationship => relationship.Id));
            package.AddRelationship(package.MainDocumentPartName!, relationshipId, OoxmlRelTypes.CommentsExtended, GetRelativeRelationshipTarget(package.MainDocumentPartName!, commentsExtendedPartName));
        }
        else if (package.GetPart(commentsExtendedPartName) is null)
        {
            byte[] bytes = Encoding.UTF8.GetBytes("""
                <?xml version="1.0" encoding="utf-8"?>
                <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" />
                """);
            package.AddPart(commentsExtendedPartName, CommentsExtendedContentType, bytes);
        }

        XDocument document = LoadDocumentPart(package, commentsExtendedPartName, cancellationToken, out XElement root);
        if (root.Name != OoxmlNs.W15 + "commentsEx")
        {
            throw new InvalidDataException($"commentsExtended part '{commentsExtendedPartName}' has root '{root.Name.LocalName}', expected 'commentsEx'.");
        }

        EnsureNamespaceDeclaration(root, "w15", OoxmlNs.W15);
        return new CommentsExtendedPartTarget(commentsExtendedPartName, document, root);
    }

    private static string AllocateCommentParaId(OoxmlPackage package, CancellationToken cancellationToken)
    {
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string partName in GetCommentsPartNames(package, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            foreach (string paraId in document
                .Descendants(OoxmlNs.W + "p")
                .Select(paragraph => (string?)paragraph.Attribute(OoxmlNs.W15 + "paraId"))
                .Where(paraId => !string.IsNullOrWhiteSpace(paraId))
                .Select(paraId => paraId!))
            {
                used.Add(paraId);
            }
        }

        foreach (string partName in GetCommentsExtendedPartNames(package, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            foreach (XElement extension in document.Descendants(OoxmlNs.W15 + "commentEx"))
            {
                foreach (string paraId in new[] { (string?)extension.Attribute(OoxmlNs.W15 + "paraId"), (string?)extension.Attribute(OoxmlNs.W15 + "paraIdParent") }
                    .Where(paraId => !string.IsNullOrWhiteSpace(paraId))
                    .Select(paraId => paraId!))
                {
                    used.Add(paraId);
                }
            }
        }

        for (uint id = 1; id < uint.MaxValue; id++)
        {
            string candidate = id.ToString("X8", System.Globalization.CultureInfo.InvariantCulture);
            if (!used.Contains(candidate))
            {
                return candidate;
            }
        }

        throw new InvalidDataException("Unable to allocate a unique comment paraId.");
    }

    private static void EnsureNamespaceDeclaration(XElement? root, string prefix, XNamespace ns)
    {
        if (root is not null && root.GetNamespaceOfPrefix(prefix) != ns)
        {
            root.SetAttributeValue(XNamespace.Xmlns + prefix, ns.NamespaceName);
        }
    }

    private static string? ReadCommentParaId(XElement comment)
    {
        return (string?)comment
            .Elements(OoxmlNs.W + "p")
            .FirstOrDefault()
            ?.Attribute(OoxmlNs.W15 + "paraId");
    }

    private static void ReplaceCommentText(XElement comment, string text)
    {
        comment.RemoveNodes();
        comment.Add(CreateSimpleParagraph(text));
    }

    private static XElement CreateComment(
        string commentId,
        string text,
        string author,
        string? initials,
        DateTimeOffset timestampUtc)
    {
        var comment = new XElement(
            OoxmlNs.W + "comment",
            new XAttribute(OoxmlNs.W + "id", commentId),
            new XAttribute(OoxmlNs.W + "author", author),
            new XAttribute(OoxmlNs.W + "date", timestampUtc.ToUniversalTime().ToString("O")));
        if (!string.IsNullOrWhiteSpace(initials))
        {
            comment.SetAttributeValue(OoxmlNs.W + "initials", initials);
        }

        comment.Add(CreateSimpleParagraph(text));
        return comment;
    }

    private static void AddCommentAnchor(XElement paragraph, string commentId)
    {
        var start = new XElement(OoxmlNs.W + "commentRangeStart", new XAttribute(OoxmlNs.W + "id", commentId));
        XElement? paragraphProperties = paragraph.Element(OoxmlNs.W + "pPr");
        if (paragraphProperties is null)
        {
            paragraph.AddFirst(start);
        }
        else
        {
            paragraphProperties.AddAfterSelf(start);
        }

        paragraph.Add(new XElement(OoxmlNs.W + "commentRangeEnd", new XAttribute(OoxmlNs.W + "id", commentId)));
        paragraph.Add(new XElement(
            OoxmlNs.W + "r",
            new XElement(OoxmlNs.W + "commentReference", new XAttribute(OoxmlNs.W + "id", commentId))));
    }

    private static void RemoveCommentExtensionRecords(OoxmlPackage package, XElement comment, CancellationToken cancellationToken)
    {
        string? paraId = ReadCommentParaId(comment);
        if (string.IsNullOrWhiteSpace(paraId))
        {
            return;
        }

        foreach (string partName in GetCommentsExtendedPartNames(package, cancellationToken))
        {
            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
            XElement[] extensionRecords = document
                .Descendants(OoxmlNs.W15 + "commentEx")
                .Where(element => string.Equals((string?)element.Attribute(OoxmlNs.W15 + "paraId"), paraId, StringComparison.Ordinal))
                .ToArray();
            if (extensionRecords.Length == 0)
            {
                continue;
            }

            foreach (XElement extensionRecord in extensionRecords)
            {
                extensionRecord.Remove();
            }

            SaveDocumentPart(package, partName, document);
        }
    }

    private static void RemoveCommentAnchors(OoxmlPackage package, string commentId, CancellationToken cancellationToken)
    {
        foreach (OoxmlPart part in package.Parts.Values
            .Where(part => part.Name.StartsWith("/word/", StringComparison.OrdinalIgnoreCase) &&
                part.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(part.Name, "/word/comments.xml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(part => part.Name, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            bool changed = false;
            foreach (XElement marker in document
                .Descendants()
                .Where(element => element.Name == OoxmlNs.W + "commentRangeStart" || element.Name == OoxmlNs.W + "commentRangeEnd")
                .Where(element => string.Equals((string?)element.Attribute(OoxmlNs.W + "id"), commentId, StringComparison.Ordinal))
                .ToArray())
            {
                marker.Remove();
                changed = true;
            }

            foreach (XElement reference in document
                .Descendants(OoxmlNs.W + "commentReference")
                .Where(element => string.Equals((string?)element.Attribute(OoxmlNs.W + "id"), commentId, StringComparison.Ordinal))
                .ToArray())
            {
                XElement? run = reference.Ancestors(OoxmlNs.W + "r").FirstOrDefault();
                reference.Remove();
                if (run is not null && !run.Elements().Any() && string.IsNullOrEmpty(run.Value))
                {
                    run.Remove();
                }

                changed = true;
            }

            if (changed)
            {
                SaveDocumentPart(package, part.Name, document);
            }
        }
    }

    private static IEnumerable<XElement> ElementAndDescendants(XElement element)
    {
        yield return element;
        foreach (XElement descendant in element.Descendants())
        {
            yield return descendant;
        }
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

    private static bool TryValidateTrackedWholeParagraphReplacement(
        XElement paragraph,
        string current,
        string replacement,
        string? style,
        out string? unsupportedReason)
    {
        unsupportedReason = null;
        if (style is not null)
        {
            unsupportedReason = "tracked paragraph replacement cannot combine text and style changes";
            return false;
        }

        if (TextContainsTrackedUnsupportedCharacters(current) ||
            TextContainsTrackedUnsupportedCharacters(replacement))
        {
            unsupportedReason = "tracked paragraph text contains tabs or line breaks";
            return false;
        }

        if (HasMixedDirectTextRunProperties(paragraph))
        {
            unsupportedReason = "paragraph contains mixed direct run formatting";
            return false;
        }

        return true;
    }

    private static bool TextContainsTrackedUnsupportedCharacters(string text)
    {
        return text.Contains('\t') || text.Contains('\n');
    }

    private static bool IsTrackedMode(DocxEditOptions options)
    {
        return options.TrackChanges is TrackChangesMode.Require or TrackChangesMode.Suggest;
    }

    private static bool TrackUnsupportedShape(
        DocxEditOptions options,
        DocxPatchOperation operation,
        string target,
        string reason,
        List<DocxDiagnostic> diagnostics)
    {
        if (options.TrackChanges == TrackChangesMode.Require)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E6002", $"Tracked-change output is not supported for {operation.OperationName} on {target}: {reason}.", operation, target));
            return false;
        }

        diagnostics.Add(Diagnostic(DocxSeverity.Warning, "W4002", $"Tracked-change output is not supported for {operation.OperationName} on {target}: {reason}; applying the edit directly.", operation, target));
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

    private static void ReplaceWholeParagraphTextWithTrackedChanges(
        OoxmlPackage package,
        XElement paragraph,
        string deletedText,
        string insertedText,
        DocxEditOptions options,
        CancellationToken cancellationToken)
    {
        int revisionCount = (deletedText.Length == 0 ? 0 : 1) + (insertedText.Length == 0 ? 0 : 1);
        string[] revisionIds = AllocateRevisionIds(package, revisionCount, cancellationToken);
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

        int revisionIndex = 0;
        if (deletedText.Length != 0)
        {
            nodes.Add(CreateDeletedRun(deletedText, firstRunProperties, revisionIds[revisionIndex++], author, timestamp));
        }

        if (insertedText.Length != 0)
        {
            nodes.Add(CreateInsertedRun(insertedText, firstRunProperties, revisionIds[revisionIndex], author, timestamp));
        }

        paragraph.RemoveNodes();
        paragraph.Add(nodes);
    }

    private static XElement CreateTrackedInsertedParagraph(
        OoxmlPackage package,
        string text,
        string? style,
        XElement? paragraphProperties,
        DocxEditOptions options,
        CancellationToken cancellationToken)
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

        string revisionId = AllocateRevisionIds(package, 1, cancellationToken)[0];
        string author = options.Author;
        string timestamp = options.TimestampUtc.ToUniversalTime().ToString("O");
        paragraph.Add(CreateInsertedRun(text, runProperties: null, revisionId, author, timestamp));
        return paragraph;
    }

    private static void SetParagraphStyleWithTrackedChange(
        OoxmlPackage package,
        XElement paragraph,
        string styleId,
        DocxEditOptions options,
        CancellationToken cancellationToken)
    {
        XElement oldParagraphProperties = paragraph.Element(OoxmlNs.W + "pPr") is { } existing
            ? new XElement(existing)
            : new XElement(OoxmlNs.W + "pPr");
        SetParagraphStyle(paragraph, styleId);
        XElement paragraphProperties = paragraph.Element(OoxmlNs.W + "pPr")
            ?? throw new InvalidDataException("Paragraph style update did not create paragraph properties.");
        paragraphProperties.Elements(OoxmlNs.W + "pPrChange").Remove();
        string revisionId = AllocateRevisionIds(package, 1, cancellationToken)[0];
        paragraphProperties.Add(new XElement(
            OoxmlNs.W + "pPrChange",
            new XAttribute(OoxmlNs.W + "id", revisionId),
            new XAttribute(OoxmlNs.W + "author", options.Author),
            new XAttribute(OoxmlNs.W + "date", options.TimestampUtc.ToUniversalTime().ToString("O")),
            oldParagraphProperties));
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

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetImageMetadata(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? alt = operation.Fields.GetValueOrDefault("alt");
        string? title = operation.Fields.GetValueOrDefault("title");
        string? name = operation.Fields.GetValueOrDefault("name");
        if (alt is null && title is null && name is null)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4202", "Operation 'set-image-metadata' requires at least one of 'alt', 'title', or 'name'.", operation, target));
        }

        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target!, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-image-metadata target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
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

        SetImageMetadata(imageContainer!, target!, alt, title, name);
        SaveDocumentPart(package, imageTarget.PartName, imageTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetImageSize(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        bool hasWidth = operation.Fields.ContainsKey("width");
        bool hasHeight = operation.Fields.ContainsKey("height");
        if (!hasWidth && !hasHeight)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4202", "Operation 'set-image-size' requires 'width', 'height', or both.", operation, target));
        }

        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target!, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-image-size target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
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

        if (!TryReadExistingImageSize(imageContainer!, imageTarget.Part, out long currentWidthEmus, out long currentHeightEmus) ||
            !TryReadImageSize(operation, currentWidthEmus, currentHeightEmus, out long widthEmus, out long heightEmus, out diagnostic))
        {
            return [diagnostic!];
        }

        if (!apply)
        {
            return [];
        }

        SetImageSize(imageContainer!, widthEmus, heightEmus);
        SaveDocumentPart(package, imageTarget.PartName, imageTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetImageWrap(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        bool hasWrapField =
            operation.Fields.ContainsKey("mode") ||
            operation.Fields.ContainsKey("dist-top") ||
            operation.Fields.ContainsKey("dist-bottom") ||
            operation.Fields.ContainsKey("dist-left") ||
            operation.Fields.ContainsKey("dist-right");
        if (!hasWrapField)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4202", "Operation 'set-image-wrap' requires 'mode' or at least one distance field.", operation, target));
        }

        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target!, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-image-wrap target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
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

        if (imageContainer!.Name != OoxmlNs.Wp + "anchor")
        {
            return [Diagnostic(DocxSeverity.Error, "E5205", $"Image '{target}' is inline; wrap metadata is only editable on anchored images.", operation, target)];
        }

        string? mode = null;
        if (operation.Fields.TryGetValue("mode", out string? modeText) &&
            !TryNormalizeWrapMode(modeText, out mode))
        {
            return [Diagnostic(DocxSeverity.Error, "E5209", $"Unsupported image wrap mode '{modeText}'.", operation, target)];
        }

        if (!TryReadWrapDistanceField(operation, "dist-top", diagnostics, out long? distanceTop) ||
            !TryReadWrapDistanceField(operation, "dist-bottom", diagnostics, out long? distanceBottom) ||
            !TryReadWrapDistanceField(operation, "dist-left", diagnostics, out long? distanceLeft) ||
            !TryReadWrapDistanceField(operation, "dist-right", diagnostics, out long? distanceRight))
        {
            return diagnostics;
        }

        if (!apply)
        {
            return [];
        }

        if (mode is not null)
        {
            SetImageWrapMode(imageContainer, mode);
        }

        SetImageWrapDistance(imageContainer, "distT", distanceTop);
        SetImageWrapDistance(imageContainer, "distB", distanceBottom);
        SetImageWrapDistance(imageContainer, "distL", distanceLeft);
        SetImageWrapDistance(imageContainer, "distR", distanceRight);
        SaveDocumentPart(package, imageTarget.PartName, imageTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetImagePosition(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        bool hasPositionField =
            operation.Fields.ContainsKey("horizontal-relative") ||
            operation.Fields.ContainsKey("horizontal-offset") ||
            operation.Fields.ContainsKey("horizontal-align") ||
            operation.Fields.ContainsKey("vertical-relative") ||
            operation.Fields.ContainsKey("vertical-offset") ||
            operation.Fields.ContainsKey("vertical-align");
        if (!hasPositionField)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4202", "Operation 'set-image-position' requires at least one position field.", operation, target));
        }

        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target!, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-image-position target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
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

        if (imageContainer!.Name != OoxmlNs.Wp + "anchor")
        {
            return [Diagnostic(DocxSeverity.Error, "E5205", $"Image '{target}' is inline; position metadata is only editable on anchored images.", operation, target)];
        }

        if (!TryReadImagePositionAxis(operation, "horizontal", diagnostics, out ImagePositionAxis horizontal) ||
            !TryReadImagePositionAxis(operation, "vertical", diagnostics, out ImagePositionAxis vertical))
        {
            return diagnostics;
        }

        if (!apply)
        {
            return [];
        }

        SetImagePositionAxis(imageContainer, "positionH", horizontal);
        SetImagePositionAxis(imageContainer, "positionV", vertical);
        SaveDocumentPart(package, imageTarget.PartName, imageTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetImageCrop(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        bool hasCropField =
            operation.Fields.ContainsKey("left-percent") ||
            operation.Fields.ContainsKey("top-percent") ||
            operation.Fields.ContainsKey("right-percent") ||
            operation.Fields.ContainsKey("bottom-percent");
        if (!hasCropField)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4202", "Operation 'set-image-crop' requires at least one crop field.", operation, target));
        }

        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target!, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-image-crop target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
        }

        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateImageContentTypeGuard(operation, target!, imageTarget.Part.ContentType, diagnostics))
        {
            return diagnostics;
        }

        XElement? blipFill = imageTarget.Blip.Ancestors(OoxmlNs.Pic + "blipFill").FirstOrDefault();
        if (blipFill is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E5205", $"Image '{target}' does not have editable DrawingML crop metadata.", operation, target)];
        }

        ImageCrop crop = ReadImageCrop(blipFill.Element(OoxmlNs.A + "srcRect"));
        if (!TryReadCropPercentField(operation, "left-percent", crop.Left, diagnostics, out int left) ||
            !TryReadCropPercentField(operation, "top-percent", crop.Top, diagnostics, out int top) ||
            !TryReadCropPercentField(operation, "right-percent", crop.Right, diagnostics, out int right) ||
            !TryReadCropPercentField(operation, "bottom-percent", crop.Bottom, diagnostics, out int bottom))
        {
            return diagnostics;
        }

        crop = new ImageCrop(left, top, right, bottom);
        if (crop.Left + crop.Right >= 100_000)
        {
            return [Diagnostic(DocxSeverity.Error, "E5208", "Image crop left-percent plus right-percent must be less than 100.", operation, target)];
        }

        if (crop.Top + crop.Bottom >= 100_000)
        {
            return [Diagnostic(DocxSeverity.Error, "E5208", "Image crop top-percent plus bottom-percent must be less than 100.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        SetImageCrop(blipFill, crop);
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
        SetImageMetadata(container, target, alt, title: null, name: null);
    }

    private static void SetImageMetadata(XElement container, string target, string? alt, string? title, string? name)
    {
        XElement? docPr = container.Element(OoxmlNs.Wp + "docPr");
        if (docPr is null)
        {
            docPr = new XElement(OoxmlNs.Wp + "docPr");
            docPr.SetAttributeValue("id", "1");
            docPr.SetAttributeValue("name", target);
            container.AddFirst(docPr);
        }

        if (alt is not null)
        {
            docPr.SetAttributeValue("descr", alt);
        }

        if (title is not null)
        {
            docPr.SetAttributeValue("title", title);
        }

        if (name is not null)
        {
            docPr.SetAttributeValue("name", name);
        }
    }

    private static bool TryReadExistingImageSize(XElement container, OoxmlPart part, out long widthEmus, out long heightEmus)
    {
        XElement? extent = container.Element(OoxmlNs.Wp + "extent");
        widthEmus = ReadLongAttribute(extent, "cx") ?? 0;
        heightEmus = ReadLongAttribute(extent, "cy") ?? 0;
        if (widthEmus > 0 && heightEmus > 0)
        {
            return true;
        }

        if (part.ContentType is not null &&
            TryReadImagePixelSize(part.Bytes, part.ContentType, out int pixelWidth, out int pixelHeight))
        {
            widthEmus = OoxmlUnits.PixelsToEmu(pixelWidth);
            heightEmus = OoxmlUnits.PixelsToEmu(pixelHeight);
            return true;
        }

        widthEmus = OoxmlUnits.InchesToEmu(1);
        heightEmus = widthEmus;
        return true;
    }

    private static bool TryReadImageSize(
        DocxPatchOperation operation,
        long currentWidthEmus,
        long currentHeightEmus,
        out long widthEmus,
        out long heightEmus,
        out DocxDiagnostic? diagnostic)
    {
        widthEmus = currentWidthEmus;
        heightEmus = currentHeightEmus;
        diagnostic = null;
        bool hasWidth = operation.Fields.TryGetValue("width", out string? width);
        bool hasHeight = operation.Fields.TryGetValue("height", out string? height);
        if (hasWidth && (!OoxmlUnits.TryParseDimension(width!, out widthEmus) || widthEmus <= 0))
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E5206", $"Invalid image width '{width}'.", operation, operation.Fields.GetValueOrDefault("target"));
            return false;
        }

        if (hasHeight && (!OoxmlUnits.TryParseDimension(height!, out heightEmus) || heightEmus <= 0))
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E5206", $"Invalid image height '{height}'.", operation, operation.Fields.GetValueOrDefault("target"));
            return false;
        }

        if (hasWidth && !hasHeight)
        {
            heightEmus = checked((long)Math.Round(widthEmus * (currentHeightEmus / (double)currentWidthEmus), MidpointRounding.AwayFromZero));
        }
        else if (!hasWidth && hasHeight)
        {
            widthEmus = checked((long)Math.Round(heightEmus * (currentWidthEmus / (double)currentHeightEmus), MidpointRounding.AwayFromZero));
        }

        return true;
    }

    private static long? ReadLongAttribute(XElement? element, string localName)
    {
        return long.TryParse((string?)element?.Attribute(localName), System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out long value)
            ? value
            : null;
    }

    private static void SetImageSize(XElement container, long widthEmus, long heightEmus)
    {
        XElement extent = container.Element(OoxmlNs.Wp + "extent") ?? new XElement(OoxmlNs.Wp + "extent");
        if (extent.Parent is null)
        {
            container.AddFirst(extent);
        }

        extent.SetAttributeValue("cx", widthEmus);
        extent.SetAttributeValue("cy", heightEmus);

        XElement? picture = container.Descendants(OoxmlNs.Pic + "pic").FirstOrDefault();
        if (picture is null)
        {
            return;
        }

        XElement shapeProperties = picture.Element(OoxmlNs.Pic + "spPr") ?? new XElement(OoxmlNs.Pic + "spPr");
        if (shapeProperties.Parent is null)
        {
            picture.Add(shapeProperties);
        }

        XElement transform = shapeProperties.Element(OoxmlNs.A + "xfrm") ?? new XElement(OoxmlNs.A + "xfrm");
        if (transform.Parent is null)
        {
            shapeProperties.AddFirst(transform);
        }

        XElement transformExtent = transform.Element(OoxmlNs.A + "ext") ?? new XElement(OoxmlNs.A + "ext");
        if (transformExtent.Parent is null)
        {
            transform.Add(transformExtent);
        }

        transformExtent.SetAttributeValue("cx", widthEmus);
        transformExtent.SetAttributeValue("cy", heightEmus);
    }

    private static bool TryNormalizeWrapMode(string text, out string? mode)
    {
        mode = text switch
        {
            "none" or "wrapNone" => "wrapNone",
            "square" or "wrapSquare" => "wrapSquare",
            "tight" or "wrapTight" => "wrapTight",
            "through" or "wrapThrough" => "wrapThrough",
            "top-bottom" or "topAndBottom" or "wrapTopAndBottom" => "wrapTopAndBottom",
            _ => null
        };
        return mode is not null;
    }

    private static bool TryReadWrapDistanceField(
        DocxPatchOperation operation,
        string fieldName,
        List<DocxDiagnostic> diagnostics,
        out long? value)
    {
        value = null;
        if (!operation.Fields.TryGetValue(fieldName, out string? text))
        {
            return true;
        }

        if (!OoxmlUnits.TryParseDimension(text, out long emus))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E5209", $"Image wrap distance field '{fieldName}' must be a non-negative dimension.", operation, operation.Fields.GetValueOrDefault("target")));
            return false;
        }

        value = emus;
        return true;
    }

    private static void SetImageWrapMode(XElement anchor, string mode)
    {
        anchor.Elements()
            .Where(element => element.Name.Namespace == OoxmlNs.Wp && element.Name.LocalName.StartsWith("wrap", StringComparison.Ordinal))
            .Remove();
        XElement wrap = new(OoxmlNs.Wp + mode);
        XElement? insertAfter = anchor.Element(OoxmlNs.Wp + "effectExtent") ??
            anchor.Element(OoxmlNs.Wp + "extent");
        if (insertAfter is null)
        {
            anchor.AddFirst(wrap);
        }
        else
        {
            insertAfter.AddAfterSelf(wrap);
        }
    }

    private static void SetImageWrapDistance(XElement anchor, string attributeName, long? value)
    {
        if (value is not null)
        {
            anchor.SetAttributeValue(attributeName, value.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
    }

    private static bool TryReadImagePositionAxis(
        DocxPatchOperation operation,
        string axis,
        List<DocxDiagnostic> diagnostics,
        out ImagePositionAxis position)
    {
        position = default;
        bool hasRelative = operation.Fields.TryGetValue($"{axis}-relative", out string? relative);
        bool hasOffset = operation.Fields.TryGetValue($"{axis}-offset", out string? offsetText);
        bool hasAlign = operation.Fields.TryGetValue($"{axis}-align", out string? align);
        if (!hasRelative && !hasOffset && !hasAlign)
        {
            return true;
        }

        if (hasOffset && hasAlign)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E5210", $"Image position axis '{axis}' cannot specify both offset and align.", operation, operation.Fields.GetValueOrDefault("target")));
            return false;
        }

        if (hasRelative && !IsValidImagePositionRelative(axis, relative!))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E5210", $"Unsupported image {axis} relative value '{relative}'.", operation, operation.Fields.GetValueOrDefault("target")));
            return false;
        }

        long? offsetEmus = null;
        if (hasOffset)
        {
            if (!TryParseSignedDimension(offsetText!, out long parsedOffset))
            {
                diagnostics.Add(Diagnostic(DocxSeverity.Error, "E5210", $"Image position field '{axis}-offset' must be a signed dimension.", operation, operation.Fields.GetValueOrDefault("target")));
                return false;
            }

            offsetEmus = parsedOffset;
        }

        if (hasAlign && !IsValidImagePositionAlign(axis, align!))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E5210", $"Unsupported image {axis} align value '{align}'.", operation, operation.Fields.GetValueOrDefault("target")));
            return false;
        }

        position = new ImagePositionAxis(true, hasRelative ? relative : null, offsetEmus, hasAlign ? align : null);
        return true;
    }

    private static bool IsValidImagePositionRelative(string axis, string value)
    {
        return axis == "horizontal"
            ? value is "page" or "margin" or "column" or "character" or "leftMargin" or "rightMargin" or "insideMargin" or "outsideMargin"
            : value is "page" or "margin" or "paragraph" or "line" or "topMargin" or "bottomMargin" or "insideMargin" or "outsideMargin";
    }

    private static bool IsValidImagePositionAlign(string axis, string value)
    {
        return axis == "horizontal"
            ? value is "left" or "center" or "right" or "inside" or "outside"
            : value is "top" or "center" or "bottom" or "inside" or "outside";
    }

    private static bool TryParseSignedDimension(string text, out long emus)
    {
        emus = 0;
        string trimmed = text.Trim();
        string[] suffixes = ["emu", "in", "cm", "pt", "px"];
        foreach (string suffix in suffixes)
        {
            if (!trimmed.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            string numberText = trimmed[..^suffix.Length];
            if (!double.TryParse(numberText, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double value))
            {
                return false;
            }

            emus = suffix.ToLowerInvariant() switch
            {
                "emu" => checked((long)Math.Round(value, MidpointRounding.AwayFromZero)),
                "in" => OoxmlUnits.InchesToEmu(value),
                "cm" => OoxmlUnits.CentimetersToEmu(value),
                "pt" => OoxmlUnits.PointsToEmu(value),
                "px" => OoxmlUnits.PixelsToEmu(value),
                _ => 0
            };
            return true;
        }

        return false;
    }

    private static void SetImagePositionAxis(XElement anchor, string elementName, ImagePositionAxis axis)
    {
        if (!axis.HasAny)
        {
            return;
        }

        XElement position = anchor.Element(OoxmlNs.Wp + elementName) ?? new XElement(OoxmlNs.Wp + elementName);
        if (position.Parent is null)
        {
            AddImagePositionElement(anchor, elementName, position);
        }

        if (axis.RelativeFrom is not null)
        {
            position.SetAttributeValue("relativeFrom", axis.RelativeFrom);
        }
        else if (position.Attribute("relativeFrom") is null)
        {
            position.SetAttributeValue("relativeFrom", elementName == "positionH" ? "column" : "paragraph");
        }

        if (axis.OffsetEmus is not null)
        {
            position.Elements(OoxmlNs.Wp + "align").Remove();
            XElement offset = position.Element(OoxmlNs.Wp + "posOffset") ?? new XElement(OoxmlNs.Wp + "posOffset");
            if (offset.Parent is null)
            {
                position.Add(offset);
            }

            offset.Value = axis.OffsetEmus.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        if (axis.Align is not null)
        {
            position.Elements(OoxmlNs.Wp + "posOffset").Remove();
            XElement align = position.Element(OoxmlNs.Wp + "align") ?? new XElement(OoxmlNs.Wp + "align");
            if (align.Parent is null)
            {
                position.Add(align);
            }

            align.Value = axis.Align;
        }

        if (!position.Elements(OoxmlNs.Wp + "posOffset").Any() &&
            !position.Elements(OoxmlNs.Wp + "align").Any())
        {
            position.Add(new XElement(OoxmlNs.Wp + "posOffset", "0"));
        }
    }

    private static void AddImagePositionElement(XElement anchor, string elementName, XElement position)
    {
        if (elementName == "positionH")
        {
            XElement? before = anchor.Element(OoxmlNs.Wp + "positionV") ?? anchor.Element(OoxmlNs.Wp + "extent");
            if (before is null)
            {
                anchor.AddFirst(position);
            }
            else
            {
                before.AddBeforeSelf(position);
            }

            return;
        }

        XElement? horizontal = anchor.Element(OoxmlNs.Wp + "positionH");
        if (horizontal is not null)
        {
            horizontal.AddAfterSelf(position);
            return;
        }

        XElement? extent = anchor.Element(OoxmlNs.Wp + "extent");
        if (extent is null)
        {
            anchor.AddFirst(position);
        }
        else
        {
            extent.AddBeforeSelf(position);
        }
    }

    private static ImageCrop ReadImageCrop(XElement? sourceRectangle)
    {
        return new ImageCrop(
            ReadCropPerThousandPercent(sourceRectangle, "l"),
            ReadCropPerThousandPercent(sourceRectangle, "t"),
            ReadCropPerThousandPercent(sourceRectangle, "r"),
            ReadCropPerThousandPercent(sourceRectangle, "b"));
    }

    private static int ReadCropPerThousandPercent(XElement? sourceRectangle, string localName)
    {
        string? value = (string?)sourceRectangle?.Attribute(localName);
        if (string.IsNullOrWhiteSpace(value))
        {
            return 0;
        }

        if (value.EndsWith("%", StringComparison.Ordinal) &&
            decimal.TryParse(value[..^1], System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out decimal percent))
        {
            return PercentToPerThousand(percent);
        }

        return int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int perThousandPercent)
            ? Math.Clamp(perThousandPercent, 0, 100_000)
            : 0;
    }

    private static bool TryReadCropPercentField(
        DocxPatchOperation operation,
        string fieldName,
        int currentValue,
        List<DocxDiagnostic> diagnostics,
        out int value)
    {
        value = currentValue;
        if (!operation.Fields.TryGetValue(fieldName, out string? text))
        {
            return true;
        }

        if (!decimal.TryParse(text, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out decimal percent) ||
            percent is < 0m or > 100m)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E5208", $"Image crop field '{fieldName}' must be a percentage from 0 to 100.", operation, operation.Fields.GetValueOrDefault("target")));
            return false;
        }

        value = PercentToPerThousand(percent);
        return true;
    }

    private static int PercentToPerThousand(decimal percent)
    {
        return (int)Math.Round(percent * 1000m, MidpointRounding.AwayFromZero);
    }

    private static void SetImageCrop(XElement blipFill, ImageCrop crop)
    {
        XElement? sourceRectangle = blipFill.Element(OoxmlNs.A + "srcRect");
        if (crop.Left == 0 && crop.Top == 0 && crop.Right == 0 && crop.Bottom == 0)
        {
            sourceRectangle?.Remove();
            return;
        }

        if (sourceRectangle is null)
        {
            sourceRectangle = new XElement(OoxmlNs.A + "srcRect");
            XElement? blip = blipFill.Element(OoxmlNs.A + "blip");
            if (blip is null)
            {
                blipFill.AddFirst(sourceRectangle);
            }
            else
            {
                blip.AddAfterSelf(sourceRectangle);
            }
        }

        SetCropAttribute(sourceRectangle, "l", crop.Left);
        SetCropAttribute(sourceRectangle, "t", crop.Top);
        SetCropAttribute(sourceRectangle, "r", crop.Right);
        SetCropAttribute(sourceRectangle, "b", crop.Bottom);
    }

    private static void SetCropAttribute(XElement sourceRectangle, string localName, int value)
    {
        sourceRectangle.SetAttributeValue(localName, value == 0 ? null : value.ToString(System.Globalization.CultureInfo.InvariantCulture));
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
        DocxEditOptions options,
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

        bool useTrackedChanges = IsTrackedMode(options);
        XElement? paragraph = cellTarget.Cell.Elements(OoxmlNs.W + "p").FirstOrDefault();
        if (useTrackedChanges)
        {
            if (force)
            {
                if (!TrackUnsupportedShape(options, operation, target!, "tracked set-cell does not support force true replacement", diagnostics))
                {
                    return diagnostics;
                }

                useTrackedChanges = false;
            }
            else if (paragraph is null)
            {
                if (!TrackUnsupportedShape(options, operation, target!, "cell has no paragraph for tracked text replacement", diagnostics))
                {
                    return diagnostics;
                }

                useTrackedChanges = false;
            }
            else if (TryGetProtectedTextEditFeature(paragraph, out string protectedFeature))
            {
                if (!TrackUnsupportedShape(options, operation, target!, $"cell paragraph contains protected OOXML boundary '{protectedFeature}'", diagnostics))
                {
                    return diagnostics;
                }

                useTrackedChanges = false;
            }
            else if (!TryValidateTrackedWholeParagraphReplacement(paragraph, current, text!, style: null, out string? trackedUnsupportedReason))
            {
                if (!TrackUnsupportedShape(options, operation, target!, trackedUnsupportedReason!, diagnostics))
                {
                    return diagnostics;
                }

                useTrackedChanges = false;
            }
        }

        if (!apply)
        {
            return diagnostics;
        }

        if (useTrackedChanges)
        {
            ReplaceWholeParagraphTextWithTrackedChanges(package, paragraph!, current, text!, options, cancellationToken);
            SaveDocumentPart(package, cellTarget.PartName, cellTarget.Document);
            return diagnostics;
        }

        ReplaceCellText(cellTarget.Cell, text!);
        SaveDocumentPart(package, cellTarget.PartName, cellTarget.Document);
        return diagnostics;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetTableStyle(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? style = ReadRequiredField(operation, "style", diagnostics);
        string? expectedStyle = operation.Fields.GetValueOrDefault("expect-style");
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        TableTarget? tableTarget = ResolveTableTarget(package, target!, cancellationToken);
        if (tableTarget is null && !IsSupportedTableTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-table-style target '{target}'. Expected a table ID such as M.T0001 or H001.T0001.", operation, target)];
        }

        if (tableTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        string? currentStyle = ReadTableStyleId(tableTarget.Table);
        if (expectedStyle is not null && !string.Equals(currentStyle, expectedStyle, StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected table style '{expectedStyle}', found '{currentStyle ?? "none"}'.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        SetTableStyle(tableTarget.Table, style!);
        SaveDocumentPart(package, tableTarget.PartName, tableTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetTableMetadata(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? expectedCaption = operation.Fields.GetValueOrDefault("expect-caption");
        string? expectedDescription = operation.Fields.GetValueOrDefault("expect-description");
        string? caption = operation.Fields.GetValueOrDefault("caption");
        string? description = operation.Fields.GetValueOrDefault("description");
        if (caption is null && description is null)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4202", "Operation 'set-table-metadata' requires at least one of 'caption' or 'description'.", operation, target));
        }

        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        TableTarget? tableTarget = ResolveTableTarget(package, target!, cancellationToken);
        if (tableTarget is null && !IsSupportedTableTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-table-metadata target '{target}'. Expected a table ID such as M.T0001 or H001.T0001.", operation, target)];
        }

        if (tableTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        string? currentCaption = ReadTableTextProperty(tableTarget.Table, "tblCaption");
        if (expectedCaption is not null && !TableMetadataEquals(currentCaption, expectedCaption))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected table caption '{expectedCaption}', found '{currentCaption ?? "none"}'.", operation, target)];
        }

        string? currentDescription = ReadTableTextProperty(tableTarget.Table, "tblDescription");
        if (expectedDescription is not null && !TableMetadataEquals(currentDescription, expectedDescription))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected table description '{expectedDescription}', found '{currentDescription ?? "none"}'.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        SetTableTextProperty(tableTarget.Table, "tblCaption", caption);
        SetTableTextProperty(tableTarget.Table, "tblDescription", description);
        SaveDocumentPart(package, tableTarget.PartName, tableTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetRowHeader(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        _ = ReadRequiredField(operation, "header", diagnostics);
        bool? header = ReadBooleanField(operation, "header", diagnostics);
        bool? expectedHeader = ReadBooleanField(operation, "expect-header", diagnostics);
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        RowTarget? rowTarget = ResolveRowTarget(package, target!, cancellationToken);
        if (rowTarget is null && !IsSupportedRowTargetShape(target!))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-row-header target '{target}'. Expected a table row ID such as M.T0001.R02 or H001.T0001.R02.", operation, target)];
        }

        if (rowTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        bool currentHeader = ReadTableRowHeader(rowTarget.Row);
        if (expectedHeader is not null && currentHeader != expectedHeader)
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected row header '{expectedHeader.Value.ToString().ToLowerInvariant()}', found '{currentHeader.ToString().ToLowerInvariant()}'.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        SetTableRowHeader(rowTarget.Row, header!.Value);
        SaveDocumentPart(package, rowTarget.PartName, rowTarget.Document);
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

    private static bool TryParseMainHyperlinkTarget(string target, out int hyperlinkOrdinal)
    {
        hyperlinkOrdinal = 0;
        return target.Length == 7 &&
            target.StartsWith("M.L", StringComparison.Ordinal) &&
            int.TryParse(target[3..], out hyperlinkOrdinal);
    }

    private static bool TryParseMainFieldTarget(string target, out int fieldOrdinal)
    {
        fieldOrdinal = 0;
        return target.Length == 7 &&
            target.StartsWith("M.F", StringComparison.Ordinal) &&
            int.TryParse(target[3..], out fieldOrdinal);
    }

    private static bool TryParseMainBookmarkTarget(string target, out int bookmarkOrdinal)
    {
        bookmarkOrdinal = 0;
        return target.Length == 7 &&
            target.StartsWith("M.B", StringComparison.Ordinal) &&
            int.TryParse(target[3..], out bookmarkOrdinal);
    }

    private static bool TryParseMainContentControlTarget(string target, out int contentControlOrdinal)
    {
        contentControlOrdinal = 0;
        return target.Length == 8 &&
            target.StartsWith("M.CC", StringComparison.Ordinal) &&
            int.TryParse(target[4..], out contentControlOrdinal);
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

    private static bool TryParseStoryHyperlinkTarget(
        string target,
        char storyPrefix,
        out int storyOrdinal,
        out int hyperlinkOrdinal)
    {
        storyOrdinal = 0;
        hyperlinkOrdinal = 0;
        if (target.Length != 10 ||
            target[0] != storyPrefix ||
            target[4..6] != ".L")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[6..], out hyperlinkOrdinal);
    }

    private static bool TryParseStoryFieldTarget(
        string target,
        char storyPrefix,
        out int storyOrdinal,
        out int fieldOrdinal)
    {
        storyOrdinal = 0;
        fieldOrdinal = 0;
        if (target.Length != 10 ||
            target[0] != storyPrefix ||
            target[4..6] != ".F")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[6..], out fieldOrdinal);
    }

    private static bool TryParseStoryBookmarkTarget(
        string target,
        char storyPrefix,
        out int storyOrdinal,
        out int bookmarkOrdinal)
    {
        storyOrdinal = 0;
        bookmarkOrdinal = 0;
        if (target.Length != 10 ||
            target[0] != storyPrefix ||
            target[4..6] != ".B")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[6..], out bookmarkOrdinal);
    }

    private static bool TryParseStoryContentControlTarget(
        string target,
        char storyPrefix,
        out int storyOrdinal,
        out int contentControlOrdinal)
    {
        storyOrdinal = 0;
        contentControlOrdinal = 0;
        if (target.Length != 11 ||
            target[0] != storyPrefix ||
            target[4..7] != ".CC")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[7..], out contentControlOrdinal);
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

    private static bool IsSupportedHyperlinkTargetShape(string target)
    {
        return TryParseMainHyperlinkTarget(target, out _) ||
            TryParseStoryHyperlinkTarget(target, 'H', out _, out _) ||
            TryParseStoryHyperlinkTarget(target, 'F', out _, out _);
    }

    private static bool IsSupportedContentControlTargetShape(string target)
    {
        return TryParseMainContentControlTarget(target, out _) ||
            TryParseStoryContentControlTarget(target, 'H', out _, out _) ||
            TryParseStoryContentControlTarget(target, 'F', out _, out _);
    }

    private static bool IsSupportedFieldTargetShape(string target)
    {
        return TryParseMainFieldTarget(target, out _) ||
            TryParseStoryFieldTarget(target, 'H', out _, out _) ||
            TryParseStoryFieldTarget(target, 'F', out _, out _);
    }

    private static bool IsSupportedBookmarkTargetShape(string target)
    {
        return TryParseMainBookmarkTarget(target, out _) ||
            TryParseStoryBookmarkTarget(target, 'H', out _, out _) ||
            TryParseStoryBookmarkTarget(target, 'F', out _, out _);
    }

    private static ContentControlTarget? ResolveContentControlTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (TryParseMainContentControlTarget(target, out int mainControlOrdinal))
        {
            return package.MainDocumentPartName is null
                ? null
                : FindContentControlTarget(package, package.MainDocumentPartName, mainControlOrdinal, cancellationToken);
        }

        if (TryParseStoryContentControlTarget(target, 'H', out int headerOrdinal, out int headerControlOrdinal))
        {
            string? partName = ResolveRelatedStoryPartName(package, OoxmlRelTypes.Header, headerOrdinal, cancellationToken);
            return partName is null ? null : FindContentControlTarget(package, partName, headerControlOrdinal, cancellationToken);
        }

        if (TryParseStoryContentControlTarget(target, 'F', out int footerOrdinal, out int footerControlOrdinal))
        {
            string? partName = ResolveRelatedStoryPartName(package, OoxmlRelTypes.Footer, footerOrdinal, cancellationToken);
            return partName is null ? null : FindContentControlTarget(package, partName, footerControlOrdinal, cancellationToken);
        }

        return null;
    }

    private static FieldTarget? ResolveFieldTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (TryParseMainFieldTarget(target, out int mainFieldOrdinal))
        {
            XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
            XElement? field = FindField(body, mainFieldOrdinal);
            return field is null || package.MainDocumentPartName is null
                ? null
                : new FieldTarget(package.MainDocumentPartName, document, field);
        }

        if (TryParseStoryFieldTarget(target, 'H', out int headerOrdinal, out int headerFieldOrdinal))
        {
            string? partName = ResolveRelatedStoryPartName(package, OoxmlRelTypes.Header, headerOrdinal, cancellationToken);
            return partName is null ? null : FindFieldTarget(package, partName, headerFieldOrdinal, cancellationToken);
        }

        if (TryParseStoryFieldTarget(target, 'F', out int footerOrdinal, out int footerFieldOrdinal))
        {
            string? partName = ResolveRelatedStoryPartName(package, OoxmlRelTypes.Footer, footerOrdinal, cancellationToken);
            return partName is null ? null : FindFieldTarget(package, partName, footerFieldOrdinal, cancellationToken);
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
        if (TryParseMainBookmarkTarget(target, out int mainBookmarkOrdinal))
        {
            return package.MainDocumentPartName is null
                ? null
                : FindBookmarkTarget(package, package.MainDocumentPartName, mainBookmarkOrdinal, cancellationToken);
        }

        if (TryParseStoryBookmarkTarget(target, 'H', out int headerOrdinal, out int headerBookmarkOrdinal))
        {
            string? partName = ResolveRelatedStoryPartName(package, OoxmlRelTypes.Header, headerOrdinal, cancellationToken);
            return partName is null ? null : FindBookmarkTarget(package, partName, headerBookmarkOrdinal, cancellationToken);
        }

        if (TryParseStoryBookmarkTarget(target, 'F', out int footerOrdinal, out int footerBookmarkOrdinal))
        {
            string? partName = ResolveRelatedStoryPartName(package, OoxmlRelTypes.Footer, footerOrdinal, cancellationToken);
            return partName is null ? null : FindBookmarkTarget(package, partName, footerBookmarkOrdinal, cancellationToken);
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
        if (TryParseMainHyperlinkTarget(target, out int mainHyperlinkOrdinal))
        {
            return package.MainDocumentPartName is null
                ? null
                : FindHyperlinkTarget(package, package.MainDocumentPartName, mainHyperlinkOrdinal, cancellationToken);
        }

        if (TryParseStoryHyperlinkTarget(target, 'H', out int headerOrdinal, out int headerHyperlinkOrdinal))
        {
            string? partName = ResolveRelatedStoryPartName(package, OoxmlRelTypes.Header, headerOrdinal, cancellationToken);
            return partName is null ? null : FindHyperlinkTarget(package, partName, headerHyperlinkOrdinal, cancellationToken);
        }

        if (TryParseStoryHyperlinkTarget(target, 'F', out int footerOrdinal, out int footerHyperlinkOrdinal))
        {
            string? partName = ResolveRelatedStoryPartName(package, OoxmlRelTypes.Footer, footerOrdinal, cancellationToken);
            return partName is null ? null : FindHyperlinkTarget(package, partName, footerHyperlinkOrdinal, cancellationToken);
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

    private static IReadOnlyList<string> GetEditableStoryPartNames(
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        if (package.MainDocumentPartName is null)
        {
            return [];
        }

        var partNames = new List<string> { package.MainDocumentPartName };
        IReadOnlyList<OoxmlRelationship> relationships = package.GetRelationships(package.MainDocumentPartName, cancellationToken);
        partNames.AddRange(relationships
            .Where(relationship => !relationship.IsExternal && relationship.Type == OoxmlRelTypes.Header && relationship.ResolvedTarget is not null)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .Select(relationship => relationship.ResolvedTarget!)
            .Where(partName => package.GetPart(partName) is not null));
        partNames.AddRange(relationships
            .Where(relationship => !relationship.IsExternal && relationship.Type == OoxmlRelTypes.Footer && relationship.ResolvedTarget is not null)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .Select(relationship => relationship.ResolvedTarget!)
            .Where(partName => package.GetPart(partName) is not null));
        return partNames;
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

    private static XElement CreateSimpleRun(string text)
    {
        var run = new XElement(OoxmlNs.W + "r");
        foreach (XNode node in CreateTextNodes(text))
        {
            run.Add(node);
        }

        return run;
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

    private static string? ReadTableStyleId(XElement table)
    {
        return (string?)table
            .Element(OoxmlNs.W + "tblPr")
            ?.Element(OoxmlNs.W + "tblStyle")
            ?.Attribute(OoxmlNs.W + "val");
    }

    private static string? ReadTableTextProperty(XElement table, string localName)
    {
        return (string?)table
            .Element(OoxmlNs.W + "tblPr")
            ?.Element(OoxmlNs.W + localName)
            ?.Attribute(OoxmlNs.W + "val");
    }

    private static bool TableMetadataEquals(string? current, string expected)
    {
        return string.Equals(current ?? string.Empty, expected, StringComparison.Ordinal);
    }

    private static void SetTableStyle(XElement table, string style)
    {
        XElement? tableProperties = table.Element(OoxmlNs.W + "tblPr");
        if (tableProperties is null)
        {
            tableProperties = new XElement(OoxmlNs.W + "tblPr");
            table.AddFirst(tableProperties);
        }

        XElement? tableStyle = tableProperties.Element(OoxmlNs.W + "tblStyle");
        if (tableStyle is null)
        {
            tableStyle = new XElement(OoxmlNs.W + "tblStyle");
            tableProperties.AddFirst(tableStyle);
        }

        tableStyle.SetAttributeValue(OoxmlNs.W + "val", style);
    }

    private static void SetTableTextProperty(XElement table, string localName, string? value)
    {
        if (value is null)
        {
            return;
        }

        XElement? tableProperties = table.Element(OoxmlNs.W + "tblPr");
        XElement? property = tableProperties?.Element(OoxmlNs.W + localName);
        if (value.Length == 0)
        {
            property?.Remove();
            return;
        }

        if (tableProperties is null)
        {
            tableProperties = new XElement(OoxmlNs.W + "tblPr");
            table.AddFirst(tableProperties);
        }

        if (property is null)
        {
            property = new XElement(OoxmlNs.W + localName);
            tableProperties.Add(property);
        }

        property.SetAttributeValue(OoxmlNs.W + "val", value);
    }

    private static bool ReadTableRowHeader(XElement row)
    {
        XElement? tableHeader = row
            .Element(OoxmlNs.W + "trPr")
            ?.Element(OoxmlNs.W + "tblHeader");
        if (tableHeader is null)
        {
            return false;
        }

        string? value = (string?)tableHeader.Attribute(OoxmlNs.W + "val");
        return value is null || value is "1" or "true" or "on";
    }

    private static void SetTableRowHeader(XElement row, bool header)
    {
        XElement? rowProperties = row.Element(OoxmlNs.W + "trPr");
        XElement? tableHeader = rowProperties?.Element(OoxmlNs.W + "tblHeader");
        if (!header)
        {
            tableHeader?.Remove();
            return;
        }

        if (rowProperties is null)
        {
            rowProperties = new XElement(OoxmlNs.W + "trPr");
            row.AddFirst(rowProperties);
        }

        if (tableHeader is null)
        {
            tableHeader = new XElement(OoxmlNs.W + "tblHeader");
            rowProperties.AddFirst(tableHeader);
        }

        tableHeader.SetAttributeValue(OoxmlNs.W + "val", null);
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
                ValidateTouchedBinaryPart(part, diagnostics);
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

    private static void ValidateTouchedBinaryPart(OoxmlPart part, List<DocxDiagnostic> diagnostics)
    {
        if (part.Bytes.Length != 0)
        {
            return;
        }

        if (string.Equals(part.ContentType, "image/png", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(part.ContentType, "image/jpeg", StringComparison.OrdinalIgnoreCase))
        {
            diagnostics.Add(PostEditValidationDiagnostic(part.Name, "Image part is empty."));
        }
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

internal sealed record ContentControlTarget(string PartName, XDocument Document, XElement ContentControl);

internal sealed record ContentControlChoice(string DisplayText);

internal sealed record BookmarkTarget(string PartName, XDocument Document, XElement Start, XElement? End);

internal sealed record CommentsPartTarget(string PartName, XDocument Document, XElement Root);

internal sealed record CommentsExtendedPartTarget(string PartName, XDocument Document, XElement Root);

internal sealed record CommentTarget(string PartName, XDocument Document, XElement Comment);

internal sealed record CommentExtensionTarget(string PartName, XDocument Document, XElement CommentExtension);

internal sealed record HyperlinkTarget(string PartName, XDocument Document, XElement Hyperlink);

internal sealed record FieldTarget(string PartName, XDocument Document, XElement Element);

internal sealed record ImageBlipTarget(string PartName, XDocument Document, XElement Blip, string RelationshipId, OoxmlPart Part);

internal readonly record struct ImageCrop(int Left, int Top, int Right, int Bottom);

internal readonly record struct ImagePositionAxis(bool HasAny, string? RelativeFrom, long? OffsetEmus, string? Align);

internal sealed record ParagraphTarget(string PartName, XDocument Document, XElement Paragraph);

internal sealed record BlockTarget(string PartName, XDocument Document, XElement Block);

internal sealed record TableTarget(string PartName, XDocument Document, XElement Table);

internal sealed record RowTarget(string PartName, XDocument Document, XElement Table, XElement Row);

internal sealed record CellTarget(string PartName, XDocument Document, XElement Table, XElement Row, XElement Cell);

internal sealed record TableOperationSnapshot(
    string TargetId,
    string? TableId,
    int? RowIndex,
    int? ColumnIndex,
    int RowCountBefore,
    int ColumnCount,
    int? CellCount);

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
