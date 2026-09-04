using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

internal static partial class DocxPatchEngine
{
    private const string SettingsContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
    private const string CommentsContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
    private const string CommentsExtendedContentType = "application/vnd.ms-word.commentsExtended+xml";
    private const string CommentsIdsContentType = "application/vnd.ms-word.commentsIds+xml";

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
        CancellationToken cancellationToken)
    {
        return Execute(package, patch, options, apply: false, cancellationToken);
    }

    public static PatchExecutionResult Apply(
        OoxmlPackage package,
        DocxPatch patch,
        DocxEditOptions options,
        CancellationToken cancellationToken)
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
        IReadOnlyList<DocxDiagnostic> trackChangeOptionDiagnostics = ValidateTrackChangeOptions(options);
        if (trackChangeOptionDiagnostics.Count != 0)
        {
            return new PatchExecutionResult(false, trackChangeOptionDiagnostics, []);
        }

        var reports = new List<DocxPatchOperationReport>();
        foreach (DocxPatchOperation operation in patch.Operations)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var operationDiagnostics = new List<DocxDiagnostic>();
            var generatedRevisionIds = new List<string>();
            TableOperationSnapshot? tableBefore = CaptureTableOperationSnapshot(package, operation, cancellationToken);
            bool supportsTrackedChanges = SupportsTrackedChangeOutput(operation.OperationName);
            if (options.TrackChanges == TrackChangesMode.Require && !supportsTrackedChanges)
            {
                operationDiagnostics.Add(Diagnostic(
                    DocxSeverity.Error,
                    "E6001",
                    BuildUnsupportedTrackedOperationMessage(options.TrackChanges, operation),
                    operation,
                    operation.Fields.GetValueOrDefault("target"),
                    "track-changes-no-revision-representation",
                    "require-failed"));
            }
            else
            {
                if (options.TrackChanges == TrackChangesMode.Suggest && !supportsTrackedChanges)
                {
                    operationDiagnostics.Add(Diagnostic(
                        DocxSeverity.Warning,
                        "W4001",
                        BuildUnsupportedTrackedOperationMessage(options.TrackChanges, operation),
                        operation,
                        operation.Fields.GetValueOrDefault("target"),
                        "track-changes-no-revision-representation",
                        "direct-edit-preserve-existing-revisions"));
                }

                IReadOnlyList<DocxDiagnostic>? executed = TryExecuteOperation(package, operation, options, apply, generatedRevisionIds, cancellationToken);
                if (executed is not null)
                {
                    operationDiagnostics.AddRange(executed);
                }
                else
                {
                    operationDiagnostics.Add(Diagnostic(DocxSeverity.Error, "E4201", $"Unsupported operation '{operation.OperationName}'.", operation));
                }
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
                AffectedTargets = operationSuccess ? BuildAffectedTargets(operation, tableBefore) : [],
                GeneratedRevisionIds = operationSuccess ? generatedRevisionIds.ToArray() : []
            });
        }

        bool shouldMarkFieldsDirty = apply &&
            options.MarkFieldsDirtyWhenEditing &&
            patch.Operations.Count != 0 &&
            patch.Operations.Any(operation => MarksFieldsDirtyAfterEdit(operation.OperationName)) &&
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

    private static string BuildUnsupportedTrackedOperationMessage(TrackChangesMode mode, DocxPatchOperation operation)
    {
        string operationName = operation.OperationName;
        string support = TrackChangesSupportValue(operationName);

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
            "set-cell" or "set-cell-shading" => CaptureCellSnapshot(package, target, cancellationToken),
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
        int columnIndex = cellTarget.VisualColumnIndex;
        return CreateTableOperationSnapshot(target, cellTarget.Table, rowIndex, columnIndex, cells.Length, cellTarget.Row);
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
        return CreateTableOperationSnapshot(target, rowTarget.Table, rowIndex, columnIndex: null, cellCount, rowTarget.Row);
    }

    private static TableOperationSnapshot? CaptureTableSnapshot(OoxmlPackage package, string target, CancellationToken cancellationToken)
    {
        TableTarget? tableTarget = ResolveTableTarget(package, target, cancellationToken);
        if (tableTarget is null)
        {
            return null;
        }

        XElement? templateRow = tableTarget.Table.Elements(OoxmlNs.W + "tr").LastOrDefault();
        return CreateTableOperationSnapshot(target, tableTarget.Table, rowIndex: null, columnIndex: null, cellCount: null, templateRow);
    }

    private static TableOperationSnapshot CreateTableOperationSnapshot(
        string target,
        XElement table,
        int? rowIndex,
        int? columnIndex,
        int? cellCount,
        XElement? row)
    {
        int rowCount = table.Elements(OoxmlNs.W + "tr").Count();
        int columnCount = TryGetConsistentVisualColumnCount(table, out int visualColumnCount)
            ? visualColumnCount
            : table.Elements(OoxmlNs.W + "tr").Select(ReadTableRowVisualColumnCount).DefaultIfEmpty(0).Max();
        string? tableId = ExtractTableId(target);
        int? gridBefore = row is null ? null : ReadTableRowGridOffset(row, "gridBefore");
        int? gridAfter = row is null ? null : ReadTableRowGridOffset(row, "gridAfter");
        IReadOnlyList<TableCellSnapshot> cells = tableId is null || row is null
            ? []
            : CreateTableCellSnapshots(tableId, table, row, rowIndex);
        return new TableOperationSnapshot(target, tableId, rowIndex, columnIndex, rowCount, columnCount, cellCount, gridBefore, gridAfter, cells);
    }

    private static IReadOnlyList<TableCellSnapshot> CreateTableCellSnapshots(
        string tableId,
        XElement table,
        XElement targetRow,
        int? targetRowIndex)
    {
        var snapshots = new List<TableCellSnapshot>();
        int mergeGroupIndex = 1;
        int rowIndex = 0;
        var activeVerticalMerges = new Dictionary<int, string>(capacity: 4);
        foreach (XElement row in table.Elements(OoxmlNs.W + "tr"))
        {
            rowIndex++;
            int gridBefore = ReadTableRowGridOffset(row, "gridBefore");
            RemoveActiveMergeGroupIds(activeVerticalMerges, 1, gridBefore);
            int columnIndex = 1 + gridBefore;
            foreach (XElement cell in row.Elements(OoxmlNs.W + "tc"))
            {
                int columnSpan = ReadTableCellColumnSpan(cell);
                string? verticalMerge = ReadTableCellVerticalMerge(cell);
                string? mergeGroupId = null;
                if (string.Equals(verticalMerge, "restart", StringComparison.Ordinal))
                {
                    mergeGroupId = AllocateTableMergeGroupId(tableId, ref mergeGroupIndex);
                    SetActiveMergeGroupId(activeVerticalMerges, columnIndex, columnSpan, mergeGroupId);
                }
                else if (verticalMerge is not null)
                {
                    mergeGroupId = FindActiveMergeGroupId(activeVerticalMerges, columnIndex, columnSpan) ?? AllocateTableMergeGroupId(tableId, ref mergeGroupIndex);
                    SetActiveMergeGroupId(activeVerticalMerges, columnIndex, columnSpan, mergeGroupId);
                }
                else
                {
                    RemoveActiveMergeGroupIds(activeVerticalMerges, columnIndex, columnSpan);
                    if (columnSpan > 1)
                    {
                        mergeGroupId = AllocateTableMergeGroupId(tableId, ref mergeGroupIndex);
                    }
                }

                if (ReferenceEquals(row, targetRow))
                {
                    string cellId = $"{tableId}.R{(targetRowIndex ?? rowIndex):00}.C{columnIndex:00}";
                    snapshots.Add(new TableCellSnapshot(
                        columnIndex,
                        columnIndex + columnSpan - 1,
                        mergeGroupId,
                        CreateNestedTablePath(cellId, cell)));
                }

                columnIndex += columnSpan;
            }

            int gridAfter = ReadTableRowGridOffset(row, "gridAfter");
            RemoveActiveMergeGroupIds(activeVerticalMerges, columnIndex, gridAfter);
        }

        return snapshots;
    }

    private static IReadOnlyList<TableCellSnapshot> CreateFallbackCellSnapshots(string tableId, int? rowIndex, int cellCount)
    {
        var snapshots = new List<TableCellSnapshot>(capacity: Math.Max(cellCount, 0));
        for (int column = 1; column <= cellCount; column++)
        {
            snapshots.Add(new TableCellSnapshot(column, column, MergeGroupId: null, NestedTablePath: null));
        }

        _ = tableId;
        _ = rowIndex;
        return snapshots;
    }

    private static string? FindActiveMergeGroupId(IReadOnlyDictionary<int, string> activeVerticalMerges, int columnIndex, int columnSpan)
    {
        string? mergeGroupId = null;
        for (int column = columnIndex; column < columnIndex + columnSpan; column++)
        {
            if (!activeVerticalMerges.TryGetValue(column, out string? current))
            {
                return null;
            }

            mergeGroupId ??= current;
            if (!string.Equals(mergeGroupId, current, StringComparison.Ordinal))
            {
                return null;
            }
        }

        return mergeGroupId;
    }

    private static void SetActiveMergeGroupId(Dictionary<int, string> activeVerticalMerges, int columnIndex, int columnSpan, string mergeGroupId)
    {
        for (int column = columnIndex; column < columnIndex + columnSpan; column++)
        {
            activeVerticalMerges[column] = mergeGroupId;
        }
    }

    private static void RemoveActiveMergeGroupIds(Dictionary<int, string> activeVerticalMerges, int columnIndex, int columnSpan)
    {
        for (int column = columnIndex; column < columnIndex + columnSpan; column++)
        {
            activeVerticalMerges.Remove(column);
        }
    }

    private static string AllocateTableMergeGroupId(string tableId, ref int mergeGroupIndex)
    {
        return $"{tableId}.MG{mergeGroupIndex++:0000}";
    }

    private static string? CreateNestedTablePath(string cellId, XElement cell)
    {
        int nestedTableCount = cell.Elements(OoxmlNs.W + "tbl").Count();
        return nestedTableCount switch
        {
            0 => null,
            1 => $"{cellId}.T0001",
            _ => $"{cellId}.T0001..T{nestedTableCount:0000}"
        };
    }

    private static IReadOnlyList<DocxPatchAffectedTarget> BuildAffectedTargets(DocxPatchOperation operation, TableOperationSnapshot? before)
    {
        if (before is null)
        {
            return [];
        }

        return operation.OperationName switch
        {
            "set-cell" or "set-cell-shading" => BuildSetCellAffectedTargets(before),
            "append-row" => BuildInsertedRowAffectedTargets(before, before.RowCountBefore + 1, operation.FieldValues.Count(field => field.Name == "cell"), "append"),
            "insert-row-before" => BuildInsertedRowAffectedTargets(before, before.RowIndex ?? 1, operation.FieldValues.Count(field => field.Name == "cell"), "insert"),
            "insert-row-after" => BuildInsertedRowAffectedTargets(before, (before.RowIndex ?? before.RowCountBefore) + 1, operation.FieldValues.Count(field => field.Name == "cell"), "insert"),
            "delete-row" => BuildDeletedRowAffectedTargets(before),
            _ => []
        };
    }

    private static IReadOnlyList<DocxPatchAffectedTarget> BuildSetCellAffectedTargets(TableOperationSnapshot before)
    {
        TableCellSnapshot? cell = before.Cells.FirstOrDefault(cell =>
            before.ColumnIndex is not null &&
            before.ColumnIndex.Value >= cell.ColumnIndex &&
            before.ColumnIndex.Value <= cell.VisualColumnEndIndex);
        return
        [
            new(before.TargetId, "cell", "update")
            {
                ParentId = before.TableId,
                RowIndex = before.RowIndex,
                ColumnIndex = before.ColumnIndex,
                VisualColumnEndIndex = cell?.VisualColumnEndIndex,
                MergeGroupId = cell?.MergeGroupId,
                NestedTablePath = cell?.NestedTablePath,
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
        IReadOnlyList<TableCellSnapshot> cells = before.Cells.Count == 0
            ? CreateFallbackCellSnapshots(tableId, before.RowIndex, requestedCellCount == 0 ? before.ColumnCount : requestedCellCount)
            : before.Cells;
        int cellCount = requestedCellCount == 0 ? cells.Count : requestedCellCount;
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
                CellCount = cellCount,
                GridBefore = before.GridBefore,
                GridAfter = before.GridAfter
            }
        };
        foreach (TableCellSnapshot cell in cells.Take(cellCount))
        {
            affected.Add(new DocxPatchAffectedTarget($"{rowId}.C{cell.ColumnIndex:00}", "cell", action)
            {
                ParentId = rowId,
                RowIndex = insertedRowIndex,
                ColumnIndex = cell.ColumnIndex,
                VisualColumnEndIndex = cell.VisualColumnEndIndex,
                NestedTablePath = cell.NestedTablePath is null ? null : $"{rowId}.C{cell.ColumnIndex:00}.T0001",
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
        IReadOnlyList<TableCellSnapshot> cells = before.Cells.Count == 0
            ? CreateFallbackCellSnapshots(tableId, rowIndex, before.CellCount ?? before.ColumnCount)
            : before.Cells;
        int cellCount = before.CellCount ?? cells.Count;
        var affected = new List<DocxPatchAffectedTarget>
        {
            new(before.TargetId, "row", "delete")
            {
                ParentId = tableId,
                RowIndex = rowIndex,
                RowCountBefore = before.RowCountBefore,
                RowCountAfter = before.RowCountBefore - 1,
                ColumnCount = before.ColumnCount,
                CellCount = cellCount,
                GridBefore = before.GridBefore,
                GridAfter = before.GridAfter
            }
        };
        foreach (TableCellSnapshot cell in cells.Take(cellCount))
        {
            affected.Add(new DocxPatchAffectedTarget($"{before.TargetId}.C{cell.ColumnIndex:00}", "cell", "delete")
            {
                ParentId = before.TargetId,
                RowIndex = rowIndex,
                ColumnIndex = cell.ColumnIndex,
                VisualColumnEndIndex = cell.VisualColumnEndIndex,
                MergeGroupId = cell.MergeGroupId,
                NestedTablePath = cell.NestedTablePath,
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
        int mergeGroupMarker = target.IndexOf(".MG", StringComparison.Ordinal);
        int marker = rowMarker switch
        {
            >= 0 when mergeGroupMarker >= 0 => Math.Min(rowMarker, mergeGroupMarker),
            >= 0 => rowMarker,
            _ => mergeGroupMarker
        };
        return marker < 0 ? target : target[..marker];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteReplaceText(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? find = ReadRequiredField(operation, "find", diagnostics);
        string? replacement = ReadRequiredField(operation, "with", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        bool? preserveRuns = ReadBooleanField(operation, "preserve-runs", diagnostics);
        int? occurrence = ReadPositiveOccurrence(operation, diagnostics);
        if (find is null || replacement is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        bool shouldPreserveRuns = preserveRuns ?? true;
        ParagraphTarget? paragraphTarget = ResolveParagraphTarget(package, operation, target, cancellationToken, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
        if (selectorDiagnostics.Count != 0)
        {
            return selectorDiagnostics;
        }

        if (paragraphTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (find.Length == 0)
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
            if (options.TrackChanges == TrackChangesMode.Require)
            {
                TrackUnsupportedShape(options, operation, target, $"paragraph contains protected OOXML boundary '{protectedFeature}'", diagnostics);
                return diagnostics;
            }

            return [Diagnostic(DocxSeverity.Error, "E4305", $"Text edit for {target} crosses protected OOXML boundary '{protectedFeature}'.", operation, target)];
        }

        IReadOnlyList<TextRange> matches = FindTextMatches(current, find, occurrence);
        if (matches.Count == 0)
        {
            return [Diagnostic(DocxSeverity.Error, "E4203", $"Find text was not found in {target}.", operation, target)];
        }

        bool useTrackedChanges = options.TrackChanges is TrackChangesMode.Require or TrackChangesMode.Suggest;
        bool canUseTrackedChanges = true;
        string? trackedUnsupportedReason = null;
        if (useTrackedChanges)
        {
            canUseTrackedChanges = TryValidateTrackedTextReplacement(paragraphTarget.Paragraph, current, matches, replacement, out trackedUnsupportedReason);
            if (!canUseTrackedChanges && trackedUnsupportedReason is not null &&
                !TrackUnsupportedShape(options, operation, target, trackedUnsupportedReason, diagnostics))
            {
                return diagnostics;
            }
        }

        if (!apply)
        {
            return diagnostics;
        }

        if (useTrackedChanges && canUseTrackedChanges)
        {
            ReplaceParagraphTextWithTrackedChanges(package, paragraphTarget.Paragraph, current, matches, replacement, options, generatedRevisionIds, cancellationToken);
            SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
            return [];
        }

        if (shouldPreserveRuns)
        {
            if (!TryReplaceParagraphTextPreservingRuns(paragraphTarget.Paragraph, matches, replacement, out string? unsupportedReason))
            {
                return [Diagnostic(DocxSeverity.Error, "E4306", $"Run-preserving replacement is not supported for {target}: {unsupportedReason}. Use preserve-runs false to allow paragraph-level rewriting.", operation, target)];
            }
        }
        else
        {
            string edited = ApplyTextReplacement(current, matches, replacement);
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
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        string? style = operation.Fields.GetValueOrDefault("style");
        if (text is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ParagraphTarget? paragraphTarget = ResolveParagraphTarget(package, operation, target, cancellationToken, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
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
            if (options.TrackChanges == TrackChangesMode.Require)
            {
                TrackUnsupportedShape(options, operation, target, $"paragraph contains protected OOXML boundary '{protectedFeature}'", diagnostics);
                return diagnostics;
            }

            return [Diagnostic(DocxSeverity.Error, "E4305", $"Paragraph replacement for {target} would remove protected OOXML boundary '{protectedFeature}'.", operation, target)];
        }

        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges &&
            !TryValidateTrackedWholeParagraphReplacement(paragraphTarget.Paragraph, current, text, style, out string? trackedUnsupportedReason))
        {
            if (!TrackUnsupportedShape(options, operation, target, trackedUnsupportedReason, diagnostics))
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
            ReplaceWholeParagraphTextWithTrackedChanges(package, paragraphTarget.Paragraph, current, text, options, generatedRevisionIds, cancellationToken);
            if (style is not null)
            {
                SetParagraphStyleWithTrackedChange(package, paragraphTarget.Paragraph, style, options, generatedRevisionIds, cancellationToken);
            }

            SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
            return diagnostics;
        }

        ReplaceParagraphText(paragraphTarget.Paragraph, text);
        if (style is not null)
        {
            SetParagraphStyle(paragraphTarget.Paragraph, style);
        }

        SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
        return diagnostics;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteInsertBlock(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool insertAfter,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        string? style = operation.Fields.GetValueOrDefault("style");
        bool copyParagraphProperties = ReadBooleanField(operation, "copy-paragraph-properties", diagnostics) ?? false;
        if (text is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        BlockTarget? blockTarget = ResolveBlockTarget(package, operation, target, cancellationToken, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
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
        if (useTrackedChanges && TextContainsTrackedUnsupportedCharacters(text))
        {
            if (!TrackUnsupportedShape(options, operation, target, "inserted paragraph text contains tabs or line breaks", diagnostics))
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
            ? CloneParagraphPropertiesForInsertion(blockTarget.Block.Element(OoxmlNs.W + "pPr"))
            : null;
        XElement paragraph = useTrackedChanges
            ? CreateTrackedInsertedParagraph(package, text, style, paragraphProperties, options, generatedRevisionIds, cancellationToken)
            : CreateSimpleParagraph(text, style, paragraphProperties);
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

    private static XElement? CloneParagraphPropertiesForInsertion(XElement? paragraphProperties)
    {
        if (paragraphProperties is null)
        {
            return null;
        }

        var clone = new XElement(paragraphProperties);
        clone.Elements(OoxmlNs.W + "pPrChange").Remove();
        clone.Elements(OoxmlNs.W + "sectPr").Remove();
        return clone.HasElements || clone.HasAttributes ? clone : null;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteDeleteBlock(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        BlockTarget? blockTarget = ResolveBlockTarget(package, operation, target, cancellationToken, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
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
                if (!TrackUnsupportedShape(options, operation, target, "tracked block deletion is supported only for paragraph targets", diagnostics))
                {
                    return diagnostics;
                }

                useTrackedChanges = false;
            }
            else if (ParagraphHasSectionProperties(blockTarget.Block))
            {
                if (!TrackUnsupportedShape(options, operation, target, "paragraph contains section properties", diagnostics))
                {
                    return diagnostics;
                }

                useTrackedChanges = false;
            }
            else if (TryGetProtectedTextEditFeature(blockTarget.Block, out string protectedFeature))
            {
                if (!TrackUnsupportedShape(options, operation, target, $"paragraph contains protected OOXML boundary '{protectedFeature}'", diagnostics))
                {
                    return diagnostics;
                }

                useTrackedChanges = false;
            }
            else if (TextContainsTrackedUnsupportedCharacters(current))
            {
                if (!TrackUnsupportedShape(options, operation, target, "deleted paragraph text contains tabs or line breaks", diagnostics))
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
            ReplaceWholeParagraphTextWithTrackedChanges(package, blockTarget.Block, current, string.Empty, options, generatedRevisionIds, cancellationToken);
            SaveDocumentPart(package, blockTarget.PartName, blockTarget.Document);
            return diagnostics;
        }

        blockTarget.Block.Remove();
        SaveDocumentPart(package, blockTarget.PartName, blockTarget.Document);
        return diagnostics;
    }

    private static bool ParagraphHasSectionProperties(XElement paragraph)
    {
        return paragraph
            .Element(OoxmlNs.W + "pPr")
            ?.Element(OoxmlNs.W + "sectPr") is not null;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetStyle(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? style = ReadRequiredField(operation, "style", diagnostics);
        if (style is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ParagraphTarget? paragraphTarget = ResolveParagraphTarget(package, operation, target, cancellationToken, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
        if (selectorDiagnostics.Count != 0)
        {
            return selectorDiagnostics;
        }

        if (paragraphTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!TryResolveStyleId(package, style, "paragraph", cancellationToken, out string? styleId, out DocxDiagnostic? styleDiagnostic, operation, target))
        {
            return [styleDiagnostic!];
        }

        if (!apply)
        {
            return [];
        }

        if (IsTrackedMode(options))
        {
            SetParagraphStyleWithTrackedChange(package, paragraphTarget.Paragraph, styleId, options, generatedRevisionIds, cancellationToken);
        }
        else
        {
            SetParagraphStyle(paragraphTarget.Paragraph, styleId);
        }

        SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetContentControlText(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        if (text is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ContentControlTarget? controlTarget = ResolveContentControlTarget(package, target, cancellationToken);
        if (controlTarget is null && !IsSupportedContentControlTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported content-control target '{target}'. Expected a content control ID such as M.CC0001 or H001.CC0001.", operation, target)];
        }

        if (controlTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        bool isPlainText = IsPlainTextContentControl(controlTarget.ContentControl);
        bool isRichText = IsRichTextContentControl(controlTarget.ContentControl);
        if (!isPlainText && !isRichText)
        {
            return [UnsupportedContentControlKindDiagnostic(controlTarget.ContentControl, operation, target, "a plain-text or rich-text content control")];
        }

        DocxDiagnostic? lockDiagnostic = ValidateContentControlUnlocked(controlTarget.ContentControl, operation, target);
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

        if (!isPlainText)
        {
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

        bool useTrackedChanges = IsTrackedMode(options);
        XElement? trackedContainer = null;
        XElement[]? trackedParagraphs = null;
        string trackedCurrent = current;
        if (useTrackedChanges)
        {
            if (!isPlainText)
            {
                if (!TryGetTrackedRichTextContentControlParagraphs(content, text, out trackedParagraphs, out string? trackedUnsupportedReason))
                {
                    if (!TrackUnsupportedShape(options, operation, target, trackedUnsupportedReason, diagnostics))
                    {
                        return diagnostics;
                    }

                    useTrackedChanges = false;
                }
            }
            else if (!TryGetTrackedContentControlTextContainer(content, text, out trackedContainer, out trackedCurrent, out string? trackedUnsupportedReason))
            {
                if (!TrackUnsupportedShape(options, operation, target, trackedUnsupportedReason, diagnostics))
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
            if (trackedParagraphs is not null)
            {
                ReplaceCellParagraphTextWithTrackedChanges(package, trackedParagraphs, text, options, generatedRevisionIds, cancellationToken);
            }
            else if (trackedContainer is not null)
            {
                ReplaceWholeParagraphTextWithTrackedChanges(package, trackedContainer, trackedCurrent, text, options, generatedRevisionIds, cancellationToken);
            }

            SaveDocumentPart(package, controlTarget.PartName, controlTarget.Document);
            return diagnostics;
        }

        ReplaceContentControlText(content, text);
        SaveDocumentPart(package, controlTarget.PartName, controlTarget.Document);
        return diagnostics;
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
        if (target is null || checkedValue is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ContentControlTarget? controlTarget = ResolveContentControlTarget(package, target, cancellationToken);
        if (controlTarget is null && !IsSupportedContentControlTargetShape(target))
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
            return [UnsupportedContentControlKindDiagnostic(controlTarget.ContentControl, operation, target, "a checkbox content control")];
        }

        DocxDiagnostic? lockDiagnostic = ValidateContentControlUnlocked(controlTarget.ContentControl, operation, target);
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

        SetCheckboxChecked(checkBox, checkedValue.Value);
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

        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ContentControlTarget? controlTarget = ResolveContentControlTarget(package, target, cancellationToken);
        if (controlTarget is null && !IsSupportedContentControlTargetShape(target))
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
            return [UnsupportedContentControlKindDiagnostic(controlTarget.ContentControl, operation, target, "a dropdown or combo box content control")];
        }

        DocxDiagnostic? lockDiagnostic = ValidateContentControlUnlocked(controlTarget.ContentControl, operation, target);
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
        if (value is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ContentControlTarget? controlTarget = ResolveContentControlTarget(package, target, cancellationToken);
        if (controlTarget is null && !IsSupportedContentControlTargetShape(target))
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
            return [UnsupportedContentControlKindDiagnostic(controlTarget.ContentControl, operation, target, "a date content control")];
        }

        DocxDiagnostic? lockDiagnostic = ValidateContentControlUnlocked(controlTarget.ContentControl, operation, target);
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

        SetContentControlDateValue(date, value);
        ReplaceContentControlText(content, displayText ?? value);
        SaveDocumentPart(package, controlTarget.PartName, controlTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteReplaceBookmarkText(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        if (text is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        BookmarkTarget? bookmarkTarget = ResolveBookmarkTarget(package, target, cancellationToken);
        if (bookmarkTarget is null && !IsSupportedBookmarkTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported bookmark target '{target}'. Expected a bookmark ID such as M.B0001 or H001.B0001.", operation, target)];
        }

        if (bookmarkTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (bookmarkTarget.End is null ||
            bookmarkTarget.Start.Parent is null ||
            bookmarkTarget.End.Parent is null ||
            bookmarkTarget.Start.Parent.Name != OoxmlNs.W + "p" ||
            bookmarkTarget.End.Parent.Name != OoxmlNs.W + "p")
        {
            return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark '{target}' is not a paragraph-bounded range.", operation, target)];
        }

        XElement startParagraph = bookmarkTarget.Start.Parent;
        XElement endParagraph = bookmarkTarget.End.Parent;
        bool sameParagraph = startParagraph == endParagraph;
        XNode[] intermediateNodes = [];
        BookmarkTextSlot[]? structuredTextSlots = null;
        string? rangeFailure = null;
        XNode[] nodes = sameParagraph
            ? bookmarkTarget.Start.NodesAfterSelf()
                .TakeWhile(node => node != bookmarkTarget.End)
                .ToArray()
            : GetMultiParagraphBookmarkReplacementNodes(
                startParagraph,
                bookmarkTarget.Start,
                endParagraph,
                bookmarkTarget.End,
                out intermediateNodes,
                out rangeFailure);
        if (!sameParagraph && rangeFailure is not null)
        {
            if (!TryBuildTableSpanningBookmarkTextSlots(
                startParagraph,
                bookmarkTarget.Start,
                endParagraph,
                bookmarkTarget.End,
                out structuredTextSlots,
                out nodes,
                out string? structuredRangeFailure))
            {
                return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark '{target}' is not a supported same-story multi-paragraph range: {structuredRangeFailure ?? rangeFailure}.", operation, target)];
            }
        }

        if (ContainsProtectedBookmarkReplacementNode(nodes, out string? protectedFeature))
        {
            return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark '{target}' replacement would remove protected OOXML boundary '{protectedFeature}'.", operation, target)];
        }

        string[]? structuredReplacementLines = null;
        if (structuredTextSlots is not null)
        {
            structuredReplacementLines = SplitReplacementParagraphText(text);
            if (structuredReplacementLines.Length != structuredTextSlots.Length)
            {
                return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark '{target}' table-spanning replacement requires {structuredTextSlots.Length} replacement lines, one per visible text slot, but received {structuredReplacementLines.Length}.", operation, target)];
            }
        }

        bool useTrackedChanges = IsTrackedMode(options);
        TrackedBookmarkReplacement? trackedReplacement = null;
        if (useTrackedChanges)
        {
            if (!sameParagraph)
            {
                if (!TrackUnsupportedShape(
                    options,
                    operation,
                    target,
                    "tracked replace-bookmark-text supports only simple same-paragraph bookmark ranges",
                    diagnostics))
                {
                    return diagnostics;
                }

                useTrackedChanges = false;
            }
            else if (!TryBuildTrackedBookmarkReplacement(nodes, text, out trackedReplacement, out string? trackedUnsupportedReason))
            {
                if (!TrackUnsupportedShape(options, operation, target, trackedUnsupportedReason, diagnostics))
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

        if (useTrackedChanges && trackedReplacement is not null)
        {
            foreach (XNode node in nodes)
            {
                node.Remove();
            }

            XNode[] trackedNodes = CreateTrackedBookmarkReplacementNodes(
                package,
                trackedReplacement.DeletedText,
                text,
                trackedReplacement.RunProperties,
                options,
                generatedRevisionIds,
                cancellationToken);
            bookmarkTarget.Start.AddAfterSelf(trackedNodes);
        }
        else if (sameParagraph)
        {
            foreach (XNode node in nodes)
            {
                node.Remove();
            }

            bookmarkTarget.Start.AddAfterSelf(CreateSimpleRun(text));
        }
        else if (structuredTextSlots is not null && structuredReplacementLines is not null)
        {
            ReplaceStructuredBookmarkTextSlots(structuredTextSlots, structuredReplacementLines);
        }
        else
        {
            ReplaceMultiParagraphBookmarkText(
                startParagraph,
                bookmarkTarget.Start,
                endParagraph,
                bookmarkTarget.End,
                intermediateNodes,
                text);
        }

        SaveDocumentPart(package, bookmarkTarget.PartName, bookmarkTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteAddBookmark(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? name = ReadRequiredField(operation, "name", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        if (name is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!IsValidBookmarkName(name))
        {
            return [Diagnostic(DocxSeverity.Error, "E4205", "Field 'name' must be a non-empty bookmark name without whitespace.", operation, target)];
        }

        ParagraphTarget? paragraphTarget = ResolveParagraphTarget(package, operation, target, cancellationToken, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
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

        if (BookmarkNameExists(paragraphTarget.Document, name))
        {
            return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark name '{name}' already exists in part '{paragraphTarget.PartName}'.", operation, target)];
        }

        if (TryGetProtectedTextEditFeature(paragraphTarget.Paragraph, out string protectedFeature))
        {
            return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark creation for {target} would cross protected OOXML boundary '{protectedFeature}'.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        string id = AllocateBookmarkId(paragraphTarget.Document);
        var start = new XElement(
            OoxmlNs.W + "bookmarkStart",
            new XAttribute(OoxmlNs.W + "id", id),
            new XAttribute(OoxmlNs.W + "name", name));
        var end = new XElement(
            OoxmlNs.W + "bookmarkEnd",
            new XAttribute(OoxmlNs.W + "id", id));

        XElement? paragraphProperties = paragraphTarget.Paragraph.Element(OoxmlNs.W + "pPr");
        if (paragraphProperties is null)
        {
            paragraphTarget.Paragraph.AddFirst(start);
        }
        else
        {
            paragraphProperties.AddAfterSelf(start);
        }

        paragraphTarget.Paragraph.Add(end);
        SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
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
        if (name is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!IsValidBookmarkName(name))
        {
            return [Diagnostic(DocxSeverity.Error, "E4205", "Field 'name' must be a non-empty bookmark name without whitespace.", operation, target)];
        }

        BookmarkTarget? bookmarkTarget = ResolveBookmarkTarget(package, target, cancellationToken);
        if (bookmarkTarget is null && !IsSupportedBookmarkTargetShape(target))
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

        if (BookmarkNameExists(bookmarkTarget.Document, bookmarkTarget.Start, name))
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
        UpdateInternalHyperlinkAnchors(bookmarkTarget.Document, oldName, name);
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
        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        BookmarkTarget? bookmarkTarget = ResolveBookmarkTarget(package, target, cancellationToken);
        if (bookmarkTarget is null && !IsSupportedBookmarkTargetShape(target))
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
            HasInternalHyperlinkAnchor(bookmarkTarget.Document, name))
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
        string? anchorText = operation.Fields.GetValueOrDefault("anchor-text");
        int? occurrence = ReadPositiveOccurrence(operation, diagnostics);
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

        if (operation.Fields.ContainsKey("occurrence") && anchorText is null)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Field 'occurrence' requires field 'anchor-text'.", operation, target));
        }

        if (anchorText is not null && anchorText.Length == 0)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Field 'anchor-text' must not be empty.", operation, target));
        }

        if (target is null || text is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ParagraphTarget? paragraphTarget = ResolveParagraphTarget(package, operation, target, cancellationToken, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
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

        TextRange? anchorRange = null;
        if (anchorText is not null)
        {
            IReadOnlyList<TextRange> matches = FindTextMatches(current, anchorText, occurrence);
            if (matches.Count == 0)
            {
                return [Diagnostic(DocxSeverity.Error, "E4203", $"Anchor text was not found in {target}.", operation, target)];
            }

            if (occurrence is null && matches.Count > 1)
            {
                return
                [
                    Diagnostic(
                        DocxSeverity.Error,
                        "E1202",
                        $"Anchor text matched {matches.Count} ranges in {target}. Specify occurrence to select one range.",
                        operation,
                        target)
                ];
            }

            if (TryGetProtectedTextEditFeature(paragraphTarget.Paragraph, out string protectedFeature))
            {
                return [Diagnostic(DocxSeverity.Error, "E4305", $"Selected comment range for {target} crosses protected OOXML boundary '{protectedFeature}'.", operation, target)];
            }

            anchorRange = matches[0];
            if (!CanAddSelectedCommentAnchor(paragraphTarget.Paragraph, anchorRange.Value, out string? unsupportedReason))
            {
                return [Diagnostic(DocxSeverity.Error, "E4317", $"Selected comment range is not supported for {target}: {unsupportedReason}.", operation, target)];
            }
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
        commentsPart.Root.Add(CreateComment(commentId, text, author, initials, timestampUtc));
        if (anchorRange is TextRange selectedRange)
        {
            AddSelectedCommentAnchor(paragraphTarget.Paragraph, commentId, selectedRange);
        }
        else
        {
            AddCommentAnchor(paragraphTarget.Paragraph, commentId);
        }

        SaveDocumentPart(package, commentsPart.PartName, commentsPart.Document);
        SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetCommentText(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        if (text is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        CommentTarget? commentTarget = ResolveCommentTarget(package, target, operation, cancellationToken, out DocxDiagnostic? diagnostic);
        if (diagnostic is not null)
        {
            return [diagnostic];
        }

        if (commentTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 comments: {target}.", operation, target)];
        }

        bool useTrackedChanges = IsTrackedMode(options);
        XElement[]? trackedParagraphs = null;
        if (useTrackedChanges &&
            !TryGetTrackedCommentParagraphs(commentTarget.Comment, text, out trackedParagraphs, out string? trackedUnsupportedReason))
        {
            if (!TrackUnsupportedShape(options, operation, target, trackedUnsupportedReason, diagnostics))
            {
                return diagnostics;
            }

            useTrackedChanges = false;
        }

        if (!apply)
        {
            return diagnostics;
        }

        if (useTrackedChanges && trackedParagraphs is not null)
        {
            ReplaceCellParagraphTextWithTrackedChanges(package, trackedParagraphs, text, options, generatedRevisionIds, cancellationToken);
        }
        else
        {
            ReplaceCommentText(commentTarget.Comment, text);
        }

        SaveDocumentPart(package, commentTarget.PartName, commentTarget.Document);
        return diagnostics;
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
        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        CommentTarget? commentTarget = ResolveCommentTarget(package, target, operation, cancellationToken, out DocxDiagnostic? diagnostic);
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
            target,
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
        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        CommentTarget? commentTarget = ResolveCommentTarget(package, target, operation, cancellationToken, out DocxDiagnostic? diagnostic);
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
        RemoveCommentIdRecords(package, commentTarget.Comment, cancellationToken);
        commentTarget.Comment.Remove();
        SaveDocumentPart(package, commentTarget.PartName, commentTarget.Document);
        RemoveCommentAnchors(package, commentId, cancellationToken);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteAddCommentReply(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
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

        if (text is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        CommentTarget? parentTarget = ResolveCommentTarget(package, target, operation, cancellationToken, out DocxDiagnostic? diagnostic);
        if (diagnostic is not null)
        {
            return [diagnostic];
        }

        if (parentTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 comments: {target}.", operation, target)];
        }

        if (parentTarget.Comment.Elements(OoxmlNs.W + "p").FirstOrDefault() is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E4314", $"Comment '{target}' has no body paragraph for reply metadata.", operation, target)];
        }

        DocxDiagnostic? commentsPartDiagnostic = ValidateExistingCommentsPart(package, cancellationToken);
        if (commentsPartDiagnostic is not null)
        {
            return [commentsPartDiagnostic];
        }

        DocxDiagnostic? commentsExtendedDiagnostic = ValidateExistingCommentsExtendedPart(package, cancellationToken);
        if (commentsExtendedDiagnostic is not null)
        {
            return [commentsExtendedDiagnostic];
        }

        DocxDiagnostic? commentsIdsDiagnostic = ValidateExistingCommentsIdsPart(package, cancellationToken);
        if (commentsIdsDiagnostic is not null)
        {
            return [commentsIdsDiagnostic];
        }

        if (!apply)
        {
            return [];
        }

        CommentExtensionTarget? parentExtensionTarget = ResolveOrCreateCommentExtensionTarget(
            package,
            parentTarget,
            operation,
            target,
            cancellationToken,
            out diagnostic,
            out _);
        if (diagnostic is not null)
        {
            return [diagnostic];
        }

        if (parentExtensionTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E4314", $"Comment '{target}' cannot host reply metadata.", operation, target)];
        }

        string parentParaId = ReadCommentParaId(parentTarget.Comment)
            ?? throw new InvalidDataException("Parent comment paraId was not created.");
        string replyParaId = AllocateCommentParaId(package, [parentParaId], cancellationToken);
        string replyCommentId = AllocateCommentId(package, cancellationToken);
        EnsureNamespaceDeclaration(parentTarget.Document.Root, "w15", OoxmlNs.W15);
        parentTarget.Comment.AddAfterSelf(CreateComment(replyCommentId, text, author, initials, timestampUtc, replyParaId));

        XElement parentExtensionRoot = parentExtensionTarget?.Document.Root
            ?? throw new InvalidDataException("commentsExtended document has no XML root.");
        EnsureNamespaceDeclaration(parentExtensionRoot, "w15", OoxmlNs.W15);
        parentExtensionRoot.Add(new XElement(
            OoxmlNs.W15 + "commentEx",
            new XAttribute(OoxmlNs.W15 + "paraId", replyParaId),
            new XAttribute(OoxmlNs.W15 + "paraIdParent", parentParaId),
            new XAttribute(OoxmlNs.W15 + "done", "0")));

        CommentsIdsPartTarget commentsIdsPart = ResolveOrCreateCommentsIdsPart(package, cancellationToken);
        EnsureNamespaceDeclaration(commentsIdsPart.Root, "w16cid", OoxmlNs.W16Cid);
        commentsIdsPart.Root.Add(new XElement(
            OoxmlNs.W16Cid + "commentId",
            new XAttribute(OoxmlNs.W16Cid + "paraId", replyParaId),
            new XAttribute(OoxmlNs.W16Cid + "durableId", AllocateCommentDurableId(package, cancellationToken))));

        SaveDocumentPart(package, parentTarget.PartName, parentTarget.Document);
        SaveDocumentPart(package, parentExtensionTarget.PartName, parentExtensionTarget.Document);
        SaveDocumentPart(package, commentsIdsPart.PartName, commentsIdsPart.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteDeleteCommentReply(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        CommentReplyTarget? replyTarget = ResolveCommentReplyTarget(package, target, operation, cancellationToken, out DocxDiagnostic? diagnostic);
        if (diagnostic is not null)
        {
            return [diagnostic];
        }

        if (replyTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 comment replies: {target}.", operation, target)];
        }

        if (HasChildCommentReplies(package, replyTarget.ParaId, cancellationToken))
        {
            return [Diagnostic(DocxSeverity.Error, "E4314", $"Comment reply '{target}' has child replies and cannot be deleted without changing thread topology.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        string replyCommentId = (string?)replyTarget.Comment.Attribute(OoxmlNs.W + "id") ?? string.Empty;
        RemoveCommentExtensionRecords(package, replyTarget.Comment, cancellationToken);
        RemoveCommentIdRecords(package, replyTarget.Comment, cancellationToken);
        replyTarget.Comment.Remove();
        SaveDocumentPart(package, replyTarget.PartName, replyTarget.Document);
        RemoveCommentAnchors(package, replyCommentId, cancellationToken);
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
        if (target is null || value is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (IsAllFieldsTarget(target))
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
                targetField.Element.SetAttributeValue(OoxmlNs.W + attributeName, value.Value ? "true" : "false");
            }

            foreach (IGrouping<string, FieldTarget> partGroup in fieldTargets.GroupBy(fieldTarget => fieldTarget.PartName, StringComparer.Ordinal))
            {
                SaveDocumentPart(package, partGroup.Key, partGroup.First().Document);
            }

            return [];
        }

        FieldTarget? fieldTarget = ResolveFieldTarget(package, target, cancellationToken);
        if (fieldTarget is null && !IsSupportedFieldTargetShape(target))
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

        fieldTarget.Element.SetAttributeValue(OoxmlNs.W + attributeName, value.Value ? "true" : "false");
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
        if (code is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        FieldTarget? fieldTarget = ResolveFieldTarget(package, target, cancellationToken);
        if (fieldTarget is null && !IsSupportedFieldTargetShape(target))
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

        fieldTarget.Element.SetAttributeValue(OoxmlNs.W + "instr", code);
        fieldTarget.Element.SetAttributeValue(OoxmlNs.W + "dirty", "true");
        SaveDocumentPart(package, fieldTarget.PartName, fieldTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetFieldResult(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-result");
        if (text is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        FieldTarget? fieldTarget = ResolveFieldTarget(package, target, cancellationToken);
        if (fieldTarget is null && !IsSupportedFieldTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported field target '{target}'. Expected a field ID such as M.F0001 or H001.F0001.", operation, target)];
        }

        if (fieldTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 fields: {target}.", operation, target)];
        }

        bool simpleField = fieldTarget.Element.Name == OoxmlNs.W + "fldSimple";
        XElement? complexSeparateRun = null;
        XElement[] complexResultRuns = [];
        if (!simpleField)
        {
            if (!TryGetSimpleComplexFieldResultRuns(
                fieldTarget.Element,
                out complexSeparateRun,
                out complexResultRuns,
                out string? complexUnsupportedReason))
            {
                return [Diagnostic(DocxSeverity.Error, "E4313", $"Cached-result replacement for {target} is not safe: {complexUnsupportedReason}.", operation, target)];
            }
        }

        string current = simpleField
            ? ReadVisibleText(fieldTarget.Element)
            : ReadVisibleText(new XElement(OoxmlNs.W + "p", complexResultRuns));
        if (expected is not null && !string.Equals(current, expected, StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected field result does not match current result.", operation, target)];
        }

        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges)
        {
            if (!simpleField)
            {
                if (!TrackUnsupportedShape(options, operation, target, "tracked complex-field result replacement is not modeled", diagnostics))
                {
                    return diagnostics;
                }

                useTrackedChanges = false;
            }
            else if (TryGetProtectedTextEditFeature(fieldTarget.Element, out string protectedFeature))
            {
                if (!TrackUnsupportedShape(options, operation, target, $"field result contains protected OOXML boundary '{protectedFeature}'", diagnostics))
                {
                    return diagnostics;
                }

                useTrackedChanges = false;
            }
            else if (!TryValidateTrackedWholeParagraphReplacement(fieldTarget.Element, current, text, style: null, out string? trackedUnsupportedReason))
            {
                if (!TrackUnsupportedShape(options, operation, target, trackedUnsupportedReason, diagnostics))
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
            ReplaceWholeParagraphTextWithTrackedChanges(package, fieldTarget.Element, current, text, options, generatedRevisionIds, cancellationToken);
            SaveDocumentPart(package, fieldTarget.PartName, fieldTarget.Document);
            return diagnostics;
        }

        if (simpleField)
        {
            ReplaceSimpleFieldResult(fieldTarget.Element, text);
        }
        else if (complexSeparateRun is not null)
        {
            ReplaceComplexFieldResult(complexSeparateRun, complexResultRuns, text);
        }

        SaveDocumentPart(package, fieldTarget.PartName, fieldTarget.Document);
        return diagnostics;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteRefreshFieldResult(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? expectedCode = operation.Fields.GetValueOrDefault("expect-code");
        string? expectedResult = operation.Fields.GetValueOrDefault("expect-result");
        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        FieldTarget? fieldTarget = ResolveFieldTarget(package, target, cancellationToken);
        if (fieldTarget is null && !IsSupportedFieldTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported field target '{target}'. Expected a field ID such as M.F0001 or H001.F0001.", operation, target)];
        }

        if (fieldTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 fields: {target}.", operation, target)];
        }

        if (fieldTarget.Element.Name != OoxmlNs.W + "fldSimple")
        {
            return [Diagnostic(DocxSeverity.Error, "E4313", $"Field refresh for {target} currently supports only simple w:fldSimple REF-style fields.", operation, target)];
        }

        string currentCode = NormalizeFieldCodeForGuard((string?)fieldTarget.Element.Attribute(OoxmlNs.W + "instr") ?? string.Empty);
        if (expectedCode is not null && !string.Equals(currentCode, NormalizeFieldCodeForGuard(expectedCode), StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected field code does not match current code.", operation, target)];
        }

        string currentResult = ReadVisibleText(fieldTarget.Element);
        if (expectedResult is not null && !string.Equals(currentResult, expectedResult, StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected field result does not match current result.", operation, target)];
        }

        if (TryReadQuoteFieldText(currentCode, out string? quoteText))
        {
            if (!apply)
            {
                return [];
            }

            ReplaceSimpleFieldResult(fieldTarget.Element, quoteText);
            fieldTarget.Element.SetAttributeValue(OoxmlNs.W + "dirty", null);
            SaveDocumentPart(package, fieldTarget.PartName, fieldTarget.Document);
            return [];
        }

        if (!TryReadRefFieldBookmarkName(currentCode, out string? bookmarkName))
        {
            return [UnsupportedFieldRefreshDiagnostic(operation, target, currentCode)];
        }

        if (!TryReadSimpleBookmarkText(fieldTarget.Document, bookmarkName, out string? bookmarkText, out string? unsupportedReason))
        {
            return [Diagnostic(DocxSeverity.Error, "E4313", $"Field refresh for {target} cannot resolve bookmark '{bookmarkName}': {unsupportedReason}.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        ReplaceSimpleFieldResult(fieldTarget.Element, bookmarkText);
        fieldTarget.Element.SetAttributeValue(OoxmlNs.W + "dirty", null);
        SaveDocumentPart(package, fieldTarget.PartName, fieldTarget.Document);
        return [];
    }

    private static string NormalizeFieldCodeForGuard(string code)
    {
        return string.Join(
            " ",
            code.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));
    }

    private static bool TryReadRefFieldBookmarkName(string code, [NotNullWhen(true)] out string? bookmarkName)
    {
        bookmarkName = null;
        string[] tokens = TokenizeFieldCodeForPatch(code);
        if (tokens.Length < 2)
        {
            return false;
        }

        string fieldType = NormalizeFieldTypeForPatch(tokens[0]);
        if (fieldType is not ("REF" or "PAGEREF" or "NOTEREF"))
        {
            return false;
        }

        foreach (string token in tokens.Skip(1))
        {
            if (token.StartsWith('\\'))
            {
                continue;
            }

            bookmarkName = token;
            return !string.IsNullOrWhiteSpace(bookmarkName);
        }

        return false;
    }

    private static bool TryReadQuoteFieldText(string code, [NotNullWhen(true)] out string? text)
    {
        text = null;
        string[] tokens = TokenizeFieldCodeForPatch(code);
        if (tokens.Length < 2 || NormalizeFieldTypeForPatch(tokens[0]) != "QUOTE")
        {
            return false;
        }

        string[] arguments = tokens
            .Skip(1)
            .TakeWhile(token => !token.StartsWith('\\'))
            .ToArray();
        if (arguments.Length == 0)
        {
            return false;
        }

        text = string.Join(" ", arguments);
        return true;
    }

    private static DocxDiagnostic UnsupportedFieldRefreshDiagnostic(
        DocxPatchOperation operation,
        string target,
        string code)
    {
        string[] tokens = TokenizeFieldCodeForPatch(code);
        string fieldType = tokens.Length == 0 ? "unknown" : NormalizeFieldTypeForPatch(tokens[0]);
        string reason = fieldType switch
        {
            "TOC" or "PAGE" or "NUMPAGES" or "SECTIONPAGES" => "requires Word layout or pagination state",
            "DOCPROPERTY" or "DOCVARIABLE" or "AUTHOR" or "TITLE" or "SUBJECT" or "KEYWORDS" => "requires document property state",
            "MERGEFIELD" or "MERGEREC" or "MERGESEQ" or "NEXT" or "NEXTIF" or "SKIPIF" => "requires mail merge data or mail merge state",
            "FORMULA" => "requires Word formula evaluation",
            "IF" => "requires Word conditional field evaluation",
            "DATE" or "TIME" or "CREATEDATE" or "SAVEDATE" or "PRINTDATE" => "requires Word date/time evaluation",
            "HYPERLINK" or "INCLUDETEXT" or "INCLUDEPICTURE" or "LINK" => "requires hyperlink or external target state",
            _ => "is not modeled for deterministic refresh"
        };
        return Diagnostic(
            DocxSeverity.Error,
            "E4313",
            $"Field refresh for {target} does not support field type '{fieldType}': {reason}. Supported deterministic refresh fields are REF, PAGEREF, NOTEREF, and QUOTE.",
            operation,
            target);
    }

    private static string NormalizeFieldTypeForPatch(string token)
    {
        return token.StartsWith('=') ? "FORMULA" : token.ToUpperInvariant();
    }

    private static string[] TokenizeFieldCodeForPatch(string code)
    {
        var tokens = new List<string>();
        var current = new StringBuilder();
        bool inQuote = false;
        foreach (char character in code)
        {
            if (character == '"')
            {
                inQuote = !inQuote;
                continue;
            }

            if (char.IsWhiteSpace(character) && !inQuote)
            {
                AddFieldCodeToken(tokens, current);
                continue;
            }

            current.Append(character);
        }

        AddFieldCodeToken(tokens, current);
        return tokens.ToArray();
    }

    private static void AddFieldCodeToken(List<string> tokens, StringBuilder current)
    {
        if (current.Length == 0)
        {
            return;
        }

        tokens.Add(current.ToString());
        current.Clear();
    }

    private static bool TryReadSimpleBookmarkText(
        XDocument document,
        string bookmarkName,
        [NotNullWhen(true)] out string? text,
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        text = null;
        unsupportedReason = null;
        XElement[] starts = document
            .Descendants(OoxmlNs.W + "bookmarkStart")
            .Where(bookmark => string.Equals((string?)bookmark.Attribute(OoxmlNs.W + "name"), bookmarkName, StringComparison.Ordinal))
            .ToArray();
        if (starts.Length == 0)
        {
            unsupportedReason = "bookmark was not found";
            return false;
        }

        if (starts.Length > 1)
        {
            unsupportedReason = "bookmark name is ambiguous";
            return false;
        }

        XElement start = starts[0];
        string? ooxmlId = (string?)start.Attribute(OoxmlNs.W + "id");
        XElement? end = ooxmlId is null
            ? null
            : document
                .Descendants(OoxmlNs.W + "bookmarkEnd")
                .FirstOrDefault(element => string.Equals((string?)element.Attribute(OoxmlNs.W + "id"), ooxmlId, StringComparison.Ordinal));
        if (end is null)
        {
            unsupportedReason = "bookmark end marker was not found";
            return false;
        }

        if (start.Parent is null || start.Parent != end.Parent || start.Parent.Name != OoxmlNs.W + "p")
        {
            unsupportedReason = "bookmark range is not a simple same-paragraph range";
            return false;
        }

        XNode[] nodes = start.NodesAfterSelf()
            .TakeWhile(node => node != end)
            .ToArray();
        if (ContainsProtectedBookmarkReplacementNode(nodes, out string? protectedFeature))
        {
            unsupportedReason = $"bookmark range contains protected OOXML boundary '{protectedFeature}'";
            return false;
        }

        text = ReadVisibleText(new XElement(OoxmlNs.W + "p", nodes));
        return true;
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

    private static bool TryGetSimpleComplexFieldResultRuns(
        XElement beginFieldChar,
        out XElement? separateRun,
        out XElement[] resultRuns,
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        separateRun = null;
        resultRuns = [];
        unsupportedReason = null;
        if (beginFieldChar.Name != OoxmlNs.W + "fldChar" ||
            !string.Equals((string?)beginFieldChar.Attribute(OoxmlNs.W + "fldCharType"), "begin", StringComparison.Ordinal))
        {
            unsupportedReason = "target is not a complex field begin marker";
            return false;
        }

        XElement? beginRun = beginFieldChar.Parent;
        XElement? paragraph = beginRun?.Parent;
        if (beginRun?.Name != OoxmlNs.W + "r" || paragraph?.Name != OoxmlNs.W + "p")
        {
            unsupportedReason = "complex field markers must be direct run children in one paragraph";
            return false;
        }

        XElement[] siblings = paragraph.Elements().ToArray();
        int beginIndex = Array.IndexOf(siblings, beginRun);
        if (beginIndex < 0)
        {
            unsupportedReason = "complex field begin marker was not found in its paragraph";
            return false;
        }

        var results = new List<XElement>();
        bool foundSeparate = false;
        for (int i = beginIndex + 1; i < siblings.Length; i++)
        {
            XElement sibling = siblings[i];
            if (sibling.Name != OoxmlNs.W + "r")
            {
                unsupportedReason = foundSeparate
                    ? $"complex field result contains non-run content '{sibling.Name.LocalName}'"
                    : $"complex field instruction contains non-run content '{sibling.Name.LocalName}'";
                return false;
            }

            XElement[] fieldChars = sibling.Descendants(OoxmlNs.W + "fldChar").ToArray();
            if (!foundSeparate)
            {
                foreach (XElement fieldChar in fieldChars)
                {
                    string? type = (string?)fieldChar.Attribute(OoxmlNs.W + "fldCharType");
                    if (string.Equals(type, "begin", StringComparison.Ordinal))
                    {
                        unsupportedReason = "complex field contains nested field topology before the result";
                        return false;
                    }

                    if (string.Equals(type, "separate", StringComparison.Ordinal))
                    {
                        separateRun = sibling;
                        foundSeparate = true;
                        break;
                    }

                    if (string.Equals(type, "end", StringComparison.Ordinal))
                    {
                        unsupportedReason = "complex field has an end marker before its separate marker";
                        return false;
                    }
                }

                continue;
            }

            foreach (XElement fieldChar in fieldChars)
            {
                string? type = (string?)fieldChar.Attribute(OoxmlNs.W + "fldCharType");
                if (string.Equals(type, "begin", StringComparison.Ordinal))
                {
                    unsupportedReason = "complex field result contains nested complex field markup";
                    return false;
                }

                if (string.Equals(type, "separate", StringComparison.Ordinal))
                {
                    unsupportedReason = "complex field has duplicate separate markers";
                    return false;
                }

                if (string.Equals(type, "end", StringComparison.Ordinal))
                {
                    resultRuns = results.ToArray();
                    return true;
                }
            }

            if (ContainsProtectedBookmarkReplacementNode([sibling], out string? protectedFeature))
            {
                unsupportedReason = $"complex field result contains protected OOXML boundary '{protectedFeature}'";
                return false;
            }

            results.Add(sibling);
        }

        unsupportedReason = foundSeparate
            ? "matching complex field end marker was not found in the same paragraph"
            : "complex field separate marker was not found in the same paragraph";
        return false;
    }

    private static void ReplaceComplexFieldResult(XElement separateRun, IReadOnlyList<XElement> resultRuns, string text)
    {
        XElement? firstRunProperties = resultRuns
            .Elements(OoxmlNs.W + "rPr")
            .FirstOrDefault();
        foreach (XElement run in resultRuns)
        {
            run.Remove();
        }

        var replacementRun = new XElement(OoxmlNs.W + "r");
        if (firstRunProperties is not null)
        {
            replacementRun.Add(new XElement(firstRunProperties));
        }

        foreach (XNode node in CreateTextNodes(text))
        {
            replacementRun.Add(node);
        }

        separateRun.AddAfterSelf(replacementRun);
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

        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        HyperlinkTarget? hyperlinkTarget = ResolveHyperlinkTarget(package, target, cancellationToken);
        if (hyperlinkTarget is null && !IsSupportedHyperlinkTargetShape(target))
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
        else if (anchor is not null)
        {
            SetInternalHyperlinkAnchor(package, hyperlinkTarget, anchor, cancellationToken);
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
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        if (text is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        HyperlinkTarget? hyperlinkTarget = ResolveHyperlinkTarget(package, target, cancellationToken);
        if (hyperlinkTarget is null && !IsSupportedHyperlinkTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported hyperlink target '{target}'. Expected a hyperlink ID such as M.L0001 or H001.L0001.", operation, target)];
        }

        if (hyperlinkTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        string current = ReadVisibleText(hyperlinkTarget.Hyperlink);
        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges)
        {
            if (TryGetProtectedTextEditFeature(hyperlinkTarget.Hyperlink, out string protectedFeature))
            {
                if (!TrackUnsupportedShape(options, operation, target, $"hyperlink contains protected OOXML boundary '{protectedFeature}'", diagnostics))
                {
                    return diagnostics;
                }

                useTrackedChanges = false;
            }
            else if (!TryValidateTrackedWholeParagraphReplacement(hyperlinkTarget.Hyperlink, current, text, style: null, out string? trackedUnsupportedReason))
            {
                if (!TrackUnsupportedShape(options, operation, target, trackedUnsupportedReason, diagnostics))
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
            ReplaceWholeParagraphTextWithTrackedChanges(package, hyperlinkTarget.Hyperlink, current, text, options, generatedRevisionIds, cancellationToken);
            SaveDocumentPart(package, hyperlinkTarget.PartName, hyperlinkTarget.Document);
            return diagnostics;
        }

        ReplaceHyperlinkText(hyperlinkTarget.Hyperlink, text);
        SaveDocumentPart(package, hyperlinkTarget.PartName, hyperlinkTarget.Document);
        return diagnostics;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteInsertHyperlinkAfter(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
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

        if (text is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        BlockTarget? blockTarget = ResolveBlockTarget(package, operation, target, cancellationToken, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
        if (selectorDiagnostics.Count != 0)
        {
            return selectorDiagnostics;
        }

        if (blockTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges && TextContainsTrackedUnsupportedCharacters(text))
        {
            if (!TrackUnsupportedShape(options, operation, target, "inserted hyperlink text contains tabs or line breaks", diagnostics))
            {
                return diagnostics;
            }

            useTrackedChanges = false;
        }

        if (!apply)
        {
            return diagnostics;
        }

        string? relationshipId = null;
        if (uri is not null)
        {
            relationshipId = OoxmlIds.AllocateRelationshipId(package.GetRelationships(blockTarget.PartName, cancellationToken).Select(relationship => relationship.Id));
            package.AddRelationship(blockTarget.PartName, relationshipId, OoxmlRelTypes.Hyperlink, uri, "External", cancellationToken);
        }

        XElement paragraph = CreateHyperlinkParagraph(
            useTrackedChanges ? string.Empty : text,
            relationshipId,
            anchor,
            operation.Fields.GetValueOrDefault("tooltip"),
            operation.Fields.GetValueOrDefault("target-frame"),
            history);
        if (useTrackedChanges)
        {
            XElement hyperlink = paragraph.Element(OoxmlNs.W + "hyperlink")
                ?? throw new InvalidDataException("Hyperlink paragraph did not contain a hyperlink element.");
            hyperlink.RemoveNodes();
            ReplaceWholeParagraphTextWithTrackedChanges(package, hyperlink, string.Empty, text, options, generatedRevisionIds, cancellationToken);
        }

        blockTarget.Block.AddAfterSelf(paragraph);
        SaveDocumentPart(package, blockTarget.PartName, blockTarget.Document);
        return diagnostics;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteRemoveHyperlink(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        HyperlinkTarget? hyperlinkTarget = ResolveHyperlinkTarget(package, target, cancellationToken);
        if (hyperlinkTarget is null && !IsSupportedHyperlinkTargetShape(target))
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
            !UsesHyperlinkRelationship(hyperlinkTarget.Document, relationshipId))
        {
            package.RemoveRelationship(hyperlinkTarget.PartName, relationshipId, cancellationToken);
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
                $"Operation '{operation.OperationName}' is not supported because threaded comment reply edits require comment-thread metadata that DocxEdit does not safely modify yet.",
                operation,
                target)
        ];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteUnsupportedRepeatingSectionOperation(DocxPatchOperation operation)
    {
        string? target = operation.Fields.GetValueOrDefault("target");
        return
        [
            Diagnostic(
                DocxSeverity.Error,
                "E4315",
                $"Operation '{operation.OperationName}' is not supported because repeating-section item edits require cloning or deleting structured document tag subtrees while preserving IDs, bindings, and section boundaries.",
                operation,
                target)
        ];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteUnsupportedColumnOperation(DocxPatchOperation operation)
    {
        string? target = operation.Fields.GetValueOrDefault("target");
        return
        [
            Diagnostic(
                DocxSeverity.Error,
                "E4316",
                $"Operation '{operation.OperationName}' is not supported because column edits require rebuilding table grids, horizontal spans, omitted cells, and vertical merge state.",
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
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", uriDiagnostic, operation, target));
            return false;
        }

        if (anchor is not null && (anchor.Length == 0 || anchor.Any(char.IsWhiteSpace)))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Field 'anchor' must be a non-empty bookmark anchor without whitespace.", operation, target));
            return false;
        }

        return true;
    }

    private static bool TryValidateExternalHyperlinkUri(string uri, [NotNullWhen(false)] out string? diagnostic)
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
            CountHyperlinkRelationshipUses(hyperlinkTarget.Document, oldRelationshipId) == 1;
        string relationshipId = canReuseRelationship && oldRelationshipId is not null
            ? oldRelationshipId
            : OoxmlIds.AllocateRelationshipId(package.GetRelationships(hyperlinkTarget.PartName, cancellationToken).Select(relationship => relationship.Id));
        if (canReuseRelationship)
        {
            package.RemoveRelationship(hyperlinkTarget.PartName, relationshipId, cancellationToken);
        }

        package.AddRelationship(hyperlinkTarget.PartName, relationshipId, OoxmlRelTypes.Hyperlink, uri, "External", cancellationToken);
        hyperlinkTarget.Hyperlink.SetAttributeValue(OoxmlNs.R + "id", relationshipId);
        hyperlinkTarget.Hyperlink.SetAttributeValue(OoxmlNs.W + "anchor", null);
    }

    private static void SetInternalHyperlinkAnchor(
        OoxmlPackage package,
        HyperlinkTarget hyperlinkTarget,
        string anchor,
        CancellationToken cancellationToken)
    {
        string? oldRelationshipId = (string?)hyperlinkTarget.Hyperlink.Attribute(OoxmlNs.R + "id");
        hyperlinkTarget.Hyperlink.SetAttributeValue(OoxmlNs.R + "id", null);
        hyperlinkTarget.Hyperlink.SetAttributeValue(OoxmlNs.W + "anchor", anchor);
        if (!string.IsNullOrWhiteSpace(oldRelationshipId) &&
            !UsesHyperlinkRelationship(hyperlinkTarget.Document, oldRelationshipId))
        {
            package.RemoveRelationship(hyperlinkTarget.PartName, oldRelationshipId, cancellationToken);
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

    private static string ReadContentControlKind(XElement contentControl)
    {
        XElement? properties = contentControl.Element(OoxmlNs.W + "sdtPr");
        if (properties is null)
        {
            return "rich-text";
        }

        string? kind = properties.Elements()
            .Select(element => element.Name.LocalName)
            .FirstOrDefault(name => name is "text" or "richText" or "checkBox" or "dropDownList" or "comboBox" or "date" or "picture" or "group" or "repeatingSection" or "repeatingSectionItem");
        return kind switch
        {
            "text" => "plain-text",
            "richText" => "rich-text",
            "checkBox" => "checkbox",
            "dropDownList" => "dropdown-list",
            "comboBox" => "combo-box",
            "date" => "date",
            "picture" => "picture",
            "group" => "group",
            "repeatingSection" => "repeating-section",
            "repeatingSectionItem" => "repeating-section-item",
            _ => "rich-text"
        };
    }

    private static DocxDiagnostic UnsupportedContentControlKindDiagnostic(
        XElement contentControl,
        DocxPatchOperation operation,
        string target,
        string expected)
    {
        string kind = ReadContentControlKind(contentControl);
        string guidance = kind switch
        {
            "picture" => "Picture content controls preserve a picture container; use read/media to inspect the contained image and target image operations when applicable.",
            "group" => "Group content controls protect a container; target an editable child content control instead.",
            "repeating-section" or "repeating-section-item" => "Repeating-section subtree edits require cloning or deleting structured document tag subtrees, which DocxEdit currently rejects with E4315.",
            "checkbox" => "Use set-content-control-checkbox for checkbox state edits.",
            "dropdown-list" or "combo-box" => "Use set-content-control-choice for dropdown or combo-box selections.",
            "date" => "Use set-content-control-date for date values.",
            _ => "Choose an operation that matches the content-control kind."
        };
        return Diagnostic(
            DocxSeverity.Error,
            "E4310",
            $"Content control '{target}' is kind '{kind}', not {expected}. {guidance}",
            operation,
            target);
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
            ? symbol
            : char.ConvertFromUtf32(checkedValue ? 0x2612 : 0x2610);
    }

    private static bool TryDecodeStateSymbol(string? value, [NotNullWhen(true)] out string? symbol)
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
            content.Add(CreateSimpleParagraph(text, style: null, paragraphProperties: null));
        }
        else
        {
            content.Add(CreateSimpleRun(text));
        }
    }

    private static XNode[] GetMultiParagraphBookmarkReplacementNodes(
        XElement startParagraph,
        XElement start,
        XElement endParagraph,
        XElement end,
        out XNode[] intermediateNodes,
        out string? failure)
    {
        intermediateNodes = [];
        failure = null;
        if (startParagraph.Parent is null || startParagraph.Parent != endParagraph.Parent)
        {
            failure = "markers must be in sibling paragraphs under the same container";
            return [];
        }

        var selected = new List<XNode>();
        selected.AddRange(start.NodesAfterSelf());

        var between = new List<XNode>();
        bool foundEndParagraph = false;
        foreach (XNode node in startParagraph.NodesAfterSelf())
        {
            if (node == endParagraph)
            {
                foundEndParagraph = true;
                break;
            }

            between.Add(node);
        }

        if (!foundEndParagraph)
        {
            failure = "end marker paragraph must follow the start marker paragraph";
            return [];
        }

        XElement? unsupportedBlock = between
            .OfType<XElement>()
            .FirstOrDefault(element => element.Name != OoxmlNs.W + "p");
        if (unsupportedBlock is not null)
        {
            failure = $"range crosses unsupported block '{unsupportedBlock.Name.LocalName}'";
            return [];
        }

        intermediateNodes = between.ToArray();
        selected.AddRange(intermediateNodes);
        selected.AddRange(endParagraph.Nodes().TakeWhile(node => node != end));
        return selected.ToArray();
    }

    private static bool TryBuildTableSpanningBookmarkTextSlots(
        XElement startParagraph,
        XElement start,
        XElement endParagraph,
        XElement end,
        out BookmarkTextSlot[] slots,
        out XNode[] selectedNodes,
        out string? failure)
    {
        slots = [];
        selectedNodes = [];
        failure = null;
        if (startParagraph.Parent is null || startParagraph.Parent != endParagraph.Parent)
        {
            failure = "markers must be in sibling paragraphs under the same container";
            return false;
        }

        var between = new List<XNode>();
        bool foundEndParagraph = false;
        foreach (XNode node in startParagraph.NodesAfterSelf())
        {
            if (node == endParagraph)
            {
                foundEndParagraph = true;
                break;
            }

            between.Add(node);
        }

        if (!foundEndParagraph)
        {
            failure = "end marker paragraph must follow the start marker paragraph";
            return false;
        }

        if (!between.OfType<XElement>().Any(element => element.Name == OoxmlNs.W + "tbl"))
        {
            failure = "range does not cross a table block";
            return false;
        }

        XElement? unsupportedBlock = between
            .OfType<XElement>()
            .FirstOrDefault(element => element.Name != OoxmlNs.W + "p" && element.Name != OoxmlNs.W + "tbl");
        if (unsupportedBlock is not null)
        {
            failure = $"range crosses unsupported block '{unsupportedBlock.Name.LocalName}'";
            return false;
        }

        var slotBuilder = new List<BookmarkTextSlot>();
        var selectedBuilder = new List<XNode>();
        AddBookmarkTextSlotIfVisible(slotBuilder, selectedBuilder, startParagraph, start.NodesAfterSelf().ToArray(), AddAfter: start, AddBefore: null);
        foreach (XElement block in between.OfType<XElement>())
        {
            if (block.Name == OoxmlNs.W + "p")
            {
                AddWholeParagraphBookmarkTextSlot(slotBuilder, selectedBuilder, block);
                continue;
            }

            if (block.Descendants(OoxmlNs.W + "tbl").Any())
            {
                failure = "table-spanning bookmark range contains a nested table";
                return false;
            }

            foreach (XElement cell in block.Descendants(OoxmlNs.W + "tc"))
            {
                XElement? unsupportedCellBlock = cell
                    .Elements()
                    .FirstOrDefault(element => element.Name != OoxmlNs.W + "tcPr" && element.Name != OoxmlNs.W + "p");
                if (unsupportedCellBlock is not null)
                {
                    failure = $"table cell contains unsupported block '{unsupportedCellBlock.Name.LocalName}'";
                    return false;
                }

                foreach (XElement paragraph in cell.Elements(OoxmlNs.W + "p"))
                {
                    AddWholeParagraphBookmarkTextSlot(slotBuilder, selectedBuilder, paragraph);
                }
            }
        }

        AddBookmarkTextSlotIfVisible(slotBuilder, selectedBuilder, endParagraph, endParagraph.Nodes().TakeWhile(node => node != end).ToArray(), AddAfter: null, AddBefore: end);
        if (slotBuilder.Count == 0)
        {
            failure = "range contains no visible text slots";
            return false;
        }

        slots = slotBuilder.ToArray();
        selectedNodes = selectedBuilder.ToArray();
        return true;
    }

    private static void AddWholeParagraphBookmarkTextSlot(
        List<BookmarkTextSlot> slots,
        List<XNode> selectedNodes,
        XElement paragraph)
    {
        XNode[] nodes = paragraph.Nodes()
            .Where(node => node is not XElement element || element.Name != OoxmlNs.W + "pPr")
            .ToArray();
        XElement? paragraphProperties = paragraph.Element(OoxmlNs.W + "pPr");
        AddBookmarkTextSlotIfVisible(slots, selectedNodes, paragraph, nodes, paragraphProperties, AddBefore: null);
    }

    private static void AddBookmarkTextSlotIfVisible(
        List<BookmarkTextSlot> slots,
        List<XNode> selectedNodes,
        XElement container,
        XNode[] nodes,
        XElement? AddAfter,
        XElement? AddBefore)
    {
        selectedNodes.AddRange(nodes);
        if (ReadVisibleText(new XElement(OoxmlNs.W + "p", nodes)).Length == 0)
        {
            return;
        }

        slots.Add(new BookmarkTextSlot(container, nodes, AddAfter, AddBefore));
    }

    private static void ReplaceStructuredBookmarkTextSlots(IReadOnlyList<BookmarkTextSlot> slots, IReadOnlyList<string> lines)
    {
        for (int i = 0; i < slots.Count; i++)
        {
            BookmarkTextSlot slot = slots[i];
            foreach (XNode node in slot.Nodes)
            {
                node.Remove();
            }

            XElement replacementRun = CreateSimpleRun(lines[i]);
            if (slot.AddAfter is not null)
            {
                slot.AddAfter.AddAfterSelf(replacementRun);
            }
            else if (slot.AddBefore is not null)
            {
                slot.AddBefore.AddBeforeSelf(replacementRun);
            }
            else if (slot.Container.Element(OoxmlNs.W + "pPr") is { } paragraphProperties)
            {
                paragraphProperties.AddAfterSelf(replacementRun);
            }
            else
            {
                slot.Container.AddFirst(replacementRun);
            }
        }
    }

    private static void ReplaceMultiParagraphBookmarkText(
        XElement startParagraph,
        XElement start,
        XElement endParagraph,
        XElement end,
        XNode[] intermediateNodes,
        string text)
    {
        string[] paragraphs = SplitReplacementParagraphText(text);

        foreach (XNode node in start.NodesAfterSelf().ToArray())
        {
            node.Remove();
        }

        start.AddAfterSelf(CreateSimpleRun(paragraphs[0]));

        foreach (XNode node in intermediateNodes)
        {
            node.Remove();
        }

        if (paragraphs.Length == 1)
        {
            XNode[] trailingNodes = end.NodesAfterSelf().ToArray();
            end.Remove();
            startParagraph.Add(end);
            foreach (XNode node in trailingNodes)
            {
                startParagraph.Add(node);
            }

            endParagraph.Remove();
            return;
        }

        for (int i = 1; i < paragraphs.Length - 1; i++)
        {
            endParagraph.AddBeforeSelf(CreateSimpleParagraph(paragraphs[i], style: null, paragraphProperties: null));
        }

        foreach (XNode node in endParagraph.Nodes().TakeWhile(node => node != end).ToArray())
        {
            if (node is XElement element && element.Name == OoxmlNs.W + "pPr")
            {
                continue;
            }

            node.Remove();
        }

        end.AddBeforeSelf(CreateSimpleRun(paragraphs[^1]));
    }

    private static string[] SplitReplacementParagraphText(string text)
    {
        return text
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Split('\n');
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

    private static bool TryBuildTrackedBookmarkReplacement(
        IReadOnlyList<XNode> nodes,
        string replacement,
        [NotNullWhen(true)] out TrackedBookmarkReplacement? trackedReplacement,
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        trackedReplacement = null;
        unsupportedReason = null;

        var clonedRuns = new List<XElement>();
        foreach (XNode node in nodes)
        {
            if (node is XText text && string.IsNullOrWhiteSpace(text.Value))
            {
                continue;
            }

            if (node is not XElement element || element.Name != OoxmlNs.W + "r")
            {
                unsupportedReason = "bookmark range contains non-run content";
                return false;
            }

            clonedRuns.Add(new XElement(element));
        }

        var rangeContainer = new XElement(OoxmlNs.W + "p", clonedRuns);
        string current = ReadVisibleText(rangeContainer);
        if (!TryValidateTrackedWholeParagraphReplacement(rangeContainer, current, replacement, style: null, out unsupportedReason))
        {
            return false;
        }

        XElement? firstRunProperties = rangeContainer
            .Elements(OoxmlNs.W + "r")
            .Elements(OoxmlNs.W + "rPr")
            .FirstOrDefault();
        trackedReplacement = new TrackedBookmarkReplacement(
            current,
            firstRunProperties is null ? null : new XElement(firstRunProperties));
        return true;
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

    private static bool BookmarkNameExists(XDocument document, string name)
    {
        return document
            .Descendants(OoxmlNs.W + "bookmarkStart")
            .Any(bookmark => string.Equals((string?)bookmark.Attribute(OoxmlNs.W + "name"), name, StringComparison.Ordinal));
    }

    private static string AllocateBookmarkId(XDocument document)
    {
        int maxId = document
            .Descendants()
            .Where(element => element.Name == OoxmlNs.W + "bookmarkStart" || element.Name == OoxmlNs.W + "bookmarkEnd")
            .Select(element => (string?)element.Attribute(OoxmlNs.W + "id"))
            .Select(value => int.TryParse(value, out int id) ? id : 0)
            .DefaultIfEmpty(0)
            .Max();
        return (maxId + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
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

    private static CommentReplyTarget? ResolveCommentReplyTarget(
        OoxmlPackage package,
        string target,
        DocxPatchOperation operation,
        CancellationToken cancellationToken,
        out DocxDiagnostic? diagnostic)
    {
        diagnostic = null;
        if (TryParseCommentReplyTarget(target, out string? parentCommentId, out int replyOrdinal))
        {
            CommentTarget? parentTarget = FindCommentTargetById(package, parentCommentId, cancellationToken);
            if (parentTarget is null)
            {
                return null;
            }

            string? parentReplyLookupParaId = ReadCommentParaId(parentTarget.Comment);
            if (string.IsNullOrWhiteSpace(parentReplyLookupParaId))
            {
                diagnostic = Diagnostic(DocxSeverity.Error, "E4314", $"Comment '{target}' parent has no w15:paraId for threaded reply lookup.", operation, target);
                return null;
            }

            return FindCommentReplyTargetByOrdinal(package, parentTarget.PartName, parentReplyLookupParaId, replyOrdinal, cancellationToken);
        }

        CommentTarget? commentTarget = ResolveCommentTarget(package, target, operation, cancellationToken, out diagnostic);
        if (diagnostic is not null || commentTarget is null)
        {
            return null;
        }

        string? paraId = ReadCommentParaId(commentTarget.Comment);
        if (string.IsNullOrWhiteSpace(paraId))
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E4314", $"Comment target '{target}' is not a threaded reply because it has no w15:paraId.", operation, target);
            return null;
        }

        string? parentParaId = ReadCommentParentParaId(package, paraId, cancellationToken);
        if (string.IsNullOrWhiteSpace(parentParaId))
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E4314", $"Comment target '{target}' is not a threaded reply.", operation, target);
            return null;
        }

        return new CommentReplyTarget(commentTarget.PartName, commentTarget.Document, commentTarget.Comment, paraId, parentParaId);
    }

    private static bool TryParseCommentReplyTarget(string target, out string parentCommentId, out int replyOrdinal)
    {
        parentCommentId = string.Empty;
        replyOrdinal = 0;
        if (!target.StartsWith("comment:", StringComparison.Ordinal))
        {
            return false;
        }

        int markerIndex = target.IndexOf(".reply:", StringComparison.Ordinal);
        if (markerIndex < 0)
        {
            return false;
        }

        parentCommentId = target["comment:".Length..markerIndex];
        string ordinalText = target[(markerIndex + ".reply:".Length)..];
        return parentCommentId.Length != 0 &&
            int.TryParse(ordinalText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out replyOrdinal) &&
            replyOrdinal > 0;
    }

    private static CommentReplyTarget? FindCommentReplyTargetByOrdinal(
        OoxmlPackage package,
        string partName,
        string parentParaId,
        int replyOrdinal,
        CancellationToken cancellationToken)
    {
        XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
        XElement? reply = document
            .Descendants(OoxmlNs.W + "comment")
            .Where(comment =>
            {
                string? paraId = ReadCommentParaId(comment);
                return paraId is not null &&
                    string.Equals(ReadCommentParentParaId(package, paraId, cancellationToken), parentParaId, StringComparison.Ordinal);
            })
            .ElementAtOrDefault(replyOrdinal - 1);
        if (reply is null)
        {
            return null;
        }

        string replyParaId = ReadCommentParaId(reply)
            ?? throw new InvalidDataException("Reply comment target has no paraId.");
        return new CommentReplyTarget(partName, document, reply, replyParaId, parentParaId);
    }

    private static string? ReadCommentParentParaId(
        OoxmlPackage package,
        string paraId,
        CancellationToken cancellationToken)
    {
        foreach (string partName in GetCommentsExtendedPartNames(package, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            XElement? extension = document
                .Descendants(OoxmlNs.W15 + "commentEx")
                .FirstOrDefault(element => string.Equals((string?)element.Attribute(OoxmlNs.W15 + "paraId"), paraId, StringComparison.Ordinal));
            if (extension is not null)
            {
                return (string?)extension.Attribute(OoxmlNs.W15 + "paraIdParent");
            }
        }

        return null;
    }

    private static bool HasChildCommentReplies(
        OoxmlPackage package,
        string paraId,
        CancellationToken cancellationToken)
    {
        foreach (string partName in GetCommentsExtendedPartNames(package, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            if (document
                .Descendants(OoxmlNs.W15 + "commentEx")
                .Any(element => string.Equals((string?)element.Attribute(OoxmlNs.W15 + "paraIdParent"), paraId, StringComparison.Ordinal)))
            {
                return true;
            }
        }

        return false;
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

        return package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => relationship.Type == OoxmlRelTypes.Comments)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .Select(relationship => relationship.ResolvedTarget)
            .ToArray();
    }

    private static IReadOnlyList<string> GetCommentsExtendedPartNames(OoxmlPackage package, CancellationToken cancellationToken)
    {

        var partNames = package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => relationship.Type == OoxmlRelTypes.CommentsExtended)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .Select(relationship => relationship.ResolvedTarget)
            .ToList();
        if (package.GetPart("/word/commentsExtended.xml") is not null &&
            !partNames.Contains("/word/commentsExtended.xml", StringComparer.Ordinal))
        {
            partNames.Add("/word/commentsExtended.xml");
        }

        return partNames;
    }

    private static IReadOnlyList<string> GetCommentsIdsPartNames(OoxmlPackage package, CancellationToken cancellationToken)
    {

        var partNames = package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => relationship.Type == OoxmlRelTypes.CommentsIds)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .Select(relationship => relationship.ResolvedTarget)
            .ToList();
        if (package.GetPart("/word/commentsIds.xml") is not null &&
            !partNames.Contains("/word/commentsIds.xml", StringComparer.Ordinal))
        {
            partNames.Add("/word/commentsIds.xml");
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

    private static DocxDiagnostic? ValidateExistingCommentsIdsPart(OoxmlPackage package, CancellationToken cancellationToken)
    {
        foreach (string partName in GetCommentsIdsPartNames(package, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out XElement root);
            if (root.Name != OoxmlNs.W16Cid + "commentsIds")
            {
                return new DocxDiagnostic(DocxSeverity.Error, "E9001", $"commentsIds part '{partName}' has root '{root.Name.LocalName}', expected 'commentsIds'.", PartName: partName);
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
                package.AddPart(commentsPartName, CommentsContentType, bytes, cancellationToken);
            }

            string relationshipId = OoxmlIds.AllocateRelationshipId(package
                .GetRelationships(package.MainDocumentPartName, cancellationToken)
                .Select(relationship => relationship.Id));
            package.AddRelationship(package.MainDocumentPartName, relationshipId, OoxmlRelTypes.Comments, GetRelativeRelationshipTarget(package.MainDocumentPartName, commentsPartName), targetMode: null, cancellationToken);
        }
        else if (package.GetPart(commentsPartName) is null)
        {
            byte[] bytes = Encoding.UTF8.GetBytes("""
                <?xml version="1.0" encoding="utf-8"?>
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />
                """);
            package.AddPart(commentsPartName, CommentsContentType, bytes, cancellationToken);
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
                package.AddPart(commentsExtendedPartName, CommentsExtendedContentType, bytes, cancellationToken);
            }

            string relationshipId = OoxmlIds.AllocateRelationshipId(package
                .GetRelationships(package.MainDocumentPartName, cancellationToken)
                .Select(relationship => relationship.Id));
            package.AddRelationship(package.MainDocumentPartName, relationshipId, OoxmlRelTypes.CommentsExtended, GetRelativeRelationshipTarget(package.MainDocumentPartName, commentsExtendedPartName), targetMode: null, cancellationToken);
        }
        else if (package.GetPart(commentsExtendedPartName) is null)
        {
            byte[] bytes = Encoding.UTF8.GetBytes("""
                <?xml version="1.0" encoding="utf-8"?>
                <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" />
                """);
            package.AddPart(commentsExtendedPartName, CommentsExtendedContentType, bytes, cancellationToken);
        }

        XDocument document = LoadDocumentPart(package, commentsExtendedPartName, cancellationToken, out XElement root);
        if (root.Name != OoxmlNs.W15 + "commentsEx")
        {
            throw new InvalidDataException($"commentsExtended part '{commentsExtendedPartName}' has root '{root.Name.LocalName}', expected 'commentsEx'.");
        }

        EnsureNamespaceDeclaration(root, "w15", OoxmlNs.W15);
        return new CommentsExtendedPartTarget(commentsExtendedPartName, document, root);
    }

    private static CommentsIdsPartTarget ResolveOrCreateCommentsIdsPart(OoxmlPackage package, CancellationToken cancellationToken)
    {
        string? commentsIdsPartName = GetCommentsIdsPartNames(package, cancellationToken).FirstOrDefault();
        if (commentsIdsPartName is null)
        {
            commentsIdsPartName = "/word/commentsIds.xml";
            if (package.GetPart(commentsIdsPartName) is null)
            {
                byte[] bytes = Encoding.UTF8.GetBytes("""
                    <?xml version="1.0" encoding="utf-8"?>
                    <w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" />
                    """);
                package.AddPart(commentsIdsPartName, CommentsIdsContentType, bytes, cancellationToken);
            }

            string relationshipId = OoxmlIds.AllocateRelationshipId(package
                .GetRelationships(package.MainDocumentPartName, cancellationToken)
                .Select(relationship => relationship.Id));
            package.AddRelationship(package.MainDocumentPartName, relationshipId, OoxmlRelTypes.CommentsIds, GetRelativeRelationshipTarget(package.MainDocumentPartName, commentsIdsPartName), targetMode: null, cancellationToken);
        }
        else if (package.GetPart(commentsIdsPartName) is null)
        {
            byte[] bytes = Encoding.UTF8.GetBytes("""
                <?xml version="1.0" encoding="utf-8"?>
                <w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" />
                """);
            package.AddPart(commentsIdsPartName, CommentsIdsContentType, bytes, cancellationToken);
        }

        XDocument document = LoadDocumentPart(package, commentsIdsPartName, cancellationToken, out XElement root);
        if (root.Name != OoxmlNs.W16Cid + "commentsIds")
        {
            throw new InvalidDataException($"commentsIds part '{commentsIdsPartName}' has root '{root.Name.LocalName}', expected 'commentsIds'.");
        }

        EnsureNamespaceDeclaration(root, "w16cid", OoxmlNs.W16Cid);
        return new CommentsIdsPartTarget(commentsIdsPartName, document, root);
    }

    private static string AllocateCommentDurableId(OoxmlPackage package, CancellationToken cancellationToken)
    {
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string partName in GetCommentsIdsPartNames(package, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            foreach (string durableId in document
                .Descendants(OoxmlNs.W16Cid + "commentId")
                .Select(element => (string?)element.Attribute(OoxmlNs.W16Cid + "durableId"))
                .Where(durableId => !string.IsNullOrWhiteSpace(durableId))
                .OfType<string>())
            {
                used.Add(durableId);
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

        throw new InvalidDataException("Unable to allocate a unique comment durableId.");
    }

    private static string AllocateCommentParaId(OoxmlPackage package, CancellationToken cancellationToken)
    {
        return AllocateCommentParaId(package, [], cancellationToken);
    }

    private static string AllocateCommentParaId(
        OoxmlPackage package,
        IEnumerable<string> reservedParaIds,
        CancellationToken cancellationToken)
    {
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string reservedParaId in reservedParaIds)
        {
            if (!string.IsNullOrWhiteSpace(reservedParaId))
            {
                used.Add(reservedParaId);
            }
        }

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
                .OfType<string>())
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
                    .OfType<string>())
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
        comment.Add(CreateSimpleParagraph(text, style: null, paragraphProperties: null));
    }

    private static XElement CreateComment(
        string commentId,
        string text,
        string author,
        string? initials,
        DateTimeOffset timestampUtc)
    {
        return CreateComment(commentId, text, author, initials, timestampUtc, paraId: null);
    }

    private static XElement CreateComment(
        string commentId,
        string text,
        string author,
        string? initials,
        DateTimeOffset timestampUtc,
        string? paraId)
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

        XElement paragraph = CreateSimpleParagraph(text, style: null, paragraphProperties: null);
        if (!string.IsNullOrWhiteSpace(paraId))
        {
            paragraph.SetAttributeValue(OoxmlNs.W15 + "paraId", paraId);
        }

        comment.Add(paragraph);
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

    private static bool CanAddSelectedCommentAnchor(XElement paragraph, TextRange range, out string? unsupportedReason)
    {
        unsupportedReason = null;
        TextPosition?[] positions = BuildTextPositions(paragraph);
        if (positions.Length != ReadVisibleText(paragraph).Length)
        {
            unsupportedReason = "selected ranges are limited to direct paragraph runs";
            return false;
        }

        if (range.Length <= 0 || range.Start < 0 || range.Start + range.Length > positions.Length)
        {
            unsupportedReason = "selected range is outside paragraph text";
            return false;
        }

        for (int i = range.Start; i < range.Start + range.Length; i++)
        {
            if (positions[i] is null)
            {
                unsupportedReason = "selected range includes tabs, line breaks, or non-text run content";
                return false;
            }
        }

        return true;
    }

    private static void AddSelectedCommentAnchor(XElement paragraph, string commentId, TextRange range)
    {
        InsertCommentBoundaryAtTextOffset(
            paragraph,
            range.Start + range.Length,
            new XElement(OoxmlNs.W + "commentRangeEnd", new XAttribute(OoxmlNs.W + "id", commentId)));
        InsertCommentBoundaryAtTextOffset(
            paragraph,
            range.Start,
            new XElement(OoxmlNs.W + "commentRangeStart", new XAttribute(OoxmlNs.W + "id", commentId)));
        paragraph.Add(new XElement(
            OoxmlNs.W + "r",
            new XElement(OoxmlNs.W + "commentReference", new XAttribute(OoxmlNs.W + "id", commentId))));
    }

    private static void InsertCommentBoundaryAtTextOffset(XElement paragraph, int textOffset, XElement boundary)
    {
        TextPosition?[] positions = BuildTextPositions(paragraph);
        if (textOffset == 0)
        {
            XElement? paragraphProperties = paragraph.Element(OoxmlNs.W + "pPr");
            if (paragraphProperties is null)
            {
                paragraph.AddFirst(boundary);
            }
            else
            {
                paragraphProperties.AddAfterSelf(boundary);
            }

            return;
        }

        if (textOffset == positions.Length)
        {
            paragraph.Add(boundary);
            return;
        }

        TextPosition position = positions[textOffset]
            ?? throw new InvalidDataException("Selected comment range boundary cannot be placed on non-text run content.");
        XElement afterRun = SplitRunAtTextPosition(position);
        afterRun.AddBeforeSelf(boundary);
    }

    private static XElement SplitRunAtTextPosition(TextPosition position)
    {
        XElement textElement = position.TextElement;
        XElement run = textElement.Parent
            ?? throw new InvalidDataException("Text element has no parent run.");
        if (run.Name != OoxmlNs.W + "r" || run.Parent?.Name != OoxmlNs.W + "p")
        {
            throw new InvalidDataException("Selected comment range boundary is not inside a direct paragraph run.");
        }

        var beforeRun = new XElement(OoxmlNs.W + "r");
        var afterRun = new XElement(OoxmlNs.W + "r");
        XElement? runProperties = run.Element(OoxmlNs.W + "rPr");
        if (runProperties is not null)
        {
            beforeRun.Add(new XElement(runProperties));
            afterRun.Add(new XElement(runProperties));
        }

        bool reachedTextElement = false;
        foreach (XElement child in run.Elements().Where(child => child.Name != OoxmlNs.W + "rPr"))
        {
            if (!ReferenceEquals(child, textElement))
            {
                (reachedTextElement ? afterRun : beforeRun).Add(new XElement(child));
                continue;
            }

            reachedTextElement = true;
            string value = textElement.Value;
            string beforeText = value[..position.Offset];
            string afterText = value[position.Offset..];
            if (beforeText.Length != 0)
            {
                XElement beforeTextElement = new XElement(textElement);
                SetTextElementValue(beforeTextElement, beforeText);
                beforeRun.Add(beforeTextElement);
            }

            if (afterText.Length != 0)
            {
                XElement afterTextElement = new XElement(textElement);
                SetTextElementValue(afterTextElement, afterText);
                afterRun.Add(afterTextElement);
            }
        }

        bool hasBeforeContent = beforeRun.Elements().Any(element => element.Name != OoxmlNs.W + "rPr");
        bool hasAfterContent = afterRun.Elements().Any(element => element.Name != OoxmlNs.W + "rPr");
        if (hasBeforeContent)
        {
            run.AddBeforeSelf(beforeRun);
        }

        if (hasAfterContent)
        {
            run.AddBeforeSelf(afterRun);
        }

        run.Remove();
        return hasAfterContent
            ? afterRun
            : beforeRun;
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

    private static void RemoveCommentIdRecords(OoxmlPackage package, XElement comment, CancellationToken cancellationToken)
    {
        string? paraId = ReadCommentParaId(comment);
        if (string.IsNullOrWhiteSpace(paraId))
        {
            return;
        }

        foreach (string partName in GetCommentsIdsPartNames(package, cancellationToken))
        {
            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
            XElement[] commentIdRecords = document
                .Descendants(OoxmlNs.W16Cid + "commentId")
                .Where(element => string.Equals((string?)element.Attribute(OoxmlNs.W16Cid + "paraId"), paraId, StringComparison.Ordinal))
                .ToArray();
            if (commentIdRecords.Length == 0)
            {
                continue;
            }

            foreach (XElement commentIdRecord in commentIdRecords)
            {
                commentIdRecord.Remove();
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
        if (asset is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported replace-image target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
        }

        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateImageContentTypeGuard(operation, target, imageTarget.Part.ContentType, diagnostics))
        {
            return diagnostics;
        }

        if (!TryReadAsset(options.AssetProvider, asset, cancellationToken, out byte[] bytes, out string? contentType, out DocxDiagnostic? assetDiagnostic, operation, target))
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
            !TryGetImageDrawingContainer(imageTarget, target, operation, out imageContainer, out DocxDiagnostic? altDiagnostic))
        {
            return [altDiagnostic!];
        }

        if (!apply)
        {
            return [];
        }

        package.ReplacePartBytes(imageTarget.Part.Name, bytes);
        if (hasAlt && imageContainer is not null)
        {
            SetImageAlt(imageContainer, alt, target);
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
        if (asset is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ParagraphTarget? paragraphTarget = ResolveParagraphTarget(package, operation, target, cancellationToken, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
        if (selectorDiagnostics.Count != 0)
        {
            return selectorDiagnostics;
        }

        if (paragraphTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!TryReadAsset(options.AssetProvider, asset, cancellationToken, out byte[] bytes, out string? contentType, out DocxDiagnostic? assetDiagnostic, operation, target))
        {
            return [assetDiagnostic!];
        }

        if (!ValidateImageContentTypeGuard(operation, target, contentType, diagnostics))
        {
            return diagnostics;
        }

        if (!TryReadImageExtent(operation, bytes, contentType, out long widthEmus, out long heightEmus, out DocxDiagnostic? dimensionDiagnostic))
        {
            return [dimensionDiagnostic!];
        }

        if (!apply)
        {
            return [];
        }

        string imagePartName = OoxmlMediaParts.AllocateImagePartName(package.Parts.Keys, contentType);
        string relationshipId = OoxmlIds.AllocateRelationshipId(package.GetRelationships(paragraphTarget.PartName, cancellationToken).Select(relationship => relationship.Id));
        package.AddPart(imagePartName, contentType, bytes, cancellationToken);
        package.AddRelationship(paragraphTarget.PartName, relationshipId, OoxmlRelTypes.Image, GetRelativeRelationshipTarget(paragraphTarget.PartName, imagePartName), targetMode: null, cancellationToken);
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
        if (width is not null && !OoxmlUnits.TryParseDimension(width, out widthEmus))
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E5206", $"Invalid image width '{width}'.", operation, operation.Fields.GetValueOrDefault("target"));
            return false;
        }

        if (height is not null && !OoxmlUnits.TryParseDimension(height, out heightEmus))
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
        [NotNullWhen(false)] out string? unsupportedReason)
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

        if (TryGetUnsupportedTrackedRunContent(paragraph, out string unsupportedContent))
        {
            unsupportedReason = $"paragraph contains unsupported run content '{unsupportedContent}'";
            return false;
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
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        unsupportedReason = null;
        if (TextContainsTrackedUnsupportedCharacters(current) ||
            TextContainsTrackedUnsupportedCharacters(replacement))
        {
            unsupportedReason = "tracked paragraph text contains tabs or line breaks";
            return false;
        }

        if (TryGetUnsupportedTrackedRunContent(paragraph, out string unsupportedContent))
        {
            unsupportedReason = $"paragraph contains unsupported run content '{unsupportedContent}'";
            return false;
        }

        if (HasMixedDirectTextRunProperties(paragraph))
        {
            unsupportedReason = "paragraph contains mixed direct run formatting";
            return false;
        }

        return true;
    }

    private static bool TryGetTrackedSetCellParagraphs(
        XElement cell,
        string replacement,
        out XElement[] paragraphs,
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        paragraphs = cell.Elements(OoxmlNs.W + "p").ToArray();
        unsupportedReason = null;
        if (paragraphs.Length == 0)
        {
            unsupportedReason = "cell has no paragraph for tracked text replacement";
            return false;
        }

        if (cell.Elements().Any(element => element.Name != OoxmlNs.W + "tcPr" && element.Name != OoxmlNs.W + "p"))
        {
            unsupportedReason = "cell contains non-paragraph block content";
            return false;
        }

        if (cell.Descendants(OoxmlNs.W + "drawing").Any())
        {
            unsupportedReason = "cell contains drawing content";
            return false;
        }

        if (cell.Descendants(OoxmlNs.W + "fldSimple").Any() ||
            cell.Descendants(OoxmlNs.W + "fldChar").Any() ||
            cell.Descendants(OoxmlNs.W + "instrText").Any())
        {
            unsupportedReason = "cell contains field content";
            return false;
        }

        for (int i = 0; i < paragraphs.Length; i++)
        {
            XElement paragraph = paragraphs[i];
            if (TryGetProtectedTextEditFeature(paragraph, out string protectedFeature))
            {
                unsupportedReason = $"cell paragraph contains protected OOXML boundary '{protectedFeature}'";
                return false;
            }

            string current = ReadVisibleText(paragraph);
            string insertedText = i == 0 ? replacement : string.Empty;
            if (!TryValidateTrackedWholeParagraphReplacement(paragraph, current, insertedText, style: null, out unsupportedReason))
            {
                return false;
            }
        }

        return true;
    }

    private static bool TryGetTrackedContentControlTextContainer(
        XElement content,
        string replacement,
        [NotNullWhen(true)] out XElement? trackedContainer,
        out string current,
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        trackedContainer = null;
        current = string.Empty;
        unsupportedReason = null;
        XElement[] paragraphs = content.Elements(OoxmlNs.W + "p").ToArray();
        if (paragraphs.Length > 1)
        {
            unsupportedReason = "plain-text content control contains multiple paragraphs";
            return false;
        }

        if (paragraphs.Length == 1)
        {
            if (content.Elements().Any(element => element.Name != OoxmlNs.W + "p"))
            {
                unsupportedReason = "plain-text content control mixes paragraph and non-paragraph content";
                return false;
            }

            trackedContainer = paragraphs[0];
        }
        else
        {
            if (!content.Elements(OoxmlNs.W + "r").Any())
            {
                unsupportedReason = "plain-text content control has no run content";
                return false;
            }

            if (content.Elements().Any(element => element.Name != OoxmlNs.W + "r"))
            {
                unsupportedReason = "plain-text content control contains non-run content";
                return false;
            }

            trackedContainer = content;
        }

        if (TryGetProtectedTextEditFeature(trackedContainer, out string protectedFeature))
        {
            unsupportedReason = $"plain-text content control contains protected OOXML boundary '{protectedFeature}'";
            return false;
        }

        current = ReadVisibleText(trackedContainer);
        if (!TryValidateTrackedWholeParagraphReplacement(trackedContainer, current, replacement, style: null, out unsupportedReason))
        {
            return false;
        }

        return true;
    }

    private static bool TryGetTrackedRichTextContentControlParagraphs(
        XElement content,
        string replacement,
        out XElement[] paragraphs,
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        paragraphs = content.Elements(OoxmlNs.W + "p").ToArray();
        unsupportedReason = null;
        if (paragraphs.Length == 0)
        {
            unsupportedReason = "rich-text content control has no paragraph for tracked text replacement";
            return false;
        }

        if (content.Elements().Any(element => element.Name != OoxmlNs.W + "p"))
        {
            unsupportedReason = "rich-text content control contains non-paragraph content";
            return false;
        }

        if (content.Descendants(OoxmlNs.W + "drawing").Any())
        {
            unsupportedReason = "rich-text content control contains drawing content";
            return false;
        }

        if (content.Descendants(OoxmlNs.W + "fldSimple").Any() ||
            content.Descendants(OoxmlNs.W + "fldChar").Any() ||
            content.Descendants(OoxmlNs.W + "instrText").Any())
        {
            unsupportedReason = "rich-text content control contains field content";
            return false;
        }

        for (int i = 0; i < paragraphs.Length; i++)
        {
            XElement paragraph = paragraphs[i];
            if (TryGetProtectedTextEditFeature(paragraph, out string protectedFeature))
            {
                unsupportedReason = $"rich-text content-control paragraph contains protected OOXML boundary '{protectedFeature}'";
                return false;
            }

            string current = ReadVisibleText(paragraph);
            string insertedText = i == 0 ? replacement : string.Empty;
            if (!TryValidateTrackedWholeParagraphReplacement(paragraph, current, insertedText, style: null, out unsupportedReason))
            {
                return false;
            }
        }

        return true;
    }

    private static bool TryGetTrackedCommentParagraphs(
        XElement comment,
        string replacement,
        out XElement[] paragraphs,
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        paragraphs = comment.Elements(OoxmlNs.W + "p").ToArray();
        unsupportedReason = null;
        if (paragraphs.Length == 0)
        {
            unsupportedReason = "comment body has no paragraph for tracked text replacement";
            return false;
        }

        if (comment.Elements().Any(element => element.Name != OoxmlNs.W + "p"))
        {
            unsupportedReason = "comment body contains non-paragraph content";
            return false;
        }

        if (comment.Descendants(OoxmlNs.W + "drawing").Any())
        {
            unsupportedReason = "comment body contains drawing content";
            return false;
        }

        if (comment.Descendants(OoxmlNs.W + "fldSimple").Any() ||
            comment.Descendants(OoxmlNs.W + "fldChar").Any() ||
            comment.Descendants(OoxmlNs.W + "instrText").Any())
        {
            unsupportedReason = "comment body contains field content";
            return false;
        }

        for (int i = 0; i < paragraphs.Length; i++)
        {
            XElement paragraph = paragraphs[i];
            if (TryGetProtectedTextEditFeature(paragraph, out string protectedFeature))
            {
                unsupportedReason = $"comment body paragraph contains protected OOXML boundary '{protectedFeature}'";
                return false;
            }

            string current = ReadVisibleText(paragraph);
            string insertedText = i == 0 ? replacement : string.Empty;
            if (!TryValidateTrackedWholeParagraphReplacement(paragraph, current, insertedText, style: null, out unsupportedReason))
            {
                return false;
            }
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
            diagnostics.Add(Diagnostic(
                DocxSeverity.Error,
                "E6002",
                BuildUnsupportedTrackedShapeMessage(options.TrackChanges, operation, target, reason),
                operation,
                target,
                "track-changes-unsupported-target-shape",
                "require-failed"));
            return false;
        }

        diagnostics.Add(Diagnostic(
            DocxSeverity.Warning,
            "W4002",
            BuildUnsupportedTrackedShapeMessage(options.TrackChanges, operation, target, reason),
            operation,
            target,
            "track-changes-unsupported-target-shape",
            "direct-edit-preserve-existing-revisions"));
        return true;
    }

    private static string BuildUnsupportedTrackedShapeMessage(
        TrackChangesMode mode,
        DocxPatchOperation operation,
        string target,
        string reason)
    {
        string operationName = operation.OperationName;
        string support = TrackChangesSupportValue(operationName);

        return mode switch
        {
            TrackChangesMode.Require =>
                $"TrackChangesMode.Require cannot apply operation '{operationName}' as tracked output for target '{target}' because catalog support is '{support}' but this target shape is unsupported: {reason}.",
            TrackChangesMode.Suggest =>
                $"TrackChangesMode.Suggest will apply operation '{operationName}' directly for target '{target}' because catalog support is '{support}' but this target shape is unsupported: {reason}. Existing tracked-change markup is preserved, but this edit will not create new revision markup.",
            _ =>
                $"TrackChangesMode.{mode} cannot generate tracked output for operation '{operationName}' on target '{target}' because catalog support is '{support}' but this target shape is unsupported: {reason}."
        };
    }

    private static bool TryGetUnsupportedTrackedRunContent(XElement paragraph, out string unsupportedContent)
    {
        foreach (XElement run in paragraph.Elements(OoxmlNs.W + "r"))
        {
            foreach (XElement child in run.Elements())
            {
                if (child.Name == OoxmlNs.W + "rPr" ||
                    child.Name == OoxmlNs.W + "t" ||
                    child.Name == OoxmlNs.W + "tab" ||
                    child.Name == OoxmlNs.W + "br")
                {
                    continue;
                }

                unsupportedContent = child.Name.LocalName;
                return true;
            }
        }

        unsupportedContent = string.Empty;
        return false;
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

            string signature = CanonicalRunPropertiesSignature(run.Element(OoxmlNs.W + "rPr"));
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

    private static string CanonicalRunPropertiesSignature(XElement? runProperties)
    {
        if (runProperties is null)
        {
            return string.Empty;
        }

        var builder = new StringBuilder();
        AppendCanonicalElement(builder, runProperties);
        return builder.ToString();
    }

    private static void AppendCanonicalElement(StringBuilder builder, XElement element)
    {
        builder.Append('{')
            .Append(element.Name.NamespaceName)
            .Append('}')
            .Append(element.Name.LocalName)
            .Append('[');

        foreach (XAttribute attribute in element.Attributes()
            .Where(attribute => !attribute.IsNamespaceDeclaration)
            .OrderBy(attribute => attribute.Name.NamespaceName, StringComparer.Ordinal)
            .ThenBy(attribute => attribute.Name.LocalName, StringComparer.Ordinal)
            .ThenBy(attribute => attribute.Value, StringComparer.Ordinal))
        {
            builder.Append('{')
                .Append(attribute.Name.NamespaceName)
                .Append('}')
                .Append(attribute.Name.LocalName)
                .Append('=')
                .Append(attribute.Value)
                .Append(';');
        }

        builder.Append(']');

        foreach (string childSignature in element.Nodes()
            .Select(CanonicalNodeSignature)
            .Where(signature => signature.Length != 0)
            .OrderBy(signature => signature, StringComparer.Ordinal))
        {
            builder.Append(childSignature);
        }
    }

    private static string CanonicalNodeSignature(XNode node)
    {
        if (node is XElement element)
        {
            var builder = new StringBuilder();
            AppendCanonicalElement(builder, element);
            return builder.ToString();
        }

        if (node is XText text && !string.IsNullOrWhiteSpace(text.Value))
        {
            return text.Value;
        }

        return string.Empty;
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
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        string[] revisionIds = AllocateRevisionIds(package, matches.Count * 2, generatedRevisionIds, cancellationToken);
        XElement? paragraphProperties = paragraph.Element(OoxmlNs.W + "pPr");
        XElement? firstRunProperties = paragraph
            .Elements(OoxmlNs.W + "r")
            .Elements(OoxmlNs.W + "rPr")
            .FirstOrDefault();
        string author = GetRevisionAuthor(options);
        string timestamp = GetRevisionTimestamp(options);

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
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        int revisionCount = (deletedText.Length == 0 ? 0 : 1) + (insertedText.Length == 0 ? 0 : 1);
        string[] revisionIds = AllocateRevisionIds(package, revisionCount, generatedRevisionIds, cancellationToken);
        XElement? paragraphProperties = paragraph.Element(OoxmlNs.W + "pPr");
        XElement? firstRunProperties = paragraph
            .Elements(OoxmlNs.W + "r")
            .Elements(OoxmlNs.W + "rPr")
            .FirstOrDefault();
        string author = GetRevisionAuthor(options);
        string timestamp = GetRevisionTimestamp(options);

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

    private static void ReplaceCellParagraphTextWithTrackedChanges(
        OoxmlPackage package,
        IReadOnlyList<XElement> paragraphs,
        string insertedText,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        for (int i = 0; i < paragraphs.Count; i++)
        {
            XElement paragraph = paragraphs[i];
            string current = ReadVisibleText(paragraph);
            ReplaceWholeParagraphTextWithTrackedChanges(
                package,
                paragraph,
                current,
                i == 0 ? insertedText : string.Empty,
                options,
                generatedRevisionIds,
                cancellationToken);
        }
    }

    private static XNode[] CreateTrackedBookmarkReplacementNodes(
        OoxmlPackage package,
        string deletedText,
        string insertedText,
        XElement? runProperties,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        int revisionCount = (deletedText.Length == 0 ? 0 : 1) + (insertedText.Length == 0 ? 0 : 1);
        string[] revisionIds = AllocateRevisionIds(package, revisionCount, generatedRevisionIds, cancellationToken);
        string author = GetRevisionAuthor(options);
        string timestamp = GetRevisionTimestamp(options);
        var nodes = new List<XNode>();
        int revisionIndex = 0;
        if (deletedText.Length != 0)
        {
            nodes.Add(CreateDeletedRun(deletedText, runProperties, revisionIds[revisionIndex++], author, timestamp));
        }

        if (insertedText.Length != 0)
        {
            nodes.Add(CreateInsertedRun(insertedText, runProperties, revisionIds[revisionIndex], author, timestamp));
        }

        return nodes.ToArray();
    }

    private static XElement CreateTrackedInsertedParagraph(
        OoxmlPackage package,
        string text,
        string? style,
        XElement? paragraphProperties,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
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

        string revisionId = AllocateRevisionIds(package, 1, generatedRevisionIds, cancellationToken)[0];
        string author = GetRevisionAuthor(options);
        string timestamp = GetRevisionTimestamp(options);
        paragraph.Add(CreateInsertedRun(text, runProperties: null, revisionId, author, timestamp));
        return paragraph;
    }

    private static void SetParagraphStyleWithTrackedChange(
        OoxmlPackage package,
        XElement paragraph,
        string styleId,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        XElement oldParagraphProperties = paragraph.Element(OoxmlNs.W + "pPr") is { } existing
            ? new XElement(existing)
            : new XElement(OoxmlNs.W + "pPr");
        SetParagraphStyle(paragraph, styleId);
        XElement paragraphProperties = paragraph.Element(OoxmlNs.W + "pPr")
            ?? throw new InvalidDataException("Paragraph style update did not create paragraph properties.");
        paragraphProperties.Elements(OoxmlNs.W + "pPrChange").Remove();
        string revisionId = AllocateRevisionIds(package, 1, generatedRevisionIds, cancellationToken)[0];
        paragraphProperties.Add(new XElement(
            OoxmlNs.W + "pPrChange",
            new XAttribute(OoxmlNs.W + "id", revisionId),
            new XAttribute(OoxmlNs.W + "author", GetRevisionAuthor(options)),
            new XAttribute(OoxmlNs.W + "date", GetRevisionTimestamp(options)),
            oldParagraphProperties));
    }

    private static void SetTableStyleWithTrackedChange(
        OoxmlPackage package,
        XElement table,
        string styleId,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        XElement oldTableProperties = table.Element(OoxmlNs.W + "tblPr") is { } existing
            ? new XElement(existing)
            : new XElement(OoxmlNs.W + "tblPr");
        oldTableProperties.Elements(OoxmlNs.W + "tblPrChange").Remove();
        SetTableStyle(table, styleId);
        XElement tableProperties = table.Element(OoxmlNs.W + "tblPr")
            ?? throw new InvalidDataException("Table style update did not create table properties.");
        tableProperties.Elements(OoxmlNs.W + "tblPrChange").Remove();
        string revisionId = AllocateRevisionIds(package, 1, generatedRevisionIds, cancellationToken)[0];
        tableProperties.Add(new XElement(
            OoxmlNs.W + "tblPrChange",
            new XAttribute(OoxmlNs.W + "id", revisionId),
            new XAttribute(OoxmlNs.W + "author", GetRevisionAuthor(options)),
            new XAttribute(OoxmlNs.W + "date", GetRevisionTimestamp(options)),
            oldTableProperties));
    }

    private static void SetTableRowHeaderWithTrackedChange(
        OoxmlPackage package,
        XElement row,
        bool header,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        XElement oldRowProperties = row.Element(OoxmlNs.W + "trPr") is { } existing
            ? new XElement(existing)
            : new XElement(OoxmlNs.W + "trPr");
        oldRowProperties.Elements(OoxmlNs.W + "trPrChange").Remove();
        SetTableRowHeader(row, header);
        XElement rowProperties = row.Element(OoxmlNs.W + "trPr") ?? new XElement(OoxmlNs.W + "trPr");
        if (rowProperties.Parent is null)
        {
            row.AddFirst(rowProperties);
        }

        rowProperties.Elements(OoxmlNs.W + "trPrChange").Remove();
        string revisionId = AllocateRevisionIds(package, 1, generatedRevisionIds, cancellationToken)[0];
        rowProperties.Add(new XElement(
            OoxmlNs.W + "trPrChange",
            new XAttribute(OoxmlNs.W + "id", revisionId),
            new XAttribute(OoxmlNs.W + "author", GetRevisionAuthor(options)),
            new XAttribute(OoxmlNs.W + "date", GetRevisionTimestamp(options)),
            oldRowProperties));
    }

    private static void SetCellShadingWithTrackedChange(
        OoxmlPackage package,
        XElement cell,
        string? fill,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        XElement oldCellProperties = cell.Element(OoxmlNs.W + "tcPr") is { } existing
            ? new XElement(existing)
            : new XElement(OoxmlNs.W + "tcPr");
        oldCellProperties.Elements(OoxmlNs.W + "tcPrChange").Remove();
        SetCellShading(cell, fill);
        XElement cellProperties = cell.Element(OoxmlNs.W + "tcPr") ?? new XElement(OoxmlNs.W + "tcPr");
        if (cellProperties.Parent is null)
        {
            cell.AddFirst(cellProperties);
        }

        cellProperties.Elements(OoxmlNs.W + "tcPrChange").Remove();
        string revisionId = AllocateRevisionIds(package, 1, generatedRevisionIds, cancellationToken)[0];
        cellProperties.Add(new XElement(
            OoxmlNs.W + "tcPrChange",
            new XAttribute(OoxmlNs.W + "id", revisionId),
            new XAttribute(OoxmlNs.W + "author", GetRevisionAuthor(options)),
            new XAttribute(OoxmlNs.W + "date", GetRevisionTimestamp(options)),
            oldCellProperties));
    }

    private static void MarkRowRevision(
        OoxmlPackage package,
        XElement row,
        XName revisionName,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        XElement rowProperties = row.Element(OoxmlNs.W + "trPr") ?? new XElement(OoxmlNs.W + "trPr");
        if (rowProperties.Parent is null)
        {
            row.AddFirst(rowProperties);
        }

        rowProperties.Elements(OoxmlNs.W + "ins").Remove();
        rowProperties.Elements(OoxmlNs.W + "del").Remove();
        rowProperties.Elements(OoxmlNs.W + "trPrChange").Remove();
        string revisionId = AllocateRevisionIds(package, 1, generatedRevisionIds, cancellationToken)[0];
        rowProperties.AddFirst(new XElement(
            revisionName,
            new XAttribute(OoxmlNs.W + "id", revisionId),
            new XAttribute(OoxmlNs.W + "author", GetRevisionAuthor(options)),
            new XAttribute(OoxmlNs.W + "date", GetRevisionTimestamp(options))));
    }

    private static void SetSectionPropertiesWithTrackedChange(
        OoxmlPackage package,
        XElement sectionProperties,
        Action<XElement> update,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        XElement oldSectionProperties = new(sectionProperties);
        oldSectionProperties.Elements(OoxmlNs.W + "sectPrChange").Remove();
        update(sectionProperties);
        sectionProperties.Elements(OoxmlNs.W + "sectPrChange").Remove();
        string revisionId = AllocateRevisionIds(package, 1, generatedRevisionIds, cancellationToken)[0];
        sectionProperties.Add(new XElement(
            OoxmlNs.W + "sectPrChange",
            new XAttribute(OoxmlNs.W + "id", revisionId),
            new XAttribute(OoxmlNs.W + "author", GetRevisionAuthor(options)),
            new XAttribute(OoxmlNs.W + "date", GetRevisionTimestamp(options)),
            oldSectionProperties));
    }

    private static IReadOnlyList<DocxDiagnostic> ValidateTrackChangeOptions(DocxEditOptions options)
    {
        if (!IsTrackedMode(options))
        {
            return [];
        }

        if (string.IsNullOrWhiteSpace(options.Author))
        {
            return
            [
                new DocxDiagnostic(
                    DocxSeverity.Error,
                    "E6003",
                    "TrackChangesMode.Suggest/Require requires a non-empty revision author; no document output was written.",
                    Feature: "track-changes-revision-metadata",
                    Fallback: "no-output-written")
            ];
        }

        return [];
    }

    private static string GetRevisionAuthor(DocxEditOptions options)
    {
        return options.Author.Trim();
    }

    private static string GetRevisionTimestamp(DocxEditOptions options)
    {
        return options.TimestampUtc.ToUniversalTime().ToString("O", System.Globalization.CultureInfo.InvariantCulture);
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
        if (alt is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-image-alt target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
        }

        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateImageContentTypeGuard(operation, target, imageTarget.Part.ContentType, diagnostics))
        {
            return diagnostics;
        }

        if (!TryGetImageDrawingContainer(imageTarget, target, operation, out XElement? imageContainer, out DocxDiagnostic? diagnostic))
        {
            return [diagnostic];
        }

        if (!apply)
        {
            return [];
        }

        SetImageAlt(imageContainer, alt, target);
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

        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-image-metadata target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
        }

        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateImageContentTypeGuard(operation, target, imageTarget.Part.ContentType, diagnostics))
        {
            return diagnostics;
        }

        if (!TryGetImageDrawingContainer(imageTarget, target, operation, out XElement? imageContainer, out DocxDiagnostic? diagnostic))
        {
            return [diagnostic];
        }

        if (!apply)
        {
            return [];
        }

        SetImageMetadata(imageContainer, target, alt, title, name);
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

        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-image-size target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
        }

        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateImageContentTypeGuard(operation, target, imageTarget.Part.ContentType, diagnostics))
        {
            return diagnostics;
        }

        if (!TryGetImageDrawingContainer(imageTarget, target, operation, out XElement? imageContainer, out DocxDiagnostic? diagnostic))
        {
            return [diagnostic];
        }

        TryReadExistingImageSize(imageContainer, imageTarget.Part, out long currentWidthEmus, out long currentHeightEmus);
        if (!TryReadImageSize(operation, currentWidthEmus, currentHeightEmus, out long widthEmus, out long heightEmus, out diagnostic))
        {
            return [diagnostic];
        }

        if (!apply)
        {
            return [];
        }

        SetImageSize(imageContainer, widthEmus, heightEmus);
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

        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-image-wrap target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
        }

        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateImageContentTypeGuard(operation, target, imageTarget.Part.ContentType, diagnostics))
        {
            return diagnostics;
        }

        if (!TryGetImageDrawingContainer(imageTarget, target, operation, out XElement? imageContainer, out DocxDiagnostic? diagnostic))
        {
            return [diagnostic];
        }

        if (imageContainer.Name != OoxmlNs.Wp + "anchor")
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

        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-image-position target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
        }

        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateImageContentTypeGuard(operation, target, imageTarget.Part.ContentType, diagnostics))
        {
            return diagnostics;
        }

        if (!TryGetImageDrawingContainer(imageTarget, target, operation, out XElement? imageContainer, out DocxDiagnostic? diagnostic))
        {
            return [diagnostic];
        }

        if (imageContainer.Name != OoxmlNs.Wp + "anchor")
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

        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-image-crop target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
        }

        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateImageContentTypeGuard(operation, target, imageTarget.Part.ContentType, diagnostics))
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
        [NotNullWhen(true)] out XElement? container,
        [NotNullWhen(false)] out DocxDiagnostic? diagnostic)
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

    private static void SetImageAlt(XElement container, string? alt, string target)
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
        [NotNullWhen(false)] out DocxDiagnostic? diagnostic)
    {
        widthEmus = currentWidthEmus;
        heightEmus = currentHeightEmus;
        diagnostic = null;
        bool hasWidth = operation.Fields.TryGetValue("width", out string? width);
        bool hasHeight = operation.Fields.TryGetValue("height", out string? height);
        if (width is not null && (!OoxmlUnits.TryParseDimension(width, out widthEmus) || widthEmus <= 0))
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E5206", $"Invalid image width '{width}'.", operation, operation.Fields.GetValueOrDefault("target"));
            return false;
        }

        if (height is not null && (!OoxmlUnits.TryParseDimension(height, out heightEmus) || heightEmus <= 0))
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

        if (relative is not null && !IsValidImagePositionRelative(axis, relative))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E5210", $"Unsupported image {axis} relative value '{relative}'.", operation, operation.Fields.GetValueOrDefault("target")));
            return false;
        }

        long? offsetEmus = null;
        if (offsetText is not null)
        {
            if (!TryParseSignedDimension(offsetText, out long parsedOffset))
            {
                diagnostics.Add(Diagnostic(DocxSeverity.Error, "E5210", $"Image position field '{axis}-offset' must be a signed dimension.", operation, operation.Fields.GetValueOrDefault("target")));
                return false;
            }

            offsetEmus = parsedOffset;
        }

        if (align is not null && !IsValidImagePositionAlign(axis, align))
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
        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ImageBlipTarget? imageTarget = ResolveImageBlipTarget(package, target, cancellationToken);
        if (imageTarget is null && !IsSupportedImageTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported delete-image target '{target}'. Expected an image ID such as M.I0001 or H001.I0001.", operation, target)];
        }

        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateImageContentTypeGuard(operation, target, imageTarget.Part.ContentType, diagnostics))
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
            package.RemoveRelationship(imageTarget.PartName, imageTarget.RelationshipId, cancellationToken);
            if (!AnyRelationshipTargetsPart(package, imageTarget.Part.Name, cancellationToken))
            {
                package.RemovePart(imageTarget.Part.Name, cancellationToken);
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
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? countText = ReadRequiredField(operation, "count", diagnostics);
        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!int.TryParse(countText, out int count) || count is < 1 or > 4)
        {
            return [Diagnostic(DocxSeverity.Error, "E6201", "Section column count must be between 1 and 4.", operation, target)];
        }

        SectionTarget? sectionTarget = ResolveMainSectionTarget(package, target, cancellationToken);
        if (sectionTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateSectionGuards(operation, target, sectionTarget.SectionProperties, diagnostics))
        {
            return diagnostics;
        }

        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges && sectionTarget.SectionProperties.Elements(OoxmlNs.W + "sectPrChange").Any())
        {
            if (!TrackUnsupportedShape(options, operation, target, "section already contains tracked section property revision markup", diagnostics))
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
            SetSectionPropertiesWithTrackedChange(
                package,
                sectionTarget.SectionProperties,
                properties => SetSectionColumns(properties, count),
                options,
                generatedRevisionIds,
                cancellationToken);
        }
        else
        {
            SetSectionColumns(sectionTarget.SectionProperties, count);
        }

        SaveMainDocument(package, sectionTarget.Document);
        return diagnostics;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetSectionOrientation(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? orientation = ReadRequiredField(operation, "orientation", diagnostics);
        if (orientation is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (orientation is not ("portrait" or "landscape"))
        {
            return [Diagnostic(DocxSeverity.Error, "E6202", "Section orientation must be portrait or landscape.", operation, target)];
        }

        SectionTarget? sectionTarget = ResolveMainSectionTarget(package, target, cancellationToken);
        if (sectionTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateSectionGuards(operation, target, sectionTarget.SectionProperties, diagnostics))
        {
            return diagnostics;
        }

        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges && sectionTarget.SectionProperties.Elements(OoxmlNs.W + "sectPrChange").Any())
        {
            if (!TrackUnsupportedShape(options, operation, target, "section already contains tracked section property revision markup", diagnostics))
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
            SetSectionPropertiesWithTrackedChange(
                package,
                sectionTarget.SectionProperties,
                properties => SetSectionOrientation(properties, orientation),
                options,
                generatedRevisionIds,
                cancellationToken);
        }
        else
        {
            SetSectionOrientation(sectionTarget.SectionProperties, orientation);
        }

        SaveMainDocument(package, sectionTarget.Document);
        return diagnostics;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetCell(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        bool force = ReadBooleanField(operation, "force", diagnostics) ?? false;
        if (text is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        CellTarget? cellTarget = ResolveCellTarget(package, target, cancellationToken);
        if (cellTarget is null && !IsSupportedCellTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-cell target '{target}'. Expected a table cell ID such as M.T0001.R02.C03 or H001.T0001.R02.C03, or a merge group ID such as M.T0001.MG0001.", operation, target)];
        }

        if (cellTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateTableGuards(operation, target, cellTarget.Table, cellTarget.Row, diagnostics))
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

        bool useTrackedChanges = IsTrackedMode(options);
        XElement[] trackedParagraphs = [];
        if (useTrackedChanges)
        {
            if (force)
            {
                if (!TrackUnsupportedShape(options, operation, target, "tracked set-cell does not support force true replacement", diagnostics))
                {
                    return diagnostics;
                }

                useTrackedChanges = false;
            }
            else if (!TryGetTrackedSetCellParagraphs(cellTarget.Cell, text, out trackedParagraphs, out string? trackedUnsupportedReason))
            {
                if (!TrackUnsupportedShape(options, operation, target, trackedUnsupportedReason, diagnostics))
                {
                    return diagnostics;
                }

                useTrackedChanges = false;
            }
        }

        if (!useTrackedChanges && !force && !IsSimpleEditableCell(cellTarget.Cell))
        {
            return [Diagnostic(DocxSeverity.Error, "E4302", $"Cell '{target}' contains unsupported content. Use force true only when replacing all cell content is intended.", operation, target)];
        }

        if (!apply)
        {
            return diagnostics;
        }

        if (useTrackedChanges)
        {
            ReplaceCellParagraphTextWithTrackedChanges(package, trackedParagraphs, text, options, generatedRevisionIds, cancellationToken);
            SaveDocumentPart(package, cellTarget.PartName, cellTarget.Document);
            return diagnostics;
        }

        ReplaceCellText(cellTarget.Cell, text);
        SaveDocumentPart(package, cellTarget.PartName, cellTarget.Document);
        return diagnostics;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetCellShading(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? fill = operation.Fields.GetValueOrDefault("fill");
        string? expectedFill = operation.Fields.GetValueOrDefault("expect-fill");
        bool clear = ReadBooleanField(operation, "clear", diagnostics) ?? false;
        if (fill is null && !clear)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4202", "Operation 'set-cell-shading' requires either 'fill' or 'clear true'.", operation, target));
        }

        if (fill is not null && clear)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4202", "Operation 'set-cell-shading' cannot combine 'fill' with 'clear true'.", operation, target));
        }

        string? normalizedFill = null;
        if (fill is not null && !TryNormalizeCellShadingFill(fill, allowNone: false, out normalizedFill))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4202", $"Field 'fill' must be a 6-digit hexadecimal color or 'auto', found '{fill}'.", operation, target));
        }

        string? normalizedExpectedFill = null;
        if (expectedFill is not null && !TryNormalizeCellShadingFill(expectedFill, allowNone: true, out normalizedExpectedFill))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4202", $"Field 'expect-fill' must be a 6-digit hexadecimal color, 'auto', or 'none', found '{expectedFill}'.", operation, target));
        }

        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        CellTarget? cellTarget = ResolveCellTarget(package, target, cancellationToken);
        if (cellTarget is null && !IsSupportedCellTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-cell-shading target '{target}'. Expected a table cell ID such as M.T0001.R02.C03 or H001.T0001.R02.C03, or a merge group ID such as M.T0001.MG0001.", operation, target)];
        }

        if (cellTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (IsVerticalMergeContinuation(cellTarget.Cell))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Unsupported merged-cell target '{target}'. Target the vertical-merge root cell instead.", operation, target)];
        }

        string? currentFill = ReadCellShadingFill(cellTarget.Cell);
        if (normalizedExpectedFill is not null && !CellShadingFillMatches(currentFill, normalizedExpectedFill))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected cell shading fill '{normalizedExpectedFill}', found '{currentFill ?? "none"}'.", operation, target)];
        }

        if (!apply)
        {
            return diagnostics;
        }

        if (IsTrackedMode(options))
        {
            SetCellShadingWithTrackedChange(package, cellTarget.Cell, clear ? null : normalizedFill, options, generatedRevisionIds, cancellationToken);
        }
        else
        {
            SetCellShading(cellTarget.Cell, clear ? null : normalizedFill);
        }

        SaveDocumentPart(package, cellTarget.PartName, cellTarget.Document);
        return diagnostics;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetTableStyle(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? style = ReadRequiredField(operation, "style", diagnostics);
        string? expectedStyle = operation.Fields.GetValueOrDefault("expect-style");
        if (style is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        TableTarget? tableTarget = ResolveTableTarget(package, target, cancellationToken);
        if (tableTarget is null && !IsSupportedTableTargetShape(target))
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

        if (IsTrackedMode(options))
        {
            SetTableStyleWithTrackedChange(package, tableTarget.Table, style, options, generatedRevisionIds, cancellationToken);
        }
        else
        {
            SetTableStyle(tableTarget.Table, style);
        }

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

        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        TableTarget? tableTarget = ResolveTableTarget(package, target, cancellationToken);
        if (tableTarget is null && !IsSupportedTableTargetShape(target))
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
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        _ = ReadRequiredField(operation, "header", diagnostics);
        bool? header = ReadBooleanField(operation, "header", diagnostics);
        bool? expectedHeader = ReadBooleanField(operation, "expect-header", diagnostics);
        if (target is null || header is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        RowTarget? rowTarget = ResolveRowTarget(package, target, cancellationToken);
        if (rowTarget is null && !IsSupportedRowTargetShape(target))
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

        if (IsTrackedMode(options))
        {
            SetTableRowHeaderWithTrackedChange(package, rowTarget.Row, header.Value, options, generatedRevisionIds, cancellationToken);
        }
        else
        {
            SetTableRowHeader(rowTarget.Row, header.Value);
        }

        SaveDocumentPart(package, rowTarget.PartName, rowTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteAppendRow(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        TableTarget? tableTarget = ResolveTableTarget(package, target, cancellationToken);
        if (tableTarget is null && !IsSupportedTableTargetShape(target))
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

        if (!ValidateTableGuards(operation, target, tableTarget.Table, row: null, diagnostics))
        {
            return diagnostics;
        }

        XElement[] rows = tableTarget.Table.Elements(OoxmlNs.W + "tr").ToArray();
        if (rows.Length == 0)
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' has no rows to clone.", operation, target)];
        }

        if (!TryGetConsistentVisualColumnCount(tableTarget.Table, out int columnCount))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' does not have a consistent visual grid and cannot be appended safely.", operation, target)];
        }

        XElement templateRow = rows[^1];
        if (!CanAppendRowWithVerticalMerges(tableTarget.Table, templateRow, out string? verticalMergeReason))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' cannot be appended safely: {verticalMergeReason}.", operation, target)];
        }

        int expectedCellCount = templateRow.Elements(OoxmlNs.W + "tc").Count();
        if (expectedCellCount == 0)
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' has no cells to clone.", operation, target)];
        }

        if (cellTexts.Length != expectedCellCount)
        {
            return [Diagnostic(DocxSeverity.Error, "E4303", $"append-row expected {expectedCellCount} cell field(s), but received {cellTexts.Length}.", operation, target)];
        }

        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges && !TryUseTrackedRowStructure(options, operation, target, tableTarget.Table, force: false, diagnostics))
        {
            if (options.TrackChanges == TrackChangesMode.Require)
            {
                return diagnostics;
            }

            useTrackedChanges = false;
        }

        if (!apply)
        {
            return diagnostics;
        }

        XElement newRow = CreateRowFromTemplate(templateRow, cellTexts);
        if (useTrackedChanges)
        {
            MarkRowRevision(package, newRow, OoxmlNs.W + "ins", options, generatedRevisionIds, cancellationToken);
        }

        rows[^1].AddAfterSelf(newRow);
        SaveDocumentPart(package, tableTarget.PartName, tableTarget.Document);
        return diagnostics;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteInsertRow(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool insertAfter,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        bool force = ReadBooleanField(operation, "force", diagnostics) ?? false;
        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        RowTarget? rowTarget = ResolveRowTarget(package, target, cancellationToken);
        if (rowTarget is null && !IsSupportedRowTargetShape(target))
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

        if (!ValidateTableGuards(operation, target, rowTarget.Table, rowTarget.Row, diagnostics))
        {
            return diagnostics;
        }

        if (!CanInsertRowWithVerticalMerges(rowTarget.Table, rowTarget.Row, insertAfter, out string? verticalMergeReason))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' cannot be edited safely by {operation.OperationName}: {verticalMergeReason}.", operation, target)];
        }

        int expectedCellCount;
        if (IsRectangular(rowTarget.Table, out int columnCount))
        {
            expectedCellCount = columnCount;
        }
        else if (TryGetConsistentVisualColumnCount(rowTarget.Table, out _))
        {
            expectedCellCount = rowTarget.Row.Elements(OoxmlNs.W + "tc").Count();
        }
        else
        {
            if (!force)
            {
                return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' does not have a consistent visual grid and cannot be edited safely without force true.", operation, target)];
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

        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges && !TryUseTrackedRowStructure(options, operation, target, rowTarget.Table, force, diagnostics))
        {
            if (options.TrackChanges == TrackChangesMode.Require)
            {
                return diagnostics;
            }

            useTrackedChanges = false;
        }

        if (!apply)
        {
            return diagnostics;
        }

        XElement newRow = CreateRowFromTemplate(rowTarget.Row, cellTexts);
        if (useTrackedChanges)
        {
            MarkRowRevision(package, newRow, OoxmlNs.W + "ins", options, generatedRevisionIds, cancellationToken);
        }

        if (insertAfter)
        {
            rowTarget.Row.AddAfterSelf(newRow);
        }
        else
        {
            rowTarget.Row.AddBeforeSelf(newRow);
        }

        SaveDocumentPart(package, rowTarget.PartName, rowTarget.Document);
        return diagnostics;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteDeleteRow(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? expectedContains = operation.Fields.GetValueOrDefault("expect-contains");
        bool force = ReadBooleanField(operation, "force", diagnostics) ?? false;
        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        RowTarget? rowTarget = ResolveRowTarget(package, target, cancellationToken);
        if (rowTarget is null && !IsSupportedRowTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported delete-row target '{target}'. Expected a table row ID such as M.T0001.R02 or H001.T0001.R02.", operation, target)];
        }

        if (rowTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!ValidateTableGuards(operation, target, rowTarget.Table, rowTarget.Row, diagnostics))
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

        if (!CanDeleteRowWithVerticalMerges(rowTarget.Table, rowTarget.Row, out string? verticalMergeReason))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' cannot be edited safely by delete-row: {verticalMergeReason}.", operation, target)];
        }

        if (!force &&
            !IsRectangular(rowTarget.Table, out _) &&
            !TryGetConsistentVisualColumnCount(rowTarget.Table, out _))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' does not have a consistent visual grid and cannot be edited safely without force true.", operation, target)];
        }

        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges && !TryUseTrackedRowStructure(options, operation, target, rowTarget.Table, force, diagnostics))
        {
            if (options.TrackChanges == TrackChangesMode.Require)
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
            MarkRowRevision(package, rowTarget.Row, OoxmlNs.W + "del", options, generatedRevisionIds, cancellationToken);
        }
        else
        {
            PromoteVerticalMergeContinuationsAfterDeletedRow(rowTarget.Table, rowTarget.Row);
            rowTarget.Row.Remove();
        }

        SaveDocumentPart(package, rowTarget.PartName, rowTarget.Document);
        return diagnostics;
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
            if (!TryGetConsistentVisualColumnCount(table, out int actualColumnCount))
            {
                diagnostics.Add(Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected {expectedColumnCount} column(s), but table does not have a consistent visual grid.", operation, target));
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

    private static bool TryUseTrackedRowStructure(
        DocxEditOptions options,
        DocxPatchOperation operation,
        string target,
        XElement table,
        bool force,
        List<DocxDiagnostic> diagnostics)
    {
        if (force)
        {
            AddTrackedRowStructureUnsupported(options, operation, target, "tracked row operations do not support force true", diagnostics);
            return false;
        }

        if (ContainsNestedTables(table))
        {
            AddTrackedRowStructureUnsupported(options, operation, target, "tracked row operations do not support nested tables", diagnostics);
            return false;
        }

        if (table.Elements(OoxmlNs.W + "tr").Any(RowHasTrackedRowRevision))
        {
            AddTrackedRowStructureUnsupported(options, operation, target, "table already contains tracked row insertion, deletion, or property revision markup", diagnostics);
            return false;
        }

        if (!IsRectangular(table, out _))
        {
            AddTrackedRowStructureUnsupported(options, operation, target, "tracked row operations require a simple rectangular table", diagnostics);
            return false;
        }

        return true;
    }

    private static void AddTrackedRowStructureUnsupported(
        DocxEditOptions options,
        DocxPatchOperation operation,
        string target,
        string reason,
        List<DocxDiagnostic> diagnostics)
    {
        _ = TrackUnsupportedShape(options, operation, target, reason, diagnostics);
    }

    private static bool ContainsNestedTables(XElement table)
    {
        return table
            .Elements(OoxmlNs.W + "tr")
            .Elements(OoxmlNs.W + "tc")
            .Descendants(OoxmlNs.W + "tbl")
            .Any();
    }

    private static bool RowHasTrackedRowRevision(XElement row)
    {
        XElement? rowProperties = row.Element(OoxmlNs.W + "trPr");
        return rowProperties is not null &&
            (rowProperties.Elements(OoxmlNs.W + "ins").Any() ||
                rowProperties.Elements(OoxmlNs.W + "del").Any() ||
                rowProperties.Elements(OoxmlNs.W + "trPrChange").Any());
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

    private static bool TryParseMainMergeGroupTarget(string target, out int tableOrdinal, out int mergeGroupOrdinal)
    {
        tableOrdinal = 0;
        mergeGroupOrdinal = 0;
        if (target.Length != 14 ||
            !target.StartsWith("M.T", StringComparison.Ordinal) ||
            target[7..10] != ".MG")
        {
            return false;
        }

        return int.TryParse(target[3..7], out tableOrdinal) &&
            int.TryParse(target[10..14], out mergeGroupOrdinal);
    }

    private static bool TryParseStoryMergeGroupTarget(
        string target,
        char storyPrefix,
        out int storyOrdinal,
        out int tableOrdinal,
        out int mergeGroupOrdinal)
    {
        storyOrdinal = 0;
        tableOrdinal = 0;
        mergeGroupOrdinal = 0;
        if (target.Length != 17 ||
            target[0] != storyPrefix ||
            target[4..6] != ".T" ||
            target[10..13] != ".MG")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[6..10], out tableOrdinal) &&
            int.TryParse(target[13..17], out mergeGroupOrdinal);
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

        return ResolveMainParagraphElementBySelector(body, selector, operation, out diagnostics);
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
            return paragraph is null
                ? null
                : new BlockTarget(package.MainDocumentPartName, document, paragraph);
        }

        if (TryParseMainTableTarget(target, out int mainTableOrdinal))
        {
            XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
            XElement? table = FindTable(body, mainTableOrdinal);
            return table is null
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
                return paragraph is null
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


        XDocument mainDocument = LoadMainDocument(package, cancellationToken, out XElement mainBody);
        XElement? selectedParagraph = ResolveMainParagraphElementBySelector(mainBody, selector, operation, out diagnostics);
        return selectedParagraph is null
            ? null
            : new ParagraphTarget(package.MainDocumentPartName, mainDocument, selectedParagraph);
    }

    private static bool TryParseTargetSelector(
        string target,
        DocxPatchOperation operation,
        [NotNullWhen(true)] out TargetSelector? selector,
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
        foreach (XElement block in body.Elements())
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement? element = block.Name == OoxmlNs.W + "p"
                ? block.Element(OoxmlNs.W + "pPr")?.Element(OoxmlNs.W + "sectPr")
                : block.Name == OoxmlNs.W + "sectPr"
                    ? block
                    : null;
            if (element is null)
            {
                continue;
            }

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
        if (storyOrdinal < 1)
        {
            return null;
        }

        ResolvedOoxmlRelationship? relationship = package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => relationship.Type == relationshipType)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .ElementAtOrDefault(storyOrdinal - 1);
        if (relationship is null)
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
        if (storyOrdinal < 1 || blockOrdinal < 1)
        {
            return null;
        }

        ResolvedOoxmlRelationship? relationship = package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => relationship.Type == relationshipType)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .ElementAtOrDefault(storyOrdinal - 1);
        if (relationship is null)
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
            TryParseMainMergeGroupTarget(target, out _, out _) ||
            TryParseStoryCellTarget(target, 'H', out _, out _, out _, out _) ||
            TryParseStoryCellTarget(target, 'F', out _, out _, out _, out _) ||
            TryParseStoryMergeGroupTarget(target, 'H', out _, out _, out _) ||
            TryParseStoryMergeGroupTarget(target, 'F', out _, out _, out _);
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
            return table is null
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
        int rowOrdinal;
        int cellOrdinal;
        int visualColumnIndex;
        if (TryParseMainCellTarget(target, out int mainTableOrdinal, out rowOrdinal, out cellOrdinal))
        {
            rowTarget = ResolveRowTarget(package, $"M.T{mainTableOrdinal:0000}.R{rowOrdinal:00}", cancellationToken);
            visualColumnIndex = cellOrdinal;
        }
        else if (TryParseStoryCellTarget(target, 'H', out int headerOrdinal, out int headerTableOrdinal, out rowOrdinal, out cellOrdinal))
        {
            rowTarget = ResolveRowTarget(package, $"H{headerOrdinal:000}.T{headerTableOrdinal:0000}.R{rowOrdinal:00}", cancellationToken);
            visualColumnIndex = cellOrdinal;
        }
        else if (TryParseStoryCellTarget(target, 'F', out int footerOrdinal, out int footerTableOrdinal, out rowOrdinal, out cellOrdinal))
        {
            rowTarget = ResolveRowTarget(package, $"F{footerOrdinal:000}.T{footerTableOrdinal:0000}.R{rowOrdinal:00}", cancellationToken);
            visualColumnIndex = cellOrdinal;
        }
        else if (TryParseMainMergeGroupTarget(target, out int mergeMainTableOrdinal, out int mainMergeGroupOrdinal))
        {
            TableTarget? tableTarget = ResolveTableTarget(package, $"M.T{mergeMainTableOrdinal:0000}", cancellationToken);
            return tableTarget is null ? null : ResolveMergeGroupCellTarget(tableTarget, mainMergeGroupOrdinal);
        }
        else if (TryParseStoryMergeGroupTarget(target, 'H', out int mergeHeaderOrdinal, out int mergeHeaderTableOrdinal, out int headerMergeGroupOrdinal))
        {
            TableTarget? tableTarget = ResolveTableTarget(package, $"H{mergeHeaderOrdinal:000}.T{mergeHeaderTableOrdinal:0000}", cancellationToken);
            return tableTarget is null ? null : ResolveMergeGroupCellTarget(tableTarget, headerMergeGroupOrdinal);
        }
        else if (TryParseStoryMergeGroupTarget(target, 'F', out int mergeFooterOrdinal, out int mergeFooterTableOrdinal, out int footerMergeGroupOrdinal))
        {
            TableTarget? tableTarget = ResolveTableTarget(package, $"F{mergeFooterOrdinal:000}.T{mergeFooterTableOrdinal:0000}", cancellationToken);
            return tableTarget is null ? null : ResolveMergeGroupCellTarget(tableTarget, footerMergeGroupOrdinal);
        }
        else
        {
            return null;
        }

        XElement? cell = rowTarget is null ? null : FindCellByVisualColumn(rowTarget.Row, visualColumnIndex);
        return rowTarget is null || cell is null
            ? null
            : new CellTarget(rowTarget.PartName, rowTarget.Document, rowTarget.Table, rowTarget.Row, cell, visualColumnIndex);
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
                string? verticalMerge = ReadTableCellVerticalMerge(cell);
                if (string.Equals(verticalMerge, "restart", StringComparison.Ordinal))
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
            return FindContentControlTarget(package, package.MainDocumentPartName, mainControlOrdinal, cancellationToken);
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
            return field is null
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
            return FindBookmarkTarget(package, package.MainDocumentPartName, mainBookmarkOrdinal, cancellationToken);
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
            return FindHyperlinkTarget(package, package.MainDocumentPartName, mainHyperlinkOrdinal, cancellationToken);
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
            return FindImageBlipTarget(package, package.MainDocumentPartName, mainImageOrdinal, cancellationToken);
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

    private static IReadOnlyList<string> GetEditableStoryPartNames(
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {

        var partNames = new List<string> { package.MainDocumentPartName };
        IReadOnlyList<ResolvedOoxmlRelationship> relationships = package.GetResolvedRelationships(package.MainDocumentPartName, cancellationToken);
        partNames.AddRange(relationships
            .Where(relationship => relationship.Type == OoxmlRelTypes.Header)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .Select(relationship => relationship.ResolvedTarget)
            .Where(partName => package.GetPart(partName) is not null));
        partNames.AddRange(relationships
            .Where(relationship => relationship.Type == OoxmlRelTypes.Footer)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .Select(relationship => relationship.ResolvedTarget)
            .Where(partName => package.GetPart(partName) is not null));
        return partNames;
    }

    private static bool TryReadAsset(
        IDocxAssetProvider? assetProvider,
        string asset,
        CancellationToken cancellationToken,
        out byte[] bytes,
        [NotNullWhen(true)] out string? contentType,
        [NotNullWhen(false)] out DocxDiagnostic? diagnostic,
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
        [NotNullWhen(true)] out string? styleId,
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
            if (ProtectedTextEditElements.TryGetValue(element.Name, out string? found) && found is not null)
            {
                feature = found;
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
        [NotNullWhen(false)] out string? unsupportedReason)
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

        cell.Add(CreateSimpleParagraph(text, style: null, paragraphProperties: paragraphProperties));
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
            cell.Add(CreateSimpleParagraph(cellTexts[i], style: null, paragraphProperties: paragraphProperties));
            row.Add(cell);
        }

        return row;
    }

    private static XElement CreateSimpleParagraph(string text, string? style, XElement? paragraphProperties)
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
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        int nextId = FindNextRevisionId(package, cancellationToken);
        foreach (string generatedRevisionId in generatedRevisionIds)
        {
            if (int.TryParse(generatedRevisionId, out int id) && id >= nextId)
            {
                nextId = id + 1;
            }
        }

        string[] ids = BuildRevisionIds(nextId, count);
        generatedRevisionIds.AddRange(ids);
        return ids;
    }

    private static string[] AllocateRevisionIds(
        OoxmlPackage package,
        int count,
        CancellationToken cancellationToken)
    {
        return BuildRevisionIds(FindNextRevisionId(package, cancellationToken), count);
    }

    private static int FindNextRevisionId(
        OoxmlPackage package,
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

        return nextId;
    }

    private static string[] BuildRevisionIds(int nextId, int count)
    {
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

    private static bool CanAppendRowWithVerticalMerges(XElement table, XElement templateRow, [NotNullWhen(false)] out string? unsupportedReason)
    {
        unsupportedReason = null;
        if (!ContainsVerticalMerges(table) || !RowHasVerticalMerge(templateRow))
        {
            return true;
        }

        unsupportedReason = "the last row contains vertical merge cells, so appending would need to decide whether to extend or terminate those merge chains";
        return false;
    }

    private static bool CanInsertRowWithVerticalMerges(
        XElement table,
        XElement templateRow,
        bool insertAfter,
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        unsupportedReason = null;
        if (!ContainsVerticalMerges(table))
        {
            return true;
        }

        if (RowHasVerticalMerge(templateRow))
        {
            unsupportedReason = "the template row contains vertical merge cells";
            return false;
        }

        XElement[] rows = table.Elements(OoxmlNs.W + "tr").ToArray();
        int rowIndex = Array.IndexOf(rows, templateRow);
        if (rowIndex < 0)
        {
            unsupportedReason = "the target row was not found in its table";
            return false;
        }

        int boundaryBeforeRowIndex = insertAfter ? rowIndex + 1 : rowIndex;
        if (VerticalMergeContinuesAcrossInsertionBoundary(rows, boundaryBeforeRowIndex))
        {
            unsupportedReason = "the insertion boundary crosses an active vertical merge chain";
            return false;
        }

        return true;
    }

    private static bool CanDeleteRowWithVerticalMerges(XElement table, XElement row, [NotNullWhen(false)] out string? unsupportedReason)
    {
        unsupportedReason = null;
        if (!ContainsVerticalMerges(table))
        {
            return true;
        }

        if (!TryGetConsistentVisualColumnCount(table, out _))
        {
            unsupportedReason = "the table does not have a consistent visual grid";
            return false;
        }

        XElement[] rows = table.Elements(OoxmlNs.W + "tr").ToArray();
        int rowIndex = Array.IndexOf(rows, row);
        if (rowIndex < 0)
        {
            unsupportedReason = "the target row was not found in its table";
            return false;
        }

        XElement? nextRow = rowIndex + 1 < rows.Length ? rows[rowIndex + 1] : null;
        if (nextRow is null)
        {
            return true;
        }

        foreach (TableCellGridSlot deletedSlot in EnumerateTableRowCells(row))
        {
            if (!string.Equals(ReadTableCellVerticalMerge(deletedSlot.Cell), "restart", StringComparison.Ordinal))
            {
                continue;
            }

            XElement? nextCell = FindCellByVisualColumn(nextRow, deletedSlot.ColumnIndex);
            if (nextCell is null || !string.Equals(ReadTableCellVerticalMerge(nextCell), "continue", StringComparison.Ordinal))
            {
                continue;
            }

            int nextCellSpan = ReadTableCellColumnSpan(nextCell);
            if (nextCellSpan != deletedSlot.ColumnSpan)
            {
                unsupportedReason = "a vertical merge continuation below the deleted root has a different column span";
                return false;
            }
        }

        return true;
    }

    private static bool VerticalMergeContinuesAcrossInsertionBoundary(IReadOnlyList<XElement> rows, int boundaryBeforeRowIndex)
    {
        if (boundaryBeforeRowIndex <= 0 || boundaryBeforeRowIndex >= rows.Count)
        {
            return false;
        }

        return rows[boundaryBeforeRowIndex]
            .Elements(OoxmlNs.W + "tc")
            .Any(cell => string.Equals(ReadTableCellVerticalMerge(cell), "continue", StringComparison.Ordinal));
    }

    private static bool RowHasVerticalMerge(XElement row)
    {
        return row.Elements(OoxmlNs.W + "tc").Any(cell => ReadTableCellVerticalMerge(cell) is not null);
    }

    private static void PromoteVerticalMergeContinuationsAfterDeletedRow(XElement table, XElement row)
    {
        if (!ContainsVerticalMerges(table))
        {
            return;
        }

        XElement[] rows = table.Elements(OoxmlNs.W + "tr").ToArray();
        int rowIndex = Array.IndexOf(rows, row);
        if (rowIndex < 0 || rowIndex + 1 >= rows.Length)
        {
            return;
        }

        XElement nextRow = rows[rowIndex + 1];
        foreach (TableCellGridSlot deletedSlot in EnumerateTableRowCells(row))
        {
            if (!string.Equals(ReadTableCellVerticalMerge(deletedSlot.Cell), "restart", StringComparison.Ordinal))
            {
                continue;
            }

            XElement? nextCell = FindCellByVisualColumn(nextRow, deletedSlot.ColumnIndex);
            if (nextCell is not null && string.Equals(ReadTableCellVerticalMerge(nextCell), "continue", StringComparison.Ordinal))
            {
                SetTableCellVerticalMerge(nextCell, "restart");
            }
        }
    }

    private static void SetTableCellVerticalMerge(XElement cell, string value)
    {
        XElement? cellProperties = cell.Element(OoxmlNs.W + "tcPr");
        if (cellProperties is null)
        {
            cellProperties = new XElement(OoxmlNs.W + "tcPr");
            cell.AddFirst(cellProperties);
        }

        XElement? verticalMerge = cellProperties.Element(OoxmlNs.W + "vMerge");
        if (verticalMerge is null)
        {
            verticalMerge = new XElement(OoxmlNs.W + "vMerge");
            cellProperties.Add(verticalMerge);
        }

        verticalMerge.SetAttributeValue(OoxmlNs.W + "val", value);
    }

    private static bool TryGetConsistentVisualColumnCount(XElement table, out int columnCount)
    {
        columnCount = 0;
        int? expectedColumnCount = null;
        foreach (XElement row in table.Elements(OoxmlNs.W + "tr"))
        {
            int rowColumnCount = ReadTableRowVisualColumnCount(row);
            if (rowColumnCount == 0)
            {
                return false;
            }

            if (expectedColumnCount is null)
            {
                expectedColumnCount = rowColumnCount;
                continue;
            }

            if (rowColumnCount != expectedColumnCount.Value)
            {
                return false;
            }
        }

        columnCount = expectedColumnCount ?? 0;
        return columnCount != 0;
    }

    private static int ReadTableRowVisualColumnCount(XElement row)
    {
        int visualColumnCount = ReadTableRowGridOffset(row, "gridBefore") + ReadTableRowGridOffset(row, "gridAfter");
        foreach (XElement cell in row.Elements(OoxmlNs.W + "tc"))
        {
            visualColumnCount += ReadTableCellColumnSpan(cell);
        }

        return visualColumnCount;
    }

    private static IEnumerable<TableCellGridSlot> EnumerateTableRowCells(XElement row)
    {
        int columnIndex = 1 + ReadTableRowGridOffset(row, "gridBefore");
        foreach (XElement cell in row.Elements(OoxmlNs.W + "tc"))
        {
            int columnSpan = ReadTableCellColumnSpan(cell);
            yield return new TableCellGridSlot(cell, columnIndex, columnSpan);
            columnIndex += columnSpan;
        }
    }

    private static bool IsRectangular(XElement table, out int columnCount)
    {
        XElement[] rows = table.Elements(OoxmlNs.W + "tr").ToArray();
        columnCount = 0;
        if (rows.Length == 0)
        {
            return false;
        }

        int? expectedColumnCount = null;
        foreach (XElement row in rows)
        {
            if (ReadTableRowGridOffset(row, "gridBefore") != 0 ||
                ReadTableRowGridOffset(row, "gridAfter") != 0)
            {
                return false;
            }

            int visualColumnCount = 0;
            foreach (XElement cell in row.Elements(OoxmlNs.W + "tc"))
            {
                int columnSpan = ReadTableCellColumnSpan(cell);
                if (columnSpan != 1 || ReadTableCellVerticalMerge(cell) is not null)
                {
                    return false;
                }

                visualColumnCount += columnSpan;
            }

            if (visualColumnCount == 0)
            {
                return false;
            }

            if (expectedColumnCount is null)
            {
                expectedColumnCount = visualColumnCount;
                continue;
            }

            if (visualColumnCount != expectedColumnCount.Value)
            {
                return false;
            }
        }

        columnCount = expectedColumnCount ?? 0;
        return columnCount != 0;
    }

    private static int ReadTableRowGridOffset(XElement row, string localName)
    {
        string? value = (string?)row
            .Element(OoxmlNs.W + "trPr")
            ?.Element(OoxmlNs.W + localName)
            ?.Attribute(OoxmlNs.W + "val");
        return int.TryParse(value, out int parsed) && parsed > 0 ? parsed : 0;
    }

    private static int ReadTableCellColumnSpan(XElement cell)
    {
        string? spanText = (string?)cell
            .Element(OoxmlNs.W + "tcPr")
            ?.Element(OoxmlNs.W + "gridSpan")
            ?.Attribute(OoxmlNs.W + "val");
        return int.TryParse(spanText, out int span) && span > 0 ? span : 1;
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

    private static string? ReadCellShadingFill(XElement cell)
    {
        string? fill = (string?)cell
            .Element(OoxmlNs.W + "tcPr")
            ?.Element(OoxmlNs.W + "shd")
            ?.Attribute(OoxmlNs.W + "fill");
        return fill is null
            ? null
            : NormalizeCellShadingFill(fill);
    }

    private static bool TableMetadataEquals(string? current, string expected)
    {
        return string.Equals(current ?? string.Empty, expected, StringComparison.Ordinal);
    }

    private static bool CellShadingFillMatches(string? current, string expected)
    {
        return string.Equals(current ?? "none", expected, StringComparison.Ordinal);
    }

    private static bool TryNormalizeCellShadingFill(string value, bool allowNone, [NotNullWhen(true)] out string? normalized)
    {
        normalized = null;
        if (allowNone && string.Equals(value, "none", StringComparison.OrdinalIgnoreCase))
        {
            normalized = "none";
            return true;
        }

        if (string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase))
        {
            normalized = "auto";
            return true;
        }

        if (value.Length == 6 && value.All(IsAsciiHexDigit))
        {
            normalized = value.ToUpperInvariant();
            return true;
        }

        return false;
    }

    private static string NormalizeCellShadingFill(string value)
    {
        return value.Length == 6 && value.All(IsAsciiHexDigit)
            ? value.ToUpperInvariant()
            : value;
    }

    private static bool IsAsciiHexDigit(char value)
    {
        return value is >= '0' and <= '9' or >= 'A' and <= 'F' or >= 'a' and <= 'f';
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

    private static void SetCellShading(XElement cell, string? fill)
    {
        XElement? cellProperties = cell.Element(OoxmlNs.W + "tcPr");
        XElement? shading = cellProperties?.Element(OoxmlNs.W + "shd");
        if (fill is null)
        {
            shading?.Remove();
            return;
        }

        if (cellProperties is null)
        {
            cellProperties = new XElement(OoxmlNs.W + "tcPr");
            cell.AddFirst(cellProperties);
        }

        if (shading is null)
        {
            shading = new XElement(OoxmlNs.W + "shd");
            cellProperties.Add(shading);
        }

        if (shading.Attribute(OoxmlNs.W + "val") is null)
        {
            shading.SetAttributeValue(OoxmlNs.W + "val", "clear");
        }

        shading.SetAttributeValue(OoxmlNs.W + "fill", fill);
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

    private static void SetSectionColumns(XElement sectionProperties, int count)
    {
        XElement? columns = sectionProperties.Element(OoxmlNs.W + "cols");
        if (columns is null)
        {
            columns = new XElement(OoxmlNs.W + "cols");
            sectionProperties.Add(columns);
        }

        columns.SetAttributeValue(OoxmlNs.W + "num", count.ToString(System.Globalization.CultureInfo.InvariantCulture));
    }

    private static void SetSectionOrientation(XElement sectionProperties, string orientation)
    {
        XElement? pageSize = sectionProperties.Element(OoxmlNs.W + "pgSz");
        if (pageSize is null)
        {
            pageSize = new XElement(OoxmlNs.W + "pgSz");
            sectionProperties.AddFirst(pageSize);
        }

        string? currentOrientation = (string?)pageSize.Attribute(OoxmlNs.W + "orient") ?? "portrait";
        if (!string.Equals(currentOrientation, orientation, StringComparison.Ordinal) &&
            pageSize.Attribute(OoxmlNs.W + "w") is XAttribute width &&
            pageSize.Attribute(OoxmlNs.W + "h") is XAttribute height)
        {
            (width.Value, height.Value) = (height.Value, width.Value);
        }

        pageSize.SetAttributeValue(OoxmlNs.W + "orient", orientation);
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
        ResolvedOoxmlRelationship? relationship = package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .FirstOrDefault(relationship => relationship.Type == OoxmlRelTypes.Settings);
        if (relationship is not null)
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
            package.AddPart(settingsPartName, SettingsContentType, bytes, cancellationToken);
        }

        string relationshipId = OoxmlIds.AllocateRelationshipId(package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .Select(relationship => relationship.Id));
        package.AddRelationship(package.MainDocumentPartName, relationshipId, OoxmlRelTypes.Settings, GetRelativeRelationshipTarget(package.MainDocumentPartName, settingsPartName), targetMode: null, cancellationToken);
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

            DocxPackageValidator.ValidateRevisionMarkup(document, part.Name, diagnostics);
            return;
        }

        if (string.Equals(part.ContentType, "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml", StringComparison.OrdinalIgnoreCase))
        {
            RequireRoot(part, root, OoxmlNs.W + "hdr", diagnostics);
            DocxPackageValidator.ValidateRevisionMarkup(document, part.Name, diagnostics);
            return;
        }

        if (string.Equals(part.ContentType, "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml", StringComparison.OrdinalIgnoreCase))
        {
            RequireRoot(part, root, OoxmlNs.W + "ftr", diagnostics);
            DocxPackageValidator.ValidateRevisionMarkup(document, part.Name, diagnostics);
            return;
        }

        if (string.Equals(part.ContentType, CommentsContentType, StringComparison.OrdinalIgnoreCase))
        {
            RequireRoot(part, root, OoxmlNs.W + "comments", diagnostics);
            DocxPackageValidator.ValidateRevisionMarkup(document, part.Name, diagnostics);
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
        XDocument document = LoadDocumentPart(package, package.MainDocumentPartName, cancellationToken, out XElement root);
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
        SaveDocumentPart(package, package.MainDocumentPartName, document);
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
        DocxPatchOperation operation)
    {
        return Diagnostic(severity, code, message, operation, targetId: null);
    }

    private static DocxDiagnostic Diagnostic(
        DocxSeverity severity,
        string code,
        string message,
        DocxPatchOperation operation,
        string? targetId)
    {
        return Diagnostic(severity, code, message, operation, targetId, feature: null, fallback: null);
    }

    private static DocxDiagnostic Diagnostic(
        DocxSeverity severity,
        string code,
        string message,
        DocxPatchOperation operation,
        string? targetId,
        string? feature,
        string? fallback)
    {
        return new DocxDiagnostic(
            severity,
            code,
            message,
            TargetId: targetId,
            Feature: feature,
            Fallback: fallback,
            OperationIndex: operation.Index);
    }
}

internal sealed record PatchExecutionResult(
    bool Success,
    IReadOnlyList<DocxDiagnostic> Diagnostics,
    IReadOnlyList<DocxPatchOperationReport> Reports);

internal sealed record ContentControlTarget(string PartName, XDocument Document, XElement ContentControl);

internal sealed record ContentControlChoice(string DisplayText);

internal sealed record BookmarkTarget(string PartName, XDocument Document, XElement Start, XElement? End);

internal sealed record TrackedBookmarkReplacement(string DeletedText, XElement? RunProperties);

internal sealed record BookmarkTextSlot(XElement Container, XNode[] Nodes, XElement? AddAfter, XElement? AddBefore);

internal sealed record CommentsPartTarget(string PartName, XDocument Document, XElement Root);

internal sealed record CommentsExtendedPartTarget(string PartName, XDocument Document, XElement Root);

internal sealed record CommentsIdsPartTarget(string PartName, XDocument Document, XElement Root);

internal sealed record CommentTarget(string PartName, XDocument Document, XElement Comment);

internal sealed record CommentReplyTarget(string PartName, XDocument Document, XElement Comment, string ParaId, string ParentParaId);

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

internal sealed record CellTarget(string PartName, XDocument Document, XElement Table, XElement Row, XElement Cell, int VisualColumnIndex);

internal sealed record TableOperationSnapshot(
    string TargetId,
    string? TableId,
    int? RowIndex,
    int? ColumnIndex,
    int RowCountBefore,
    int ColumnCount,
    int? CellCount,
    int? GridBefore,
    int? GridAfter,
    IReadOnlyList<TableCellSnapshot> Cells);

internal sealed record TableCellSnapshot(
    int ColumnIndex,
    int VisualColumnEndIndex,
    string? MergeGroupId,
    string? NestedTablePath);

internal sealed record TableCellGridSlot(XElement Cell, int ColumnIndex, int ColumnSpan);

internal sealed record SectionTarget(XDocument Document, XElement SectionProperties);

internal sealed record MergeGroupRootState(XElement Row, XElement Cell, int VisualColumnIndex);

internal readonly record struct TextRange(int Start, int Length);

internal sealed record TextPosition(XElement TextElement, int Offset);

/// <summary>
/// Parsed form of a patch <c>target</c> field. Implementors describe <i>which</i>
/// paragraphs a target can resolve to; resolution itself lives in
/// <c>ResolveMainParagraphElementBySelector</c> and the explicit-ID parsers.
/// </summary>
/// <remarks>
/// <para>
/// <c>Raw</c> is always the original target text and is used verbatim in
/// diagnostics (<c>E1201</c> no match, <c>E1202</c> ambiguous match,
/// <c>E1203</c> malformed selector), so it must survive parsing unchanged.
/// </para>
/// <para>
/// Predicate selectors (<c>Heading</c>, <c>ParagraphText</c>, <c>Bookmark</c>,
/// <c>ContentControl</c>) only ever match direct <c>w:p</c> children of the
/// <i>main</i> document story and require exactly one match: zero matches fail
/// with <c>E1201</c> (with up to 3 suggestions), two or more fail with
/// <c>E1202</c> (listing the candidate IDs). All text/name comparisons are
/// ordinal and case-sensitive. Add a new subtype only together with its branch
/// in <c>TryParseTargetSelector</c> and in
/// <c>ResolveMainParagraphElementBySelector</c>; the compiler does not enforce
/// the pairing, so cover it with a parser/resolver round-trip test.
/// </para>
/// </remarks>
/// <param name="Raw">Original target text, preserved for diagnostics.</param>
internal abstract record TargetSelector(string Raw);

/// <summary>
/// Catch-all for targets without a <c>prefix:</c> selector: a stable explicit ID
/// (<c>M.P0001</c>, <c>H001.P0002</c>, table/cell/image/section/comment IDs, …).
/// Never handled by the predicate path — callers dispatch on ID shape instead
/// (<c>TryParseMainParagraphTarget</c>, <c>TryParseStoryTableTarget</c>, …).
/// </summary>
internal sealed record ExplicitIdTargetSelector(string Raw) : TargetSelector(Raw);

/// <summary>
/// <c>heading:"Text"</c> or <c>heading:L:"Text"</c>: matches a main-story paragraph
/// whose heading level equals <c>Level</c> (any heading level when null) and whose
/// visible text equals <c>Text</c> ordinally. <c>Level</c> is validated to 1–9 at
/// parse time (<c>E1203</c> otherwise).
/// </summary>
internal sealed record HeadingTargetSelector(string Raw, int? Level, string Text) : TargetSelector(Raw);

/// <summary>
/// <c>text:"..."</c>: matches a main-story paragraph whose visible text <i>contains</i>
/// <c>Text</c> as an ordinal substring. Prefer an explicit ID when the text is not unique.
/// </summary>
internal sealed record ParagraphTextTargetSelector(string Raw, string Text) : TargetSelector(Raw);

/// <summary>
/// <c>bookmark:"Name"</c>: matches a main-story paragraph containing a descendant
/// <c>bookmarkStart</c> whose name equals <c>Name</c> ordinally.
/// </summary>
internal sealed record BookmarkTargetSelector(string Raw, string Name) : TargetSelector(Raw);

/// <summary>
/// <c>content-control:"TagOrAlias"</c>: matches a main-story paragraph containing a
/// descendant <c>sdt</c> whose tag or alias equals <c>Name</c> ordinally.
/// </summary>
internal sealed record ContentControlTargetSelector(string Raw, string Name) : TargetSelector(Raw);

internal sealed record ParagraphSelectorMatch(string Id, XElement Paragraph);
