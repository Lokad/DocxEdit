using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

internal static partial class DocxPatchEngine
{
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
}
