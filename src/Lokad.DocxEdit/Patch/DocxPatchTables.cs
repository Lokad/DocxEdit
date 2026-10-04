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
            "set-cell" or "set-cell-shading" or "replace-text" => CaptureCellSnapshot(package, target, cancellationToken),
            "append-row" => CaptureTableSnapshot(package, target, cancellationToken),
            "insert-row-before" or "insert-row-after" => CaptureInsertAnchorSnapshot(package, target, cancellationToken),
            "delete-row" or "set-row-header" => CaptureRowSnapshot(package, target, cancellationToken),
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

        if (!DocxTargetId.TryParse(target, out DocxTargetId cellId) ||
            cellId.Kind is not (DocxTargetKind.Cell or DocxTargetKind.MergeGroup))
        {
            return null;
        }

        XElement[] cells = cellTarget.Row.Elements(OoxmlNs.W + "tc").ToArray();
        (int historicalRow, int historicalColumn) = HistoricalCellIndex(cellId, cellTarget);
        return CreateTableOperationSnapshot(cellId, cellTarget.Table, historicalRow, historicalColumn, cells.Length, cellTarget.Row);
    }

    private static TableOperationSnapshot? CaptureRowSnapshot(OoxmlPackage package, string target, CancellationToken cancellationToken)
    {
        if (IsAliasReference(target))
        {
            return CreateAliasRowSnapshot(package, target, cancellationToken);
        }

        RowTarget? rowTarget = ResolveRowTarget(package, target, cancellationToken);
        if (rowTarget is null)
        {
            return null;
        }

        if (!DocxTargetId.TryParse(target, out DocxTargetId rowId) || rowId.Kind != DocxTargetKind.Row)
        {
            return null;
        }

        int cellCount = rowTarget.Row.Elements(OoxmlNs.W + "tc").Count();
        return CreateTableOperationSnapshot(rowId, rowTarget.Table, rowId.Secondary, columnIndex: null, cellCount, rowTarget.Row);
    }

    // Insertion anchors report in operation-time space: the new row lands at
    // the anchor live position, so the anchor index and the table identity
    // are both measured live. Explicit and alias anchors share this path.
    private static TableOperationSnapshot? CaptureInsertAnchorSnapshot(OoxmlPackage package, string target, CancellationToken cancellationToken)
    {
        RowTarget? rowTarget = ResolveRowTarget(package, target, cancellationToken);
        return rowTarget is null ? null : CaptureLiveRowSnapshot(package, rowTarget, cancellationToken);
    }

    // Alias rows carry their own provenance. Created rows report in
    // operation-time space with their creation mark; snapshot-only rows fall
    // back to their historical identity.
    private static TableOperationSnapshot? CreateAliasRowSnapshot(OoxmlPackage package, string target, CancellationToken cancellationToken)
    {
        RowTarget? rowTarget = ResolveAliasRowTarget(package, target, cancellationToken);
        if (rowTarget is null)
        {
            return null;
        }

        if (rowTarget.Row.Attribute(SnapshotCreatedName) is null)
        {
            string? snapshotId = (string?)rowTarget.Row.Attribute(SnapshotIdName);
            if (snapshotId is null || !DocxTargetId.TryParse(snapshotId, out DocxTargetId historical) || historical.Kind != DocxTargetKind.Row)
            {
                return null;
            }

            int cellCount = rowTarget.Row.Elements(OoxmlNs.W + "tc").Count();
            return CreateTableOperationSnapshot(historical, rowTarget.Table, historical.Secondary, columnIndex: null, cellCount, rowTarget.Row);
        }

        return CaptureLiveRowSnapshot(package, rowTarget, cancellationToken);
    }

    private static TableOperationSnapshot? CaptureLiveRowSnapshot(OoxmlPackage package, RowTarget rowTarget, CancellationToken cancellationToken)
    {
        DocxTargetId? liveTable = LiveTableId(package, rowTarget.PartName, rowTarget.Document, rowTarget.Table, cancellationToken);
        if (liveTable is null)
        {
            return null;
        }

        XElement[] rows = rowTarget.Table.Elements(OoxmlNs.W + "tr").ToArray();
        int liveRowIndex = Array.IndexOf(rows, rowTarget.Row) + 1;
        if (liveRowIndex < 1)
        {
            return null;
        }

        int cellCount = rowTarget.Row.Elements(OoxmlNs.W + "tc").Count();
        return CreateTableOperationSnapshot(
            liveTable.Value with { Kind = DocxTargetKind.Row, Secondary = liveRowIndex },
            rowTarget.Table,
            liveRowIndex,
            columnIndex: null,
            cellCount,
            rowTarget.Row,
            inputCoordinates: false) with
        {
            CreationMark = (string?)rowTarget.Row.Attribute(SnapshotCreatedName)
        };
    }

    private static DocxTargetId? LiveTableId(OoxmlPackage package, string partName, XDocument document, XElement table, CancellationToken cancellationToken)
    {
        if (!DocxPartRoles.GetStoryPrefixes(package, cancellationToken).TryGetValue(partName, out string? prefix))
        {
            return null;
        }

        int? liveOrdinal = PhysicalTableOrdinal(document, table);
        if (liveOrdinal is null)
        {
            return null;
        }

        (char story, int storyPart) = DocxTargetId.ParseStoryPrefix(prefix);
        return new DocxTargetId(story, storyPart, DocxTargetKind.Table, liveOrdinal.Value, 0, 0);
    }

    private static int? PhysicalTableOrdinal(XDocument document, XElement table)
    {
        XElement? root = document.Root;
        if (root is null)
        {
            return null;
        }

        XElement container = root.Element(OoxmlNs.W + "body") ?? root;
        int ordinal = 0;
        foreach (DocxStoryBlocks.StoryBlock entry in DocxStoryBlocks.EnumeratePhysicalBlocks(container))
        {
            if (entry.Block.Name != OoxmlNs.W + "tbl")
            {
                continue;
            }

            ordinal++;
            if (ReferenceEquals(entry.Block, table))
            {
                return ordinal;
            }
        }

        return null;
    }

    // Input-space indexes for explicitly addressed cells: the parsed cell ID names
    // the historical row and visual column. Merge-group targets resolve to their
    // root cell, whose row carries the historical row mark.
    private static (int rowIndex, int columnIndex) HistoricalCellIndex(DocxTargetId cellId, CellTarget cellTarget)
    {
        if (cellId.Kind == DocxTargetKind.Cell)
        {
            return (cellId.Secondary, cellId.Tertiary);
        }

        string? rootSnapshot = (string?)cellTarget.Row.Attribute(SnapshotIdName);
        if (rootSnapshot is not null && DocxTargetId.TryParse(rootSnapshot, out DocxTargetId rootId) && rootId.Kind == DocxTargetKind.Row)
        {
            return (rootId.Secondary, cellTarget.VisualColumnIndex);
        }

        XElement[] liveRows = cellTarget.Table.Elements(OoxmlNs.W + "tr").ToArray();
        return (Array.IndexOf(liveRows, cellTarget.Row) + 1, cellTarget.VisualColumnIndex);
    }

    private static TableOperationSnapshot? CaptureTableSnapshot(OoxmlPackage package, string target, CancellationToken cancellationToken)
    {
        TableTarget? tableTarget = ResolveTableTarget(package, target, cancellationToken);
        if (tableTarget is null)
        {
            return null;
        }

        if (!DocxTargetId.TryParse(target, out DocxTargetId tableId) || tableId.Kind != DocxTargetKind.Table)
        {
            return null;
        }

        // Appended rows report in operation-time space, so the table identity is
        // measured live: earlier table deletions can renumber it.
        DocxTargetId liveTable = LiveTableId(package, tableTarget.PartName, tableTarget.Document, tableTarget.Table, cancellationToken) ?? tableId;

        XElement? templateRow = tableTarget.Table.Elements(OoxmlNs.W + "tr").LastOrDefault();
        return CreateTableOperationSnapshot(liveTable, tableTarget.Table, rowIndex: null, columnIndex: null, cellCount: null, templateRow, inputCoordinates: false);
    }

    private static TableOperationSnapshot CreateTableOperationSnapshot(
        DocxTargetId resolvedId,
        XElement table,
        int? rowIndex,
        int? columnIndex,
        int? cellCount,
        XElement? row,
        bool inputCoordinates = true)
    {
        // Counts and grid offsets always describe live execution state. Identity
        // and row/column indexes arrive in report-space provenance from the
        // caller: historical for explicit targets, live for created anchors.
        int rowCount = table.Elements(OoxmlNs.W + "tr").Count();
        int columnCount = TryGetConsistentVisualColumnCount(table, out int visualColumnCount)
            ? visualColumnCount
            : table.Elements(OoxmlNs.W + "tr").Select(ReadTableRowVisualColumnCount).DefaultIfEmpty(0).Max();
        int? gridBefore = row is null ? null : DocxTableGrid.ReadGridOffset(row, "gridBefore");
        int? gridAfter = row is null ? null : DocxTableGrid.ReadGridOffset(row, "gridAfter");
        IReadOnlyList<TableCellSnapshot> cells = row is null
            ? []
            : CreateTableCellSnapshots(resolvedId.TableId, table, row, rowIndex, inputCoordinates);
        return new TableOperationSnapshot(resolvedId, rowIndex, columnIndex, rowCount, columnCount, cellCount, gridBefore, gridAfter, cells);
    }

    private static IReadOnlyList<TableCellSnapshot> CreateTableCellSnapshots(
        DocxTargetId tableId,
        XElement table,
        XElement targetRow,
        int? targetRowIndex,
        bool inputCoordinates)
    {
        var snapshots = new List<TableCellSnapshot>();
        foreach (TableMergeCell entry in EnumerateTableMergeCells(table))
        {
            if (!ReferenceEquals(entry.Row, targetRow))
            {
                continue;
            }

            DocxTargetId? mergeGroupId = null;
            if (inputCoordinates)
            {
                string? historical = (string?)entry.Cell.Attribute(SnapshotMergeGroupName)
                    ?? (string?)entry.Cell.Attribute(SnapshotMergeReferenceName);
                if (DocxTargetId.TryParse(historical, out DocxTargetId inputMerge))
                {
                    mergeGroupId = inputMerge;
                }
            }
            else if (entry.Group is { } group)
            {
                mergeGroupId = tableId with { Kind = DocxTargetKind.MergeGroup, Secondary = group.Ordinal };
            }

            DocxTargetId cellId = tableId with { Kind = DocxTargetKind.Cell, Secondary = targetRowIndex ?? entry.RowOrdinal, Tertiary = entry.VisualColumn };
            snapshots.Add(new TableCellSnapshot(
                entry.VisualColumn,
                entry.VisualColumn + DocxTableGrid.ReadColumnSpan(entry.Cell) - 1,
                mergeGroupId,
                CreateNestedTablePath(cellId.ToWireValue(), entry.Cell)));
        }

        return snapshots;
    }

    private static IReadOnlyList<TableCellSnapshot> CreateFallbackCellSnapshots(int cellCount)
    {
        var snapshots = new List<TableCellSnapshot>(capacity: Math.Max(cellCount, 0));
        for (int column = 1; column <= cellCount; column++)
        {
            snapshots.Add(new TableCellSnapshot(column, column, MergeGroupId: null, NestedTablePath: null));
        }

        return snapshots;
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

    private static IReadOnlyList<DocxPatchAffectedTarget> BuildAffectedTargets(DocxPatchOperation operation, OoxmlPackage package, TableOperationSnapshot? before, DocxTargetId? resolved, string? resolvedMark, CancellationToken cancellationToken)
    {
        if (before is not null)
        {
            return operation.OperationName switch
            {
                "set-cell" or "set-cell-shading" => BuildSetCellAffectedTargets(before),
                "replace-text" => BuildSetCellAffectedTargets(before),
                "append-row" => BuildInsertedRowAffectedTargets(operation, package, before, "append", cancellationToken),
                "insert-row-before" or "insert-row-after" => BuildInsertedRowAffectedTargets(operation, package, before, "insert", cancellationToken),
                "delete-row" => BuildDeletedRowAffectedTargets(before),
                "set-row-header" => BuildRowUpdateAffectedTargets(before),
                _ => []
            };
        }

        if (resolved is not { } resolvedId)
        {
            return [];
        }

        string action = operation.OperationName switch
        {
            "insert-before" or "insert-after" or "insert-equation" => "insert",
            "delete-block" or "delete-image" or "delete-bookmark" or "delete-comment" or "delete-comment-reply" or "remove-hyperlink" or "delete-row" or "delete-equation" => "delete",
            _ => "update"
        };
        string kind = TargetKindWord(resolvedId.Kind);
        if (resolvedMark is not null)
        {
            return [new DocxPatchAffectedTarget(resolvedId, kind, action) { CreationMark = resolvedMark, Coordinate = "operation-time", FinalId = resolvedId }];
        }
        return [new DocxPatchAffectedTarget(resolvedId, kind, action) { Coordinate = ResolveTargetCoordinate(operation), FinalId = resolvedId }];
    }

    // Snapshot identities are historical input coordinates. A resolved target
    // without a creation mark always names an input snapshot object, even when
    // the selector was semantic. Created objects always carry a creation mark
    // and use the operation time path above, so this path is always input.
    private static string ResolveTargetCoordinate(DocxPatchOperation operation)
    {
        _ = operation;
        return "input";
    }

    private static IReadOnlyList<DocxPatchAffectedTarget> BuildSetCellAffectedTargets(TableOperationSnapshot before)
    {
        TableCellSnapshot? cell = before.Cells.FirstOrDefault(cell =>
            before.ColumnIndex is not null &&
            before.ColumnIndex.Value >= cell.ColumnIndex &&
            before.ColumnIndex.Value <= cell.VisualColumnEndIndex);
        return
        [
            new(before.ResolvedTarget, "cell", "update")
            {
                ParentId = before.ResolvedTarget.TableId,
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
        DocxPatchOperation operation,
        OoxmlPackage package,
        TableOperationSnapshot before,
        string action,
        CancellationToken cancellationToken)
    {
        var found = FindMarkedStoryElement(package, SnapshotCreatedName, CreatedMarkValue(operation, 0), OoxmlNs.W + "tr", cancellationToken)
            ?? throw new InvalidDataException("Inserted row is missing its creation mark.");
        XElement table = found.Element.Parent
            ?? throw new InvalidDataException("Inserted row is missing its table.");
        var rowTarget = new RowTarget(found.Story.PartName, found.Document, table, found.Element);
        TableOperationSnapshot created = CaptureLiveRowSnapshot(package, rowTarget, cancellationToken)
            ?? throw new InvalidDataException("Inserted row has no live coordinate.");
        DocxTargetId rowId = created.ResolvedTarget;
        var affected = new List<DocxPatchAffectedTarget>
        {
            new(rowId, "row", action)
            {
                Coordinate = "operation-time",
                CreationMark = created.CreationMark,
                ParentId = rowId.TableId,
                RowIndex = created.RowIndex,
                RowCountBefore = before.RowCountBefore,
                RowCountAfter = before.RowCountBefore + 1,
                ColumnCount = created.ColumnCount,
                CellCount = created.CellCount,
                GridBefore = created.GridBefore,
                GridAfter = created.GridAfter
            }
        };
        foreach (TableCellSnapshot cell in created.Cells)
        {
            affected.Add(new DocxPatchAffectedTarget(rowId with { Kind = DocxTargetKind.Cell, Tertiary = cell.ColumnIndex }, "cell", action)
            {
                Coordinate = "operation-time",
                CreationMark = created.CreationMark,
                ParentId = rowId,
                RowIndex = created.RowIndex,
                ColumnIndex = cell.ColumnIndex,
                VisualColumnEndIndex = cell.VisualColumnEndIndex,
                MergeGroupId = cell.MergeGroupId,
                NestedTablePath = cell.NestedTablePath,
                RowCountBefore = before.RowCountBefore,
                RowCountAfter = before.RowCountBefore + 1,
                ColumnCount = created.ColumnCount
            });
        }

        return affected;
    }

    private static IReadOnlyList<DocxPatchAffectedTarget> BuildDeletedRowAffectedTargets(TableOperationSnapshot before)
    {
        int rowIndex = before.RowIndex ?? 1;
        IReadOnlyList<TableCellSnapshot> cells = before.Cells.Count == 0
            ? CreateFallbackCellSnapshots(before.CellCount ?? before.ColumnCount)
            : before.Cells;
        int cellCount = before.CellCount ?? cells.Count;
        var affected = new List<DocxPatchAffectedTarget>
        {
            new(before.ResolvedTarget, "row", "delete")
            {
                CreationMark = before.CreationMark,
                Coordinate = before.CreationMark is null ? "input" : "operation-time",
                ParentId = before.ResolvedTarget.TableId,
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
            affected.Add(new DocxPatchAffectedTarget(before.ResolvedTarget with { Kind = DocxTargetKind.Cell, Tertiary = cell.ColumnIndex }, "cell", "delete")
            {
                CreationMark = before.CreationMark,
                Coordinate = before.CreationMark is null ? "input" : "operation-time",
                ParentId = before.ResolvedTarget,
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

    private static IReadOnlyList<DocxPatchAffectedTarget> BuildRowUpdateAffectedTargets(TableOperationSnapshot before)
    {
        return
        [
            new(before.ResolvedTarget, "row", "update")
            {
                CreationMark = before.CreationMark,
                Coordinate = before.CreationMark is null ? "input" : "operation-time",
                FinalId = before.ResolvedTarget,
                ParentId = before.ResolvedTarget.TableId,
                RowIndex = before.RowIndex,
                RowCountBefore = before.RowCountBefore,
                RowCountAfter = before.RowCountBefore,
                ColumnCount = before.ColumnCount
            }
        ];
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
                    target, fieldName: "expect-text")
            ];
        }

        if (string.Equals(text, current, StringComparison.Ordinal) && IsSimpleEditableCell(cellTarget.Cell))
        {
            return NoOpResult(operation, target, "Set-cell for " + target + " leaves the cell unchanged; nothing was written and no revisions were generated.");
        }
        if (TryGetStoryEnteringProtectedFeature(cellTarget.Cell, out string cellProtected))
        {
            if (options.TrackChanges == TrackChangesMode.Require)
            {
                TrackUnsupportedShape(options, operation, target, "paragraph contains protected OOXML boundary " + Quote(cellProtected), diagnostics);
                return diagnostics;
            }
            return [Diagnostic(DocxSeverity.Error, "E4305", "Set-cell for " + target + " crosses protected OOXML boundary " + Quote(cellProtected) + ".", operation, target)];
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
                if (!TryFallbackToDirectEdit(options, operation, target, "tracked set-cell does not support force true replacement", diagnostics, ref useTrackedChanges))
                {
                    return diagnostics;
                }
            }
            else if (!TryGetTrackedSetCellParagraphs(cellTarget.Cell, text, out trackedParagraphs, out string? trackedUnsupportedReason))
            {
                if (!TryFallbackToDirectEdit(options, operation, target, trackedUnsupportedReason, diagnostics, ref useTrackedChanges))
                {
                    return diagnostics;
                }
            }
        }

        if (!useTrackedChanges && !force && !IsSimpleEditableCell(cellTarget.Cell))
        {
            return [Diagnostic(DocxSeverity.Error, "E4302", $"Cell '{target}' contains unsupported content. Use force true only when replacing all cell content is intended.", operation, target)];
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
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected cell shading fill '{normalizedExpectedFill}', found '{currentFill ?? "none"}'.", operation, target, fieldName: "expect-fill")];
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
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected table style '{expectedStyle}', found '{currentStyle ?? "none"}'.", operation, target, fieldName: "expect-style")];
        }

        if (!TryResolveStyleId(package, style, "table", cancellationToken, out string? styleId, out DocxDiagnostic? styleDiagnostic, operation, target))
        {
            return [styleDiagnostic];
        }


        if (IsTrackedMode(options))
        {
            SetTableStyleWithTrackedChange(package, tableTarget.Table, styleId, options, generatedRevisionIds, cancellationToken);
        }
        else
        {
            SetTableStyle(tableTarget.Table, styleId);
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
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected table caption '{expectedCaption}', found '{currentCaption ?? "none"}'.", operation, target, fieldName: "expect-caption")];
        }

        string? currentDescription = ReadTableTextProperty(tableTarget.Table, "tblDescription");
        if (expectedDescription is not null && !TableMetadataEquals(currentDescription, expectedDescription))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected table description '{expectedDescription}', found '{currentDescription ?? "none"}'.", operation, target, fieldName: "expect-description")];
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
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected row header '{expectedHeader.Value.ToString().ToLowerInvariant()}', found '{currentHeader.ToString().ToLowerInvariant()}'.", operation, target, fieldName: "expect-header")];
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


        XElement newRow = CreateRowFromTemplate(templateRow, cellTexts);
        newRow.SetAttributeValue(SnapshotCreatedName, CreatedMarkValue(operation, 0));
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
                return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' does not have a consistent visual grid and cannot be edited safely without force true. With force true, the new row clones the target row raw cell count, which may misalign merged or spanned columns.", operation, target)];
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


        XElement newRow = CreateRowFromTemplate(rowTarget.Row, cellTexts);
        newRow.SetAttributeValue(SnapshotCreatedName, CreatedMarkValue(operation, 0));
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
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected row text to contain '{expectedContains}'.", operation, target, fieldName: "expect-contains")];
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
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' does not have a consistent visual grid and cannot be edited safely without force true. With force true, the row is deleted anyway, which may leave the remaining grid ragged.", operation, target)];
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


        if (useTrackedChanges)
        {
            MarkRowRevision(package, rowTarget.Row, OoxmlNs.W + "del", options, generatedRevisionIds, cancellationToken);
        }
        else
        {
            if (FindOrphanedRangeBoundary(rowTarget.Document, rowTarget.Row) is { } orphanedRowRange)
            {
                return [Diagnostic(DocxSeverity.Error, "E4305", $"Delete for {target} would orphan {orphanedRowRange} outside the deleted element. Delete the range first or choose another target.", operation, target)];
            }

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
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected {expectedRowCount} row(s), found {actualRowCount}.", operation, target, fieldName: "expect-row-count"));
        }

        if (expectedColumnCount is not null)
        {
            if (!TryGetConsistentVisualColumnCount(table, out int actualColumnCount))
            {
                diagnostics.Add(Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected {expectedColumnCount} column(s), but table does not have a consistent visual grid.", operation, target, fieldName: "expect-column-count"));
            }
            else if (actualColumnCount != expectedColumnCount)
            {
                diagnostics.Add(Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected {expectedColumnCount} column(s), found {actualColumnCount}.", operation, target, fieldName: "expect-column-count"));
            }
        }

        if (expectedCellCount is not null)
        {
            if (row is null)
            {
                diagnostics.Add(Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Field 'expect-cell-count' requires a row or cell target.", operation, target, fieldName: "expect-cell-count"));
            }
            else
            {
                int actualCellCount = row.Elements(OoxmlNs.W + "tc").Count();
                if (actualCellCount != expectedCellCount)
                {
                    diagnostics.Add(Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected {expectedCellCount} cell(s), found {actualCellCount}.", operation, target, fieldName: "expect-cell-count"));
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

        static bool RowHasTrackedRowRevision(XElement row)
        {
            XElement? rowProperties = row.Element(OoxmlNs.W + "trPr");
            return rowProperties is not null &&
                (rowProperties.Elements(OoxmlNs.W + "ins").Any() ||
                    rowProperties.Elements(OoxmlNs.W + "del").Any() ||
                    rowProperties.Elements(OoxmlNs.W + "trPrChange").Any());
        }
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
            if (ReadTableCellVerticalMerge(deletedSlot.Cell) != DocxVerticalMerge.Restart)
            {
                continue;
            }

            XElement? nextCell = FindCellByVisualColumn(nextRow, deletedSlot.ColumnIndex);
            if (nextCell is null || ReadTableCellVerticalMerge(nextCell) != DocxVerticalMerge.Continue)
            {
                continue;
            }

            int nextCellSpan = DocxTableGrid.ReadColumnSpan(nextCell);
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
            .Any(cell => ReadTableCellVerticalMerge(cell) == DocxVerticalMerge.Continue);
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
            if (ReadTableCellVerticalMerge(deletedSlot.Cell) != DocxVerticalMerge.Restart)
            {
                continue;
            }

            XElement? nextCell = FindCellByVisualColumn(nextRow, deletedSlot.ColumnIndex);
            if (nextCell is not null && ReadTableCellVerticalMerge(nextCell) == DocxVerticalMerge.Continue)
            {
                SetTableCellVerticalMerge(nextCell, DocxVerticalMerge.Restart);
            }
        }
    }

    private static void SetTableCellVerticalMerge(XElement cell, DocxVerticalMerge value)
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

        verticalMerge.SetAttributeValue(OoxmlNs.W + "val", value.ToWireValue());
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
        int visualColumnCount = DocxTableGrid.ReadGridOffset(row, "gridBefore") + DocxTableGrid.ReadGridOffset(row, "gridAfter");
        foreach (XElement cell in row.Elements(OoxmlNs.W + "tc"))
        {
            visualColumnCount += DocxTableGrid.ReadColumnSpan(cell);
        }

        return visualColumnCount;
    }

    private static IEnumerable<TableCellGridSlot> EnumerateTableRowCells(XElement row)
    {
        int columnIndex = 1 + DocxTableGrid.ReadGridOffset(row, "gridBefore");
        foreach (XElement cell in row.Elements(OoxmlNs.W + "tc"))
        {
            int columnSpan = DocxTableGrid.ReadColumnSpan(cell);
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
            if (DocxTableGrid.ReadGridOffset(row, "gridBefore") != 0 ||
                DocxTableGrid.ReadGridOffset(row, "gridAfter") != 0)
            {
                return false;
            }

            int visualColumnCount = 0;
            foreach (XElement cell in row.Elements(OoxmlNs.W + "tc"))
            {
                int columnSpan = DocxTableGrid.ReadColumnSpan(cell);
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

    

    

    private static DocxVerticalMerge? ReadTableCellVerticalMerge(XElement cell)
    {
        XElement? verticalMerge = cell
            .Element(OoxmlNs.W + "tcPr")
            ?.Element(OoxmlNs.W + "vMerge");
        if (verticalMerge is null)
        {
            return null;
        }

        string? value = (string?)verticalMerge.Attribute(OoxmlNs.W + "val");
        return string.Equals(value, "restart", StringComparison.Ordinal) ? DocxVerticalMerge.Restart : DocxVerticalMerge.Continue;
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
