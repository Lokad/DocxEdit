using System.Globalization;
using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

internal static partial class DocxDocumentScanner
{
    private static DocxTableInfo ReadTable(
        XElement table,
        string id,
        string story,
        DocxTextView textView,
        OoxmlPackage package,
        IReadOnlyDictionary<string, OoxmlRelationship> relationships,
        List<DocxImageInfo> images,
        string imageIdPrefix,
        Dictionary<XElement, string> targets,
        ref int imageIndex)
    {
        var cells = new List<DocxTableCellInfo>();
        var rows = new List<DocxTableRowInfo>();
        XElement? tableProperties = table.Element(OoxmlNs.W + "tblPr");
        string? styleId = (string?)tableProperties
            ?.Element(OoxmlNs.W + "tblStyle")
            ?.Attribute(OoxmlNs.W + "val");
        int? gridColumnCount = table
            .Element(OoxmlNs.W + "tblGrid")
            ?.Elements(OoxmlNs.W + "gridCol")
            .Count();
        int rowIndex = 1;
        int mergeGroupIndex = 1;
        int maxColumns = 0;
        var activeVerticalMerges = new Dictionary<int, TableMergeState>();
        foreach (XElement row in table.Elements(OoxmlNs.W + "tr"))
        {
            bool rowInserted = IsTableRowInserted(row);
            bool rowDeleted = IsTableRowDeleted(row);
            if (textView == DocxTextView.Final && rowDeleted ||
                textView == DocxTextView.Original && rowInserted)
            {
                continue;
            }

            int gridBefore = ReadRowGridOffset(row, "gridBefore");
            int gridAfter = ReadRowGridOffset(row, "gridAfter");
            RemoveActiveVerticalMerges(activeVerticalMerges, 1, gridBefore);
            int columnIndex = 1 + gridBefore;
            int physicalColumnIndex = 1;
            string rowId = $"{id}.R{rowIndex:00}";
            foreach (XElement cell in row.Elements(OoxmlNs.W + "tc"))
            {
                int columnSpan = ReadCellColumnSpan(cell);
                string cellId = $"{id}.R{rowIndex:00}.C{columnIndex:00}";
                DocxVerticalMerge? verticalMerge = ReadCellVerticalMerge(cell);
                string? mergeGroupId = null;
                string? verticalMergeRootCellId = null;
                if (verticalMerge == DocxVerticalMerge.Restart)
                {
                    mergeGroupId = AllocateMergeGroupId(id, ref mergeGroupIndex);
                    verticalMergeRootCellId = cellId;
                    SetActiveVerticalMerge(activeVerticalMerges, columnIndex, columnSpan, new TableMergeState(mergeGroupId, cellId));
                }
                else if (verticalMerge is not null)
                {
                    TableMergeState? activeMerge = FindActiveVerticalMerge(activeVerticalMerges, columnIndex, columnSpan);
                    mergeGroupId = activeMerge?.MergeGroupId ?? AllocateMergeGroupId(id, ref mergeGroupIndex);
                    verticalMergeRootCellId = activeMerge?.RootCellId;
                    SetActiveVerticalMerge(activeVerticalMerges, columnIndex, columnSpan, new TableMergeState(mergeGroupId, verticalMergeRootCellId));
                }
                else
                {
                    RemoveActiveVerticalMerges(activeVerticalMerges, columnIndex, columnSpan);
                    if (columnSpan > 1)
                    {
                        mergeGroupId = AllocateMergeGroupId(id, ref mergeGroupIndex);
                    }
                }

                foreach (XElement drawing in cell.Descendants(OoxmlNs.W + "drawing"))
                {
                    AddDrawingImages(drawing, package, relationships, images, imageIdPrefix, cellId, ref imageIndex);
                }

                targets[cell] = cellId;
                string cellText = ReadText(cell, textView);
                if (textView == DocxTextView.Markup)
                {
                    if (rowInserted)
                    {
                        cellText = $"[+{cellText}+]";
                    }
                    else if (rowDeleted)
                    {
                        cellText = $"[-{cellText}-]";
                    }
                }

                cells.Add(new DocxTableCellInfo(
                    cellId,
                    rowIndex,
                    columnIndex,
                    cellText,
                    columnSpan,
                    verticalMerge,
                    cell.Elements(OoxmlNs.W + "tbl").Any())
                {
                    PhysicalColumnIndex = physicalColumnIndex,
                    VisualColumnEndIndex = columnIndex + columnSpan - 1,
                    MergeGroupId = mergeGroupId,
                    VerticalMergeRootCellId = verticalMergeRootCellId
                });
                columnIndex += columnSpan;
                physicalColumnIndex++;
            }

            RemoveActiveVerticalMerges(activeVerticalMerges, columnIndex, gridAfter);
            rows.Add(new DocxTableRowInfo
            {
                Id = rowId,
                RowIndex = rowIndex,
                CellCount = physicalColumnIndex - 1,
                GridBefore = gridBefore,
                GridAfter = gridAfter,
                IsHeader = ReadRowFlag(row, "tblHeader"),
                CantSplit = ReadRowFlag(row, "cantSplit")
            });
            maxColumns = Math.Max(maxColumns, columnIndex - 1 + gridAfter);
            rowIndex++;
        }

        return new DocxTableInfo(id, story, rowIndex - 1, Math.Max(maxColumns, gridColumnCount ?? 0), cells)
        {
            StyleId = styleId,
            Caption = (string?)tableProperties?.Element(OoxmlNs.W + "tblCaption")?.Attribute(OoxmlNs.W + "val"),
            Description = (string?)tableProperties?.Element(OoxmlNs.W + "tblDescription")?.Attribute(OoxmlNs.W + "val"),
            GridColumnCount = gridColumnCount,
            HasHeaderRow = rows.Any(row => row.IsHeader),
            HasMergedCells = cells.Any(cell => cell.ColumnSpan > 1 || cell.VerticalMerge is not null) ||
                rows.Any(row => row.GridBefore > 0 || row.GridAfter > 0),
            HasNestedTables = cells.Any(cell => cell.HasNestedTable),
            Rows = rows
        };
    }

    private static string AllocateMergeGroupId(string tableId, ref int mergeGroupIndex)
    {
        return $"{tableId}.MG{mergeGroupIndex++:0000}";
    }

    private static void SetActiveVerticalMerge(
        Dictionary<int, TableMergeState> activeVerticalMerges,
        int columnIndex,
        int columnSpan,
        TableMergeState state)
    {
        for (int column = columnIndex; column < columnIndex + columnSpan; column++)
        {
            activeVerticalMerges[column] = state;
        }
    }

    private static TableMergeState? FindActiveVerticalMerge(
        IReadOnlyDictionary<int, TableMergeState> activeVerticalMerges,
        int columnIndex,
        int columnSpan)
    {
        for (int column = columnIndex; column < columnIndex + columnSpan; column++)
        {
            if (activeVerticalMerges.TryGetValue(column, out TableMergeState? state))
            {
                return state;
            }
        }

        return null;
    }

    private static void RemoveActiveVerticalMerges(Dictionary<int, TableMergeState> activeVerticalMerges, int columnIndex, int columnSpan)
    {
        for (int column = columnIndex; column < columnIndex + columnSpan; column++)
        {
            activeVerticalMerges.Remove(column);
        }
    }

    private static int ReadRowGridOffset(XElement row, string localName)
    {
        string? value = (string?)row
            .Element(OoxmlNs.W + "trPr")
            ?.Element(OoxmlNs.W + localName)
            ?.Attribute(OoxmlNs.W + "val");
        return int.TryParse(value, out int parsed) && parsed > 0 ? parsed : 0;
    }

    private static bool ReadRowFlag(XElement row, string localName)
    {
        XElement? element = row
            .Element(OoxmlNs.W + "trPr")
            ?.Element(OoxmlNs.W + localName);
        if (element is null)
        {
            return false;
        }

        string? value = (string?)element.Attribute(OoxmlNs.W + "val");
        return value is null || value is "1" or "true" or "on";
    }

    private static bool IsTableRowInserted(XElement row)
    {
        return row
            .Element(OoxmlNs.W + "trPr")
            ?.Element(OoxmlNs.W + "ins") is not null;
    }

    private static bool IsTableRowDeleted(XElement row)
    {
        return row
            .Element(OoxmlNs.W + "trPr")
            ?.Element(OoxmlNs.W + "del") is not null;
    }
}
