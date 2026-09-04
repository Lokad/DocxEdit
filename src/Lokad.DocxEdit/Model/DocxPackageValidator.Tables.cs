using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

internal static partial class DocxPackageValidator
{
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
}
