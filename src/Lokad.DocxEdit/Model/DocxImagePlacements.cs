using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

// R01: single owner for final-view image placement membership and order.
// Discovery, snapshot marking, resolution, and reporting must use this
// enumeration so an advertised image ID always addresses the drawing it
// describes. Final-view identity is preserved across text views: deleted
// drawings own no editable ID, and changing text view never renames an
// editable image. Unsupported top-level wrappers (custom XML, block-level
// sdt) own no placements because their paragraphs are not enumerated by
// DocxStoryBlocks. Table rows hidden by trPr deletion own no placements.
// Drawings inside revision runs within visible content still count.
internal static class DocxImagePlacements
{
    internal readonly record struct Placement(
        XElement Drawing,
        int PhysicalBlockIndex,
        bool IsTableCell,
        int PhysicalRowIndex,
        int VisualColumnIndex,
        XElement ContainingBlock,
        XElement? ContainingCell);

    public static IEnumerable<Placement> EnumerateFinalDrawings(XElement container)
    {
        int paragraphIndex = 1;
        int tableIndex = 1;
        foreach (DocxStoryBlocks.StoryBlock entry in DocxStoryBlocks.EnumeratePhysicalBlocks(container))
        {
            if (!DocxStoryBlocks.IsVisibleInView(entry.Wrapper, DocxTextView.Final))
            {
                if (entry.Block.Name == OoxmlNs.W + "p")
                {
                    paragraphIndex++;
                }
                else if (entry.Block.Name == OoxmlNs.W + "tbl")
                {
                    tableIndex++;
                }

                continue;
            }

            XElement block = entry.Block;
            if (block.Name == OoxmlNs.W + "p")
            {
                int currentParagraph = paragraphIndex++;
                foreach (XElement drawing in block.Descendants(OoxmlNs.W + "drawing"))
                {
                    yield return new Placement(drawing, currentParagraph, false, 0, 0, block, null);
                }
            }
            else if (block.Name == OoxmlNs.W + "tbl")
            {
                int currentTable = tableIndex++;
                int rowIndex = 1;
                foreach (XElement row in block.Elements(OoxmlNs.W + "tr"))
                {
                    int currentRow = rowIndex++;
                    if (row.Element(OoxmlNs.W + "trPr")?.Element(OoxmlNs.W + "del") is not null)
                    {
                        continue;
                    }

                    int gridBefore = DocxTableGrid.ReadGridOffset(row, "gridBefore");
                    int columnIndex = 1 + gridBefore;
                    foreach (XElement cell in row.Elements(OoxmlNs.W + "tc"))
                    {
                        int columnSpan = DocxTableGrid.ReadColumnSpan(cell);
                        int currentColumn = columnIndex;
                        foreach (XElement drawing in cell.Descendants(OoxmlNs.W + "drawing"))
                        {
                            yield return new Placement(drawing, currentTable, true, currentRow, currentColumn, block, cell);
                        }

                        columnIndex += columnSpan;
                    }
                }
            }
        }
    }

}
