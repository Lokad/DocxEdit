using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

// D01: explicit target IDs bind to the input snapshot for one patch.
// Capture annotates addressable elements with their discovered wire IDs so
// later operations resolve the same element instead of sliding onto a
// neighbour after an earlier insert/delete. Semantic selectors keep resolving
// live against current content. Newly inserted blocks carry no snapshot mark
// and are not addressable by pre-discovered explicit IDs in the same patch;
// deleted targets fail instead of retargeting. Covers paragraphs, tables,
// rows, cells, sections, hyperlinks, bookmarks, content controls, fields, images, and merge groups; comments keep live resolution.
internal static partial class DocxPatchEngine
{
    private static readonly XNamespace SnapshotNs = "http://schemas.lokad.com/docxedit/snapshot";
    private static readonly XName SnapshotIdName = SnapshotNs + "sid";
    private static readonly XName SnapshotAliasName = SnapshotNs + "alias";
    private static readonly XName SnapshotMergeGroupName = SnapshotNs + "mgid";

    private static void CaptureTargetSnapshot(OoxmlPackage package, CancellationToken cancellationToken)
    {
        foreach (StoryPartRef story in DocxPartRoles.GetOrderedStories(package, includeHeadersFooters: true, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            OoxmlPart? part = package.GetPart(story.PartName);
            if (part is null)
            {
                continue;
            }

            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            XElement? root = document.Root;
            if (root is null)
            {
                continue;
            }

            XElement container = root.Element(OoxmlNs.W + "body") ?? root;
            (char storyLetter, int storyPart) = DocxTargetId.ParseStoryPrefix(story.Prefix);
            int paragraphIndex = 1;
            int tableIndex = 1;
            int sectionIndex = 1;
            bool annotated = false;
            foreach (DocxStoryBlocks.StoryBlock entry in DocxStoryBlocks.EnumeratePhysicalBlocks(container))
            {
                cancellationToken.ThrowIfCancellationRequested();
                XElement block = entry.Block;
                if (block.Name == OoxmlNs.W + "p")
                {
                    var id = new DocxTargetId(storyLetter, storyPart, DocxTargetKind.Paragraph, paragraphIndex++, 0, 0);
                    block.SetAttributeValue(SnapshotIdName, id.ToWireValue());
                    annotated = true;
                    if (story.Prefix == "M")
                    {
                        XElement? sectionProperties = block.Element(OoxmlNs.W + "pPr")?.Element(OoxmlNs.W + "sectPr");
                        if (sectionProperties is not null)
                        {
                            var sectionId = new DocxTargetId(storyLetter, storyPart, DocxTargetKind.Section, sectionIndex++, 0, 0);
                            sectionProperties.SetAttributeValue(SnapshotIdName, sectionId.ToWireValue());
                        }
                    }
                }
                else if (block.Name == OoxmlNs.W + "tbl")
                {
                    var tableId = new DocxTargetId(storyLetter, storyPart, DocxTargetKind.Table, tableIndex++, 0, 0);
                    block.SetAttributeValue(SnapshotIdName, tableId.ToWireValue());
                    annotated = true;
                    AnnotateTableRowsAndCells(block, tableId);
                }
                else if (block.Name == OoxmlNs.W + "sectPr")
                {
                    if (story.Prefix == "M")
                    {
                        var sectionId = new DocxTargetId(storyLetter, storyPart, DocxTargetKind.Section, sectionIndex++, 0, 0);
                        block.SetAttributeValue(SnapshotIdName, sectionId.ToWireValue());
                        annotated = true;
                    }
                }
            }

            int bookmarkIndex = 1;
            foreach (XElement start in document.Descendants(OoxmlNs.W + "bookmarkStart"))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (string.IsNullOrWhiteSpace((string?)start.Attribute(OoxmlNs.W + "name")))
                {
                    continue;
                }

                var bookmarkId = new DocxTargetId(storyLetter, storyPart, DocxTargetKind.Bookmark, bookmarkIndex++, 0, 0);
                start.SetAttributeValue(SnapshotIdName, bookmarkId.ToWireValue());
                annotated = true;
            }

            int controlIndex = 1;
            foreach (XElement control in document.Descendants(OoxmlNs.W + "sdt"))
            {
                cancellationToken.ThrowIfCancellationRequested();
                var controlId = new DocxTargetId(storyLetter, storyPart, DocxTargetKind.ContentControl, controlIndex++, 0, 0);
                control.SetAttributeValue(SnapshotIdName, controlId.ToWireValue());
                annotated = true;
            }

            int fieldIndex = 1;
            foreach (XElement field in FindFields(root))
            {
                cancellationToken.ThrowIfCancellationRequested();
                var fieldId = new DocxTargetId(storyLetter, storyPart, DocxTargetKind.Field, fieldIndex++, 0, 0);
                field.SetAttributeValue(SnapshotIdName, fieldId.ToWireValue());
                annotated = true;
            }

            int hyperlinkIndex = 1;
            foreach (XElement hyperlink in document.Descendants(OoxmlNs.W + "hyperlink"))
            {
                cancellationToken.ThrowIfCancellationRequested();
                var hyperlinkId = new DocxTargetId(storyLetter, storyPart, DocxTargetKind.Hyperlink, hyperlinkIndex++, 0, 0);
                hyperlink.SetAttributeValue(SnapshotIdName, hyperlinkId.ToWireValue());
                annotated = true;
            }

            int imageIndex = 1;
            foreach (ImageBlipEntry entry in FindImageBlipEntries(package, story.PartName, document, cancellationToken))
            {
                var imageId = new DocxTargetId(storyLetter, storyPart, DocxTargetKind.Image, imageIndex++, 0, 0);
                entry.Blip.SetAttributeValue(SnapshotIdName, imageId.ToWireValue());
                annotated = true;
            }
            if (annotated)
            {
                SaveDocumentPart(package, story.PartName, document);
            }
        }
    }

    private static void AnnotateTableRowsAndCells(XElement table, DocxTargetId tableId)
    {
        int rowIndex = 1;
        foreach (XElement row in table.Elements(OoxmlNs.W + "tr"))
        {
            int currentRow = rowIndex++;
            var rowId = tableId with { Kind = DocxTargetKind.Row, Secondary = currentRow };
            row.SetAttributeValue(SnapshotIdName, rowId.ToWireValue());
            int columnIndex = 1 + ReadTableRowGridOffset(row, "gridBefore");
            foreach (XElement cell in row.Elements(OoxmlNs.W + "tc"))
            {
                int columnSpan = ReadTableCellColumnSpan(cell);
                var cellId = tableId with { Kind = DocxTargetKind.Cell, Secondary = currentRow, Tertiary = columnIndex };
                cell.SetAttributeValue(SnapshotIdName, cellId.ToWireValue());
                columnIndex += columnSpan;
            }
        }
        int mergeGroupOrdinal = 1;
        while (TryFindMergeGroupRoot(table, mergeGroupOrdinal, out _, out XElement? mergeRoot, out _, out _))
        {
            var mergeGroupId = tableId with { Kind = DocxTargetKind.MergeGroup, Secondary = mergeGroupOrdinal };
            mergeRoot?.SetAttributeValue(SnapshotMergeGroupName, mergeGroupId.ToWireValue());
            mergeGroupOrdinal++;
        }
    }

    private static Dictionary<string, byte[]> RecordStoryPartBytes(OoxmlPackage package, CancellationToken cancellationToken)
    {
        var originals = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        foreach (StoryPartRef story in DocxPartRoles.GetOrderedStories(package, includeHeadersFooters: true, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(story.PartName);
            if (part is not null)
            {
                originals[story.PartName] = part.Bytes;
            }
        }

        return originals;
    }

    private static void UnmarkStoryParts(OoxmlPackage package, CancellationToken cancellationToken)
    {
        foreach (StoryPartRef story in DocxPartRoles.GetOrderedStories(package, includeHeadersFooters: true, cancellationToken))
        {
            package.UnmarkPartTouched(story.PartName);
        }
    }

    private static void StripTargetSnapshot(
        OoxmlPackage package,
        IReadOnlyDictionary<string, byte[]> originals,
        CancellationToken cancellationToken)
    {
        var operationTouched = new HashSet<string>(package.TouchedPartNames, StringComparer.OrdinalIgnoreCase);
        foreach (StoryPartRef story in DocxPartRoles.GetOrderedStories(package, includeHeadersFooters: true, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!operationTouched.Contains(story.PartName))
            {
                // No operation edited this part: restore input bytes exactly so
                // snapshot bookkeeping never rewrites otherwise-untouched output.
                if (originals.TryGetValue(story.PartName, out byte[]? original))
                {
                    OoxmlPart? part = package.GetPart(story.PartName);
                    if (part is not null && !part.Bytes.SequenceEqual(original))
                    {
                        package.ReplacePartBytes(story.PartName, original);
                    }

                    package.UnmarkPartTouched(story.PartName);
                }

                continue;
            }

            OoxmlPart? edited = package.GetPart(story.PartName);
            if (edited is null)
            {
                continue;
            }

            using Stream stream = edited.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            bool dirty = false;
            foreach (XElement element in document.Descendants())
            {
                foreach (XName markName in new[] { SnapshotIdName, SnapshotAliasName, SnapshotMergeGroupName })
                {
                    XAttribute? attribute = element.Attribute(markName);
                    if (attribute is not null)
                    {
                        attribute.Remove();
                        dirty = true;
                    }
                }

                // Capture materializes snapshot namespace declarations alongside
                // the marks; drop the now-unused declarations so output bytes
                // carry no snapshot residue.
                foreach (XAttribute declaration in element.Attributes().Where(static a => a.IsNamespaceDeclaration).ToArray())
                {
                    if (string.Equals(declaration.Value, SnapshotNs.NamespaceName, StringComparison.Ordinal))
                    {
                        declaration.Remove();
                        dirty = true;
                    }
                }
            }

            if (dirty)
            {
                SaveDocumentPart(package, story.PartName, document);
            }
        }
    }

    private static XElement? FindSnapshotCellByVisualColumn(XDocument document, DocxTargetId cellId)
    {
        string wireId = cellId.ToWireValue();
        XElement? exact = FindSnapshotElement(document, OoxmlNs.W + "tc", wireId);
        if (exact is not null)
        {
            return exact;
        }

        // Visual columns inside a horizontal span share one w:tc: the snapshot
        // marks the span start, so fall back to span coverage within the row.
        XElement? row = FindSnapshotElement(document, OoxmlNs.W + "tr", cellId.RowId.ToWireValue());
        if (row is null)
        {
            return null;
        }

        foreach (XElement cell in row.Elements(OoxmlNs.W + "tc"))
        {
            string? snapshotId = (string?)cell.Attribute(SnapshotIdName);
            if (snapshotId is null ||
                !DocxTargetId.TryParse(snapshotId, out DocxTargetId annotated) ||
                annotated.Kind != DocxTargetKind.Cell ||
                annotated.Secondary != cellId.Secondary)
            {
                continue;
            }

            int span = ReadTableCellColumnSpan(cell);
            if (cellId.Tertiary >= annotated.Tertiary && cellId.Tertiary < annotated.Tertiary + span)
            {
                return cell;
            }
        }

        return null;
    }
    private static string? NormalizeSnapshotWireId(string target)
    {
        return DocxTargetId.TryParse(target, out DocxTargetId parsed) ? parsed.ToWireValue() : null;
    }

    private static XElement? FindSnapshotElementInContainer(XElement container, XName name, string canonicalWireId)
    {
        foreach (XElement element in container.Descendants(name))
        {
            if (string.Equals((string?)element.Attribute(SnapshotIdName), canonicalWireId, StringComparison.Ordinal))
            {
                return element;
            }
        }

        return container.Name == name &&
            string.Equals((string?)container.Attribute(SnapshotIdName), canonicalWireId, StringComparison.Ordinal)
            ? container
            : null;
    }

    private static XElement? FindSnapshotElement(XDocument document, XName name, string canonicalWireId)
    {
        foreach (XElement element in document.Descendants(name))
        {
            if (string.Equals((string?)element.Attribute(SnapshotIdName), canonicalWireId, StringComparison.Ordinal))
            {
                return element;
            }
        }

        return null;
    }
}
