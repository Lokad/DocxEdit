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
// rows, cells, sections, hyperlinks, bookmarks, content controls, fields, images, merge groups, and comment bodies; threaded replies and semantic selectors keep live resolution.
internal static partial class DocxPatchEngine
{
    private static readonly XNamespace SnapshotNs = "http://schemas.lokad.com/docxedit/snapshot";
    private static readonly XName SnapshotIdName = SnapshotNs + "sid";
    private static readonly XName SnapshotAliasName = SnapshotNs + "alias";
    private static readonly XName SnapshotMergeGroupName = SnapshotNs + "mgid";
    private static readonly XName SnapshotCreatedName = SnapshotNs + "created";

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

        int commentsPartOrdinal = 1;
        foreach (string commentsPartName in GetCommentsPartNames(package, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            OoxmlPart? commentsPart = package.GetPart(commentsPartName);
            if (commentsPart is null)
            {
                continue;
            }

            using Stream commentsStream = commentsPart.OpenRead();
            XDocument commentsDocument = SafeXml.Load(commentsStream, cancellationToken);
            XElement? commentsRoot = commentsDocument.Root;
            if (commentsRoot is null)
            {
                commentsPartOrdinal++;
                continue;
            }

            int commentOrdinal = 1;
            bool commentsAnnotated = false;
            foreach (XElement comment in commentsRoot.Elements(OoxmlNs.W + "comment"))
            {
                cancellationToken.ThrowIfCancellationRequested();
                string wireId = "C" + commentsPartOrdinal.ToString("D3", System.Globalization.CultureInfo.InvariantCulture) + ".C" + commentOrdinal++.ToString("D4", System.Globalization.CultureInfo.InvariantCulture);
                comment.SetAttributeValue(SnapshotIdName, wireId);
                commentsAnnotated = true;
            }

            if (commentsAnnotated)
            {
                SaveDocumentPart(package, commentsPartName, commentsDocument);
            }

            commentsPartOrdinal++;
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
        int mergeGroupOrdinal = 0;
        foreach (var merge in EnumerateMergeGroupRoots(table))
        {
            mergeGroupOrdinal++;
            var mergeGroupId = tableId with { Kind = DocxTargetKind.MergeGroup, Secondary = mergeGroupOrdinal };
            merge.Cell.SetAttributeValue(SnapshotMergeGroupName, mergeGroupId.ToWireValue());
        }
    }

    private static List<string> SnapshotPartNames(OoxmlPackage package, CancellationToken cancellationToken)
    {
        var names = new List<string>();
        foreach (StoryPartRef story in DocxPartRoles.GetOrderedStories(package, includeHeadersFooters: true, cancellationToken))
        {
            names.Add(story.PartName);
        }

        foreach (string commentsPartName in GetCommentsPartNames(package, cancellationToken))
        {
            names.Add(commentsPartName);
        }

        return names;
    }

    internal static string CreatedMarkValue(DocxPatchOperation operation, int index)
    {
        return "op" + operation.Index.ToString(System.Globalization.CultureInfo.InvariantCulture) + "-" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    internal static int? PhysicalParagraphOrdinal(XDocument document, XElement paragraph)
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
            if (entry.Block.Name != OoxmlNs.W + "p")
            {
                continue;
            }

            ordinal++;
            if (ReferenceEquals(entry.Block, paragraph))
            {
                return ordinal;
            }
        }

        return null;
    }

    // R04: rebase report IDs to final-output coordinates. Affected entries name
    // input-snapshot IDs at capture time and created entries name operation-time
    // IDs; later structural operations can renumber both. Before publication each
    // ID is resolved back to its element (snapshot marks for pre-existing objects,
    // creation marks for new ones) and rewritten in final physical order, so a
    // reported ID always matches a fresh read. IDs whose object no longer exists
    // are kept as-is: they name the input or operation-time object, which fails
    // loudly on lookup instead of naming a different live object.
    internal static void RebaseReportTargetIds(
        OoxmlPackage package,
        List<DocxPatchOperationReport> reports,
        CancellationToken cancellationToken)
    {
        if (reports.Count == 0)
        {
            return;
        }

        // C04: reports without identities need no coordinate rewrite; skip the
        // story sweeps entirely for validation-only or fully refused patches.
        if (reports.All(static report => report.AffectedTargets.Count == 0 && report.CreatedTargetIds.Count == 0))
        {
            return;
        }

        IReadOnlyDictionary<string, string> prefixes = DocxPartRoles.GetStoryPrefixes(package, cancellationToken);
        var finals = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (StoryPartRef story in DocxPartRoles.GetOrderedStories(package, includeHeadersFooters: true, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!prefixes.TryGetValue(story.PartName, out string? prefix))
            {
                continue;
            }

            OoxmlPart? part = package.GetPart(story.PartName);
            if (part is null)
            {
                continue;
            }

            XDocument document = LoadDocumentPart(package, story.PartName, cancellationToken, out _);
            if (document.Root is null)
            {
                continue;
            }

            (char storyLetter, int storyPart) = DocxTargetId.ParseStoryPrefix(prefix);
            CollectFinalTargetIds(package, story.PartName, document, storyLetter, storyPart, finals, cancellationToken);
        }

        for (int reportIndex = 0; reportIndex < reports.Count; reportIndex++)
        {
            DocxPatchOperationReport report = reports[reportIndex];
            if (report.AffectedTargets.Count == 0 && report.CreatedTargetIds.Count == 0)
            {
                continue;
            }

            var affected = new List<DocxPatchAffectedTarget>();
            foreach (DocxPatchAffectedTarget entry in report.AffectedTargets)
            {
                affected.Add(RebaseAffectedTarget(report, entry, finals));
            }

            var created = new List<string>();
            for (int createdIndex = 0; createdIndex < report.CreatedTargetIds.Count; createdIndex++)
            {
                string createdId = report.CreatedTargetIds[createdIndex];
                string mark = "op" + report.Index.ToString(System.Globalization.CultureInfo.InvariantCulture) + "-" + createdIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (finals.TryGetValue(mark, out string? finalWire))
                {
                    created.Add(finalWire);
                }
                else if (createdId.StartsWith("comment:", StringComparison.Ordinal))
                {
                    created.Add(createdId);
                }
            }

            reports[reportIndex] = report with { AffectedTargets = affected, CreatedTargetIds = created };
        }
    }

    private static void CollectFinalTargetIds(
        OoxmlPackage package,
        string partName,
        XDocument document,
        char storyLetter,
        int storyPart,
        Dictionary<string, string> finals,
        CancellationToken cancellationToken)
    {
        XElement root = document.Root!;
        XElement container = root.Element(OoxmlNs.W + "body") ?? root;
        int paragraphOrdinal = 0;
        int tableOrdinal = 0;
        foreach (DocxStoryBlocks.StoryBlock entry in DocxStoryBlocks.EnumeratePhysicalBlocks(container))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (entry.Block.Name == OoxmlNs.W + "p")
            {
                paragraphOrdinal++;
                RecordFinalId(entry.Block, new DocxTargetId(storyLetter, storyPart, DocxTargetKind.Paragraph, paragraphOrdinal, 0, 0), finals);
            }
            else if (entry.Block.Name == OoxmlNs.W + "tbl")
            {
                tableOrdinal++;
                var tableId = new DocxTargetId(storyLetter, storyPart, DocxTargetKind.Table, tableOrdinal, 0, 0);
                RecordFinalId(entry.Block, tableId, finals);
                CollectFinalTableIds(entry.Block, tableId, finals, cancellationToken);
            }
        }

        CollectFinalDescendantIds(document, OoxmlNs.W + "hyperlink", DocxTargetKind.Hyperlink, storyLetter, storyPart, finals, static _ => true);
        CollectFinalDescendantIds(document, OoxmlNs.W + "sdt", DocxTargetKind.ContentControl, storyLetter, storyPart, finals, static _ => true);
        CollectFinalDescendantIds(document, OoxmlNs.W + "bookmarkStart", DocxTargetKind.Bookmark, storyLetter, storyPart, finals, static element => !string.IsNullOrWhiteSpace((string?)element.Attribute(OoxmlNs.W + "name")));
        int fieldOrdinal = 0;
        foreach (XElement field in FindFields(root))
        {
            fieldOrdinal++;
            RecordFinalId(field, new DocxTargetId(storyLetter, storyPart, DocxTargetKind.Field, fieldOrdinal, 0, 0), finals);
        }

        int imageOrdinal = 0;
        foreach (ImageBlipEntry blip in FindImageBlipEntries(package, partName, document, cancellationToken))
        {
            imageOrdinal++;
            RecordFinalId(blip.Blip, new DocxTargetId(storyLetter, storyPart, DocxTargetKind.Image, imageOrdinal, 0, 0), finals);
        }
    }

    private static void CollectFinalDescendantIds(
        XDocument document,
        XName name,
        DocxTargetKind kind,
        char storyLetter,
        int storyPart,
        Dictionary<string, string> finals,
        Func<XElement, bool> accept)
    {
        int ordinal = 0;
        foreach (XElement element in document.Descendants(name))
        {
            if (!accept(element))
            {
                continue;
            }

            ordinal++;
            RecordFinalId(element, new DocxTargetId(storyLetter, storyPart, kind, ordinal, 0, 0), finals);
        }
    }

    private static void CollectFinalTableIds(
        XElement table,
        DocxTargetId tableId,
        Dictionary<string, string> finals,
        CancellationToken cancellationToken)
    {
        int rowOrdinal = 0;
        foreach (XElement row in table.Elements(OoxmlNs.W + "tr"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            rowOrdinal++;
            RecordFinalId(row, tableId with { Kind = DocxTargetKind.Row, Secondary = rowOrdinal }, finals);
            int visualColumn = 1 + ReadTableRowGridOffset(row, "gridBefore");
            foreach (XElement cell in row.Elements(OoxmlNs.W + "tc"))
            {
                RecordFinalId(cell, tableId with { Kind = DocxTargetKind.Cell, Secondary = rowOrdinal, Tertiary = visualColumn }, finals);
                visualColumn += ReadTableCellColumnSpan(cell);
            }
        }

        int mergeOrdinal = 0;
        foreach (var merge in EnumerateMergeGroupRoots(table))
        {
            mergeOrdinal++;
            RecordFinalId(merge.Cell, tableId with { Kind = DocxTargetKind.MergeGroup, Secondary = mergeOrdinal }, finals);
        }
    }

    private static void RecordFinalId(XElement element, DocxTargetId finalId, Dictionary<string, string> finals)
    {
        string finalWire = finalId.ToWireValue();
        string? snapshotId = (string?)element.Attribute(SnapshotIdName);
        if (snapshotId is not null && finalId.Kind != DocxTargetKind.MergeGroup)
        {
            finals[snapshotId] = finalWire;
        }

        string? mergeId = (string?)element.Attribute(SnapshotMergeGroupName);
        if (mergeId is not null && finalId.Kind == DocxTargetKind.MergeGroup)
        {
            finals[mergeId] = finalWire;
        }

        string? createdMark = (string?)element.Attribute(SnapshotCreatedName);
        if (createdMark is not null)
        {
            finals[createdMark] = finalWire;
        }
    }

    private static DocxPatchAffectedTarget RebaseAffectedTarget(
        DocxPatchOperationReport report,
        DocxPatchAffectedTarget entry,
        Dictionary<string, string> finals)
    {
        if (entry.Kind is "row" or "cell" &&
            report.OperationName is "append-row" or "insert-row-before" or "insert-row-after")
        {
            return RebaseInsertedRowEntry(report, entry, finals) ?? entry;
        }

        if (entry.CreatedMark is not null)
        {
            if (finals.TryGetValue(entry.CreatedMark, out string? markWire) &&
                DocxTargetId.TryParse(markWire, out DocxTargetId markId))
            {
                return entry with { Id = markId, FinalId = markId };
            }

            return entry with { FinalId = null };
        }
        if (finals.TryGetValue(entry.Id.ToWireValue(), out string? finalWire) &&
            DocxTargetId.TryParse(finalWire, out DocxTargetId finalId))
        {
            DocxTargetId? finalParent = entry.ParentId;
            if (entry.ParentId is not null)
            {
                if (finals.TryGetValue(entry.ParentId.Value.ToWireValue(), out string? parentWire))
                {
                    if (DocxTargetId.TryParse(parentWire, out DocxTargetId parentId))
                    {
                        finalParent = parentId;
                    }
                }
            }
            DocxTargetId? finalMerge = entry.MergeGroupId;
            if (entry.MergeGroupId is not null)
            {
                if (finals.TryGetValue(entry.MergeGroupId.Value.ToWireValue(), out string? mergeWire))
                {
                    if (DocxTargetId.TryParse(mergeWire, out DocxTargetId mergeId))
                    {
                        finalMerge = mergeId;
                    }
                }
            }
            return entry with { Id = finalId, ParentId = finalParent, MergeGroupId = finalMerge, FinalId = finalId };
        }

        return entry with { FinalId = null };
    }

    private static DocxPatchAffectedTarget? RebaseInsertedRowEntry(
        DocxPatchOperationReport report,
        DocxPatchAffectedTarget entry,
        Dictionary<string, string> finals)
    {
        string mark = "op" + report.Index.ToString(System.Globalization.CultureInfo.InvariantCulture) + "-0";
        if (!finals.TryGetValue(mark, out string? rowWire) ||
            !DocxTargetId.TryParse(rowWire, out DocxTargetId rowId) ||
            rowId.Kind != DocxTargetKind.Row)
        {
            return null;
        }

        if (entry.Kind == "row")
        {
            return entry with { Id = rowId, FinalId = rowId };
        }

        DocxTargetId cellId = rowId with { Kind = DocxTargetKind.Cell, Tertiary = entry.Id.Tertiary };
        return entry with { Id = cellId, FinalId = cellId };
    }

    private static Dictionary<string, byte[]> RecordStoryPartBytes(OoxmlPackage package, CancellationToken cancellationToken)
    {
        var originals = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        foreach (string partName in SnapshotPartNames(package, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(partName);
            if (part is not null)
            {
                originals[partName] = part.Bytes;
            }
        }

        return originals;
    }

    private static void UnmarkStoryParts(OoxmlPackage package, CancellationToken cancellationToken)
    {
        foreach (string partName in SnapshotPartNames(package, cancellationToken))
        {
            package.UnmarkPartTouched(partName);
        }
    }

    private static void StripTargetSnapshot(
        OoxmlPackage package,
        IReadOnlyDictionary<string, byte[]> originals,
        CancellationToken cancellationToken)
    {
        var operationTouched = new HashSet<string>(package.TouchedPartNames, StringComparer.OrdinalIgnoreCase);
        foreach (string partName in SnapshotPartNames(package, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!operationTouched.Contains(partName))
            {
                // No operation edited this part: restore input bytes exactly so
                // snapshot bookkeeping never rewrites otherwise-untouched output.
                if (originals.TryGetValue(partName, out byte[]? original))
                {
                    OoxmlPart? part = package.GetPart(partName);
                    if (part is not null && !part.Bytes.SequenceEqual(original))
                    {
                        package.ReplacePartBytes(partName, original);
                    }

                    package.UnmarkPartTouched(partName);
                }

                continue;
            }

            OoxmlPart? edited = package.GetPart(partName);
            if (edited is null)
            {
                continue;
            }

            using Stream stream = edited.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            bool dirty = false;
            foreach (XElement element in document.Descendants())
            {
                foreach (XName markName in new[] { SnapshotIdName, SnapshotAliasName, SnapshotMergeGroupName, SnapshotCreatedName })
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
                SaveDocumentPart(package, partName, document);
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
