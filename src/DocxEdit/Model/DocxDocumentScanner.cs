using System.Xml.Linq;
using DocxEdit.Ooxml;

namespace DocxEdit.Model;

internal static class DocxDocumentScanner
{
    public static DocxDocumentModel Scan(
        OoxmlPackage package,
        bool includeHeadersFooters = false,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (package.MainDocumentPartName is null)
        {
            return DocxDocumentModel.Empty;
        }

        var paragraphs = new List<DocxParagraphInfo>();
        var tables = new List<DocxTableInfo>();
        var images = new List<DocxImageInfo>();
        var sections = new List<DocxSectionInfo>();
        ScanStory(package, package.MainDocumentPartName, "M", "main", paragraphs, tables, images, sections, cancellationToken);

        if (includeHeadersFooters)
        {
            IReadOnlyList<OoxmlRelationship> relationships = package.GetRelationships(package.MainDocumentPartName, cancellationToken);
            int headerIndex = 1;
            foreach (OoxmlRelationship relationship in relationships
                .Where(relationship => !relationship.IsExternal && relationship.Type == OoxmlRelTypes.Header && relationship.ResolvedTarget is not null)
                .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
            {
                if (package.GetPart(relationship.ResolvedTarget!) is not null)
                {
                    string prefix = $"H{headerIndex++:000}";
                    ScanStory(package, relationship.ResolvedTarget!, prefix, $"header[{headerIndex - 1}]", paragraphs, tables, images, sections, cancellationToken);
                }
            }

            int footerIndex = 1;
            foreach (OoxmlRelationship relationship in relationships
                .Where(relationship => !relationship.IsExternal && relationship.Type == OoxmlRelTypes.Footer && relationship.ResolvedTarget is not null)
                .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
            {
                if (package.GetPart(relationship.ResolvedTarget!) is not null)
                {
                    string prefix = $"F{footerIndex++:000}";
                    ScanStory(package, relationship.ResolvedTarget!, prefix, $"footer[{footerIndex - 1}]", paragraphs, tables, images, sections, cancellationToken);
                }
            }
        }

        return new DocxDocumentModel(paragraphs, tables, images, sections);
    }

    private static void ScanStory(
        OoxmlPackage package,
        string partName,
        string idPrefix,
        string story,
        List<DocxParagraphInfo> paragraphs,
        List<DocxTableInfo> tables,
        List<DocxImageInfo> images,
        List<DocxSectionInfo> sections,
        CancellationToken cancellationToken)
    {
        OoxmlPart part = package.GetPart(partName)
            ?? throw new InvalidDataException($"Story part '{partName}' does not exist.");
        using Stream stream = part.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        XElement body = document.Root?.Element(OoxmlNs.W + "body")
            ?? document.Root
            ?? throw new InvalidDataException($"Story part '{partName}' has no XML root.");

        IReadOnlyDictionary<string, OoxmlRelationship> relationships = package
            .GetRelationships(partName, cancellationToken)
            .ToDictionary(relationship => relationship.Id, StringComparer.Ordinal);

        int paragraphIndex = 1;
        int tableIndex = 1;
        int imageIndex = 1;
        int sectionIndex = sections.Count(section => section.Id.StartsWith($"{idPrefix}.S", StringComparison.Ordinal)) + 1;
        foreach (XElement block in body.Elements())
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement? sectionProperties = null;
            if (block.Name == OoxmlNs.W + "p")
            {
                paragraphs.Add(ReadParagraph(block, $"{idPrefix}.P{paragraphIndex++:0000}", story, package, relationships, images, idPrefix, ref imageIndex));
                sectionProperties = block.Element(OoxmlNs.W + "pPr")?.Element(OoxmlNs.W + "sectPr");
            }
            else if (block.Name == OoxmlNs.W + "tbl")
            {
                tables.Add(ReadTable(block, $"{idPrefix}.T{tableIndex++:0000}", story, package, relationships, images, idPrefix, ref imageIndex));
            }
            else if (block.Name == OoxmlNs.W + "sectPr")
            {
                sectionProperties = block;
            }

            if (sectionProperties is not null && idPrefix == "M")
            {
                sections.Add(ReadSection(sectionProperties, $"{idPrefix}.S{sectionIndex++:0000}", story));
            }
        }
    }

    private static DocxParagraphInfo ReadParagraph(
        XElement paragraph,
        string id,
        string story,
        OoxmlPackage package,
        IReadOnlyDictionary<string, OoxmlRelationship> relationships,
        List<DocxImageInfo> images,
        string imageIdPrefix,
        ref int imageIndex)
    {
        var runs = paragraph
            .Elements(OoxmlNs.W + "r")
            .Select(run => new DocxRunInfo(ReadVisibleText(run)))
            .Where(run => run.Text.Length != 0)
            .ToArray();
        foreach (XElement drawing in paragraph.Descendants(OoxmlNs.W + "drawing"))
        {
            AddDrawingImages(drawing, package, relationships, images, imageIdPrefix, ref imageIndex);
        }

        return new DocxParagraphInfo(
            id,
            story,
            ReadVisibleText(paragraph),
            ReadHeadingLevel(paragraph),
            ReadListInfo(paragraph),
            runs);
    }

    private static DocxTableInfo ReadTable(
        XElement table,
        string id,
        string story,
        OoxmlPackage package,
        IReadOnlyDictionary<string, OoxmlRelationship> relationships,
        List<DocxImageInfo> images,
        string imageIdPrefix,
        ref int imageIndex)
    {
        var cells = new List<DocxTableCellInfo>();
        int rowIndex = 1;
        int maxColumns = 0;
        foreach (XElement row in table.Elements(OoxmlNs.W + "tr"))
        {
            int columnIndex = 1;
            foreach (XElement cell in row.Elements(OoxmlNs.W + "tc"))
            {
                foreach (XElement drawing in cell.Descendants(OoxmlNs.W + "drawing"))
                {
                    AddDrawingImages(drawing, package, relationships, images, imageIdPrefix, ref imageIndex);
                }

                cells.Add(new DocxTableCellInfo(
                    $"{id}.R{rowIndex:00}.C{columnIndex:00}",
                    rowIndex,
                    columnIndex,
                    ReadVisibleText(cell)));
                columnIndex++;
            }

            maxColumns = Math.Max(maxColumns, columnIndex - 1);
            rowIndex++;
        }

        return new DocxTableInfo(id, story, rowIndex - 1, maxColumns, cells);
    }

    private static string ReadVisibleText(XElement container)
    {
        var buffer = new List<string>();
        foreach (XElement element in container.Descendants())
        {
            if (IsInsideDeletedRevision(element))
            {
                continue;
            }

            if (element.Name == OoxmlNs.W + "t")
            {
                buffer.Add(element.Value);
            }
            else if (element.Name == OoxmlNs.W + "tab")
            {
                buffer.Add("\t");
            }
            else if (element.Name == OoxmlNs.W + "br")
            {
                buffer.Add("\n");
            }
        }

        return string.Concat(buffer);
    }

    private static bool IsInsideDeletedRevision(XElement element)
    {
        return element.Ancestors(OoxmlNs.W + "del").Any() ||
            element.Ancestors(OoxmlNs.W + "moveFrom").Any();
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

    private static DocxListInfo? ReadListInfo(XElement paragraph)
    {
        XElement? numberingProperties = paragraph
            .Element(OoxmlNs.W + "pPr")
            ?.Element(OoxmlNs.W + "numPr");
        string? numberingId = (string?)numberingProperties
            ?.Element(OoxmlNs.W + "numId")
            ?.Attribute(OoxmlNs.W + "val");
        if (string.IsNullOrWhiteSpace(numberingId))
        {
            return null;
        }

        string? levelText = (string?)numberingProperties
            ?.Element(OoxmlNs.W + "ilvl")
            ?.Attribute(OoxmlNs.W + "val");
        int level = int.TryParse(levelText, out int parsedLevel) && parsedLevel >= 0
            ? parsedLevel
            : 0;
        return new DocxListInfo(numberingId, level);
    }

    private static DocxSectionInfo ReadSection(XElement sectionProperties, string id, string story)
    {
        string? columnCountText = (string?)sectionProperties
            .Element(OoxmlNs.W + "cols")
            ?.Attribute(OoxmlNs.W + "num");
        int columns = int.TryParse(columnCountText, out int parsedColumns) && parsedColumns > 0
            ? parsedColumns
            : 1;
        string orientation = (string?)sectionProperties
            .Element(OoxmlNs.W + "pgSz")
            ?.Attribute(OoxmlNs.W + "orient") ?? "portrait";
        return new DocxSectionInfo(id, story, columns, orientation);
    }

    private static void AddDrawingImages(
        XElement drawing,
        OoxmlPackage package,
        IReadOnlyDictionary<string, OoxmlRelationship> relationships,
        List<DocxImageInfo> images,
        string imageIdPrefix,
        ref int imageIndex)
    {
        foreach (XElement blip in drawing.Descendants(OoxmlNs.A + "blip"))
        {
            string? relationshipId = (string?)blip.Attribute(OoxmlNs.R + "embed");
            if (relationshipId is null ||
                !relationships.TryGetValue(relationshipId, out OoxmlRelationship? relationship) ||
                relationship.IsExternal ||
                relationship.ResolvedTarget is null)
            {
                continue;
            }

            OoxmlPart? imagePart = package.GetPart(relationship.ResolvedTarget);
            if (imagePart is null)
            {
                continue;
            }

            if (images.Any(image => string.Equals(image.PartName, imagePart.Name, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            images.Add(new DocxImageInfo($"{imageIdPrefix}.I{imageIndex++:0000}", imagePart.Name, imagePart.ContentType, imagePart.Bytes.Length));
        }
    }
}
