using System.Xml.Linq;
using DocxEdit.Ooxml;

namespace DocxEdit.Model;

internal static class DocxDocumentScanner
{
    public static DocxDocumentModel Scan(OoxmlPackage package, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (package.MainDocumentPartName is null)
        {
            return DocxDocumentModel.Empty;
        }

        OoxmlPart documentPart = package.GetPart(package.MainDocumentPartName)
            ?? throw new InvalidDataException($"Main document part '{package.MainDocumentPartName}' does not exist.");
        using Stream stream = documentPart.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        XElement body = document.Root?.Element(OoxmlNs.W + "body")
            ?? throw new InvalidDataException("Main document part is missing w:body.");

        IReadOnlyDictionary<string, OoxmlRelationship> relationships = package
            .GetRelationships(package.MainDocumentPartName, cancellationToken)
            .ToDictionary(relationship => relationship.Id, StringComparer.Ordinal);

        var paragraphs = new List<DocxParagraphInfo>();
        var tables = new List<DocxTableInfo>();
        var images = new List<DocxImageInfo>();
        int paragraphIndex = 1;
        int tableIndex = 1;
        int imageIndex = 1;

        foreach (XElement block in body.Elements())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (block.Name == OoxmlNs.W + "p")
            {
                paragraphs.Add(ReadParagraph(block, $"M.P{paragraphIndex++:0000}", "main", package, relationships, images, ref imageIndex));
            }
            else if (block.Name == OoxmlNs.W + "tbl")
            {
                tables.Add(ReadTable(block, $"M.T{tableIndex++:0000}", package, relationships, images, ref imageIndex));
            }
        }

        return new DocxDocumentModel(paragraphs, tables, images);
    }

    private static DocxParagraphInfo ReadParagraph(
        XElement paragraph,
        string id,
        string story,
        OoxmlPackage package,
        IReadOnlyDictionary<string, OoxmlRelationship> relationships,
        List<DocxImageInfo> images,
        ref int imageIndex)
    {
        var runs = paragraph
            .Elements(OoxmlNs.W + "r")
            .Select(run => new DocxRunInfo(ReadVisibleText(run)))
            .Where(run => run.Text.Length != 0)
            .ToArray();
        foreach (XElement drawing in paragraph.Descendants(OoxmlNs.W + "drawing"))
        {
            AddDrawingImages(drawing, package, relationships, images, ref imageIndex);
        }

        return new DocxParagraphInfo(
            id,
            story,
            ReadVisibleText(paragraph),
            ReadHeadingLevel(paragraph),
            runs);
    }

    private static DocxTableInfo ReadTable(
        XElement table,
        string id,
        OoxmlPackage package,
        IReadOnlyDictionary<string, OoxmlRelationship> relationships,
        List<DocxImageInfo> images,
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
                    AddDrawingImages(drawing, package, relationships, images, ref imageIndex);
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

        return new DocxTableInfo(id, "main", rowIndex - 1, maxColumns, cells);
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

    private static void AddDrawingImages(
        XElement drawing,
        OoxmlPackage package,
        IReadOnlyDictionary<string, OoxmlRelationship> relationships,
        List<DocxImageInfo> images,
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

            images.Add(new DocxImageInfo($"M.I{imageIndex++:0000}", imagePart.Name, imagePart.ContentType, imagePart.Bytes.Length));
        }
    }
}

