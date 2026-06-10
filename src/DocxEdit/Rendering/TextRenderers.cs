using System.Text;
using DocxEdit.Model;

namespace DocxEdit.Rendering;

internal static class TextRenderers
{
    public static string RenderRead(DocxDocumentModel model)
    {
        var builder = new StringBuilder();
        foreach (DocxParagraphInfo paragraph in model.Paragraphs)
        {
            string kind = paragraph.HeadingLevel is null ? "paragraph" : $"heading level={paragraph.HeadingLevel}";
            builder.Append(paragraph.Id).Append(' ').Append(kind).Append(" text=\"").Append(Escape(paragraph.Text)).AppendLine("\"");
        }

        foreach (DocxTableInfo table in model.Tables)
        {
            builder.Append(table.Id).Append(" table rows=").Append(table.RowCount).Append(" columns=").Append(table.ColumnCount).AppendLine();
            foreach (DocxTableCellInfo cell in table.Cells)
            {
                builder.Append("  ").Append(cell.Id).Append(" text=\"").Append(Escape(cell.Text)).AppendLine("\"");
            }
        }

        foreach (DocxImageInfo image in model.Images)
        {
            builder.Append(image.Id).Append(" image part=").Append(image.PartName).Append(" content-type=").Append(image.ContentType ?? "unknown").Append(" bytes=").Append(image.ByteLength).AppendLine();
        }

        return builder.ToString();
    }

    public static IReadOnlyList<string> RenderOutline(DocxDocumentModel model)
    {
        var lines = new List<string>();
        foreach (DocxParagraphInfo paragraph in model.Paragraphs.Where(paragraph => paragraph.HeadingLevel is not null))
        {
            lines.Add($"{paragraph.Id} heading level={paragraph.HeadingLevel} text=\"{Escape(paragraph.Text)}\"");
        }

        foreach (DocxTableInfo table in model.Tables)
        {
            lines.Add($"{table.Id} table rows={table.RowCount} columns={table.ColumnCount}");
        }

        foreach (DocxImageInfo image in model.Images)
        {
            lines.Add($"{image.Id} image part={image.PartName}");
        }

        return lines;
    }

    public static IReadOnlyList<string> Find(DocxDocumentModel model, string query)
    {
        var matches = new List<string>();
        foreach (DocxParagraphInfo paragraph in model.Paragraphs)
        {
            if (paragraph.Text.Contains(query, StringComparison.OrdinalIgnoreCase))
            {
                matches.Add($"{paragraph.Id} text=\"{Escape(paragraph.Text)}\"");
            }
        }

        foreach (DocxTableInfo table in model.Tables)
        {
            foreach (DocxTableCellInfo cell in table.Cells)
            {
                if (cell.Text.Contains(query, StringComparison.OrdinalIgnoreCase))
                {
                    matches.Add($"{cell.Id} text=\"{Escape(cell.Text)}\"");
                }
            }
        }

        return matches;
    }

    public static string? Dump(DocxDocumentModel model, string targetId)
    {
        DocxParagraphInfo? paragraph = model.Paragraphs.FirstOrDefault(paragraph => string.Equals(paragraph.Id, targetId, StringComparison.Ordinal));
        if (paragraph is not null)
        {
            return paragraph.Text;
        }

        DocxTableInfo? table = model.Tables.FirstOrDefault(table => string.Equals(table.Id, targetId, StringComparison.Ordinal));
        if (table is not null)
        {
            return string.Join(Environment.NewLine, table.Cells.Select(cell => $"{cell.Id}: {cell.Text}"));
        }

        DocxTableCellInfo? cell = model.Tables
            .SelectMany(table => table.Cells)
            .FirstOrDefault(cell => string.Equals(cell.Id, targetId, StringComparison.Ordinal));
        return cell?.Text;
    }

    private static string Escape(string text)
    {
        return text.Replace("\\", "\\\\", StringComparison.Ordinal).Replace("\"", "\\\"", StringComparison.Ordinal).Replace("\n", "\\n", StringComparison.Ordinal).Replace("\t", "\\t", StringComparison.Ordinal);
    }
}

