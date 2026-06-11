using System.Text;
using DocxEdit.Model;

namespace DocxEdit.Rendering;

internal static class TextRenderers
{
    public static string RenderRead(DocxDocumentModel model, int maxText)
    {
        var builder = new StringBuilder();
        foreach (DocxSectionInfo section in model.Sections)
        {
            builder.Append(section.Id).Append(" section columns=").Append(section.Columns).Append(" orientation=").Append(section.Orientation).AppendLine();
        }

        foreach (DocxParagraphInfo paragraph in model.Paragraphs)
        {
            string kind = paragraph.HeadingLevel is null ? "paragraph" : $"heading level={paragraph.HeadingLevel}";
            builder.Append(paragraph.Id).Append(' ').Append(kind).Append(" text=\"").Append(Escape(Truncate(paragraph.Text, maxText))).AppendLine("\"");
        }

        foreach (DocxTableInfo table in model.Tables)
        {
            builder.Append(table.Id).Append(" table rows=").Append(table.RowCount).Append(" columns=").Append(table.ColumnCount).AppendLine();
            foreach (DocxTableCellInfo cell in table.Cells)
            {
                builder.Append("  ").Append(cell.Id).Append(" text=\"").Append(Escape(Truncate(cell.Text, maxText))).AppendLine("\"");
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

        foreach (DocxSectionInfo section in model.Sections)
        {
            lines.Add($"{section.Id} section columns={section.Columns} orientation={section.Orientation}");
        }

        foreach (DocxImageInfo image in model.Images)
        {
            lines.Add($"{image.Id} image part={image.PartName}");
        }

        return lines;
    }

    public static IReadOnlyList<string> Find(DocxDocumentModel model, string query, int maxText)
    {
        var matches = new List<string>();
        foreach (DocxParagraphInfo paragraph in model.Paragraphs)
        {
            if (paragraph.Text.Contains(query, StringComparison.OrdinalIgnoreCase))
            {
                matches.Add($"{paragraph.Id} text=\"{Escape(Truncate(paragraph.Text, maxText))}\"");
            }
        }

        foreach (DocxTableInfo table in model.Tables)
        {
            foreach (DocxTableCellInfo cell in table.Cells)
            {
                if (cell.Text.Contains(query, StringComparison.OrdinalIgnoreCase))
                {
                    matches.Add($"{cell.Id} text=\"{Escape(Truncate(cell.Text, maxText))}\"");
                }
            }
        }

        return matches;
    }

    public static string? Dump(DocxDocumentModel model, string targetId, bool includeRuns, int maxText)
    {
        DocxParagraphInfo? paragraph = model.Paragraphs.FirstOrDefault(paragraph => string.Equals(paragraph.Id, targetId, StringComparison.Ordinal));
        if (paragraph is not null)
        {
            if (!includeRuns)
            {
                return Truncate(paragraph.Text, maxText);
            }

            var builder = new StringBuilder();
            builder.Append("text=\"").Append(Escape(Truncate(paragraph.Text, maxText))).AppendLine("\"");
            builder.AppendLine("runs:");
            for (int i = 0; i < paragraph.Runs.Count; i++)
            {
                builder.Append("  ")
                    .Append(paragraph.Id)
                    .Append(".R")
                    .Append((i + 1).ToString("0000"))
                    .Append(" text=\"")
                    .Append(Escape(Truncate(paragraph.Runs[i].Text, maxText)))
                    .AppendLine("\"");
            }

            return builder.ToString();
        }

        DocxTableInfo? table = model.Tables.FirstOrDefault(table => string.Equals(table.Id, targetId, StringComparison.Ordinal));
        if (table is not null)
        {
            return string.Join(Environment.NewLine, table.Cells.Select(cell => $"{cell.Id}: {Truncate(cell.Text, maxText)}"));
        }

        DocxTableCellInfo? cell = model.Tables
            .SelectMany(table => table.Cells)
            .FirstOrDefault(cell => string.Equals(cell.Id, targetId, StringComparison.Ordinal));
        return cell is null ? null : Truncate(cell.Text, maxText);
    }

    public static IReadOnlyList<string> RenderStyles(IReadOnlyList<DocxStyleInfo> styles)
    {
        return styles
            .OrderBy(style => style.Type, StringComparer.Ordinal)
            .ThenBy(style => style.StyleId, StringComparer.Ordinal)
            .Select(style =>
            {
                string defaultText = style.IsDefault ? " default=true" : string.Empty;
                return $"{style.Type} styleId={style.StyleId} name=\"{Escape(style.Name)}\"{defaultText}";
            })
            .ToArray();
    }

    private static string Escape(string text)
    {
        return text.Replace("\\", "\\\\", StringComparison.Ordinal).Replace("\"", "\\\"", StringComparison.Ordinal).Replace("\n", "\\n", StringComparison.Ordinal).Replace("\t", "\\t", StringComparison.Ordinal);
    }

    private static string Truncate(string text, int maxText)
    {
        if (maxText < 0 || text.Length <= maxText)
        {
            return text;
        }

        if (maxText <= 3)
        {
            return text[..maxText];
        }

        return text[..(maxText - 3)] + "...";
    }
}
