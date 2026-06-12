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
            string style = paragraph.StyleId is null
                ? string.Empty
                : $" styleId={Escape(paragraph.StyleId)}";
            string list = paragraph.List is null
                ? string.Empty
                : RenderList(paragraph.List);
            builder.Append(paragraph.Id).Append(' ').Append(kind).Append(style).Append(list).Append(" text=\"").Append(Escape(Truncate(paragraph.Text, maxText))).AppendLine("\"");
        }

        foreach (DocxTableInfo table in model.Tables)
        {
            builder.Append(table.Id).Append(" table rows=").Append(table.RowCount).Append(" columns=").Append(table.ColumnCount).AppendLine();
            foreach (DocxTableCellInfo cell in table.Cells)
            {
                string columnSpan = cell.ColumnSpan == 1 ? string.Empty : $" column-span={cell.ColumnSpan}";
                string verticalMerge = cell.VerticalMerge is null ? string.Empty : $" vertical-merge={cell.VerticalMerge}";
                string nestedTable = cell.HasNestedTable ? " nested-table=true" : string.Empty;
                builder.Append("  ").Append(cell.Id).Append(columnSpan).Append(verticalMerge).Append(nestedTable).Append(" text=\"").Append(Escape(Truncate(cell.Text, maxText))).AppendLine("\"");
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
                DocxRunInfo run = paragraph.Runs[i];
                string markup = run.MarkupType is null ? string.Empty : $" markup={run.MarkupType}";
                string revisionId = run.RevisionId is null ? string.Empty : $" revision-id={Escape(run.RevisionId)}";
                string author = run.Author is null ? string.Empty : $" author=\"{Escape(run.Author)}\"";
                string timestamp = run.TimestampUtc is null ? string.Empty : $" timestamp-utc={run.TimestampUtc:O}";
                string commentId = run.CommentId is null ? string.Empty : $" comment-id={Escape(run.CommentId)}";
                builder.Append("  ")
                    .Append(paragraph.Id)
                    .Append(".R")
                    .Append((i + 1).ToString("0000"))
                    .Append(markup)
                    .Append(revisionId)
                    .Append(author)
                    .Append(timestamp)
                    .Append(commentId)
                    .Append(" text=\"")
                    .Append(Escape(Truncate(run.Text, maxText)))
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

    public static IReadOnlyList<DocxDumpRunInfo> DumpRuns(DocxDocumentModel model, string targetId, int maxText)
    {
        DocxParagraphInfo? paragraph = model.Paragraphs.FirstOrDefault(paragraph => string.Equals(paragraph.Id, targetId, StringComparison.Ordinal));
        if (paragraph is null)
        {
            return [];
        }

        return paragraph.Runs
            .Select((run, index) => new DocxDumpRunInfo
            {
                Id = $"{paragraph.Id}.R{index + 1:0000}",
                Text = Truncate(run.Text, maxText),
                MarkupType = run.MarkupType,
                RevisionId = run.RevisionId,
                Author = run.Author,
                TimestampUtc = run.TimestampUtc,
                CommentId = run.CommentId
            })
            .ToArray();
    }

    public static IReadOnlyList<DocxContextItem> Context(DocxDocumentModel model, string targetId, int radius, int maxText)
    {
        radius = Math.Max(0, radius);

        DocxParagraphInfo? paragraph = model.Paragraphs.FirstOrDefault(paragraph => string.Equals(paragraph.Id, targetId, StringComparison.Ordinal));
        if (paragraph is not null)
        {
            DocxParagraphInfo[] storyParagraphs = model.Paragraphs
                .Where(candidate => string.Equals(candidate.Story, paragraph.Story, StringComparison.Ordinal))
                .ToArray();
            int index = Array.FindIndex(storyParagraphs, candidate => string.Equals(candidate.Id, targetId, StringComparison.Ordinal));
            return Window(storyParagraphs, index, radius)
                .Select(item => ToContextItem(item.Value, Relation(item.Offset), maxText))
                .ToArray();
        }

        DocxTableInfo? table = model.Tables.FirstOrDefault(table => string.Equals(table.Id, targetId, StringComparison.Ordinal));
        if (table is not null)
        {
            DocxTableInfo[] storyTables = model.Tables
                .Where(candidate => string.Equals(candidate.Story, table.Story, StringComparison.Ordinal))
                .ToArray();
            int index = Array.FindIndex(storyTables, candidate => string.Equals(candidate.Id, targetId, StringComparison.Ordinal));
            return Window(storyTables, index, radius)
                .Select(item => ToContextItem(item.Value, Relation(item.Offset)))
                .ToArray();
        }

        foreach (DocxTableInfo candidateTable in model.Tables)
        {
            DocxTableCellInfo? cell = candidateTable.Cells.FirstOrDefault(cell => string.Equals(cell.Id, targetId, StringComparison.Ordinal));
            if (cell is not null)
            {
                return CellContext(candidateTable, cell, radius, maxText);
            }

            DocxTableCellInfo[] rowCells = candidateTable.Cells
                .Where(cell => cell.Id.StartsWith($"{targetId}.C", StringComparison.Ordinal))
                .OrderBy(cell => cell.ColumnIndex)
                .ToArray();
            if (rowCells.Length > 0)
            {
                return RowContext(candidateTable, targetId, rowCells, radius, maxText);
            }
        }

        DocxSectionInfo? section = model.Sections.FirstOrDefault(section => string.Equals(section.Id, targetId, StringComparison.Ordinal));
        if (section is not null)
        {
            DocxSectionInfo[] storySections = model.Sections
                .Where(candidate => string.Equals(candidate.Story, section.Story, StringComparison.Ordinal))
                .ToArray();
            int index = Array.FindIndex(storySections, candidate => string.Equals(candidate.Id, targetId, StringComparison.Ordinal));
            return Window(storySections, index, radius)
                .Select(item => ToContextItem(item.Value, Relation(item.Offset)))
                .ToArray();
        }

        DocxImageInfo? image = model.Images.FirstOrDefault(image => string.Equals(image.Id, targetId, StringComparison.Ordinal));
        return image is null
            ? []
            : [ToContextItem(image, "target")];
    }

    public static string RenderContext(IReadOnlyList<DocxContextItem> items)
    {
        var builder = new StringBuilder();
        foreach (DocxContextItem item in items)
        {
            string story = string.IsNullOrWhiteSpace(item.Story) ? string.Empty : $" story=\"{Escape(item.Story)}\"";
            string parent = item.ParentId is null ? string.Empty : $" parent={item.ParentId}";
            string heading = item.HeadingLevel is null ? string.Empty : $" heading-level={item.HeadingLevel}";
            string style = item.StyleId is null ? string.Empty : $" styleId={Escape(item.StyleId)}";
            string list = item.List is null ? string.Empty : RenderList(item.List);
            string rowCount = item.RowCount is null ? string.Empty : $" rows={item.RowCount}";
            string columnCount = item.ColumnCount is null ? string.Empty : $" columns={item.ColumnCount}";
            string row = item.RowIndex is null ? string.Empty : $" row={item.RowIndex}";
            string column = item.ColumnIndex is null ? string.Empty : $" column={item.ColumnIndex}";
            string columnSpan = item.ColumnSpan is null or 1 ? string.Empty : $" column-span={item.ColumnSpan}";
            string verticalMerge = item.VerticalMerge is null ? string.Empty : $" vertical-merge={item.VerticalMerge}";
            string nestedTable = item.HasNestedTable ? " nested-table=true" : string.Empty;
            string text = item.Kind is "paragraph" or "cell"
                ? $" text=\"{Escape(item.Text)}\""
                : string.Empty;
            builder.Append(item.Relation)
                .Append(' ')
                .Append(item.Id)
                .Append(' ')
                .Append(item.Kind)
                .Append(story)
                .Append(parent)
                .Append(heading)
                .Append(style)
                .Append(list)
                .Append(rowCount)
                .Append(columnCount)
                .Append(row)
                .Append(column)
                .Append(columnSpan)
                .Append(verticalMerge)
                .Append(nestedTable)
                .Append(text)
                .AppendLine();
        }

        return builder.ToString();
    }

    public static IReadOnlyList<string> RenderStyles(IReadOnlyList<DocxStyleInfo> styles)
    {
        return styles
            .OrderBy(style => style.Type, StringComparer.Ordinal)
            .ThenBy(style => style.StyleId, StringComparer.Ordinal)
            .Select(style =>
            {
                string defaultText = style.IsDefault ? " default=true" : string.Empty;
                string basedOn = style.BasedOnStyleId is null ? string.Empty : $" based-on={Escape(style.BasedOnStyleId)}";
                string next = style.NextStyleId is null ? string.Empty : $" next={Escape(style.NextStyleId)}";
                string linked = style.LinkedStyleId is null ? string.Empty : $" linked={Escape(style.LinkedStyleId)}";
                string numbering = style.NumberingId is null
                    ? string.Empty
                    : $" numbering numId={Escape(style.NumberingId)} level={style.NumberingLevel ?? 0}";
                return $"{style.Type} styleId={style.StyleId} name=\"{Escape(style.Name)}\"{defaultText}{basedOn}{next}{linked}{numbering}";
            })
            .ToArray();
    }

    private static string RenderList(DocxListInfo list)
    {
        string abstractId = list.AbstractNumberingId is null ? string.Empty : $" abstractNumId={Escape(list.AbstractNumberingId)}";
        string format = list.Format is null ? string.Empty : $" format={Escape(list.Format)}";
        string levelText = list.LevelText is null ? string.Empty : $" level-text=\"{Escape(list.LevelText)}\"";
        string paragraphStyle = list.ParagraphStyleId is null ? string.Empty : $" paragraph-style={Escape(list.ParagraphStyleId)}";
        string source = string.Equals(list.Source, "direct", StringComparison.Ordinal)
            ? string.Empty
            : $" source={Escape(list.Source)}";
        return $" list numId={Escape(list.NumberingId)} level={list.Level}{abstractId}{format}{levelText}{paragraphStyle}{source}";
    }

    private static IReadOnlyList<DocxContextItem> CellContext(DocxTableInfo table, DocxTableCellInfo cell, int radius, int maxText)
    {
        var items = new List<DocxContextItem>
        {
            ToContextItem(table, "parent")
        };
        DocxTableCellInfo[] rowCells = table.Cells
            .Where(candidate => candidate.RowIndex == cell.RowIndex)
            .OrderBy(candidate => candidate.ColumnIndex)
            .ToArray();
        int index = Array.FindIndex(rowCells, candidate => string.Equals(candidate.Id, cell.Id, StringComparison.Ordinal));
        items.AddRange(Window(rowCells, index, radius)
            .Select(item => ToContextItem(item.Value, Relation(item.Offset), maxText, table.Id, table.Story)));
        return items;
    }

    private static IReadOnlyList<DocxContextItem> RowContext(DocxTableInfo table, string rowId, IReadOnlyList<DocxTableCellInfo> rowCells, int radius, int maxText)
    {
        var items = new List<DocxContextItem>
        {
            ToContextItem(table, "parent"),
            new()
            {
                Id = rowId,
                Kind = "row",
                Relation = "target",
                Story = table.Story,
                ParentId = table.Id,
                RowIndex = rowCells[0].RowIndex,
                ColumnCount = rowCells.Count
            }
        };

        items.AddRange(rowCells
            .Take(Math.Max(1, radius * 2 + 1))
            .Select(cell => ToContextItem(cell, "child", maxText, table.Id, table.Story)));
        return items;
    }

    private static IEnumerable<(T Value, int Offset)> Window<T>(IReadOnlyList<T> items, int index, int radius)
    {
        if (index < 0)
        {
            yield break;
        }

        int start = Math.Max(0, index - radius);
        int end = Math.Min(items.Count - 1, index + radius);
        for (int i = start; i <= end; i++)
        {
            yield return (items[i], i - index);
        }
    }

    private static string Relation(int offset)
    {
        return offset < 0 ? "before" : offset > 0 ? "after" : "target";
    }

    private static DocxContextItem ToContextItem(DocxParagraphInfo paragraph, string relation, int maxText)
    {
        return new DocxContextItem
        {
            Id = paragraph.Id,
            Kind = "paragraph",
            Relation = relation,
            Story = paragraph.Story,
            Text = Truncate(paragraph.Text, maxText),
            HeadingLevel = paragraph.HeadingLevel,
            StyleId = paragraph.StyleId,
            StyleName = paragraph.StyleName,
            List = paragraph.List
        };
    }

    private static DocxContextItem ToContextItem(DocxTableInfo table, string relation)
    {
        return new DocxContextItem
        {
            Id = table.Id,
            Kind = "table",
            Relation = relation,
            Story = table.Story,
            RowCount = table.RowCount,
            ColumnCount = table.ColumnCount
        };
    }

    private static DocxContextItem ToContextItem(DocxTableCellInfo cell, string relation, int maxText, string parentId, string story)
    {
        return new DocxContextItem
        {
            Id = cell.Id,
            Kind = "cell",
            Relation = relation,
            Story = story,
            ParentId = parentId,
            Text = Truncate(cell.Text, maxText),
            RowIndex = cell.RowIndex,
            ColumnIndex = cell.ColumnIndex,
            ColumnSpan = cell.ColumnSpan,
            VerticalMerge = cell.VerticalMerge,
            HasNestedTable = cell.HasNestedTable
        };
    }

    private static DocxContextItem ToContextItem(DocxSectionInfo section, string relation)
    {
        return new DocxContextItem
        {
            Id = section.Id,
            Kind = "section",
            Relation = relation,
            Story = section.Story,
            ColumnCount = section.Columns
        };
    }

    private static DocxContextItem ToContextItem(DocxImageInfo image, string relation)
    {
        return new DocxContextItem
        {
            Id = image.Id,
            Kind = "image",
            Relation = relation,
            ParentId = image.PartName,
            Text = image.ContentType ?? string.Empty
        };
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
