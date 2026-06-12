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

        foreach (DocxBookmarkInfo bookmark in model.Bookmarks)
        {
            string ooxmlId = bookmark.OoxmlId is null ? string.Empty : $" ooxml-id={Escape(bookmark.OoxmlId)}";
            string start = bookmark.StartTargetId is null ? " start=unknown" : $" start={bookmark.StartTargetId}";
            string end = bookmark.EndTargetId is null ? " end=unknown" : $" end={bookmark.EndTargetId}";
            builder.Append(bookmark.Id)
                .Append(" bookmark name=\"")
                .Append(Escape(bookmark.Name))
                .Append('"')
                .Append(ooxmlId)
                .Append(" story=\"")
                .Append(Escape(bookmark.Story))
                .Append("\" part=")
                .Append(bookmark.PartName)
                .Append(start)
                .Append(end)
                .Append(" complete=")
                .Append(bookmark.IsComplete)
                .AppendLine();
        }

        foreach (DocxContentControlInfo control in model.ContentControls)
        {
            string target = control.TargetId is null ? " target=unknown" : $" target={control.TargetId}";
            string ooxmlId = control.OoxmlId is null ? string.Empty : $" ooxml-id={Escape(control.OoxmlId)}";
            string tag = control.Tag is null ? string.Empty : $" tag=\"{Escape(control.Tag)}\"";
            string alias = control.Alias is null ? string.Empty : $" alias=\"{Escape(control.Alias)}\"";
            string locked = control.Lock is null ? string.Empty : $" lock={Escape(control.Lock)}";
            builder.Append(control.Id)
                .Append(" content-control kind=")
                .Append(Escape(control.Kind))
                .Append(" story=\"")
                .Append(Escape(control.Story))
                .Append("\" part=")
                .Append(control.PartName)
                .Append(target)
                .Append(ooxmlId)
                .Append(tag)
                .Append(alias)
                .Append(locked)
                .Append(" text-length=")
                .Append(control.TextLength)
                .AppendLine();
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

        foreach (DocxBookmarkInfo bookmark in model.Bookmarks)
        {
            string start = bookmark.StartTargetId is null ? "unknown" : bookmark.StartTargetId;
            string end = bookmark.EndTargetId is null ? "unknown" : bookmark.EndTargetId;
            lines.Add($"{bookmark.Id} bookmark name=\"{Escape(bookmark.Name)}\" start={start} end={end}");
        }

        foreach (DocxContentControlInfo control in model.ContentControls)
        {
            string target = control.TargetId is null ? "unknown" : control.TargetId;
            string tag = control.Tag is null ? string.Empty : $" tag=\"{Escape(control.Tag)}\"";
            string alias = control.Alias is null ? string.Empty : $" alias=\"{Escape(control.Alias)}\"";
            lines.Add($"{control.Id} content-control kind={Escape(control.Kind)} target={target}{tag}{alias}");
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
        TargetAnnotations annotations = BuildTargetAnnotations(model);

        DocxParagraphInfo? paragraph = model.Paragraphs.FirstOrDefault(paragraph => string.Equals(paragraph.Id, targetId, StringComparison.Ordinal));
        if (paragraph is not null)
        {
            DocxParagraphInfo[] storyParagraphs = model.Paragraphs
                .Where(candidate => string.Equals(candidate.Story, paragraph.Story, StringComparison.Ordinal))
                .ToArray();
            int index = Array.FindIndex(storyParagraphs, candidate => string.Equals(candidate.Id, targetId, StringComparison.Ordinal));
            return Window(storyParagraphs, index, radius)
                .Select(item => ApplyAnnotations(ToContextItem(item.Value, Relation(item.Offset), maxText), annotations))
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
                .Select(item => ApplyAnnotations(ToContextItem(item.Value, Relation(item.Offset)), annotations))
                .ToArray();
        }

        foreach (DocxTableInfo candidateTable in model.Tables)
        {
            DocxTableCellInfo? cell = candidateTable.Cells.FirstOrDefault(cell => string.Equals(cell.Id, targetId, StringComparison.Ordinal));
            if (cell is not null)
            {
                return CellContext(candidateTable, cell, radius, maxText, annotations);
            }

            DocxTableCellInfo[] rowCells = candidateTable.Cells
                .Where(cell => cell.Id.StartsWith($"{targetId}.C", StringComparison.Ordinal))
                .OrderBy(cell => cell.ColumnIndex)
                .ToArray();
            if (rowCells.Length > 0)
            {
                return RowContext(candidateTable, targetId, rowCells, radius, maxText, annotations);
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
                .Select(item => ApplyAnnotations(ToContextItem(item.Value, Relation(item.Offset)), annotations))
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
            string bookmarks = item.BookmarkNames.Count == 0 ? string.Empty : $" bookmark-names=\"{Escape(string.Join(",", item.BookmarkNames))}\"";
            string contentControlIds = item.ContentControlIds.Count == 0 ? string.Empty : $" content-controls=\"{Escape(string.Join(",", item.ContentControlIds))}\"";
            string contentControlTags = item.ContentControlTags.Count == 0 ? string.Empty : $" content-control-tags=\"{Escape(string.Join(",", item.ContentControlTags))}\"";
            string contentControlAliases = item.ContentControlAliases.Count == 0 ? string.Empty : $" content-control-aliases=\"{Escape(string.Join(",", item.ContentControlAliases))}\"";
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
                .Append(bookmarks)
                .Append(contentControlIds)
                .Append(contentControlTags)
                .Append(contentControlAliases)
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

    private static IReadOnlyList<DocxContextItem> CellContext(DocxTableInfo table, DocxTableCellInfo cell, int radius, int maxText, TargetAnnotations annotations)
    {
        var items = new List<DocxContextItem>
        {
            ApplyAnnotations(ToContextItem(table, "parent"), annotations)
        };
        DocxTableCellInfo[] rowCells = table.Cells
            .Where(candidate => candidate.RowIndex == cell.RowIndex)
            .OrderBy(candidate => candidate.ColumnIndex)
            .ToArray();
        int index = Array.FindIndex(rowCells, candidate => string.Equals(candidate.Id, cell.Id, StringComparison.Ordinal));
        items.AddRange(Window(rowCells, index, radius)
            .Select(item => ApplyAnnotations(ToContextItem(item.Value, Relation(item.Offset), maxText, table.Id, table.Story), annotations)));
        return items;
    }

    private static IReadOnlyList<DocxContextItem> RowContext(DocxTableInfo table, string rowId, IReadOnlyList<DocxTableCellInfo> rowCells, int radius, int maxText, TargetAnnotations annotations)
    {
        var items = new List<DocxContextItem>
        {
            ApplyAnnotations(ToContextItem(table, "parent"), annotations),
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
            .Select(cell => ApplyAnnotations(ToContextItem(cell, "child", maxText, table.Id, table.Story), annotations)));
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

    private static TargetAnnotations BuildTargetAnnotations(DocxDocumentModel model)
    {
        var bookmarkNames = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        foreach (DocxBookmarkInfo bookmark in model.Bookmarks)
        {
            AddAnnotation(bookmarkNames, bookmark.StartTargetId, bookmark.Name);
            AddAnnotation(bookmarkNames, bookmark.EndTargetId, bookmark.Name);
        }

        var contentControlIds = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var contentControlTags = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var contentControlAliases = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        foreach (DocxContentControlInfo control in model.ContentControls)
        {
            AddAnnotation(contentControlIds, control.TargetId, control.Id);
            AddAnnotation(contentControlTags, control.TargetId, control.Tag);
            AddAnnotation(contentControlAliases, control.TargetId, control.Alias);
        }

        return new TargetAnnotations(
            ToArrayDictionary(bookmarkNames),
            ToArrayDictionary(contentControlIds),
            ToArrayDictionary(contentControlTags),
            ToArrayDictionary(contentControlAliases));
    }

    private static DocxContextItem ApplyAnnotations(DocxContextItem item, TargetAnnotations annotations)
    {
        return item with
        {
            BookmarkNames = LookupAnnotations(annotations.BookmarkNamesByTarget, item.Id),
            ContentControlIds = LookupAnnotations(annotations.ContentControlIdsByTarget, item.Id),
            ContentControlTags = LookupAnnotations(annotations.ContentControlTagsByTarget, item.Id),
            ContentControlAliases = LookupAnnotations(annotations.ContentControlAliasesByTarget, item.Id)
        };
    }

    private static void AddAnnotation(Dictionary<string, List<string>> annotations, string? targetId, string? value)
    {
        if (string.IsNullOrWhiteSpace(targetId) || string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        if (!annotations.TryGetValue(targetId, out List<string>? values))
        {
            values = [];
            annotations[targetId] = values;
        }

        if (!values.Contains(value, StringComparer.Ordinal))
        {
            values.Add(value);
        }
    }

    private static IReadOnlyDictionary<string, string[]> ToArrayDictionary(Dictionary<string, List<string>> annotations)
    {
        return annotations.ToDictionary(
            pair => pair.Key,
            pair => pair.Value.Order(StringComparer.Ordinal).ToArray(),
            StringComparer.Ordinal);
    }

    private static IReadOnlyList<string> LookupAnnotations(IReadOnlyDictionary<string, string[]> annotations, string targetId)
    {
        return annotations.TryGetValue(targetId, out string[]? values) ? values : [];
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

    private sealed record TargetAnnotations(
        IReadOnlyDictionary<string, string[]> BookmarkNamesByTarget,
        IReadOnlyDictionary<string, string[]> ContentControlIdsByTarget,
        IReadOnlyDictionary<string, string[]> ContentControlTagsByTarget,
        IReadOnlyDictionary<string, string[]> ContentControlAliasesByTarget);
}
