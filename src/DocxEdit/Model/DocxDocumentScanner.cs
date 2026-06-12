using System.Xml.Linq;
using DocxEdit.Ooxml;

namespace DocxEdit.Model;

internal static class DocxDocumentScanner
{
    public static DocxDocumentModel Scan(
        OoxmlPackage package,
        bool includeHeadersFooters = false,
        DocxTextView textView = DocxTextView.Final,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (package.MainDocumentPartName is null)
        {
            return DocxDocumentModel.Empty;
        }

        IReadOnlyList<DocxStyleInfo> styles = DocxStyleScanner.Scan(package, cancellationToken);
        IReadOnlyDictionary<string, DocxStyleInfo> stylesById = styles.ToDictionary(style => style.StyleId, StringComparer.Ordinal);
        DocxNumberingCatalog numbering = DocxNumberingCatalog.Scan(package, cancellationToken);
        var paragraphs = new List<DocxParagraphInfo>();
        var tables = new List<DocxTableInfo>();
        var images = new List<DocxImageInfo>();
        var sections = new List<DocxSectionInfo>();
        var bookmarks = new List<DocxBookmarkInfo>();
        var contentControls = new List<DocxContentControlInfo>();
        ScanStory(package, package.MainDocumentPartName, "M", "main", textView, stylesById, numbering, paragraphs, tables, images, sections, bookmarks, contentControls, cancellationToken);

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
                    ScanStory(package, relationship.ResolvedTarget!, prefix, $"header[{headerIndex - 1}]", textView, stylesById, numbering, paragraphs, tables, images, sections, bookmarks, contentControls, cancellationToken);
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
                    ScanStory(package, relationship.ResolvedTarget!, prefix, $"footer[{footerIndex - 1}]", textView, stylesById, numbering, paragraphs, tables, images, sections, bookmarks, contentControls, cancellationToken);
                }
            }
        }

        return new DocxDocumentModel(paragraphs, tables, images, sections, bookmarks, contentControls);
    }

    private static void ScanStory(
        OoxmlPackage package,
        string partName,
        string idPrefix,
        string story,
        DocxTextView textView,
        IReadOnlyDictionary<string, DocxStyleInfo> stylesById,
        DocxNumberingCatalog numbering,
        List<DocxParagraphInfo> paragraphs,
        List<DocxTableInfo> tables,
        List<DocxImageInfo> images,
        List<DocxSectionInfo> sections,
        List<DocxBookmarkInfo> bookmarks,
        List<DocxContentControlInfo> contentControls,
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
        var targets = new Dictionary<XElement, string>();
        foreach (XElement block in body.Elements())
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement? sectionProperties = null;
            if (block.Name == OoxmlNs.W + "p")
            {
                string paragraphId = $"{idPrefix}.P{paragraphIndex++:0000}";
                targets[block] = paragraphId;
                paragraphs.Add(ReadParagraph(block, paragraphId, story, textView, package, relationships, stylesById, numbering, images, idPrefix, ref imageIndex));
                sectionProperties = block.Element(OoxmlNs.W + "pPr")?.Element(OoxmlNs.W + "sectPr");
            }
            else if (block.Name == OoxmlNs.W + "tbl")
            {
                string tableId = $"{idPrefix}.T{tableIndex++:0000}";
                targets[block] = tableId;
                tables.Add(ReadTable(block, tableId, story, textView, package, relationships, images, idPrefix, targets, ref imageIndex));
            }
            else if (block.Name == OoxmlNs.W + "sectPr")
            {
                sectionProperties = block;
            }

            if (sectionProperties is not null && idPrefix == "M")
            {
                string sectionId = $"{idPrefix}.S{sectionIndex++:0000}";
                targets[sectionProperties] = sectionId;
                sections.Add(ReadSection(sectionProperties, sectionId, story));
            }
        }

        bookmarks.AddRange(ReadBookmarks(document, partName, story, idPrefix, targets));
        contentControls.AddRange(ReadContentControls(document, partName, story, idPrefix, textView, targets));
    }

    private static DocxParagraphInfo ReadParagraph(
        XElement paragraph,
        string id,
        string story,
        DocxTextView textView,
        OoxmlPackage package,
        IReadOnlyDictionary<string, OoxmlRelationship> relationships,
        IReadOnlyDictionary<string, DocxStyleInfo> stylesById,
        DocxNumberingCatalog numbering,
        List<DocxImageInfo> images,
        string imageIdPrefix,
        ref int imageIndex)
    {
        DocxRunInfo[] runs = ReadRuns(paragraph, textView);
        foreach (XElement drawing in paragraph.Descendants(OoxmlNs.W + "drawing"))
        {
            AddDrawingImages(drawing, package, relationships, images, imageIdPrefix, ref imageIndex);
        }

        string? styleId = ReadParagraphStyleId(paragraph);
        stylesById.TryGetValue(styleId ?? string.Empty, out DocxStyleInfo? style);
        return new DocxParagraphInfo(
            id,
            story,
            ReadText(paragraph, textView),
            ReadHeadingLevel(paragraph),
            ReadListInfo(paragraph, styleId, stylesById, numbering),
            runs)
        {
            StyleId = styleId,
            StyleName = style?.Name
        };
    }

    private static DocxRunInfo[] ReadRuns(XElement paragraph, DocxTextView textView)
    {
        var runs = new List<DocxRunInfo>();
        foreach (XElement child in paragraph.Elements())
        {
            if (child.Name == OoxmlNs.W + "r")
            {
                AddRunIfVisible(runs, child, textView, null);
            }
            else if (IsRevisionRunContainer(child))
            {
                string? markupType = GetRevisionMarkupType(child);
                foreach (XElement run in child.Elements(OoxmlNs.W + "r"))
                {
                    AddRunIfVisible(runs, run, textView, new RunMarkup(
                        markupType,
                        (string?)child.Attribute(OoxmlNs.W + "id"),
                        (string?)child.Attribute(OoxmlNs.W + "author"),
                        ParseDate((string?)child.Attribute(OoxmlNs.W + "date")),
                        null));
                }
            }
            else if (IsCommentMarker(child))
            {
                runs.Add(new DocxRunInfo(string.Empty)
                {
                    MarkupType = GetCommentMarkupType(child),
                    CommentId = (string?)child.Attribute(OoxmlNs.W + "id")
                });
            }
        }

        return runs.ToArray();
    }

    private static void AddRunIfVisible(List<DocxRunInfo> runs, XElement run, DocxTextView textView, RunMarkup? inheritedMarkup)
    {
        RunMarkup markup = ReadRunMarkup(run) ?? inheritedMarkup ?? RunMarkup.Empty;
        string text = ReadText(run, textView);
        if (text.Length == 0 && markup.IsEmpty)
        {
            return;
        }

        runs.Add(new DocxRunInfo(text)
        {
            MarkupType = markup.MarkupType,
            RevisionId = markup.RevisionId,
            Author = markup.Author,
            TimestampUtc = markup.TimestampUtc,
            CommentId = markup.CommentId
        });
    }

    private static RunMarkup? ReadRunMarkup(XElement run)
    {
        XElement? commentReference = run.Descendants(OoxmlNs.W + "commentReference").FirstOrDefault();
        if (commentReference is not null)
        {
            return new RunMarkup("comment-reference", null, null, null, (string?)commentReference.Attribute(OoxmlNs.W + "id"));
        }

        return null;
    }

    private static bool IsRevisionRunContainer(XElement element)
    {
        return element.Name.Namespace == OoxmlNs.W &&
            element.Name.LocalName is "ins" or "del" or "moveFrom" or "moveTo";
    }

    private static string? GetRevisionMarkupType(XElement element)
    {
        return element.Name.LocalName switch
        {
            "ins" => "inserted-run",
            "del" => "deleted-run",
            "moveFrom" => "move-from-run",
            "moveTo" => "move-to-run",
            _ => null
        };
    }

    private static bool IsCommentMarker(XElement element)
    {
        return element.Name.Namespace == OoxmlNs.W &&
            element.Name.LocalName is "commentRangeStart" or "commentRangeEnd";
    }

    private static string? GetCommentMarkupType(XElement element)
    {
        return element.Name.LocalName switch
        {
            "commentRangeStart" => "comment-range-start",
            "commentRangeEnd" => "comment-range-end",
            _ => null
        };
    }

    private static DocxTableInfo ReadTable(
        XElement table,
        string id,
        string story,
        DocxTextView textView,
        OoxmlPackage package,
        IReadOnlyDictionary<string, OoxmlRelationship> relationships,
        List<DocxImageInfo> images,
        string imageIdPrefix,
        Dictionary<XElement, string> targets,
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

                int columnSpan = ReadCellColumnSpan(cell);
                string cellId = $"{id}.R{rowIndex:00}.C{columnIndex:00}";
                targets[cell] = cellId;
                cells.Add(new DocxTableCellInfo(
                    cellId,
                    rowIndex,
                    columnIndex,
                    ReadText(cell, textView),
                    columnSpan,
                    ReadCellVerticalMerge(cell),
                    cell.Elements(OoxmlNs.W + "tbl").Any()));
                columnIndex += columnSpan;
            }

            maxColumns = Math.Max(maxColumns, columnIndex - 1);
            rowIndex++;
        }

        return new DocxTableInfo(id, story, rowIndex - 1, maxColumns, cells);
    }

    private static IReadOnlyList<DocxBookmarkInfo> ReadBookmarks(
        XDocument document,
        string partName,
        string story,
        string idPrefix,
        IReadOnlyDictionary<XElement, string> targets)
    {
        var bookmarks = new List<DocxBookmarkInfo>();
        var endsByOoxmlId = document
            .Descendants(OoxmlNs.W + "bookmarkEnd")
            .Select(end => ((string?)end.Attribute(OoxmlNs.W + "id"), end))
            .Where(pair => !string.IsNullOrWhiteSpace(pair.Item1))
            .GroupBy(pair => pair.Item1!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First().end, StringComparer.Ordinal);

        int bookmarkIndex = 1;
        foreach (XElement start in document.Descendants(OoxmlNs.W + "bookmarkStart"))
        {
            string? name = (string?)start.Attribute(OoxmlNs.W + "name");
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            string? ooxmlId = (string?)start.Attribute(OoxmlNs.W + "id");
            endsByOoxmlId.TryGetValue(ooxmlId ?? string.Empty, out XElement? end);
            bookmarks.Add(new DocxBookmarkInfo
            {
                Id = $"{idPrefix}.B{bookmarkIndex++:0000}",
                Name = name,
                OoxmlId = ooxmlId,
                Story = story,
                PartName = partName,
                StartTargetId = FindTargetId(start, targets),
                EndTargetId = end is null ? null : FindTargetId(end, targets),
                IsComplete = end is not null
            });
        }

        return bookmarks;
    }

    private static IReadOnlyList<DocxContentControlInfo> ReadContentControls(
        XDocument document,
        string partName,
        string story,
        string idPrefix,
        DocxTextView textView,
        IReadOnlyDictionary<XElement, string> targets)
    {
        var contentControls = new List<DocxContentControlInfo>();
        int controlIndex = 1;
        foreach (XElement control in document.Descendants(OoxmlNs.W + "sdt"))
        {
            XElement? properties = control.Element(OoxmlNs.W + "sdtPr");
            XElement content = control.Element(OoxmlNs.W + "sdtContent") ?? control;
            contentControls.Add(new DocxContentControlInfo
            {
                Id = $"{idPrefix}.CC{controlIndex++:0000}",
                Story = story,
                PartName = partName,
                TargetId = FindTargetId(control, targets),
                Kind = ReadContentControlKind(properties),
                OoxmlId = ReadSdtProperty(properties, "id"),
                Tag = ReadSdtProperty(properties, "tag"),
                Alias = ReadSdtProperty(properties, "alias"),
                Lock = ReadSdtProperty(properties, "lock"),
                TextLength = ReadText(content, textView).Length
            });
        }

        return contentControls;
    }

    private static string? FindTargetId(XElement element, IReadOnlyDictionary<XElement, string> targets)
    {
        foreach (XElement candidate in element.AncestorsAndSelf())
        {
            if (targets.TryGetValue(candidate, out string? id))
            {
                return id;
            }
        }

        foreach (XElement descendant in element.Descendants())
        {
            if (targets.TryGetValue(descendant, out string? id))
            {
                return id;
            }
        }

        return null;
    }

    private static string ReadContentControlKind(XElement? properties)
    {
        if (properties is null)
        {
            return "rich-text";
        }

        string? kind = properties.Elements()
            .Select(element => element.Name.LocalName)
            .FirstOrDefault(name => name is "text" or "richText" or "checkBox" or "dropDownList" or "comboBox" or "date" or "repeatingSection");
        return kind switch
        {
            "text" => "plain-text",
            "richText" => "rich-text",
            "checkBox" => "checkbox",
            "dropDownList" => "dropdown-list",
            "comboBox" => "combo-box",
            "date" => "date",
            "repeatingSection" => "repeating-section",
            _ => "rich-text"
        };
    }

    private static string? ReadSdtProperty(XElement? properties, string localName)
    {
        XElement? element = properties?.Elements().FirstOrDefault(element => element.Name.LocalName == localName);
        return (string?)element?.Attribute(OoxmlNs.W + "val");
    }

    private static string ReadText(XElement container, DocxTextView textView)
    {
        var buffer = new List<string>();
        foreach (XElement element in container.Descendants())
        {
            if (!ShouldIncludeTextElement(element, textView))
            {
                continue;
            }

            if (element.Name == OoxmlNs.W + "t" || element.Name == OoxmlNs.W + "delText")
            {
                buffer.Add(ApplyMarkupTextView(element, element.Value, textView));
            }
            else if (element.Name == OoxmlNs.W + "tab")
            {
                buffer.Add(ApplyMarkupTextView(element, "\t", textView));
            }
            else if (element.Name == OoxmlNs.W + "br")
            {
                buffer.Add(ApplyMarkupTextView(element, "\n", textView));
            }
        }

        return string.Concat(buffer);
    }

    private static bool ShouldIncludeTextElement(XElement element, DocxTextView textView)
    {
        return textView switch
        {
            DocxTextView.Final => !IsInsideDeletedRevision(element),
            DocxTextView.Original => !IsInsideInsertedRevision(element),
            DocxTextView.Markup => true,
            _ => !IsInsideDeletedRevision(element)
        };
    }

    private static string ApplyMarkupTextView(XElement element, string text, DocxTextView textView)
    {
        if (textView != DocxTextView.Markup)
        {
            return text;
        }

        if (IsInsideInsertedRevision(element))
        {
            return $"[+{text}+]";
        }

        if (IsInsideDeletedRevision(element))
        {
            return $"[-{text}-]";
        }

        return text;
    }

    private static int ReadCellColumnSpan(XElement cell)
    {
        string? spanText = (string?)cell
            .Element(OoxmlNs.W + "tcPr")
            ?.Element(OoxmlNs.W + "gridSpan")
            ?.Attribute(OoxmlNs.W + "val");
        return int.TryParse(spanText, out int span) && span > 0 ? span : 1;
    }

    private static string? ReadCellVerticalMerge(XElement cell)
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

    private static bool IsInsideDeletedRevision(XElement element)
    {
        return element.Ancestors(OoxmlNs.W + "del").Any() ||
            element.Ancestors(OoxmlNs.W + "moveFrom").Any();
    }

    private static bool IsInsideInsertedRevision(XElement element)
    {
        return element.Ancestors(OoxmlNs.W + "ins").Any() ||
            element.Ancestors(OoxmlNs.W + "moveTo").Any();
    }

    private static DateTimeOffset? ParseDate(string? value)
    {
        return DateTimeOffset.TryParse(value, out DateTimeOffset parsed)
            ? parsed.ToUniversalTime()
            : null;
    }

    private static int? ReadHeadingLevel(XElement paragraph)
    {
        string? styleId = ReadParagraphStyleId(paragraph);
        if (styleId is null)
        {
            return null;
        }

        string digits = new(styleId.Where(char.IsDigit).ToArray());
        return int.TryParse(digits, out int level) && level is >= 1 and <= 9
            ? level
            : null;
    }

    private static string? ReadParagraphStyleId(XElement paragraph)
    {
        return (string?)paragraph
            .Element(OoxmlNs.W + "pPr")
            ?.Element(OoxmlNs.W + "pStyle")
            ?.Attribute(OoxmlNs.W + "val");
    }

    private static DocxListInfo? ReadListInfo(
        XElement paragraph,
        string? styleId,
        IReadOnlyDictionary<string, DocxStyleInfo> stylesById,
        DocxNumberingCatalog numbering)
    {
        XElement? numberingProperties = paragraph
            .Element(OoxmlNs.W + "pPr")
            ?.Element(OoxmlNs.W + "numPr");
        string? numberingId = (string?)numberingProperties
            ?.Element(OoxmlNs.W + "numId")
            ?.Attribute(OoxmlNs.W + "val");
        string? levelText = (string?)numberingProperties
            ?.Element(OoxmlNs.W + "ilvl")
            ?.Attribute(OoxmlNs.W + "val");
        if (!string.IsNullOrWhiteSpace(numberingId))
        {
            int level = int.TryParse(levelText, out int parsedLevel) && parsedLevel >= 0
                ? parsedLevel
                : ResolveStyleNumbering(styleId, stylesById)?.Level ?? 0;
            return numbering.Resolve(numberingId, level, "direct");
        }

        StyleNumbering? styleNumbering = ResolveStyleNumbering(styleId, stylesById);
        return styleNumbering is null
            ? null
            : numbering.Resolve(styleNumbering.NumberingId, styleNumbering.Level, styleNumbering.Source);
    }

    private static StyleNumbering? ResolveStyleNumbering(
        string? styleId,
        IReadOnlyDictionary<string, DocxStyleInfo> stylesById)
    {
        if (string.IsNullOrWhiteSpace(styleId))
        {
            return null;
        }

        var visited = new HashSet<string>(StringComparer.Ordinal);
        string? current = styleId;
        bool inherited = false;
        while (!string.IsNullOrWhiteSpace(current) && visited.Add(current))
        {
            if (!stylesById.TryGetValue(current, out DocxStyleInfo? style))
            {
                return null;
            }

            if (!string.IsNullOrWhiteSpace(style.NumberingId))
            {
                return new StyleNumbering(
                    style.NumberingId,
                    style.NumberingLevel ?? 0,
                    inherited ? "style-inherited" : "style");
            }

            current = style.BasedOnStyleId;
            inherited = true;
        }

        return null;
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

    private sealed record RunMarkup(
        string? MarkupType,
        string? RevisionId,
        string? Author,
        DateTimeOffset? TimestampUtc,
        string? CommentId)
    {
        public static RunMarkup Empty { get; } = new(null, null, null, null, null);

        public bool IsEmpty =>
            MarkupType is null &&
            RevisionId is null &&
            Author is null &&
            TimestampUtc is null &&
            CommentId is null;
    }

    private sealed record StyleNumbering(string NumberingId, int Level, string Source);
}
