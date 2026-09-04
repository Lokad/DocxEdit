using System.Globalization;
using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

internal static partial class DocxDocumentScanner
{
    public static DocxDocumentModel Scan(
        OoxmlPackage package,
        bool includeHeadersFooters,
        DocxTextView textView,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        IReadOnlyList<DocxStyleInfo> styles = DocxStyleScanner.Scan(package, cancellationToken);
        IReadOnlyDictionary<string, DocxStyleInfo> stylesById = styles.ToDictionary(style => style.StyleId, StringComparer.Ordinal);
        DocxNumberingCatalog numbering = DocxNumberingCatalog.Scan(package, cancellationToken);
        var paragraphs = new List<DocxParagraphInfo>();
        var tables = new List<DocxTableInfo>();
        var images = new List<DocxImageInfo>();
        var sections = new List<DocxSectionInfo>();
        var bookmarks = new List<DocxBookmarkInfo>();
        var contentControls = new List<DocxContentControlInfo>();
        var fields = new List<DocxFieldInfo>();
        var hyperlinks = new List<DocxHyperlinkInfo>();
        ScanStory(package, package.MainDocumentPartName, "M", "main", textView, stylesById, numbering, paragraphs, tables, images, sections, bookmarks, contentControls, fields, hyperlinks, cancellationToken);

        if (includeHeadersFooters)
        {
            IReadOnlyList<ResolvedOoxmlRelationship> relationships = package.GetResolvedRelationships(package.MainDocumentPartName, cancellationToken);
            int headerIndex = 1;
            foreach (ResolvedOoxmlRelationship relationship in relationships
                .Where(relationship => relationship.Type == OoxmlRelTypes.Header)
                .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
            {
                if (package.GetPart(relationship.ResolvedTarget) is not null)
                {
                    string prefix = $"H{headerIndex++:000}";
                    ScanStory(package, relationship.ResolvedTarget, prefix, $"header[{headerIndex - 1}]", textView, stylesById, numbering, paragraphs, tables, images, sections, bookmarks, contentControls, fields, hyperlinks, cancellationToken);
                }
            }

            int footerIndex = 1;
            foreach (ResolvedOoxmlRelationship relationship in relationships
                .Where(relationship => relationship.Type == OoxmlRelTypes.Footer)
                .OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
            {
                if (package.GetPart(relationship.ResolvedTarget) is not null)
                {
                    string prefix = $"F{footerIndex++:000}";
                    ScanStory(package, relationship.ResolvedTarget, prefix, $"footer[{footerIndex - 1}]", textView, stylesById, numbering, paragraphs, tables, images, sections, bookmarks, contentControls, fields, hyperlinks, cancellationToken);
                }
            }
        }

        return new DocxDocumentModel(paragraphs, tables, images, sections, bookmarks, contentControls, fields, hyperlinks);
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
        List<DocxFieldInfo> fields,
        List<DocxHyperlinkInfo> hyperlinks,
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
        var numberingLabeler = new DocxNumberingLabeler(numbering);
        foreach (XElement block in EnumerateStoryBlocks(body, textView))
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement? sectionProperties = null;
            if (block.Name == OoxmlNs.W + "p")
            {
                string paragraphId = $"{idPrefix}.P{paragraphIndex++:0000}";
                targets[block] = paragraphId;
                paragraphs.Add(ReadParagraph(block, paragraphId, story, textView, package, relationships, stylesById, numbering, numberingLabeler, images, idPrefix, ref imageIndex));
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
        fields.AddRange(ReadFields(document, partName, story, idPrefix, textView, targets));
        hyperlinks.AddRange(ReadHyperlinks(document, partName, story, idPrefix, textView, targets, relationships));
    }

    private static IEnumerable<XElement> EnumerateStoryBlocks(XElement body, DocxTextView textView)
    {
        foreach (XElement block in body.Elements())
        {
            if (IsStoryBlock(block))
            {
                yield return block;
                continue;
            }

            if (!IsRevisionBlockContainer(block) || !ShouldIncludeRevisionBlock(block, textView))
            {
                continue;
            }

            foreach (XElement revisionBlock in block.Elements().Where(IsStoryBlock))
            {
                yield return revisionBlock;
            }
        }
    }

    private static bool IsStoryBlock(XElement element)
    {
        return element.Name == OoxmlNs.W + "p" ||
            element.Name == OoxmlNs.W + "tbl" ||
            element.Name == OoxmlNs.W + "sectPr";
    }

    private static bool IsRevisionBlockContainer(XElement element)
    {
        return element.Name.Namespace == OoxmlNs.W &&
            element.Name.LocalName is "ins" or "del" or "moveFrom" or "moveTo" &&
            element.Elements().Any(IsStoryBlock);
    }

    private static bool ShouldIncludeRevisionBlock(XElement element, DocxTextView textView)
    {
        return textView switch
        {
            DocxTextView.Final => element.Name.LocalName is not ("del" or "moveFrom"),
            DocxTextView.Original => element.Name.LocalName is not ("ins" or "moveTo"),
            DocxTextView.Markup => true,
            _ => element.Name.LocalName is not ("del" or "moveFrom")
        };
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
        DocxNumberingLabeler numberingLabeler,
        List<DocxImageInfo> images,
        string imageIdPrefix,
        ref int imageIndex)
    {
        DocxRunInfo[] runs = ReadRuns(paragraph, textView);
        foreach (XElement drawing in paragraph.Descendants(OoxmlNs.W + "drawing"))
        {
            AddDrawingImages(drawing, package, relationships, images, imageIdPrefix, id, ref imageIndex);
        }

        string? styleId = ReadParagraphStyleId(paragraph);
        stylesById.TryGetValue(styleId ?? string.Empty, out DocxStyleInfo? style);
        return new DocxParagraphInfo(
            id,
            story,
            ReadText(paragraph, textView),
            ReadHeadingLevel(paragraph),
            numberingLabeler.ApplyLabel(ReadListInfo(paragraph, styleId, stylesById, numbering)),
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
                        null,
                        null,
                        null));
                }
            }
            else if (child.Name == OoxmlNs.W + "hyperlink")
            {
                var hyperlinkMarkup = new RunMarkup(
                    "hyperlink",
                    null,
                    null,
                    null,
                    null,
                    (string?)child.Attribute(OoxmlNs.R + "id"),
                    (string?)child.Attribute(OoxmlNs.W + "anchor"));
                foreach (XElement run in child.Elements(OoxmlNs.W + "r"))
                {
                    AddRunIfVisible(runs, run, textView, hyperlinkMarkup);
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
            CommentId = markup.CommentId,
            HyperlinkRelationshipId = markup.HyperlinkRelationshipId,
            HyperlinkAnchor = markup.HyperlinkAnchor
        });
    }

    private static RunMarkup? ReadRunMarkup(XElement run)
    {
        XElement? commentReference = run.Descendants(OoxmlNs.W + "commentReference").FirstOrDefault();
        if (commentReference is not null)
        {
            return new RunMarkup("comment-reference", null, null, null, (string?)commentReference.Attribute(OoxmlNs.W + "id"), null, null);
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
        string containingTargetId,
        ref int imageIndex)
    {
        XElement? layout = drawing.Element(OoxmlNs.Wp + "inline") ?? drawing.Element(OoxmlNs.Wp + "anchor");
        string layoutKind = layout?.Name.LocalName ?? "unknown";
        XElement? extent = layout?.Element(OoxmlNs.Wp + "extent");
        XElement? docProperties = layout?.Element(OoxmlNs.Wp + "docPr");
        XElement? horizontalPosition = layout?.Element(OoxmlNs.Wp + "positionH");
        XElement? verticalPosition = layout?.Element(OoxmlNs.Wp + "positionV");
        XElement? lockProperties = layout?.Descendants(OoxmlNs.A + "graphicFrameLocks").FirstOrDefault() ??
            layout?.Descendants(OoxmlNs.A + "picLocks").FirstOrDefault();
        string? wrapMode = layout?.Elements().FirstOrDefault(element => element.Name.LocalName.StartsWith("wrap", StringComparison.Ordinal))?.Name.LocalName;
        bool behindDoc = WmlBoolean.IsTrue((string?)layout?.Attribute("behindDoc"), valueWhenMissing: false);

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

            XElement? sourceRectangle = blip
                .Ancestors(OoxmlNs.Pic + "blipFill")
                .FirstOrDefault()
                ?.Element(OoxmlNs.A + "srcRect");
            images.Add(new DocxImageInfo($"{imageIdPrefix}.I{imageIndex++:0000}", imagePart.Name, imagePart.ContentType, imagePart.Bytes.Length)
            {
                LayoutKind = layoutKind,
                RelationshipId = relationshipId,
                ContainingTargetId = containingTargetId,
                WidthEmu = ReadLongAttribute(extent, "cx"),
                HeightEmu = ReadLongAttribute(extent, "cy"),
                Name = (string?)docProperties?.Attribute("name"),
                Description = (string?)docProperties?.Attribute("descr"),
                Title = (string?)docProperties?.Attribute("title"),
                WrapMode = wrapMode,
                BehindDoc = behindDoc,
                WrapDistanceTopEmu = ReadLongAttribute(layout, "distT"),
                WrapDistanceBottomEmu = ReadLongAttribute(layout, "distB"),
                WrapDistanceLeftEmu = ReadLongAttribute(layout, "distL"),
                WrapDistanceRightEmu = ReadLongAttribute(layout, "distR"),
                RelativeHeight = ReadLongAttribute(layout, "relativeHeight"),
                AllowOverlap = ReadBooleanAttribute(layout, "allowOverlap"),
                LockAspectRatio = ReadBooleanAttribute(lockProperties, "noChangeAspect"),
                HorizontalPositionRelativeFrom = (string?)horizontalPosition?.Attribute("relativeFrom"),
                HorizontalPositionOffsetEmu = ReadPositionOffset(horizontalPosition),
                HorizontalPositionAlign = ReadPositionAlign(horizontalPosition),
                VerticalPositionRelativeFrom = (string?)verticalPosition?.Attribute("relativeFrom"),
                VerticalPositionOffsetEmu = ReadPositionOffset(verticalPosition),
                VerticalPositionAlign = ReadPositionAlign(verticalPosition),
                CropLeftPercent = ReadCropPercent(sourceRectangle, "l"),
                CropTopPercent = ReadCropPercent(sourceRectangle, "t"),
                CropRightPercent = ReadCropPercent(sourceRectangle, "r"),
                CropBottomPercent = ReadCropPercent(sourceRectangle, "b")
            });
        }
    }

    private static long? ReadLongAttribute(XElement? element, string localName)
    {
        return long.TryParse((string?)element?.Attribute(localName), NumberStyles.Integer, CultureInfo.InvariantCulture, out long value) ? value : null;
    }

    private static bool? ReadBooleanAttribute(XElement? element, string localName)
    {
        string? value = (string?)element?.Attribute(localName);
        if (value is null)
        {
            return null;
        }

        return value is "1" or "true" or "on";
    }

    private static long? ReadPositionOffset(XElement? element)
    {
        string? value = element?.Element(OoxmlNs.Wp + "posOffset")?.Value;
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long offset)
            ? offset
            : null;
    }

    private static string? ReadPositionAlign(XElement? element)
    {
        string? value = element?.Element(OoxmlNs.Wp + "align")?.Value;
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    private static decimal? ReadCropPercent(XElement? element, string localName)
    {
        string? value = (string?)element?.Attribute(localName);
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        if (value.EndsWith("%", StringComparison.Ordinal))
        {
            return decimal.TryParse(value[..^1], NumberStyles.Number, CultureInfo.InvariantCulture, out decimal percent)
                ? percent
                : null;
        }

        return decimal.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out decimal perThousandPercent)
            ? perThousandPercent / 1000m
            : null;
    }

    private sealed record RunMarkup(
        string? MarkupType,
        string? RevisionId,
        string? Author,
        DateTimeOffset? TimestampUtc,
        string? CommentId,
        string? HyperlinkRelationshipId,
        string? HyperlinkAnchor)
    {
        public static RunMarkup Empty { get; } = new(null, null, null, null, null, null, null);

        public bool IsEmpty =>
            MarkupType is null &&
            RevisionId is null &&
            Author is null &&
            TimestampUtc is null &&
            CommentId is null &&
            HyperlinkRelationshipId is null &&
            HyperlinkAnchor is null;
    }

    private sealed record StyleNumbering(string NumberingId, int Level, string Source);

    private sealed record TableMergeState(string MergeGroupId, string? RootCellId);

    private sealed record ContentControlScanEntry(XElement Element, DocxContentControlInfo Info);

    private sealed record HyperlinkUriValidation(string? Scheme, bool? IsValid, string? Reason)
    {
        public static HyperlinkUriValidation None { get; } = new(null, null, null);
    }

    private sealed record FieldCodeMetadata(
        string? FieldType,
        IReadOnlyList<string> Arguments,
        IReadOnlyList<string> Switches,
        IReadOnlyList<string> BookmarkDependencies,
        IReadOnlyList<string> HyperlinkDependencies,
        string RefreshPolicy,
        string? RefreshReason,
        bool CanRefreshDeterministically);

    private sealed record FieldRefreshProfile(string Policy, string? Reason, bool CanRefreshDeterministically);

    private sealed class ComplexFieldBuilder(XElement startElement, string? targetId, int nestingDepth)
    {
        public XElement StartElement { get; } = startElement;
        public string? TargetId { get; } = targetId;
        public int NestingDepth { get; } = nestingDepth;
        public System.Text.StringBuilder Code { get; } = new();
        public System.Text.StringBuilder ResultText { get; } = new();
        public bool HasSeparate { get; set; }
        public bool? IsDirty { get; init; }
        public bool? IsLocked { get; init; }

        public DocxFieldInfo ToInfo(string id, string story, string partName, bool complete)
        {
            _ = StartElement;
            string code = NormalizeFieldCode(Code.ToString());
            string resultText = ResultText.ToString();
            FieldCodeMetadata metadata = AnalyzeFieldCode(code);
            return new DocxFieldInfo
            {
                Id = id,
                Story = story,
                PartName = partName,
                TargetId = TargetId,
                Kind = "complex",
                FieldType = metadata.FieldType,
                Code = code,
                Arguments = metadata.Arguments,
                Switches = metadata.Switches,
                CachedResultText = resultText,
                ResultTextLength = resultText.Length,
                NestingDepth = NestingDepth,
                BookmarkDependencies = metadata.BookmarkDependencies,
                HyperlinkDependencies = metadata.HyperlinkDependencies,
                RefreshPolicy = metadata.RefreshPolicy,
                RefreshReason = metadata.RefreshReason,
                CanRefreshDeterministically = metadata.CanRefreshDeterministically,
                SafeEditStatus = DetermineFieldSafeEditStatus("complex", complete, IsLocked),
                IsDirty = IsDirty,
                IsLocked = IsLocked,
                IsComplete = complete
            };
        }
    }
}
