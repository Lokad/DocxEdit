using System.Globalization;
using System.Xml.Linq;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit.Model;

// Reads the package into an inspection model. Missing parts and unrecognized
// entries yield empty collections; malformed document data throws a document
// exception (InvalidDataException, XmlException, IOException) at an explicit
// boundary, which the editor converts into failed results with diagnostics.
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
        // Style IDs come from the document, so uniqueness is validated here at
        // the dictionary-construction boundary instead of throwing ArgumentException.
        string? duplicateStyleId = styles
            .GroupBy(style => style.StyleId, StringComparer.Ordinal)
            .FirstOrDefault(group => group.Count() > 1)
            ?.Key;
        if (duplicateStyleId is not null)
        {
            throw new InvalidDataException($"Duplicate style ID '{duplicateStyleId}'.");
        }

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
            foreach (StoryPartRef story in DocxPartRoles.GetOrderedStories(package, includeHeadersFooters: true, cancellationToken))
            {
                if (story.Prefix == "M")
                {
                    continue;
                }

                ScanStory(package, story.PartName, story.Prefix, story.StoryLabel, textView, stylesById, numbering, paragraphs, tables, images, sections, bookmarks, contentControls, fields, hyperlinks, cancellationToken);
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
        (char storyLetter, int storyPartNumber) = DocxTargetId.ParseStoryPrefix(idPrefix);
        int sectionIndex = sections.Count(section => section.Id.Kind == DocxTargetKind.Section && section.Id.Story == storyLetter && section.Id.StoryPart == storyPartNumber) + 1;
        var targets = new Dictionary<XElement, string>();
        var numberingLabeler = new DocxNumberingLabeler(numbering);
        foreach (DocxStoryBlocks.StoryBlock entry in DocxStoryBlocks.EnumeratePhysicalBlocks(body))
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement block = entry.Block;
            bool visible = DocxStoryBlocks.IsVisibleInView(entry.Wrapper, textView);
            XElement? sectionProperties = null;
            if (block.Name == OoxmlNs.W + "p")
            {
                DocxTargetId paragraphId = new DocxTargetId(storyLetter, storyPartNumber, DocxTargetKind.Paragraph, paragraphIndex++, 0, 0);
                targets[block] = paragraphId.ToWireValue();
                if (visible)
                {
                    paragraphs.Add(ReadParagraph(block, paragraphId, story, textView, package, relationships, stylesById, numbering, numberingLabeler, images, ref imageIndex, cancellationToken));
                    sectionProperties = block.Element(OoxmlNs.W + "pPr")?.Element(OoxmlNs.W + "sectPr");
                }
            }
            else if (block.Name == OoxmlNs.W + "tbl")
            {
                DocxTargetId tableId = new DocxTargetId(storyLetter, storyPartNumber, DocxTargetKind.Table, tableIndex++, 0, 0);
                targets[block] = tableId.ToWireValue();
                if (visible)
                {
                    tables.Add(ReadTable(block, tableId, story, textView, package, relationships, images, targets, ref imageIndex));
                }
            }
            else if (block.Name == OoxmlNs.W + "sectPr")
            {
                sectionProperties = block;
            }

            if (sectionProperties is not null && idPrefix == "M")
            {
                DocxTargetId sectionId = new DocxTargetId(storyLetter, storyPartNumber, DocxTargetKind.Section, sectionIndex++, 0, 0);
                targets[sectionProperties] = sectionId.ToWireValue();
                if (visible)
                {
                    sections.Add(ReadSection(sectionProperties, sectionId, story));
                }
            }
        }

        bookmarks.AddRange(ReadBookmarks(document, partName, story, idPrefix, targets));
        contentControls.AddRange(ReadContentControls(document, partName, story, idPrefix, textView, targets));
        fields.AddRange(ReadFields(document, partName, story, idPrefix, textView, targets));
        hyperlinks.AddRange(ReadHyperlinks(document, partName, story, idPrefix, textView, targets, relationships));
    }

    private static DocxParagraphInfo ReadParagraph(
        XElement paragraph,
        DocxTargetId id,
        string story,
        DocxTextView textView,
        OoxmlPackage package,
        IReadOnlyDictionary<string, OoxmlRelationship> relationships,
        IReadOnlyDictionary<string, DocxStyleInfo> stylesById,
        DocxNumberingCatalog numbering,
        DocxNumberingLabeler numberingLabeler,
        List<DocxImageInfo> images,
        ref int imageIndex,
        CancellationToken cancellationToken)
    {
        DocxRunInfo[] runs = ReadRuns(paragraph, textView);
        foreach (XElement drawing in paragraph.Descendants(OoxmlNs.W + "drawing"))
        {
            AddDrawingImages(drawing, package, relationships, images, id, ref imageIndex);
        }

        string? styleId = ReadParagraphStyleId(paragraph);
        stylesById.TryGetValue(styleId ?? string.Empty, out DocxStyleInfo? style);
        return new DocxParagraphInfo(
            id,
            story,
            ReadText(paragraph, textView),
            DocxHeadingLevels.GetHeadingLevel(package, paragraph, cancellationToken),
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

    private static DocxVerticalMerge? ReadCellVerticalMerge(XElement cell)
    {
        XElement? verticalMerge = cell
            .Element(OoxmlNs.W + "tcPr")
            ?.Element(OoxmlNs.W + "vMerge");
        if (verticalMerge is null)
        {
            return null;
        }

        string? value = (string?)verticalMerge.Attribute(OoxmlNs.W + "val");
        return string.Equals(value, "restart", StringComparison.Ordinal) ? DocxVerticalMerge.Restart : DocxVerticalMerge.Continue;
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
            return numbering.Resolve(numberingId, level, DocxLabelSource.Direct);
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
                    inherited ? DocxLabelSource.StyleInherited : DocxLabelSource.Style);
            }

            current = style.BasedOnStyleId;
            inherited = true;
        }

        return null;
    }

    private static DocxSectionInfo ReadSection(XElement sectionProperties, DocxTargetId id, string story)
    {
        string? columnCountText = (string?)sectionProperties
            .Element(OoxmlNs.W + "cols")
            ?.Attribute(OoxmlNs.W + "num");
        int columns = int.TryParse(columnCountText, out int parsedColumns) && parsedColumns > 0
            ? parsedColumns
            : 1;
        string? orientationText = (string?)sectionProperties
            .Element(OoxmlNs.W + "pgSz")
            ?.Attribute(OoxmlNs.W + "orient");
        // Absence means portrait per the OOXML default; an unrecognized value is coerced
        // here and still reported by package validation (E9120).
        DocxOrientation orientation = DocxOrientationExtensions.TryParseWireValue(orientationText, out DocxOrientation parsedOrientation)
            ? parsedOrientation
            : DocxOrientation.Portrait;
        return new DocxSectionInfo(id, story, columns, orientation);
    }
    // C02: image identity is placement based. Each drawing placement in a
    // story gets its own public ID in document order. The media part is
    // a property of the placement, not the identity. Alt text, position,
    // and size belong to the placement. Repeated use of one media part
    // yields one ID per placement. Discovery, snapshot binding, mutation,
    // and reporting must use this same enumeration so an ID from discovery
    // always addresses the drawing it describes.
    // Public IDs for repeated media change under this contract.

    private static void AddDrawingImages(
        XElement drawing,
        OoxmlPackage package,
        IReadOnlyDictionary<string, OoxmlRelationship> relationships,
        List<DocxImageInfo> images,
        DocxTargetId containingTarget,
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

            XElement? sourceRectangle = blip
                .Ancestors(OoxmlNs.Pic + "blipFill")
                .FirstOrDefault()
                ?.Element(OoxmlNs.A + "srcRect");
            images.Add(new DocxImageInfo(new DocxTargetId(containingTarget.Story, containingTarget.StoryPart, DocxTargetKind.Image, imageIndex++, 0, 0), imagePart.Name, imagePart.ContentType, imagePart.Bytes.Length)
            {
                LayoutKind = layoutKind,
                RelationshipId = relationshipId,
                ContainingTargetId = containingTarget,
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

    private sealed record StyleNumbering(string NumberingId, int Level, DocxLabelSource Source);

    private sealed record TableMergeState(DocxTargetId? MergeGroupId, DocxTargetId? RootCellId);

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
        DocxRefreshPolicy RefreshPolicy,
        string? RefreshReason,
        bool CanRefreshDeterministically);

    private sealed record FieldRefreshProfile(DocxRefreshPolicy Policy, string? Reason, bool CanRefreshDeterministically);

    private sealed class ComplexFieldBuilder(XElement startElement, DocxTargetId? targetId, int nestingDepth)
    {
        public XElement StartElement { get; } = startElement;
        public DocxTargetId? TargetId { get; } = targetId;
        public int NestingDepth { get; } = nestingDepth;
        public System.Text.StringBuilder Code { get; } = new();
        public System.Text.StringBuilder ResultText { get; } = new();
        public bool HasSeparate { get; set; }
        public bool? IsDirty { get; init; }
        public bool? IsLocked { get; init; }

        public DocxFieldInfo ToInfo(DocxTargetId id, string story, string partName, bool complete)
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
