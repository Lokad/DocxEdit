using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

internal static partial class DocxPatchEngine
{
    private static bool TryParseMainParagraphTarget(string target, out int paragraphOrdinal)
    {
        paragraphOrdinal = 0;
        return target.Length == 7 &&
            target.StartsWith("M.P", StringComparison.Ordinal) &&
            int.TryParse(target[3..], out paragraphOrdinal);
    }

    private static bool TryParseMainTableTarget(string target, out int tableOrdinal)
    {
        tableOrdinal = 0;
        return target.Length == 7 &&
            target.StartsWith("M.T", StringComparison.Ordinal) &&
            int.TryParse(target[3..], out tableOrdinal);
    }

    private static bool TryParseMainCellTarget(string target, out int tableOrdinal, out int rowOrdinal, out int cellOrdinal)
    {
        tableOrdinal = 0;
        rowOrdinal = 0;
        cellOrdinal = 0;
        if (target.Length != 15 ||
            !target.StartsWith("M.T", StringComparison.Ordinal) ||
            target[7..9] != ".R" ||
            target[11..13] != ".C")
        {
            return false;
        }

        return int.TryParse(target[3..7], out tableOrdinal) &&
            int.TryParse(target[9..11], out rowOrdinal) &&
            int.TryParse(target[13..15], out cellOrdinal);
    }

    private static bool TryParseMainMergeGroupTarget(string target, out int tableOrdinal, out int mergeGroupOrdinal)
    {
        tableOrdinal = 0;
        mergeGroupOrdinal = 0;
        if (target.Length != 14 ||
            !target.StartsWith("M.T", StringComparison.Ordinal) ||
            target[7..10] != ".MG")
        {
            return false;
        }

        return int.TryParse(target[3..7], out tableOrdinal) &&
            int.TryParse(target[10..14], out mergeGroupOrdinal);
    }

    private static bool TryParseStoryMergeGroupTarget(
        string target,
        char storyPrefix,
        out int storyOrdinal,
        out int tableOrdinal,
        out int mergeGroupOrdinal)
    {
        storyOrdinal = 0;
        tableOrdinal = 0;
        mergeGroupOrdinal = 0;
        if (target.Length != 17 ||
            target[0] != storyPrefix ||
            target[4..6] != ".T" ||
            target[10..13] != ".MG")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[6..10], out tableOrdinal) &&
            int.TryParse(target[13..17], out mergeGroupOrdinal);
    }

    private static bool TryParseMainRowTarget(string target, out int tableOrdinal, out int rowOrdinal)
    {
        tableOrdinal = 0;
        rowOrdinal = 0;
        if (target.Length != 11 ||
            !target.StartsWith("M.T", StringComparison.Ordinal) ||
            target[7..9] != ".R")
        {
            return false;
        }

        return int.TryParse(target[3..7], out tableOrdinal) &&
            int.TryParse(target[9..11], out rowOrdinal);
    }

    private static bool TryParseMainImageTarget(string target, out int imageOrdinal)
    {
        imageOrdinal = 0;
        return target.Length == 7 &&
            target.StartsWith("M.I", StringComparison.Ordinal) &&
            int.TryParse(target[3..], out imageOrdinal);
    }

    private static bool TryParseMainHyperlinkTarget(string target, out int hyperlinkOrdinal)
    {
        hyperlinkOrdinal = 0;
        return target.Length == 7 &&
            target.StartsWith("M.L", StringComparison.Ordinal) &&
            int.TryParse(target[3..], out hyperlinkOrdinal);
    }

    private static bool TryParseMainFieldTarget(string target, out int fieldOrdinal)
    {
        fieldOrdinal = 0;
        return target.Length == 7 &&
            target.StartsWith("M.F", StringComparison.Ordinal) &&
            int.TryParse(target[3..], out fieldOrdinal);
    }

    private static bool TryParseMainBookmarkTarget(string target, out int bookmarkOrdinal)
    {
        bookmarkOrdinal = 0;
        return target.Length == 7 &&
            target.StartsWith("M.B", StringComparison.Ordinal) &&
            int.TryParse(target[3..], out bookmarkOrdinal);
    }

    private static bool TryParseMainContentControlTarget(string target, out int contentControlOrdinal)
    {
        contentControlOrdinal = 0;
        return target.Length == 8 &&
            target.StartsWith("M.CC", StringComparison.Ordinal) &&
            int.TryParse(target[4..], out contentControlOrdinal);
    }

    private static bool TryParseStoryImageTarget(
        string target,
        char storyPrefix,
        out int storyOrdinal,
        out int imageOrdinal)
    {
        storyOrdinal = 0;
        imageOrdinal = 0;
        if (target.Length != 10 ||
            target[0] != storyPrefix ||
            target[4..6] != ".I")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[6..], out imageOrdinal);
    }

    private static bool TryParseStoryHyperlinkTarget(
        string target,
        char storyPrefix,
        out int storyOrdinal,
        out int hyperlinkOrdinal)
    {
        storyOrdinal = 0;
        hyperlinkOrdinal = 0;
        if (target.Length != 10 ||
            target[0] != storyPrefix ||
            target[4..6] != ".L")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[6..], out hyperlinkOrdinal);
    }

    private static bool TryParseStoryFieldTarget(
        string target,
        char storyPrefix,
        out int storyOrdinal,
        out int fieldOrdinal)
    {
        storyOrdinal = 0;
        fieldOrdinal = 0;
        if (target.Length != 10 ||
            target[0] != storyPrefix ||
            target[4..6] != ".F")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[6..], out fieldOrdinal);
    }

    private static bool TryParseStoryBookmarkTarget(
        string target,
        char storyPrefix,
        out int storyOrdinal,
        out int bookmarkOrdinal)
    {
        storyOrdinal = 0;
        bookmarkOrdinal = 0;
        if (target.Length != 10 ||
            target[0] != storyPrefix ||
            target[4..6] != ".B")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[6..], out bookmarkOrdinal);
    }

    private static bool TryParseStoryContentControlTarget(
        string target,
        char storyPrefix,
        out int storyOrdinal,
        out int contentControlOrdinal)
    {
        storyOrdinal = 0;
        contentControlOrdinal = 0;
        if (target.Length != 11 ||
            target[0] != storyPrefix ||
            target[4..7] != ".CC")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[7..], out contentControlOrdinal);
    }

    private static bool TryParseMainSectionTarget(string target, out int sectionOrdinal)
    {
        sectionOrdinal = 0;
        return target.Length == 7 &&
            target.StartsWith("M.S", StringComparison.Ordinal) &&
            int.TryParse(target[3..], out sectionOrdinal);
    }

    private static XElement? FindTable(XElement body, int tableOrdinal)
    {
        return tableOrdinal < 1
            ? null
            : body.Elements(OoxmlNs.W + "tbl").ElementAtOrDefault(tableOrdinal - 1);
    }

    private static XElement? FindParagraph(XElement body, int paragraphOrdinal)
    {
        return paragraphOrdinal < 1
            ? null
            : body.Elements(OoxmlNs.W + "p").ElementAtOrDefault(paragraphOrdinal - 1);
    }

    private static XElement? ResolveMainBlock(
        XElement body,
        DocxPatchOperation operation,
        string target,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        diagnostics = [];
        if (!TryParseTargetSelector(target, operation, out TargetSelector? selector, out DocxDiagnostic? diagnostic))
        {
            diagnostics = [diagnostic!];
            return null;
        }

        if (selector is not ExplicitIdTargetSelector explicitId)
        {
            return ResolveMainParagraphElementBySelector(body, selector, operation, out diagnostics);
        }

        if (explicitId.TargetId is not { } blockId)
        {
            return null;
        }

        if (blockId is { Story: 'M', Kind: DocxTargetKind.Paragraph })
        {
            return FindParagraph(body, blockId.Primary);
        }

        if (blockId is { Story: 'M', Kind: DocxTargetKind.Table })
        {
            return FindTable(body, blockId.Primary);
        }

        return null;
    }

    private static BlockTarget? ResolveBlockTarget(
        OoxmlPackage package,
        DocxPatchOperation operation,
        string target,
        CancellationToken cancellationToken,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        diagnostics = [];
        if (!TryParseTargetSelector(target, operation, out TargetSelector? selector, out DocxDiagnostic? diagnostic))
        {
            diagnostics = [diagnostic!];
            return null;
        }

        if (selector is not ExplicitIdTargetSelector)
        {

            XDocument mainDocument = LoadMainDocument(package, cancellationToken, out XElement mainBody);
            XElement? selectedBlock = ResolveMainBlock(mainBody, operation, target, out diagnostics);
            return selectedBlock is null
                ? null
                : new BlockTarget(package.MainDocumentPartName, mainDocument, selectedBlock);
        }

        if (selector is ExplicitIdTargetSelector { TargetId: { } blockId })
        {
            if (blockId is { Story: 'M', Kind: DocxTargetKind.Paragraph })
            {
                XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
                XElement? paragraph = FindParagraph(body, blockId.Primary);
                return paragraph is null
                    ? null
                    : new BlockTarget(package.MainDocumentPartName, document, paragraph);
            }

            if (blockId is { Story: 'M', Kind: DocxTargetKind.Table })
            {
                XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
                XElement? table = FindTable(body, blockId.Primary);
                return table is null
                    ? null
                    : new BlockTarget(package.MainDocumentPartName, document, table);
            }

            bool isStoryBlock = blockId is { Kind: DocxTargetKind.Paragraph, Story: 'H' or 'F' }
                or { Kind: DocxTargetKind.Table, Story: 'H' or 'F' };
            if (isStoryBlock)
            {
                string relationshipType = blockId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
                XName blockName = blockId.Kind == DocxTargetKind.Paragraph ? OoxmlNs.W + "p" : OoxmlNs.W + "tbl";
                return ResolveRelatedStoryBlockTarget(package, relationshipType, blockId.StoryPart, blockName, blockId.Primary, cancellationToken);
            }
        }

        return null;
    }

    private static ParagraphTarget? ResolveParagraphTarget(
        OoxmlPackage package,
        DocxPatchOperation operation,
        string target,
        CancellationToken cancellationToken,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        diagnostics = [];
        if (!TryParseTargetSelector(target, operation, out TargetSelector? selector, out DocxDiagnostic? diagnostic))
        {
            diagnostics = [diagnostic!];
            return null;
        }

        if (selector is ExplicitIdTargetSelector { TargetId: { } paragraphId })
        {
            if (paragraphId is { Story: 'M', Kind: DocxTargetKind.Paragraph })
            {
                XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
                XElement? paragraph = FindParagraph(body, paragraphId.Primary);
                return paragraph is null
                    ? null
                    : new ParagraphTarget(package.MainDocumentPartName, document, paragraph);
            }

            bool isStoryParagraph = paragraphId is { Kind: DocxTargetKind.Paragraph, Story: 'H' or 'F' };
            if (isStoryParagraph)
            {
                string relationshipType = paragraphId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
                return ResolveRelatedStoryParagraphTarget(package, relationshipType, paragraphId.StoryPart, paragraphId.Primary, cancellationToken);
            }

            return null;
        }


        XDocument mainDocument = LoadMainDocument(package, cancellationToken, out XElement mainBody);
        XElement? selectedParagraph = ResolveMainParagraphElementBySelector(mainBody, selector, operation, out diagnostics);
        return selectedParagraph is null
            ? null
            : new ParagraphTarget(package.MainDocumentPartName, mainDocument, selectedParagraph);
    }

    internal static bool TryParseTargetId(string? target, out DocxTargetId targetId)
    {
        targetId = default;
        if (string.IsNullOrEmpty(target))
        {
            return false;
        }

        if (TryParseMainParagraphTarget(target, out int paragraphOrdinal))
        {
            targetId = new DocxTargetId('M', 0, DocxTargetKind.Paragraph, paragraphOrdinal, 0, 0);
            return true;
        }

        if (TryParseMainTableTarget(target, out int tableOrdinal))
        {
            targetId = new DocxTargetId('M', 0, DocxTargetKind.Table, tableOrdinal, 0, 0);
            return true;
        }

        if (TryParseMainCellTarget(target, out int cellTableOrdinal, out int rowOrdinal, out int cellOrdinal))
        {
            targetId = new DocxTargetId('M', 0, DocxTargetKind.Cell, cellTableOrdinal, rowOrdinal, cellOrdinal);
            return true;
        }

        if (TryParseMainMergeGroupTarget(target, out int mergeGroupTableOrdinal, out int mergeGroupOrdinal))
        {
            targetId = new DocxTargetId('M', 0, DocxTargetKind.MergeGroup, mergeGroupTableOrdinal, mergeGroupOrdinal, 0);
            return true;
        }

        if (TryParseMainRowTarget(target, out int rowTableOrdinal, out int mainRowOrdinal))
        {
            targetId = new DocxTargetId('M', 0, DocxTargetKind.Row, rowTableOrdinal, mainRowOrdinal, 0);
            return true;
        }

        if (TryParseMainImageTarget(target, out int imageOrdinal))
        {
            targetId = new DocxTargetId('M', 0, DocxTargetKind.Image, imageOrdinal, 0, 0);
            return true;
        }

        if (TryParseMainHyperlinkTarget(target, out int hyperlinkOrdinal))
        {
            targetId = new DocxTargetId('M', 0, DocxTargetKind.Hyperlink, hyperlinkOrdinal, 0, 0);
            return true;
        }

        if (TryParseMainFieldTarget(target, out int fieldOrdinal))
        {
            targetId = new DocxTargetId('M', 0, DocxTargetKind.Field, fieldOrdinal, 0, 0);
            return true;
        }

        if (TryParseMainBookmarkTarget(target, out int bookmarkOrdinal))
        {
            targetId = new DocxTargetId('M', 0, DocxTargetKind.Bookmark, bookmarkOrdinal, 0, 0);
            return true;
        }

        if (TryParseMainContentControlTarget(target, out int contentControlOrdinal))
        {
            targetId = new DocxTargetId('M', 0, DocxTargetKind.ContentControl, contentControlOrdinal, 0, 0);
            return true;
        }

        if (TryParseMainSectionTarget(target, out int sectionOrdinal))
        {
            targetId = new DocxTargetId('M', 0, DocxTargetKind.Section, sectionOrdinal, 0, 0);
            return true;
        }

        foreach (char storyPrefix in new[] { 'H', 'F' })
        {
            if (TryParseStoryParagraphTarget(target, storyPrefix, out int storyOrdinal, out int storyParagraphOrdinal))
            {
                targetId = new DocxTargetId(storyPrefix, storyOrdinal, DocxTargetKind.Paragraph, storyParagraphOrdinal, 0, 0);
                return true;
            }

            if (TryParseStoryTableTarget(target, storyPrefix, out storyOrdinal, out int storyTableOrdinal))
            {
                targetId = new DocxTargetId(storyPrefix, storyOrdinal, DocxTargetKind.Table, storyTableOrdinal, 0, 0);
                return true;
            }

            if (TryParseStoryCellTarget(target, storyPrefix, out storyOrdinal, out int storyCellTableOrdinal, out int storyRowOrdinal, out int storyCellOrdinal))
            {
                targetId = new DocxTargetId(storyPrefix, storyOrdinal, DocxTargetKind.Cell, storyCellTableOrdinal, storyRowOrdinal, storyCellOrdinal);
                return true;
            }

            if (TryParseStoryMergeGroupTarget(target, storyPrefix, out storyOrdinal, out int storyMergeGroupTableOrdinal, out int storyMergeGroupOrdinal))
            {
                targetId = new DocxTargetId(storyPrefix, storyOrdinal, DocxTargetKind.MergeGroup, storyMergeGroupTableOrdinal, storyMergeGroupOrdinal, 0);
                return true;
            }

            if (TryParseStoryRowTarget(target, storyPrefix, out storyOrdinal, out int storyRowTableOrdinal, out int storyRowRowOrdinal))
            {
                targetId = new DocxTargetId(storyPrefix, storyOrdinal, DocxTargetKind.Row, storyRowTableOrdinal, storyRowRowOrdinal, 0);
                return true;
            }

            if (TryParseStoryImageTarget(target, storyPrefix, out storyOrdinal, out int storyImageOrdinal))
            {
                targetId = new DocxTargetId(storyPrefix, storyOrdinal, DocxTargetKind.Image, storyImageOrdinal, 0, 0);
                return true;
            }

            if (TryParseStoryHyperlinkTarget(target, storyPrefix, out storyOrdinal, out int storyHyperlinkOrdinal))
            {
                targetId = new DocxTargetId(storyPrefix, storyOrdinal, DocxTargetKind.Hyperlink, storyHyperlinkOrdinal, 0, 0);
                return true;
            }

            if (TryParseStoryFieldTarget(target, storyPrefix, out storyOrdinal, out int storyFieldOrdinal))
            {
                targetId = new DocxTargetId(storyPrefix, storyOrdinal, DocxTargetKind.Field, storyFieldOrdinal, 0, 0);
                return true;
            }

            if (TryParseStoryBookmarkTarget(target, storyPrefix, out storyOrdinal, out int storyBookmarkOrdinal))
            {
                targetId = new DocxTargetId(storyPrefix, storyOrdinal, DocxTargetKind.Bookmark, storyBookmarkOrdinal, 0, 0);
                return true;
            }

            if (TryParseStoryContentControlTarget(target, storyPrefix, out storyOrdinal, out int storyContentControlOrdinal))
            {
                targetId = new DocxTargetId(storyPrefix, storyOrdinal, DocxTargetKind.ContentControl, storyContentControlOrdinal, 0, 0);
                return true;
            }
        }

        return false;
    }

    private static bool TryParseTargetSelector(
        string target,
        DocxPatchOperation operation,
        [NotNullWhen(true)] out TargetSelector? selector,
        out DocxDiagnostic? diagnostic)
    {
        selector = null;
        diagnostic = null;
        if (target.StartsWith("heading:", StringComparison.Ordinal))
        {
            string value = target["heading:".Length..].Trim();
            int? level = null;
            int separator = value.IndexOf(':');
            if (separator > 0 && value[..separator].All(char.IsDigit))
            {
                if (!int.TryParse(value[..separator], out int parsedLevel) || parsedLevel is < 1 or > 9)
                {
                    diagnostic = Diagnostic(DocxSeverity.Error, "E1203", $"Invalid heading selector '{target}'. Heading level must be between 1 and 9.", operation, target);
                    return false;
                }

                level = parsedLevel;
                value = value[(separator + 1)..].Trim();
            }

            if (!TryReadSelectorText(value, out string selectorText))
            {
                diagnostic = Diagnostic(DocxSeverity.Error, "E1203", $"Invalid heading selector '{target}'. Expected heading:\"Text\" or heading:2:\"Text\".", operation, target);
                return false;
            }

            selector = new HeadingTargetSelector(target, level, selectorText);
            return true;
        }

        if (target.StartsWith("text:", StringComparison.Ordinal))
        {
            string value = target["text:".Length..].Trim();
            if (!TryReadSelectorText(value, out string selectorText))
            {
                diagnostic = Diagnostic(DocxSeverity.Error, "E1203", $"Invalid text selector '{target}'. Expected text:\"Text\".", operation, target);
                return false;
            }

            selector = new ParagraphTextTargetSelector(target, selectorText);
            return true;
        }

        if (target.StartsWith("bookmark:", StringComparison.Ordinal))
        {
            string value = target["bookmark:".Length..].Trim();
            if (!TryReadSelectorText(value, out string selectorText))
            {
                diagnostic = Diagnostic(DocxSeverity.Error, "E1203", $"Invalid bookmark selector '{target}'. Expected bookmark:\"Name\".", operation, target);
                return false;
            }

            selector = new BookmarkTargetSelector(target, selectorText);
            return true;
        }

        if (target.StartsWith("content-control:", StringComparison.Ordinal))
        {
            string value = target["content-control:".Length..].Trim();
            if (!TryReadSelectorText(value, out string selectorText))
            {
                diagnostic = Diagnostic(DocxSeverity.Error, "E1203", $"Invalid content-control selector '{target}'. Expected content-control:\"TagOrAlias\".", operation, target);
                return false;
            }

            selector = new ContentControlTargetSelector(target, selectorText);
            return true;
        }

        DocxTargetId? parsedTargetId = TryParseTargetId(target, out DocxTargetId parsedTargetIdValue) ? parsedTargetIdValue : null;
        selector = new ExplicitIdTargetSelector(target, parsedTargetId);
        return true;
    }

    private static bool TryReadSelectorText(string value, out string selectorText)
    {
        selectorText = string.Empty;
        if (value.Length >= 2 && value[0] == '"' && value[^1] == '"')
        {
            selectorText = value[1..^1];
            return selectorText.Length != 0;
        }

        if (value.StartsWith('"') || value.EndsWith('"'))
        {
            return false;
        }

        selectorText = value;
        return selectorText.Length != 0;
    }

    private static XElement? ResolveMainParagraphElementBySelector(
        XElement body,
        TargetSelector selector,
        DocxPatchOperation operation,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        diagnostics = [];
        if (selector is HeadingTargetSelector headingSelector)
        {
            return ResolveMainParagraphElementByPredicate(
                body,
                paragraph =>
                {
                    int? headingLevel = ReadHeadingLevel(paragraph);
                    return headingLevel is not null &&
                        (headingSelector.Level is null || headingSelector.Level == headingLevel) &&
                        string.Equals(ReadVisibleText(paragraph), headingSelector.Text, StringComparison.Ordinal);
                },
                selector.Raw,
                operation,
                out diagnostics);
        }

        if (selector is ParagraphTextTargetSelector paragraphTextSelector)
        {
            return ResolveMainParagraphElementByPredicate(
                body,
                paragraph => ReadVisibleText(paragraph).Contains(paragraphTextSelector.Text, StringComparison.Ordinal),
                selector.Raw,
                operation,
                out diagnostics);
        }

        if (selector is BookmarkTargetSelector bookmarkSelector)
        {
            return ResolveMainParagraphElementByPredicate(
                body,
                paragraph => ParagraphHasBookmark(paragraph, bookmarkSelector.Name),
                selector.Raw,
                operation,
                out diagnostics);
        }

        if (selector is ContentControlTargetSelector contentControlSelector)
        {
            return ResolveMainParagraphElementByPredicate(
                body,
                paragraph => ParagraphHasContentControl(paragraph, contentControlSelector.Name),
                selector.Raw,
                operation,
                out diagnostics);
        }

        return null;
    }

    private static bool ParagraphHasBookmark(XElement paragraph, string name)
    {
        return paragraph
            .Descendants(OoxmlNs.W + "bookmarkStart")
            .Any(bookmark => string.Equals((string?)bookmark.Attribute(OoxmlNs.W + "name"), name, StringComparison.Ordinal));
    }

    private static bool ParagraphHasContentControl(XElement paragraph, string name)
    {
        return paragraph
            .Descendants(OoxmlNs.W + "sdt")
            .Any(contentControl =>
            {
                XElement? properties = contentControl.Element(OoxmlNs.W + "sdtPr");
                string? tag = (string?)properties
                    ?.Element(OoxmlNs.W + "tag")
                    ?.Attribute(OoxmlNs.W + "val");
                string? alias = (string?)properties
                    ?.Element(OoxmlNs.W + "alias")
                    ?.Attribute(OoxmlNs.W + "val");
                return string.Equals(tag, name, StringComparison.Ordinal) ||
                    string.Equals(alias, name, StringComparison.Ordinal);
            });
    }

    private static XElement? ResolveMainParagraphElementByPredicate(
        XElement body,
        Func<XElement, bool> predicate,
        string rawSelector,
        DocxPatchOperation operation,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        diagnostics = [];
        var matches = new List<ParagraphSelectorMatch>();
        int paragraphOrdinal = 0;
        foreach (XElement paragraph in body.Elements(OoxmlNs.W + "p"))
        {
            paragraphOrdinal++;
            if (predicate(paragraph))
            {
                matches.Add(new ParagraphSelectorMatch($"M.P{paragraphOrdinal:0000}", paragraph));
            }
        }

        if (matches.Count > 1)
        {
            diagnostics =
            [
                Diagnostic(
                    DocxSeverity.Error,
                    "E1202",
                    $"Selector matched {matches.Count} targets: {string.Join(", ", matches.Select(match => match.Id))}. Use a more specific selector or an explicit ID.",
                    operation,
                    rawSelector)
            ];
            return null;
        }

        if (matches.Count == 0)
        {
            diagnostics =
            [
                Diagnostic(
                    DocxSeverity.Error,
                    "E1201",
                    $"Selector matched 0 targets: {rawSelector}.{BuildNoMatchSuggestion(body, rawSelector)}",
                    operation,
                    rawSelector)
            ];
            return null;
        }

        return matches[0].Paragraph;
    }

    private static string BuildNoMatchSuggestion(XElement body, string rawSelector)
    {
        string[] suggestions = rawSelector.StartsWith("heading:", StringComparison.Ordinal)
            ? EnumerateMainParagraphs(body)
                .Select(match => (match.Id, HeadingLevel: ReadHeadingLevel(match.Paragraph)))
                .Where(match => match.HeadingLevel is not null)
                .Take(3)
                .Select(match => $"{match.Id} heading level={match.HeadingLevel}")
                .ToArray()
            : EnumerateMainParagraphs(body)
                .Take(3)
                .Select(match => match.Id)
                .ToArray();
        if (suggestions.Length == 0)
        {
            return " No nearby paragraph targets are available.";
        }

        string label = rawSelector.StartsWith("heading:", StringComparison.Ordinal)
            ? "Nearby headings"
            : "Nearby paragraphs";
        return $" {label}: {string.Join(", ", suggestions)}.";
    }

    private static IEnumerable<ParagraphSelectorMatch> EnumerateMainParagraphs(XElement body)
    {
        int paragraphOrdinal = 0;
        foreach (XElement paragraph in body.Elements(OoxmlNs.W + "p"))
        {
            paragraphOrdinal++;
            yield return new ParagraphSelectorMatch($"M.P{paragraphOrdinal:0000}", paragraph);
        }
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

    private static SectionTarget? ResolveMainSectionTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (!TryParseTargetId(target, out DocxTargetId sectionId) ||
            sectionId is not { Story: 'M', Kind: DocxTargetKind.Section } ||
            sectionId.Primary < 1)
        {
            return null;
        }

        XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
        int currentOrdinal = 0;
        foreach (XElement block in body.Elements())
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement? element = block.Name == OoxmlNs.W + "p"
                ? block.Element(OoxmlNs.W + "pPr")?.Element(OoxmlNs.W + "sectPr")
                : block.Name == OoxmlNs.W + "sectPr"
                    ? block
                    : null;
            if (element is null)
            {
                continue;
            }

            currentOrdinal++;
            if (currentOrdinal == sectionId.Primary)
            {
                return new SectionTarget(document, element);
            }
        }

        return null;
    }

    private static ParagraphTarget? ResolveRelatedStoryParagraphTarget(
        OoxmlPackage package,
        string relationshipType,
        int storyOrdinal,
        int paragraphOrdinal,
        CancellationToken cancellationToken)
    {
        if (storyOrdinal < 1)
        {
            return null;
        }

        ResolvedOoxmlRelationship? relationship = package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => relationship.Type == relationshipType)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .ElementAtOrDefault(storyOrdinal - 1);
        if (relationship is null)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, relationship.ResolvedTarget, cancellationToken, out XElement root);
        XElement? paragraph = FindParagraph(root, paragraphOrdinal);
        return paragraph is null ? null : new ParagraphTarget(relationship.ResolvedTarget, document, paragraph);
    }

    private static bool TryParseStoryParagraphTarget(
        string target,
        char storyPrefix,
        out int storyOrdinal,
        out int paragraphOrdinal)
    {
        storyOrdinal = 0;
        paragraphOrdinal = 0;
        if (target.Length != 10 ||
            target[0] != storyPrefix ||
            target[4..6] != ".P")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[6..], out paragraphOrdinal);
    }

    private static bool TryParseStoryTableTarget(
        string target,
        char storyPrefix,
        out int storyOrdinal,
        out int tableOrdinal)
    {
        storyOrdinal = 0;
        tableOrdinal = 0;
        if (target.Length != 10 ||
            target[0] != storyPrefix ||
            target[4..6] != ".T")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[6..], out tableOrdinal);
    }

    private static bool TryParseStoryRowTarget(
        string target,
        char storyPrefix,
        out int storyOrdinal,
        out int tableOrdinal,
        out int rowOrdinal)
    {
        storyOrdinal = 0;
        tableOrdinal = 0;
        rowOrdinal = 0;
        if (target.Length != 14 ||
            target[0] != storyPrefix ||
            target[4..6] != ".T" ||
            target[10..12] != ".R")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[6..10], out tableOrdinal) &&
            int.TryParse(target[12..], out rowOrdinal);
    }

    private static bool TryParseStoryCellTarget(
        string target,
        char storyPrefix,
        out int storyOrdinal,
        out int tableOrdinal,
        out int rowOrdinal,
        out int cellOrdinal)
    {
        storyOrdinal = 0;
        tableOrdinal = 0;
        rowOrdinal = 0;
        cellOrdinal = 0;
        if (target.Length != 18 ||
            target[0] != storyPrefix ||
            target[4..6] != ".T" ||
            target[10..12] != ".R" ||
            target[14..16] != ".C")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[6..10], out tableOrdinal) &&
            int.TryParse(target[12..14], out rowOrdinal) &&
            int.TryParse(target[16..], out cellOrdinal);
    }

    private static BlockTarget? ResolveRelatedStoryBlockTarget(
        OoxmlPackage package,
        string relationshipType,
        int storyOrdinal,
        XName blockName,
        int blockOrdinal,
        CancellationToken cancellationToken)
    {
        if (storyOrdinal < 1 || blockOrdinal < 1)
        {
            return null;
        }

        ResolvedOoxmlRelationship? relationship = package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => relationship.Type == relationshipType)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .ElementAtOrDefault(storyOrdinal - 1);
        if (relationship is null)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, relationship.ResolvedTarget, cancellationToken, out XElement root);
        XElement? block = root.Elements(blockName).ElementAtOrDefault(blockOrdinal - 1);
        return block is null ? null : new BlockTarget(relationship.ResolvedTarget, document, block);
    }

    private static bool IsSupportedTableTargetShape(string target)
    {
        return TryParseTargetId(target, out DocxTargetId tableId) && tableId.Kind == DocxTargetKind.Table;
    }

    private static bool IsSupportedRowTargetShape(string target)
    {
        return TryParseTargetId(target, out DocxTargetId rowId) && rowId.Kind == DocxTargetKind.Row;
    }

    private static bool IsSupportedCellTargetShape(string target)
    {
        return TryParseTargetId(target, out DocxTargetId cellId)
            && cellId.Kind is (DocxTargetKind.Cell or DocxTargetKind.MergeGroup);
    }

    private static TableTarget? ResolveTableTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (!TryParseTargetId(target, out DocxTargetId tableId) || tableId.Kind != DocxTargetKind.Table)
        {
            return null;
        }

        return ResolveTableTarget(package, tableId, cancellationToken);
    }

    private static TableTarget? ResolveTableTarget(
        OoxmlPackage package,
        DocxTargetId tableId,
        CancellationToken cancellationToken)
    {
        if (tableId.Story == 'M')
        {
            XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
            XElement? table = FindTable(body, tableId.Primary);
            return table is null
                ? null
                : new TableTarget(package.MainDocumentPartName, document, table);
        }

        if (tableId.Story is 'H' or 'F')
        {
            string relationshipType = tableId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
            return ResolveRelatedStoryTableTarget(package, relationshipType, tableId.StoryPart, tableId.Primary, cancellationToken);
        }

        return null;
    }

    private static RowTarget? ResolveRowTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (!TryParseTargetId(target, out DocxTargetId rowId) || rowId.Kind != DocxTargetKind.Row)
        {
            return null;
        }

        return ResolveRowTarget(package, rowId, cancellationToken);
    }

    private static RowTarget? ResolveRowTarget(
        OoxmlPackage package,
        DocxTargetId rowId,
        CancellationToken cancellationToken)
    {
        TableTarget? tableTarget = ResolveTableTarget(package, rowId.TableId, cancellationToken);
        XElement? row = tableTarget?.Table.Elements(OoxmlNs.W + "tr").ElementAtOrDefault(rowId.Secondary - 1);
        return tableTarget is null || row is null
            ? null
            : new RowTarget(tableTarget.PartName, tableTarget.Document, tableTarget.Table, row);
    }

    private static CellTarget? ResolveCellTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (!TryParseTargetId(target, out DocxTargetId cellId) ||
            cellId.Kind is not (DocxTargetKind.Cell or DocxTargetKind.MergeGroup))
        {
            return null;
        }

        if (cellId.Kind == DocxTargetKind.MergeGroup)
        {
            TableTarget? tableTarget = ResolveTableTarget(package, cellId.TableId, cancellationToken);
            return tableTarget is null ? null : ResolveMergeGroupCellTarget(tableTarget, cellId.Secondary);
        }

        RowTarget? rowTarget = ResolveRowTarget(package, cellId.RowId, cancellationToken);
        XElement? cell = rowTarget is null ? null : FindCellByVisualColumn(rowTarget.Row, cellId.Tertiary);
        return rowTarget is null || cell is null
            ? null
            : new CellTarget(rowTarget.PartName, rowTarget.Document, rowTarget.Table, rowTarget.Row, cell, cellId.Tertiary);
    }

    private static XElement? FindCellByVisualColumn(XElement row, int visualColumnIndex)
    {
        if (visualColumnIndex < 1)
        {
            return null;
        }

        int columnIndex = 1 + ReadTableRowGridOffset(row, "gridBefore");
        foreach (XElement cell in row.Elements(OoxmlNs.W + "tc"))
        {
            int columnSpan = ReadTableCellColumnSpan(cell);
            if (visualColumnIndex >= columnIndex && visualColumnIndex < columnIndex + columnSpan)
            {
                return cell;
            }

            columnIndex += columnSpan;
        }

        return null;
    }

    private static CellTarget? ResolveMergeGroupCellTarget(TableTarget tableTarget, int mergeGroupOrdinal)
    {
        if (mergeGroupOrdinal < 1)
        {
            return null;
        }

        int mergeGroupIndex = 1;
        var activeVerticalMerges = new Dictionary<int, MergeGroupRootState>();
        foreach (XElement row in tableTarget.Table.Elements(OoxmlNs.W + "tr"))
        {
            int gridBefore = ReadTableRowGridOffset(row, "gridBefore");
            RemoveActiveMergeGroups(activeVerticalMerges, 1, gridBefore);
            int columnIndex = 1 + gridBefore;
            foreach (XElement cell in row.Elements(OoxmlNs.W + "tc"))
            {
                int columnSpan = ReadTableCellColumnSpan(cell);
                string? verticalMerge = ReadTableCellVerticalMerge(cell);
                if (string.Equals(verticalMerge, "restart", StringComparison.Ordinal))
                {
                    int currentMergeGroup = mergeGroupIndex++;
                    SetActiveMergeGroup(activeVerticalMerges, columnIndex, columnSpan, new MergeGroupRootState(row, cell, columnIndex));
                    if (currentMergeGroup == mergeGroupOrdinal)
                    {
                        return new CellTarget(tableTarget.PartName, tableTarget.Document, tableTarget.Table, row, cell, columnIndex);
                    }
                }
                else if (verticalMerge is not null)
                {
                    MergeGroupRootState? root = FindActiveMergeGroup(activeVerticalMerges, columnIndex, columnSpan);
                    if (root is null)
                    {
                        int currentMergeGroup = mergeGroupIndex++;
                        if (currentMergeGroup == mergeGroupOrdinal)
                        {
                            return new CellTarget(tableTarget.PartName, tableTarget.Document, tableTarget.Table, row, cell, columnIndex);
                        }

                        SetActiveMergeGroup(activeVerticalMerges, columnIndex, columnSpan, new MergeGroupRootState(row, cell, columnIndex));
                    }
                }
                else
                {
                    RemoveActiveMergeGroups(activeVerticalMerges, columnIndex, columnSpan);
                    if (columnSpan > 1)
                    {
                        int currentMergeGroup = mergeGroupIndex++;
                        if (currentMergeGroup == mergeGroupOrdinal)
                        {
                            return new CellTarget(tableTarget.PartName, tableTarget.Document, tableTarget.Table, row, cell, columnIndex);
                        }
                    }
                }

                columnIndex += columnSpan;
            }

            int gridAfter = ReadTableRowGridOffset(row, "gridAfter");
            RemoveActiveMergeGroups(activeVerticalMerges, columnIndex, gridAfter);
        }

        return null;
    }

    private static MergeGroupRootState? FindActiveMergeGroup(
        IReadOnlyDictionary<int, MergeGroupRootState> activeVerticalMerges,
        int columnIndex,
        int columnSpan)
    {
        MergeGroupRootState? root = null;
        for (int column = columnIndex; column < columnIndex + columnSpan; column++)
        {
            if (!activeVerticalMerges.TryGetValue(column, out MergeGroupRootState? current))
            {
                return null;
            }

            root ??= current;
            if (!ReferenceEquals(root.Cell, current.Cell))
            {
                return null;
            }
        }

        return root;
    }

    private static void SetActiveMergeGroup(
        Dictionary<int, MergeGroupRootState> activeVerticalMerges,
        int columnIndex,
        int columnSpan,
        MergeGroupRootState root)
    {
        for (int column = columnIndex; column < columnIndex + columnSpan; column++)
        {
            activeVerticalMerges[column] = root;
        }
    }

    private static void RemoveActiveMergeGroups(
        Dictionary<int, MergeGroupRootState> activeVerticalMerges,
        int columnIndex,
        int columnSpan)
    {
        for (int column = columnIndex; column < columnIndex + columnSpan; column++)
        {
            activeVerticalMerges.Remove(column);
        }
    }

    private static TableTarget? ResolveRelatedStoryTableTarget(
        OoxmlPackage package,
        string relationshipType,
        int storyOrdinal,
        int tableOrdinal,
        CancellationToken cancellationToken)
    {
        BlockTarget? blockTarget = ResolveRelatedStoryBlockTarget(package, relationshipType, storyOrdinal, OoxmlNs.W + "tbl", tableOrdinal, cancellationToken);
        return blockTarget is null ? null : new TableTarget(blockTarget.PartName, blockTarget.Document, blockTarget.Block);
    }

    private static bool IsSupportedImageTargetShape(string target)
    {
        return TryParseTargetId(target, out DocxTargetId imageId) && imageId.Kind == DocxTargetKind.Image;
    }

    private static bool IsSupportedHyperlinkTargetShape(string target)
    {
        return TryParseTargetId(target, out DocxTargetId hyperlinkId) && hyperlinkId.Kind == DocxTargetKind.Hyperlink;
    }

    private static bool IsSupportedContentControlTargetShape(string target)
    {
        return TryParseTargetId(target, out DocxTargetId contentcontrolId) && contentcontrolId.Kind == DocxTargetKind.ContentControl;
    }

    private static bool IsSupportedFieldTargetShape(string target)
    {
        return TryParseTargetId(target, out DocxTargetId fieldId) && fieldId.Kind == DocxTargetKind.Field;
    }

    private static bool IsSupportedBookmarkTargetShape(string target)
    {
        return TryParseTargetId(target, out DocxTargetId bookmarkId) && bookmarkId.Kind == DocxTargetKind.Bookmark;
    }

    private static ContentControlTarget? ResolveContentControlTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (!TryParseTargetId(target, out DocxTargetId controlId) || controlId.Kind != DocxTargetKind.ContentControl)
        {
            return null;
        }

        if (controlId.Story == 'M')
        {
            return FindContentControlTarget(package, package.MainDocumentPartName, controlId.Primary, cancellationToken);
        }

        if (controlId.Story is 'H' or 'F')
        {
            string relationshipType = controlId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
            string? partName = ResolveRelatedStoryPartName(package, relationshipType, controlId.StoryPart, cancellationToken);
            return partName is null ? null : FindContentControlTarget(package, partName, controlId.Primary, cancellationToken);
        }

        return null;
    }

    private static FieldTarget? ResolveFieldTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (!TryParseTargetId(target, out DocxTargetId fieldId) || fieldId.Kind != DocxTargetKind.Field)
        {
            return null;
        }

        if (fieldId.Story == 'M')
        {
            XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
            XElement? field = FindField(body, fieldId.Primary);
            return field is null
                ? null
                : new FieldTarget(package.MainDocumentPartName, document, field);
        }

        if (fieldId.Story is 'H' or 'F')
        {
            string relationshipType = fieldId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
            string? partName = ResolveRelatedStoryPartName(package, relationshipType, fieldId.StoryPart, cancellationToken);
            return partName is null ? null : FindFieldTarget(package, partName, fieldId.Primary, cancellationToken);
        }

        return null;
    }

    private static bool IsAllFieldsTarget(string target)
    {
        return target is "all" or "all-fields";
    }

    private static IReadOnlyList<FieldTarget> ResolveAllFieldTargets(
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        var targets = new List<FieldTarget>();
        foreach (string partName in GetEditableStoryPartNames(package, cancellationToken))
        {
            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out XElement root);
            targets.AddRange(FindFields(root).Select(field => new FieldTarget(partName, document, field)));
        }

        return targets;
    }

    private static FieldTarget? FindFieldTarget(
        OoxmlPackage package,
        string partName,
        int fieldOrdinal,
        CancellationToken cancellationToken)
    {
        if (fieldOrdinal < 1)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, partName, cancellationToken, out XElement root);
        XElement? field = FindField(root, fieldOrdinal);
        return field is null ? null : new FieldTarget(partName, document, field);
    }

    private static XElement? FindField(XElement root, int fieldOrdinal)
    {
        if (fieldOrdinal < 1)
        {
            return null;
        }

        return FindFields(root).ElementAtOrDefault(fieldOrdinal - 1);
    }

    private static IReadOnlyList<XElement> FindFields(XElement root)
    {
        var fields = new List<XElement>();
        var stack = new Stack<XElement>();
        foreach (XElement element in root.Descendants())
        {
            if (element.Name == OoxmlNs.W + "fldSimple")
            {
                fields.Add(element);
                continue;
            }

            if (element.Name != OoxmlNs.W + "fldChar")
            {
                continue;
            }

            string? fieldCharType = (string?)element.Attribute(OoxmlNs.W + "fldCharType");
            if (string.Equals(fieldCharType, "begin", StringComparison.Ordinal))
            {
                stack.Push(element);
            }
            else if (string.Equals(fieldCharType, "end", StringComparison.Ordinal) && stack.Count > 0)
            {
                XElement begin = stack.Pop();
                fields.Add(begin);
            }
        }

        foreach (XElement incompleteBegin in stack)
        {
            fields.Add(incompleteBegin);
        }

        return fields;
    }

    private static ContentControlTarget? FindContentControlTarget(
        OoxmlPackage package,
        string partName,
        int contentControlOrdinal,
        CancellationToken cancellationToken)
    {
        if (contentControlOrdinal < 1)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
        XElement? contentControl = document
            .Descendants(OoxmlNs.W + "sdt")
            .ElementAtOrDefault(contentControlOrdinal - 1);
        return contentControl is null ? null : new ContentControlTarget(partName, document, contentControl);
    }

    private static BookmarkTarget? ResolveBookmarkTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (!TryParseTargetId(target, out DocxTargetId bookmarkId) || bookmarkId.Kind != DocxTargetKind.Bookmark)
        {
            return null;
        }

        if (bookmarkId.Story == 'M')
        {
            return FindBookmarkTarget(package, package.MainDocumentPartName, bookmarkId.Primary, cancellationToken);
        }

        if (bookmarkId.Story is 'H' or 'F')
        {
            string relationshipType = bookmarkId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
            string? partName = ResolveRelatedStoryPartName(package, relationshipType, bookmarkId.StoryPart, cancellationToken);
            return partName is null ? null : FindBookmarkTarget(package, partName, bookmarkId.Primary, cancellationToken);
        }

        return null;
    }

    private static BookmarkTarget? FindBookmarkTarget(
        OoxmlPackage package,
        string partName,
        int bookmarkOrdinal,
        CancellationToken cancellationToken)
    {
        if (bookmarkOrdinal < 1)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
        XElement? start = document
            .Descendants(OoxmlNs.W + "bookmarkStart")
            .ElementAtOrDefault(bookmarkOrdinal - 1);
        string? ooxmlId = (string?)start?.Attribute(OoxmlNs.W + "id");
        XElement? end = ooxmlId is null
            ? null
            : document
                .Descendants(OoxmlNs.W + "bookmarkEnd")
                .FirstOrDefault(element => string.Equals((string?)element.Attribute(OoxmlNs.W + "id"), ooxmlId, StringComparison.Ordinal));
        return start is null ? null : new BookmarkTarget(partName, document, start, end);
    }

    private static HyperlinkTarget? ResolveHyperlinkTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (!TryParseTargetId(target, out DocxTargetId hyperlinkId) || hyperlinkId.Kind != DocxTargetKind.Hyperlink)
        {
            return null;
        }

        if (hyperlinkId.Story == 'M')
        {
            return FindHyperlinkTarget(package, package.MainDocumentPartName, hyperlinkId.Primary, cancellationToken);
        }

        if (hyperlinkId.Story is 'H' or 'F')
        {
            string relationshipType = hyperlinkId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
            string? partName = ResolveRelatedStoryPartName(package, relationshipType, hyperlinkId.StoryPart, cancellationToken);
            return partName is null ? null : FindHyperlinkTarget(package, partName, hyperlinkId.Primary, cancellationToken);
        }

        return null;
    }

    private static HyperlinkTarget? FindHyperlinkTarget(
        OoxmlPackage package,
        string partName,
        int hyperlinkOrdinal,
        CancellationToken cancellationToken)
    {
        if (hyperlinkOrdinal < 1)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
        XElement? hyperlink = document
            .Descendants(OoxmlNs.W + "hyperlink")
            .ElementAtOrDefault(hyperlinkOrdinal - 1);
        return hyperlink is null ? null : new HyperlinkTarget(partName, document, hyperlink);
    }

    private static ImageBlipTarget? ResolveImageBlipTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (!TryParseTargetId(target, out DocxTargetId imageId) || imageId.Kind != DocxTargetKind.Image)
        {
            return null;
        }

        if (imageId.Story == 'M')
        {
            return FindImageBlipTarget(package, package.MainDocumentPartName, imageId.Primary, cancellationToken);
        }

        if (imageId.Story is 'H' or 'F')
        {
            string relationshipType = imageId.Story == 'H' ? OoxmlRelTypes.Header : OoxmlRelTypes.Footer;
            string? partName = ResolveRelatedStoryPartName(package, relationshipType, imageId.StoryPart, cancellationToken);
            return partName is null ? null : FindImageBlipTarget(package, partName, imageId.Primary, cancellationToken);
        }

        return null;
    }

    private static ImageBlipTarget? FindImageBlipTarget(
        OoxmlPackage package,
        string partName,
        int imageOrdinal,
        CancellationToken cancellationToken)
    {
        if (imageOrdinal < 1)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
        IReadOnlyDictionary<string, OoxmlRelationship> relationships = package
            .GetRelationships(partName, cancellationToken)
            .ToDictionary(relationship => relationship.Id, StringComparer.Ordinal);
        var seenParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int currentOrdinal = 0;
        foreach (XElement blip in document.Descendants(OoxmlNs.A + "blip"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? relationshipId = (string?)blip.Attribute(OoxmlNs.R + "embed");
            if (relationshipId is null ||
                !relationships.TryGetValue(relationshipId, out OoxmlRelationship? relationship) ||
                relationship.IsExternal ||
                relationship.ResolvedTarget is null)
            {
                continue;
            }

            OoxmlPart? part = package.GetPart(relationship.ResolvedTarget);
            if (part is null || !seenParts.Add(part.Name))
            {
                continue;
            }

            currentOrdinal++;
            if (currentOrdinal == imageOrdinal)
            {
                return new ImageBlipTarget(partName, document, blip, relationshipId, part);
            }
        }

        return null;
    }

    private static string? ResolveRelatedStoryPartName(
        OoxmlPackage package,
        string relationshipType,
        int storyOrdinal,
        CancellationToken cancellationToken)
    {
        if (storyOrdinal < 1)
        {
            return null;
        }

        return package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => relationship.Type == relationshipType)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .ElementAtOrDefault(storyOrdinal - 1)
            ?.ResolvedTarget;
    }
}
