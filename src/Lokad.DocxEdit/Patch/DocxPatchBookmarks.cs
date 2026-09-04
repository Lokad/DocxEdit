using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

internal static partial class DocxPatchEngine
{
    private static IReadOnlyList<DocxDiagnostic> ExecuteReplaceBookmarkText(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        if (text is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        BookmarkTarget? bookmarkTarget = ResolveBookmarkTarget(package, target, cancellationToken);
        if (bookmarkTarget is null && !IsSupportedBookmarkTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported bookmark target '{target}'. Expected a bookmark ID such as M.B0001 or H001.B0001.", operation, target)];
        }

        if (bookmarkTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (bookmarkTarget.End is null ||
            bookmarkTarget.Start.Parent is null ||
            bookmarkTarget.End.Parent is null ||
            bookmarkTarget.Start.Parent.Name != OoxmlNs.W + "p" ||
            bookmarkTarget.End.Parent.Name != OoxmlNs.W + "p")
        {
            return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark '{target}' is not a paragraph-bounded range.", operation, target)];
        }

        XElement startParagraph = bookmarkTarget.Start.Parent;
        XElement endParagraph = bookmarkTarget.End.Parent;
        bool sameParagraph = startParagraph == endParagraph;
        XNode[] intermediateNodes = [];
        BookmarkTextSlot[]? structuredTextSlots = null;
        string? rangeFailure = null;
        XNode[] nodes = sameParagraph
            ? bookmarkTarget.Start.NodesAfterSelf()
                .TakeWhile(node => node != bookmarkTarget.End)
                .ToArray()
            : GetMultiParagraphBookmarkReplacementNodes(
                startParagraph,
                bookmarkTarget.Start,
                endParagraph,
                bookmarkTarget.End,
                out intermediateNodes,
                out rangeFailure);
        if (!sameParagraph && rangeFailure is not null)
        {
            if (!TryBuildTableSpanningBookmarkTextSlots(
                startParagraph,
                bookmarkTarget.Start,
                endParagraph,
                bookmarkTarget.End,
                out structuredTextSlots,
                out nodes,
                out string? structuredRangeFailure))
            {
                return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark '{target}' is not a supported same-story multi-paragraph range: {structuredRangeFailure ?? rangeFailure}.", operation, target)];
            }
        }

        if (ContainsProtectedBookmarkReplacementNode(nodes, out string? protectedFeature))
        {
            return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark '{target}' replacement would remove protected OOXML boundary '{protectedFeature}'.", operation, target)];
        }

        string[]? structuredReplacementLines = null;
        if (structuredTextSlots is not null)
        {
            structuredReplacementLines = SplitReplacementParagraphText(text);
            if (structuredReplacementLines.Length != structuredTextSlots.Length)
            {
                return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark '{target}' table-spanning replacement requires {structuredTextSlots.Length} replacement lines, one per visible text slot, but received {structuredReplacementLines.Length}.", operation, target)];
            }
        }

        bool useTrackedChanges = IsTrackedMode(options);
        TrackedBookmarkReplacement? trackedReplacement = null;
        if (useTrackedChanges)
        {
            if (!sameParagraph)
            {
                if (!TrackUnsupportedShape(
                    options,
                    operation,
                    target,
                    "tracked replace-bookmark-text supports only simple same-paragraph bookmark ranges",
                    diagnostics))
                {
                    return diagnostics;
                }

                useTrackedChanges = false;
            }
            else if (!TryBuildTrackedBookmarkReplacement(nodes, text, out trackedReplacement, out string? trackedUnsupportedReason))
            {
                if (!TrackUnsupportedShape(options, operation, target, trackedUnsupportedReason, diagnostics))
                {
                    return diagnostics;
                }

                useTrackedChanges = false;
            }
        }

        if (!apply)
        {
            return diagnostics;
        }

        if (useTrackedChanges && trackedReplacement is not null)
        {
            foreach (XNode node in nodes)
            {
                node.Remove();
            }

            XNode[] trackedNodes = CreateTrackedBookmarkReplacementNodes(
                package,
                trackedReplacement.DeletedText,
                text,
                trackedReplacement.RunProperties,
                options,
                generatedRevisionIds,
                cancellationToken);
            bookmarkTarget.Start.AddAfterSelf(trackedNodes);
        }
        else if (sameParagraph)
        {
            foreach (XNode node in nodes)
            {
                node.Remove();
            }

            bookmarkTarget.Start.AddAfterSelf(CreateSimpleRun(text));
        }
        else if (structuredTextSlots is not null && structuredReplacementLines is not null)
        {
            ReplaceStructuredBookmarkTextSlots(structuredTextSlots, structuredReplacementLines);
        }
        else
        {
            ReplaceMultiParagraphBookmarkText(
                startParagraph,
                bookmarkTarget.Start,
                endParagraph,
                bookmarkTarget.End,
                intermediateNodes,
                text);
        }

        SaveDocumentPart(package, bookmarkTarget.PartName, bookmarkTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteAddBookmark(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? name = ReadRequiredField(operation, "name", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        if (name is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!IsValidBookmarkName(name))
        {
            return [Diagnostic(DocxSeverity.Error, "E4205", "Field 'name' must be a non-empty bookmark name without whitespace.", operation, target)];
        }

        ParagraphTarget? paragraphTarget = ResolveParagraphTarget(package, operation, target, cancellationToken, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
        if (selectorDiagnostics.Count != 0)
        {
            return selectorDiagnostics;
        }

        if (paragraphTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        string current = ReadVisibleText(paragraphTarget.Paragraph);
        if (expected is not null && !string.Equals(current, expected, StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected text does not match current text.", operation, target)];
        }

        if (BookmarkNameExists(paragraphTarget.Document, name))
        {
            return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark name '{name}' already exists in part '{paragraphTarget.PartName}'.", operation, target)];
        }

        if (TryGetProtectedTextEditFeature(paragraphTarget.Paragraph, out string protectedFeature))
        {
            return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark creation for {target} would cross protected OOXML boundary '{protectedFeature}'.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        string id = AllocateBookmarkId(paragraphTarget.Document);
        var start = new XElement(
            OoxmlNs.W + "bookmarkStart",
            new XAttribute(OoxmlNs.W + "id", id),
            new XAttribute(OoxmlNs.W + "name", name));
        var end = new XElement(
            OoxmlNs.W + "bookmarkEnd",
            new XAttribute(OoxmlNs.W + "id", id));

        XElement? paragraphProperties = paragraphTarget.Paragraph.Element(OoxmlNs.W + "pPr");
        if (paragraphProperties is null)
        {
            paragraphTarget.Paragraph.AddFirst(start);
        }
        else
        {
            paragraphProperties.AddAfterSelf(start);
        }

        paragraphTarget.Paragraph.Add(end);
        SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteRenameBookmark(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? name = ReadRequiredField(operation, "name", diagnostics);
        if (name is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!IsValidBookmarkName(name))
        {
            return [Diagnostic(DocxSeverity.Error, "E4205", "Field 'name' must be a non-empty bookmark name without whitespace.", operation, target)];
        }

        BookmarkTarget? bookmarkTarget = ResolveBookmarkTarget(package, target, cancellationToken);
        if (bookmarkTarget is null && !IsSupportedBookmarkTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported bookmark target '{target}'. Expected a bookmark ID such as M.B0001 or H001.B0001.", operation, target)];
        }

        if (bookmarkTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        string? oldName = (string?)bookmarkTarget.Start.Attribute(OoxmlNs.W + "name");
        if (string.IsNullOrWhiteSpace(oldName))
        {
            return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark '{target}' has no current name.", operation, target)];
        }

        if (string.Equals(oldName, name, StringComparison.Ordinal))
        {
            return [];
        }

        if (BookmarkNameExists(bookmarkTarget.Document, bookmarkTarget.Start, name))
        {
            return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark name '{name}' already exists in part '{bookmarkTarget.PartName}'.", operation, target)];
        }

        if (CountBookmarkName(bookmarkTarget.Document, oldName) > 1 &&
            HasInternalHyperlinkAnchor(bookmarkTarget.Document, oldName))
        {
            return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark '{target}' has duplicate name '{oldName}' and same-part hyperlink anchors; rename would be ambiguous.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        bookmarkTarget.Start.SetAttributeValue(OoxmlNs.W + "name", name);
        UpdateInternalHyperlinkAnchors(bookmarkTarget.Document, oldName, name);
        SaveDocumentPart(package, bookmarkTarget.PartName, bookmarkTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteDeleteBookmark(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        BookmarkTarget? bookmarkTarget = ResolveBookmarkTarget(package, target, cancellationToken);
        if (bookmarkTarget is null && !IsSupportedBookmarkTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported bookmark target '{target}'. Expected a bookmark ID such as M.B0001 or H001.B0001.", operation, target)];
        }

        if (bookmarkTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (bookmarkTarget.End is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark '{target}' is incomplete and cannot be deleted safely.", operation, target)];
        }

        string? name = (string?)bookmarkTarget.Start.Attribute(OoxmlNs.W + "name");
        if (!string.IsNullOrWhiteSpace(name) &&
            HasInternalHyperlinkAnchor(bookmarkTarget.Document, name))
        {
            return [Diagnostic(DocxSeverity.Error, "E4311", $"Bookmark '{target}' is referenced by same-part hyperlink anchors; update or remove those hyperlinks before deleting the bookmark.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        bookmarkTarget.Start.Remove();
        bookmarkTarget.End.Remove();
        SaveDocumentPart(package, bookmarkTarget.PartName, bookmarkTarget.Document);
        return [];
    }
    private static XNode[] GetMultiParagraphBookmarkReplacementNodes(
        XElement startParagraph,
        XElement start,
        XElement endParagraph,
        XElement end,
        out XNode[] intermediateNodes,
        out string? failure)
    {
        intermediateNodes = [];
        failure = null;
        if (startParagraph.Parent is null || startParagraph.Parent != endParagraph.Parent)
        {
            failure = "markers must be in sibling paragraphs under the same container";
            return [];
        }

        var selected = new List<XNode>();
        selected.AddRange(start.NodesAfterSelf());

        var between = new List<XNode>();
        bool foundEndParagraph = false;
        foreach (XNode node in startParagraph.NodesAfterSelf())
        {
            if (node == endParagraph)
            {
                foundEndParagraph = true;
                break;
            }

            between.Add(node);
        }

        if (!foundEndParagraph)
        {
            failure = "end marker paragraph must follow the start marker paragraph";
            return [];
        }

        XElement? unsupportedBlock = between
            .OfType<XElement>()
            .FirstOrDefault(element => element.Name != OoxmlNs.W + "p");
        if (unsupportedBlock is not null)
        {
            failure = $"range crosses unsupported block '{unsupportedBlock.Name.LocalName}'";
            return [];
        }

        intermediateNodes = between.ToArray();
        selected.AddRange(intermediateNodes);
        selected.AddRange(endParagraph.Nodes().TakeWhile(node => node != end));
        return selected.ToArray();
    }

    private static bool TryBuildTableSpanningBookmarkTextSlots(
        XElement startParagraph,
        XElement start,
        XElement endParagraph,
        XElement end,
        out BookmarkTextSlot[] slots,
        out XNode[] selectedNodes,
        out string? failure)
    {
        slots = [];
        selectedNodes = [];
        failure = null;
        if (startParagraph.Parent is null || startParagraph.Parent != endParagraph.Parent)
        {
            failure = "markers must be in sibling paragraphs under the same container";
            return false;
        }

        var between = new List<XNode>();
        bool foundEndParagraph = false;
        foreach (XNode node in startParagraph.NodesAfterSelf())
        {
            if (node == endParagraph)
            {
                foundEndParagraph = true;
                break;
            }

            between.Add(node);
        }

        if (!foundEndParagraph)
        {
            failure = "end marker paragraph must follow the start marker paragraph";
            return false;
        }

        if (!between.OfType<XElement>().Any(element => element.Name == OoxmlNs.W + "tbl"))
        {
            failure = "range does not cross a table block";
            return false;
        }

        XElement? unsupportedBlock = between
            .OfType<XElement>()
            .FirstOrDefault(element => element.Name != OoxmlNs.W + "p" && element.Name != OoxmlNs.W + "tbl");
        if (unsupportedBlock is not null)
        {
            failure = $"range crosses unsupported block '{unsupportedBlock.Name.LocalName}'";
            return false;
        }

        var slotBuilder = new List<BookmarkTextSlot>();
        var selectedBuilder = new List<XNode>();
        AddBookmarkTextSlotIfVisible(slotBuilder, selectedBuilder, startParagraph, start.NodesAfterSelf().ToArray(), AddAfter: start, AddBefore: null);
        foreach (XElement block in between.OfType<XElement>())
        {
            if (block.Name == OoxmlNs.W + "p")
            {
                AddWholeParagraphBookmarkTextSlot(slotBuilder, selectedBuilder, block);
                continue;
            }

            if (block.Descendants(OoxmlNs.W + "tbl").Any())
            {
                failure = "table-spanning bookmark range contains a nested table";
                return false;
            }

            foreach (XElement cell in block.Descendants(OoxmlNs.W + "tc"))
            {
                XElement? unsupportedCellBlock = cell
                    .Elements()
                    .FirstOrDefault(element => element.Name != OoxmlNs.W + "tcPr" && element.Name != OoxmlNs.W + "p");
                if (unsupportedCellBlock is not null)
                {
                    failure = $"table cell contains unsupported block '{unsupportedCellBlock.Name.LocalName}'";
                    return false;
                }

                foreach (XElement paragraph in cell.Elements(OoxmlNs.W + "p"))
                {
                    AddWholeParagraphBookmarkTextSlot(slotBuilder, selectedBuilder, paragraph);
                }
            }
        }

        AddBookmarkTextSlotIfVisible(slotBuilder, selectedBuilder, endParagraph, endParagraph.Nodes().TakeWhile(node => node != end).ToArray(), AddAfter: null, AddBefore: end);
        if (slotBuilder.Count == 0)
        {
            failure = "range contains no visible text slots";
            return false;
        }

        slots = slotBuilder.ToArray();
        selectedNodes = selectedBuilder.ToArray();
        return true;
    }

    private static void AddWholeParagraphBookmarkTextSlot(
        List<BookmarkTextSlot> slots,
        List<XNode> selectedNodes,
        XElement paragraph)
    {
        XNode[] nodes = paragraph.Nodes()
            .Where(node => node is not XElement element || element.Name != OoxmlNs.W + "pPr")
            .ToArray();
        XElement? paragraphProperties = paragraph.Element(OoxmlNs.W + "pPr");
        AddBookmarkTextSlotIfVisible(slots, selectedNodes, paragraph, nodes, paragraphProperties, AddBefore: null);
    }

    private static void AddBookmarkTextSlotIfVisible(
        List<BookmarkTextSlot> slots,
        List<XNode> selectedNodes,
        XElement container,
        XNode[] nodes,
        XElement? AddAfter,
        XElement? AddBefore)
    {
        selectedNodes.AddRange(nodes);
        if (ReadVisibleText(new XElement(OoxmlNs.W + "p", nodes)).Length == 0)
        {
            return;
        }

        slots.Add(new BookmarkTextSlot(container, nodes, AddAfter, AddBefore));
    }

    private static void ReplaceStructuredBookmarkTextSlots(IReadOnlyList<BookmarkTextSlot> slots, IReadOnlyList<string> lines)
    {
        for (int i = 0; i < slots.Count; i++)
        {
            BookmarkTextSlot slot = slots[i];
            foreach (XNode node in slot.Nodes)
            {
                node.Remove();
            }

            XElement replacementRun = CreateSimpleRun(lines[i]);
            if (slot.AddAfter is not null)
            {
                slot.AddAfter.AddAfterSelf(replacementRun);
            }
            else if (slot.AddBefore is not null)
            {
                slot.AddBefore.AddBeforeSelf(replacementRun);
            }
            else if (slot.Container.Element(OoxmlNs.W + "pPr") is { } paragraphProperties)
            {
                paragraphProperties.AddAfterSelf(replacementRun);
            }
            else
            {
                slot.Container.AddFirst(replacementRun);
            }
        }
    }

    private static void ReplaceMultiParagraphBookmarkText(
        XElement startParagraph,
        XElement start,
        XElement endParagraph,
        XElement end,
        XNode[] intermediateNodes,
        string text)
    {
        string[] paragraphs = SplitReplacementParagraphText(text);

        foreach (XNode node in start.NodesAfterSelf().ToArray())
        {
            node.Remove();
        }

        start.AddAfterSelf(CreateSimpleRun(paragraphs[0]));

        foreach (XNode node in intermediateNodes)
        {
            node.Remove();
        }

        if (paragraphs.Length == 1)
        {
            XNode[] trailingNodes = end.NodesAfterSelf().ToArray();
            end.Remove();
            startParagraph.Add(end);
            foreach (XNode node in trailingNodes)
            {
                startParagraph.Add(node);
            }

            endParagraph.Remove();
            return;
        }

        for (int i = 1; i < paragraphs.Length - 1; i++)
        {
            endParagraph.AddBeforeSelf(CreateSimpleParagraph(paragraphs[i], style: null, paragraphProperties: null));
        }

        foreach (XNode node in endParagraph.Nodes().TakeWhile(node => node != end).ToArray())
        {
            if (node is XElement element && element.Name == OoxmlNs.W + "pPr")
            {
                continue;
            }

            node.Remove();
        }

        end.AddBeforeSelf(CreateSimpleRun(paragraphs[^1]));
    }

    private static string[] SplitReplacementParagraphText(string text)
    {
        return text
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Split('\n');
    }

    private static bool ContainsProtectedBookmarkReplacementNode(IEnumerable<XNode> nodes, out string? feature)
    {
        foreach (XElement element in nodes.OfType<XElement>().SelectMany(ElementAndDescendants))
        {
            if (ProtectedTextEditElements.TryGetValue(element.Name, out string? protectedFeature) &&
                element.Name != OoxmlNs.W + "bookmarkStart" &&
                element.Name != OoxmlNs.W + "bookmarkEnd")
            {
                feature = protectedFeature;
                return true;
            }
        }

        feature = null;
        return false;
    }

    private static bool TryBuildTrackedBookmarkReplacement(
        IReadOnlyList<XNode> nodes,
        string replacement,
        [NotNullWhen(true)] out TrackedBookmarkReplacement? trackedReplacement,
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        trackedReplacement = null;
        unsupportedReason = null;

        var clonedRuns = new List<XElement>();
        foreach (XNode node in nodes)
        {
            if (node is XText text && string.IsNullOrWhiteSpace(text.Value))
            {
                continue;
            }

            if (node is not XElement element || element.Name != OoxmlNs.W + "r")
            {
                unsupportedReason = "bookmark range contains non-run content";
                return false;
            }

            clonedRuns.Add(new XElement(element));
        }

        var rangeContainer = new XElement(OoxmlNs.W + "p", clonedRuns);
        string current = ReadVisibleText(rangeContainer);
        if (!TryValidateTrackedWholeParagraphReplacement(rangeContainer, current, replacement, style: null, out unsupportedReason))
        {
            return false;
        }

        XElement? firstRunProperties = rangeContainer
            .Elements(OoxmlNs.W + "r")
            .Elements(OoxmlNs.W + "rPr")
            .FirstOrDefault();
        trackedReplacement = new TrackedBookmarkReplacement(
            current,
            firstRunProperties is null ? null : new XElement(firstRunProperties));
        return true;
    }

    private static bool IsValidBookmarkName(string name)
    {
        return !string.IsNullOrWhiteSpace(name) && !name.Any(char.IsWhiteSpace);
    }

    private static int CountBookmarkName(XDocument document, string name)
    {
        return document
            .Descendants(OoxmlNs.W + "bookmarkStart")
            .Count(bookmark => string.Equals((string?)bookmark.Attribute(OoxmlNs.W + "name"), name, StringComparison.Ordinal));
    }

    private static bool BookmarkNameExists(XDocument document, XElement excludedStart, string name)
    {
        return document
            .Descendants(OoxmlNs.W + "bookmarkStart")
            .Any(bookmark => bookmark != excludedStart &&
                string.Equals((string?)bookmark.Attribute(OoxmlNs.W + "name"), name, StringComparison.Ordinal));
    }

    private static bool BookmarkNameExists(XDocument document, string name)
    {
        return document
            .Descendants(OoxmlNs.W + "bookmarkStart")
            .Any(bookmark => string.Equals((string?)bookmark.Attribute(OoxmlNs.W + "name"), name, StringComparison.Ordinal));
    }

    private static string AllocateBookmarkId(XDocument document)
    {
        int maxId = document
            .Descendants()
            .Where(element => element.Name == OoxmlNs.W + "bookmarkStart" || element.Name == OoxmlNs.W + "bookmarkEnd")
            .Select(element => (string?)element.Attribute(OoxmlNs.W + "id"))
            .Select(value => int.TryParse(value, out int id) ? id : 0)
            .DefaultIfEmpty(0)
            .Max();
        return (maxId + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    private static bool HasInternalHyperlinkAnchor(XDocument document, string anchor)
    {
        return document
            .Descendants(OoxmlNs.W + "hyperlink")
            .Any(hyperlink => string.Equals((string?)hyperlink.Attribute(OoxmlNs.W + "anchor"), anchor, StringComparison.Ordinal));
    }

    private static void UpdateInternalHyperlinkAnchors(XDocument document, string oldName, string newName)
    {
        foreach (XElement hyperlink in document.Descendants(OoxmlNs.W + "hyperlink"))
        {
            if (string.Equals((string?)hyperlink.Attribute(OoxmlNs.W + "anchor"), oldName, StringComparison.Ordinal))
            {
                hyperlink.SetAttributeValue(OoxmlNs.W + "anchor", newName);
            }
        }
    }
}
