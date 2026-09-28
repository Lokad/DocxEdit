using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

internal static partial class DocxPatchEngine
{
    private static bool TryValidateTrackedTextReplacement(
        XElement paragraph,
        string current,
        IReadOnlyList<TextRange> matches,
        string replacement,
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        unsupportedReason = null;
        if (replacement.Contains('\t') || replacement.Contains('\n'))
        {
            unsupportedReason = "replacement contains tabs or line breaks";
            return false;
        }

        foreach (TextRange match in matches)
        {
            string deletedText = current.Substring(match.Start, match.Length);
            if (deletedText.Contains('\t') || deletedText.Contains('\n'))
            {
                unsupportedReason = "matched text contains tabs or line breaks";
                return false;
            }
        }

        if (TryGetUnsupportedTrackedRunContent(paragraph, out string unsupportedContent))
        {
            unsupportedReason = $"paragraph contains unsupported run content '{unsupportedContent}'";
            return false;
        }

        if (HasMixedDirectTextRunProperties(paragraph))
        {
            unsupportedReason = "paragraph contains mixed direct run formatting";
            return false;
        }

        return true;
    }

    private static bool TryValidateTrackedWholeParagraphReplacement(
        XElement paragraph,
        string current,
        string replacement,
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        unsupportedReason = null;
        if (TextContainsTrackedUnsupportedCharacters(current) ||
            TextContainsTrackedUnsupportedCharacters(replacement))
        {
            unsupportedReason = "tracked paragraph text contains tabs or line breaks";
            return false;
        }

        if (TryGetUnsupportedTrackedRunContent(paragraph, out string unsupportedContent))
        {
            unsupportedReason = $"paragraph contains unsupported run content '{unsupportedContent}'";
            return false;
        }

        if (HasMixedDirectTextRunProperties(paragraph))
        {
            unsupportedReason = "paragraph contains mixed direct run formatting";
            return false;
        }

        return true;
    }

    private static bool TryGetTrackedSetCellParagraphs(
        XElement cell,
        string replacement,
        out XElement[] paragraphs,
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        paragraphs = cell.Elements(OoxmlNs.W + "p").ToArray();
        unsupportedReason = null;
        if (paragraphs.Length == 0)
        {
            unsupportedReason = "cell has no paragraph for tracked text replacement";
            return false;
        }

        if (cell.Elements().Any(element => element.Name != OoxmlNs.W + "tcPr" && element.Name != OoxmlNs.W + "p"))
        {
            unsupportedReason = "cell contains non-paragraph block content";
            return false;
        }

        if (cell.Descendants(OoxmlNs.W + "drawing").Any())
        {
            unsupportedReason = "cell contains drawing content";
            return false;
        }

        if (cell.Descendants(OoxmlNs.W + "fldSimple").Any() ||
            cell.Descendants(OoxmlNs.W + "fldChar").Any() ||
            cell.Descendants(OoxmlNs.W + "instrText").Any())
        {
            unsupportedReason = "cell contains field content";
            return false;
        }

        for (int i = 0; i < paragraphs.Length; i++)
        {
            XElement paragraph = paragraphs[i];
            if (TryGetProtectedTextEditFeature(paragraph, out string protectedFeature))
            {
                unsupportedReason = $"cell paragraph contains protected OOXML boundary '{protectedFeature}'";
                return false;
            }

            string current = ReadVisibleText(paragraph);
            string insertedText = i == 0 ? replacement : string.Empty;
            if (!TryValidateTrackedWholeParagraphReplacement(paragraph, current, insertedText, out unsupportedReason))
            {
                return false;
            }
        }

        return true;
    }

    private static bool TryGetTrackedContentControlTextContainer(
        XElement content,
        string replacement,
        [NotNullWhen(true)] out XElement? trackedContainer,
        out string current,
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        trackedContainer = null;
        current = string.Empty;
        unsupportedReason = null;
        XElement[] paragraphs = content.Elements(OoxmlNs.W + "p").ToArray();
        if (paragraphs.Length > 1)
        {
            unsupportedReason = "plain-text content control contains multiple paragraphs";
            return false;
        }

        if (paragraphs.Length == 1)
        {
            if (content.Elements().Any(element => element.Name != OoxmlNs.W + "p"))
            {
                unsupportedReason = "plain-text content control mixes paragraph and non-paragraph content";
                return false;
            }

            trackedContainer = paragraphs[0];
        }
        else
        {
            if (!content.Elements(OoxmlNs.W + "r").Any())
            {
                unsupportedReason = "plain-text content control has no run content";
                return false;
            }

            if (content.Elements().Any(element => element.Name != OoxmlNs.W + "r"))
            {
                unsupportedReason = "plain-text content control contains non-run content";
                return false;
            }

            trackedContainer = content;
        }

        if (TryGetProtectedTextEditFeature(trackedContainer, out string protectedFeature))
        {
            unsupportedReason = $"plain-text content control contains protected OOXML boundary '{protectedFeature}'";
            return false;
        }

        current = ReadVisibleText(trackedContainer);
        if (!TryValidateTrackedWholeParagraphReplacement(trackedContainer, current, replacement, out unsupportedReason))
        {
            return false;
        }

        return true;
    }

    private static bool TryGetTrackedRichTextContentControlParagraphs(
        XElement content,
        string replacement,
        out XElement[] paragraphs,
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        paragraphs = content.Elements(OoxmlNs.W + "p").ToArray();
        unsupportedReason = null;
        if (paragraphs.Length == 0)
        {
            unsupportedReason = "rich-text content control has no paragraph for tracked text replacement";
            return false;
        }

        if (content.Elements().Any(element => element.Name != OoxmlNs.W + "p"))
        {
            unsupportedReason = "rich-text content control contains non-paragraph content";
            return false;
        }

        if (content.Descendants(OoxmlNs.W + "drawing").Any())
        {
            unsupportedReason = "rich-text content control contains drawing content";
            return false;
        }

        if (content.Descendants(OoxmlNs.W + "fldSimple").Any() ||
            content.Descendants(OoxmlNs.W + "fldChar").Any() ||
            content.Descendants(OoxmlNs.W + "instrText").Any())
        {
            unsupportedReason = "rich-text content control contains field content";
            return false;
        }

        for (int i = 0; i < paragraphs.Length; i++)
        {
            XElement paragraph = paragraphs[i];
            if (TryGetProtectedTextEditFeature(paragraph, out string protectedFeature))
            {
                unsupportedReason = $"rich-text content-control paragraph contains protected OOXML boundary '{protectedFeature}'";
                return false;
            }

            string current = ReadVisibleText(paragraph);
            string insertedText = i == 0 ? replacement : string.Empty;
            if (!TryValidateTrackedWholeParagraphReplacement(paragraph, current, insertedText, out unsupportedReason))
            {
                return false;
            }
        }

        return true;
    }

    private static bool TryGetTrackedCommentParagraphs(
        XElement comment,
        string replacement,
        out XElement[] paragraphs,
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        paragraphs = comment.Elements(OoxmlNs.W + "p").ToArray();
        unsupportedReason = null;
        if (paragraphs.Length == 0)
        {
            unsupportedReason = "comment body has no paragraph for tracked text replacement";
            return false;
        }

        if (comment.Elements().Any(element => element.Name != OoxmlNs.W + "p"))
        {
            unsupportedReason = "comment body contains non-paragraph content";
            return false;
        }

        if (comment.Descendants(OoxmlNs.W + "drawing").Any())
        {
            unsupportedReason = "comment body contains drawing content";
            return false;
        }

        if (comment.Descendants(OoxmlNs.W + "fldSimple").Any() ||
            comment.Descendants(OoxmlNs.W + "fldChar").Any() ||
            comment.Descendants(OoxmlNs.W + "instrText").Any())
        {
            unsupportedReason = "comment body contains field content";
            return false;
        }

        for (int i = 0; i < paragraphs.Length; i++)
        {
            XElement paragraph = paragraphs[i];
            if (TryGetProtectedTextEditFeature(paragraph, out string protectedFeature))
            {
                unsupportedReason = $"comment body paragraph contains protected OOXML boundary '{protectedFeature}'";
                return false;
            }

            string current = ReadVisibleText(paragraph);
            string insertedText = i == 0 ? replacement : string.Empty;
            if (!TryValidateTrackedWholeParagraphReplacement(paragraph, current, insertedText, out unsupportedReason))
            {
                return false;
            }
        }

        return true;
    }

    private static bool TextContainsTrackedUnsupportedCharacters(string text)
    {
        return text.Contains('\t') || text.Contains('\n');
    }

    private static bool IsTrackedMode(DocxEditOptions options)
    {
        return options.TrackChanges is TrackChangesMode.Require or TrackChangesMode.Suggest;
    }

    private static bool TrackUnsupportedShape(
        DocxEditOptions options,
        DocxPatchOperation operation,
        string target,
        string reason,
        List<DocxDiagnostic> diagnostics)
    {
        if (options.TrackChanges == TrackChangesMode.Require)
        {
            diagnostics.Add(Diagnostic(
                DocxSeverity.Error,
                "E6002",
                BuildUnsupportedTrackedShapeMessage(options.TrackChanges, operation, target, reason),
                operation,
                target,
                "track-changes-unsupported-target-shape",
                "require-failed"));
            return false;
        }

        diagnostics.Add(Diagnostic(
            DocxSeverity.Warning,
            "W4002",
            BuildUnsupportedTrackedShapeMessage(options.TrackChanges, operation, target, reason),
            operation,
            target,
            "track-changes-unsupported-target-shape",
            "direct-edit-preserve-existing-revisions"));
        return true;
    }

    /// <summary>Records an unsupported tracked shape and falls back to a direct edit. Returns false when the caller must return immediately.</summary>
    private static bool TryFallbackToDirectEdit(
        DocxEditOptions options,
        DocxPatchOperation operation,
        string target,
        string reason,
        List<DocxDiagnostic> diagnostics,
        ref bool useTrackedChanges)
    {
        if (!TrackUnsupportedShape(options, operation, target, reason, diagnostics))
        {
            return false;
        }

        useTrackedChanges = false;
        return true;
    }

    private static string BuildUnsupportedTrackedShapeMessage(
        TrackChangesMode mode,
        DocxPatchOperation operation,
        string target,
        string reason)
    {
        string operationName = operation.OperationName;
        string support = TrackChangesSupportValue(operationName);

        return mode switch
        {
            TrackChangesMode.Require =>
                $"TrackChangesMode.Require cannot apply operation '{operationName}' as tracked output for target '{target}' because catalog support is '{support}' but this target shape is unsupported: {reason}.",
            TrackChangesMode.Suggest =>
                $"TrackChangesMode.Suggest will apply operation '{operationName}' directly for target '{target}' because catalog support is '{support}' but this target shape is unsupported: {reason}. Existing tracked-change markup is preserved, but this edit will not create new revision markup.",
            _ =>
                $"TrackChangesMode.{mode} cannot generate tracked output for operation '{operationName}' on target '{target}' because catalog support is '{support}' but this target shape is unsupported: {reason}."
        };
    }

    private static bool TryGetUnsupportedTrackedRunContent(XElement paragraph, out string unsupportedContent)
    {
        foreach (XElement run in paragraph.Elements(OoxmlNs.W + "r"))
        {
            foreach (XElement child in run.Elements())
            {
                if (child.Name == OoxmlNs.W + "rPr" ||
                    child.Name == OoxmlNs.W + "t" ||
                    child.Name == OoxmlNs.W + "tab" ||
                    child.Name == OoxmlNs.W + "br")
                {
                    continue;
                }

                unsupportedContent = child.Name.LocalName;
                return true;
            }
        }

        unsupportedContent = string.Empty;
        return false;
    }

    private static bool HasMixedDirectTextRunProperties(XElement paragraph)
    {
        string? firstSignature = null;
        foreach (XElement run in paragraph.Elements(OoxmlNs.W + "r"))
        {
            if (!RunHasVisibleText(run))
            {
                continue;
            }

            string signature = CanonicalRunPropertiesSignature(run.Element(OoxmlNs.W + "rPr"));
            if (firstSignature is null)
            {
                firstSignature = signature;
                continue;
            }

            if (!string.Equals(firstSignature, signature, StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    private static string CanonicalRunPropertiesSignature(XElement? runProperties)
    {
        if (runProperties is null)
        {
            return string.Empty;
        }

        var builder = new StringBuilder();
        AppendCanonicalElement(builder, runProperties);
        return builder.ToString();
    }

    private static void AppendCanonicalElement(StringBuilder builder, XElement element)
    {
        builder.Append('{')
            .Append(element.Name.NamespaceName)
            .Append('}')
            .Append(element.Name.LocalName)
            .Append('[');

        foreach (XAttribute attribute in element.Attributes()
            .Where(attribute => !attribute.IsNamespaceDeclaration)
            .OrderBy(attribute => attribute.Name.NamespaceName, StringComparer.Ordinal)
            .ThenBy(attribute => attribute.Name.LocalName, StringComparer.Ordinal)
            .ThenBy(attribute => attribute.Value, StringComparer.Ordinal))
        {
            builder.Append('{')
                .Append(attribute.Name.NamespaceName)
                .Append('}')
                .Append(attribute.Name.LocalName)
                .Append('=')
                .Append(attribute.Value)
                .Append(';');
        }

        builder.Append(']');

        foreach (string childSignature in element.Nodes()
            .Select(CanonicalNodeSignature)
            .Where(signature => signature.Length != 0)
            .OrderBy(signature => signature, StringComparer.Ordinal))
        {
            builder.Append(childSignature);
        }

        static string CanonicalNodeSignature(XNode node)
        {
            if (node is XElement child)
            {
                var childBuilder = new StringBuilder();
                AppendCanonicalElement(childBuilder, child);
                return childBuilder.ToString();
            }

            if (node is XText text && !string.IsNullOrWhiteSpace(text.Value))
            {
                return text.Value;
            }

            return string.Empty;
        }
    }

    private static bool RunHasVisibleText(XElement run)
    {
        return run.Elements().Any(element =>
            element.Name == OoxmlNs.W + "t" ||
            element.Name == OoxmlNs.W + "tab" ||
            element.Name == OoxmlNs.W + "br");
    }

    private static void ReplaceParagraphTextWithTrackedChanges(
        OoxmlPackage package,
        XElement paragraph,
        string current,
        IReadOnlyList<TextRange> matches,
        string replacement,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        string[] revisionIds = AllocateRevisionIds(package, matches.Count * 2, generatedRevisionIds, cancellationToken);
        XElement? paragraphProperties = paragraph.Element(OoxmlNs.W + "pPr");
        XElement? firstRunProperties = paragraph
            .Elements(OoxmlNs.W + "r")
            .Elements(OoxmlNs.W + "rPr")
            .FirstOrDefault();
        string author = GetRevisionAuthor(options);
        string timestamp = GetRevisionTimestamp(options);

        var nodes = new List<XNode>();
        if (paragraphProperties is not null)
        {
            nodes.Add(new XElement(paragraphProperties));
        }

        int cursor = 0;
        int revisionIndex = 0;
        foreach (TextRange match in matches)
        {
            if (match.Start > cursor)
            {
                AddTextRun(nodes, current[cursor..match.Start], firstRunProperties);
            }

            string deletedText = current.Substring(match.Start, match.Length);
            nodes.Add(CreateDeletedRun(deletedText, firstRunProperties, revisionIds[revisionIndex++], author, timestamp));
            nodes.Add(CreateInsertedRun(replacement, firstRunProperties, revisionIds[revisionIndex++], author, timestamp));
            cursor = match.Start + match.Length;
        }

        if (cursor < current.Length)
        {
            AddTextRun(nodes, current[cursor..], firstRunProperties);
        }

        paragraph.RemoveNodes();
        paragraph.Add(nodes);
    }

    // D05: span-scoped validation for the preserving tracked rewrite below.
    // Tab rules reuse the whole-paragraph messages; run-content rules apply to
    // runs the replacement actually touches, while surrounding markup is
    // preserved untouched.
    private static bool TryValidateTrackedSpanReplacement(
        List<VisibleCharEntry> map,
        string current,
        IReadOnlyList<TextRange> matches,
        string replacement,
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        unsupportedReason = null;
        if (TextContainsTrackedUnsupportedCharacters(replacement))
        {
            unsupportedReason = "replacement contains tabs or line breaks";
            return false;
        }

        foreach (TextRange match in matches)
        {
            string deletedText = current.Substring(match.Start, match.Length);
            if (TextContainsTrackedUnsupportedCharacters(deletedText))
            {
                unsupportedReason = "matched text contains tabs or line breaks";
                return false;
            }

            if (!TryCheckTrackedSpanRuns(map, match, out unsupportedReason))
            {
                return false;
            }

            if (SpanHasMixedRunProperties(map, match))
            {
                unsupportedReason = "paragraph contains mixed direct run formatting";
                return false;
            }
        }

        return true;
    }

    // Boundary runs keep unmatched siblings through splits, so only children
    // positioned inside the match can be lost; interior runs are removed whole
    // and keep no non-text children beyond text markers.
    private static bool TryCheckTrackedSpanRuns(
        List<VisibleCharEntry> map,
        TextRange match,
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        unsupportedReason = null;
        XElement startElement = map[match.Start].TextElement;
        XElement endElement = map[match.Start + match.Length - 1].TextElement;
        XElement? startRun = startElement.Parent;
        XElement? endRun = endElement.Parent;
        var seen = new HashSet<XElement>(ReferenceEqualityComparer.Instance);
        for (int i = match.Start; i < match.Start + match.Length; i++)
        {
            VisibleCharEntry entry = map[i];
            if (!entry.IsText)
            {
                continue;
            }

            XElement run = entry.TextElement.Parent!;
            if (!seen.Add(run))
            {
                continue;
            }

            foreach (XElement child in SpanMatchedChildren(run, startElement, endElement, startRun, endRun))
            {
                if (child.Name == OoxmlNs.W + "rPr" ||
                    child.Name == OoxmlNs.W + "t" ||
                    child.Name == OoxmlNs.W + "tab" ||
                    child.Name == OoxmlNs.W + "br")
                {
                    continue;
                }

                unsupportedReason = "paragraph contains unsupported run content " + Quote(child.Name.LocalName);
                return false;
            }
        }

        return true;
    }

    // Children of one span run positioned inside the match: from the start
    // element onward in the first run, up to the end element in the last run,
    // and everything in between for interior runs.
    private static IEnumerable<XElement> SpanMatchedChildren(
        XElement run,
        XElement startElement,
        XElement endElement,
        XElement? startRun,
        XElement? endRun)
    {
        IEnumerable<XElement> children = run.Elements();
        if (ReferenceEquals(run, startRun))
        {
            children = children.SkipWhile(child => !ReferenceEquals(child, startElement));
        }

        if (ReferenceEquals(run, endRun))
        {
            children = children.TakeWhile(child => !ReferenceEquals(child, endElement)).Concat([endElement]);
        }

        return children;
    }

    // Whether the match spans runs with different direct formatting; each
    // replacement carries the first span run properties.
    private static bool SpanHasMixedRunProperties(List<VisibleCharEntry> map, TextRange match)
    {
        string? firstSignature = null;
        XElement? previousRun = null;
        for (int i = match.Start; i < match.Start + match.Length; i++)
        {
            VisibleCharEntry entry = map[i];
            if (!entry.IsText)
            {
                continue;
            }

            XElement run = entry.TextElement.Parent!;
            if (previousRun is not null && ReferenceEquals(previousRun, run))
            {
                continue;
            }

            previousRun = run;
            string signature = CanonicalRunPropertiesSignature(run.Element(OoxmlNs.W + "rPr"));
            if (firstSignature is null)
            {
                firstSignature = signature;
                continue;
            }

            if (!string.Equals(firstSignature, signature, StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }



    // Isolates exactly the runs covering one match by splitting boundary runs,
    // preserving unmatched siblings and run properties in place. Each split
    // relocates through a fresh map, so no detached references escape.
    // Returns null when the span cannot be resolved to run boundaries.
    private static List<XElement>? IsolateMatchRuns(XElement paragraph, TextRange match)
    {
        List<VisibleCharEntry> map = BuildVisibleTextMap(paragraph);
        if (match.Start < 0 || match.Length <= 0 || match.Start + match.Length > map.Count)
        {
            return null;
        }

        VisibleCharEntry startEntry = map[match.Start];
        if (!startEntry.IsText)
        {
            return null;
        }

        XElement? startRun = startEntry.TextElement.Parent;
        if (startRun is null || startRun.Name != OoxmlNs.W + "r")
        {
            return null;
        }

        if (startEntry.OffsetInElement > 0)
        {
            SplitRunAtTextPosition(new TextPosition(startEntry.TextElement, startEntry.OffsetInElement));
            map = BuildVisibleTextMap(paragraph);
        }

        VisibleCharEntry endEntry = map[match.Start + match.Length - 1];
        if (!endEntry.IsText)
        {
            return null;
        }

        XElement? endRun = endEntry.TextElement.Parent;
        if (endRun is null || endRun.Name != OoxmlNs.W + "r")
        {
            return null;
        }

        if (endEntry.OffsetInElement + 1 < endEntry.TextElement.Value.Length)
        {
            SplitRunAtTextPosition(new TextPosition(endEntry.TextElement, endEntry.OffsetInElement + 1));
            map = BuildVisibleTextMap(paragraph);
        }

        VisibleCharEntry firstEntry = map[match.Start];
        VisibleCharEntry lastEntry = map[match.Start + match.Length - 1];
        if (!firstEntry.IsText || !lastEntry.IsText)
        {
            return null;
        }

        XElement? firstRun = firstEntry.TextElement.Parent;
        XElement? lastRun = lastEntry.TextElement.Parent;
        if (firstRun is null || lastRun is null)
        {
            return null;
        }

        var mids = new List<XElement>();
        XElement? cursor = firstRun;
        while (cursor is not null)
        {
            if (cursor.Name == OoxmlNs.W + "r" && cursor.Elements(OoxmlNs.W + "t").Any(static element => element.Value.Length != 0))
            {
                mids.Add(cursor);
            }

            if (ReferenceEquals(cursor, lastRun))
            {
                break;
            }

            cursor = cursor.ElementsAfterSelf().FirstOrDefault();
        }

        if (mids.Count == 0 || !ReferenceEquals(mids[mids.Count - 1], lastRun))
        {
            return null;
        }

        return mids;
    }


    // D05: tracked replacement that preserves surrounding markup in place.
    // Matched spans become delete/insert revision runs; every other node keeps
    // its position and content, so existing revisions, markers, and wrappers
    // survive with their history intact.
    private static bool ReplaceParagraphTextWithTrackedSpans(
        OoxmlPackage package,
        XElement paragraph,
        string current,
        IReadOnlyList<TextRange> matches,
        string replacement,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken,
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        unsupportedReason = null;
        string[] revisionIds = AllocateRevisionIds(package, matches.Count * 2, generatedRevisionIds, cancellationToken);
        string author = GetRevisionAuthor(options);
        string timestamp = GetRevisionTimestamp(options);
        for (int i = matches.Count - 1; i >= 0; i--)
        {
            TextRange match = matches[i];
            List<XElement>? mids = IsolateMatchRuns(paragraph, match);
            if (mids is null || mids.Count == 0)
            {
                unsupportedReason = "match spans cannot be isolated to editable run boundaries";
                return false;
            }

            XElement? spanProperties = mids[0].Element(OoxmlNs.W + "rPr");
            string deletedText = current.Substring(match.Start, match.Length);
            XElement deletedRun = CreateDeletedRun(deletedText, spanProperties is null ? null : new XElement(spanProperties), revisionIds[i * 2], author, timestamp);
            XElement insertedRun = CreateInsertedRun(replacement, spanProperties is null ? null : new XElement(spanProperties), revisionIds[i * 2 + 1], author, timestamp);
            mids[0].AddBeforeSelf(deletedRun);
            mids[0].AddBeforeSelf(insertedRun);
            foreach (XElement mid in mids)
            {
                mid.Remove();
            }
        }

        return true;
    }

    private static void ReplaceWholeParagraphTextWithTrackedChanges(
        OoxmlPackage package,
        XElement paragraph,
        string deletedText,
        string insertedText,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        int revisionCount = (deletedText.Length == 0 ? 0 : 1) + (insertedText.Length == 0 ? 0 : 1);
        string[] revisionIds = AllocateRevisionIds(package, revisionCount, generatedRevisionIds, cancellationToken);
        XElement? paragraphProperties = paragraph.Element(OoxmlNs.W + "pPr");
        XElement? firstRunProperties = paragraph
            .Elements(OoxmlNs.W + "r")
            .Elements(OoxmlNs.W + "rPr")
            .FirstOrDefault();
        string author = GetRevisionAuthor(options);
        string timestamp = GetRevisionTimestamp(options);

        var nodes = new List<XNode>();
        if (paragraphProperties is not null)
        {
            nodes.Add(new XElement(paragraphProperties));
        }

        int revisionIndex = 0;
        if (deletedText.Length != 0)
        {
            nodes.Add(CreateDeletedRun(deletedText, firstRunProperties, revisionIds[revisionIndex++], author, timestamp));
        }

        if (insertedText.Length != 0)
        {
            nodes.Add(CreateInsertedRun(insertedText, firstRunProperties, revisionIds[revisionIndex], author, timestamp));
        }

        paragraph.RemoveNodes();
        paragraph.Add(nodes);
    }

    private static void ReplaceCellParagraphTextWithTrackedChanges(
        OoxmlPackage package,
        IReadOnlyList<XElement> paragraphs,
        string insertedText,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        for (int i = 0; i < paragraphs.Count; i++)
        {
            XElement paragraph = paragraphs[i];
            string current = ReadVisibleText(paragraph);
            ReplaceWholeParagraphTextWithTrackedChanges(
                package,
                paragraph,
                current,
                i == 0 ? insertedText : string.Empty,
                options,
                generatedRevisionIds,
                cancellationToken);
        }
    }

    private static XNode[] CreateTrackedBookmarkReplacementNodes(
        OoxmlPackage package,
        string deletedText,
        string insertedText,
        XElement? runProperties,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        int revisionCount = (deletedText.Length == 0 ? 0 : 1) + (insertedText.Length == 0 ? 0 : 1);
        string[] revisionIds = AllocateRevisionIds(package, revisionCount, generatedRevisionIds, cancellationToken);
        string author = GetRevisionAuthor(options);
        string timestamp = GetRevisionTimestamp(options);
        var nodes = new List<XNode>();
        int revisionIndex = 0;
        if (deletedText.Length != 0)
        {
            nodes.Add(CreateDeletedRun(deletedText, runProperties, revisionIds[revisionIndex++], author, timestamp));
        }

        if (insertedText.Length != 0)
        {
            nodes.Add(CreateInsertedRun(insertedText, runProperties, revisionIds[revisionIndex], author, timestamp));
        }

        return nodes.ToArray();
    }

    private static XElement CreateTrackedInsertedParagraph(
        OoxmlPackage package,
        string text,
        string? style,
        XElement? paragraphProperties,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var paragraph = new XElement(OoxmlNs.W + "p");
        if (paragraphProperties is not null)
        {
            paragraph.Add(new XElement(paragraphProperties));
        }

        if (style is not null)
        {
            SetParagraphStyle(paragraph, style);
        }

        string revisionId = AllocateRevisionIds(package, 1, generatedRevisionIds, cancellationToken)[0];
        string author = GetRevisionAuthor(options);
        string timestamp = GetRevisionTimestamp(options);
        paragraph.Add(CreateInsertedRun(text, runProperties: null, revisionId, author, timestamp));
        return paragraph;
    }

    private static void SetParagraphStyleWithTrackedChange(
        OoxmlPackage package,
        XElement paragraph,
        string styleId,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        XElement oldParagraphProperties = paragraph.Element(OoxmlNs.W + "pPr") is { } existing
            ? new XElement(existing)
            : new XElement(OoxmlNs.W + "pPr");
        SetParagraphStyle(paragraph, styleId);
        XElement paragraphProperties = paragraph.Element(OoxmlNs.W + "pPr")
            ?? throw new InvalidDataException("Paragraph style update did not create paragraph properties.");
        paragraphProperties.Elements(OoxmlNs.W + "pPrChange").Remove();
        string revisionId = AllocateRevisionIds(package, 1, generatedRevisionIds, cancellationToken)[0];
        paragraphProperties.Add(new XElement(
            OoxmlNs.W + "pPrChange",
            new XAttribute(OoxmlNs.W + "id", revisionId),
            new XAttribute(OoxmlNs.W + "author", GetRevisionAuthor(options)),
            new XAttribute(OoxmlNs.W + "date", GetRevisionTimestamp(options)),
            oldParagraphProperties));
    }

    private static void SetTableStyleWithTrackedChange(
        OoxmlPackage package,
        XElement table,
        string styleId,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        XElement oldTableProperties = table.Element(OoxmlNs.W + "tblPr") is { } existing
            ? new XElement(existing)
            : new XElement(OoxmlNs.W + "tblPr");
        oldTableProperties.Elements(OoxmlNs.W + "tblPrChange").Remove();
        SetTableStyle(table, styleId);
        XElement tableProperties = table.Element(OoxmlNs.W + "tblPr")
            ?? throw new InvalidDataException("Table style update did not create table properties.");
        tableProperties.Elements(OoxmlNs.W + "tblPrChange").Remove();
        string revisionId = AllocateRevisionIds(package, 1, generatedRevisionIds, cancellationToken)[0];
        tableProperties.Add(new XElement(
            OoxmlNs.W + "tblPrChange",
            new XAttribute(OoxmlNs.W + "id", revisionId),
            new XAttribute(OoxmlNs.W + "author", GetRevisionAuthor(options)),
            new XAttribute(OoxmlNs.W + "date", GetRevisionTimestamp(options)),
            oldTableProperties));
    }

    private static void SetTableRowHeaderWithTrackedChange(
        OoxmlPackage package,
        XElement row,
        bool header,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        XElement oldRowProperties = row.Element(OoxmlNs.W + "trPr") is { } existing
            ? new XElement(existing)
            : new XElement(OoxmlNs.W + "trPr");
        oldRowProperties.Elements(OoxmlNs.W + "trPrChange").Remove();
        SetTableRowHeader(row, header);
        XElement rowProperties = row.Element(OoxmlNs.W + "trPr") ?? new XElement(OoxmlNs.W + "trPr");
        if (rowProperties.Parent is null)
        {
            row.AddFirst(rowProperties);
        }

        rowProperties.Elements(OoxmlNs.W + "trPrChange").Remove();
        string revisionId = AllocateRevisionIds(package, 1, generatedRevisionIds, cancellationToken)[0];
        rowProperties.Add(new XElement(
            OoxmlNs.W + "trPrChange",
            new XAttribute(OoxmlNs.W + "id", revisionId),
            new XAttribute(OoxmlNs.W + "author", GetRevisionAuthor(options)),
            new XAttribute(OoxmlNs.W + "date", GetRevisionTimestamp(options)),
            oldRowProperties));
    }

    private static void SetCellShadingWithTrackedChange(
        OoxmlPackage package,
        XElement cell,
        string? fill,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        XElement oldCellProperties = cell.Element(OoxmlNs.W + "tcPr") is { } existing
            ? new XElement(existing)
            : new XElement(OoxmlNs.W + "tcPr");
        oldCellProperties.Elements(OoxmlNs.W + "tcPrChange").Remove();
        SetCellShading(cell, fill);
        XElement cellProperties = cell.Element(OoxmlNs.W + "tcPr") ?? new XElement(OoxmlNs.W + "tcPr");
        if (cellProperties.Parent is null)
        {
            cell.AddFirst(cellProperties);
        }

        cellProperties.Elements(OoxmlNs.W + "tcPrChange").Remove();
        string revisionId = AllocateRevisionIds(package, 1, generatedRevisionIds, cancellationToken)[0];
        cellProperties.Add(new XElement(
            OoxmlNs.W + "tcPrChange",
            new XAttribute(OoxmlNs.W + "id", revisionId),
            new XAttribute(OoxmlNs.W + "author", GetRevisionAuthor(options)),
            new XAttribute(OoxmlNs.W + "date", GetRevisionTimestamp(options)),
            oldCellProperties));
    }

    private static void MarkRowRevision(
        OoxmlPackage package,
        XElement row,
        XName revisionName,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        XElement rowProperties = row.Element(OoxmlNs.W + "trPr") ?? new XElement(OoxmlNs.W + "trPr");
        if (rowProperties.Parent is null)
        {
            row.AddFirst(rowProperties);
        }

        rowProperties.Elements(OoxmlNs.W + "ins").Remove();
        rowProperties.Elements(OoxmlNs.W + "del").Remove();
        rowProperties.Elements(OoxmlNs.W + "trPrChange").Remove();
        string revisionId = AllocateRevisionIds(package, 1, generatedRevisionIds, cancellationToken)[0];
        rowProperties.AddFirst(new XElement(
            revisionName,
            new XAttribute(OoxmlNs.W + "id", revisionId),
            new XAttribute(OoxmlNs.W + "author", GetRevisionAuthor(options)),
            new XAttribute(OoxmlNs.W + "date", GetRevisionTimestamp(options))));
    }

    private static void SetSectionPropertiesWithTrackedChange(
        OoxmlPackage package,
        XElement sectionProperties,
        Action<XElement> update,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        XElement oldSectionProperties = new(sectionProperties);
        oldSectionProperties.Elements(OoxmlNs.W + "sectPrChange").Remove();
        update(sectionProperties);
        sectionProperties.Elements(OoxmlNs.W + "sectPrChange").Remove();
        string revisionId = AllocateRevisionIds(package, 1, generatedRevisionIds, cancellationToken)[0];
        sectionProperties.Add(new XElement(
            OoxmlNs.W + "sectPrChange",
            new XAttribute(OoxmlNs.W + "id", revisionId),
            new XAttribute(OoxmlNs.W + "author", GetRevisionAuthor(options)),
            new XAttribute(OoxmlNs.W + "date", GetRevisionTimestamp(options)),
            oldSectionProperties));
    }

    private static IReadOnlyList<DocxDiagnostic> ValidateTrackChangeOptions(DocxEditOptions options)
    {
        if (!IsTrackedMode(options))
        {
            return [];
        }

        if (string.IsNullOrWhiteSpace(options.Author))
        {
            return
            [
                new DocxDiagnostic(
                    DocxSeverity.Error,
                    "E6003",
                    "TrackChangesMode.Suggest/Require requires a non-empty revision author; no document output was written.") with
                {
                    Feature = "track-changes-revision-metadata",
                    Fallback = "no-output-written"
                }
            ];
        }

        return [];
    }

    private static string GetRevisionAuthor(DocxEditOptions options)
    {
        return options.Author.Trim();
    }

    private static string GetRevisionTimestamp(DocxEditOptions options)
    {
        return options.TimestampUtc.ToUniversalTime().ToString("O", System.Globalization.CultureInfo.InvariantCulture);
    }

    private static void AddTextRun(List<XNode> nodes, string text, XElement? runProperties)
    {
        if (text.Length == 0)
        {
            return;
        }

        var run = new XElement(OoxmlNs.W + "r");
        if (runProperties is not null)
        {
            run.Add(new XElement(runProperties));
        }

        foreach (XNode node in CreateTextNodes(text))
        {
            run.Add(node);
        }

        nodes.Add(run);
    }

    private static XElement CreateDeletedRun(
        string text,
        XElement? runProperties,
        string revisionId,
        string author,
        string timestamp)
    {
        var run = new XElement(OoxmlNs.W + "r");
        if (runProperties is not null)
        {
            run.Add(new XElement(runProperties));
        }

        run.Add(CreateDeletedTextElement(text));
        return new XElement(
            OoxmlNs.W + "del",
            new XAttribute(OoxmlNs.W + "id", revisionId),
            new XAttribute(OoxmlNs.W + "author", author),
            new XAttribute(OoxmlNs.W + "date", timestamp),
            run);
    }

    private static XElement CreateInsertedRun(
        string text,
        XElement? runProperties,
        string revisionId,
        string author,
        string timestamp)
    {
        var run = new XElement(OoxmlNs.W + "r");
        if (runProperties is not null)
        {
            run.Add(new XElement(runProperties));
        }

        run.Add(CreateTextElement(text));
        return new XElement(
            OoxmlNs.W + "ins",
            new XAttribute(OoxmlNs.W + "id", revisionId),
            new XAttribute(OoxmlNs.W + "author", author),
            new XAttribute(OoxmlNs.W + "date", timestamp),
            run);
    }
}
