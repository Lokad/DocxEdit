using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

internal static partial class DocxPatchEngine
{
    private static IReadOnlyList<DocxDiagnostic> ExecuteReplaceText(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? find = ReadRequiredField(operation, "find", diagnostics);
        string? replacement = ReadRequiredField(operation, "with", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        bool? preserveRuns = ReadBooleanField(operation, "preserve-runs", diagnostics);
        bool replaceAll = string.Equals(operation.Fields.GetValueOrDefault("occurrence"), "all", StringComparison.Ordinal);
        int? occurrence = replaceAll ? null : ReadPositiveOccurrence(operation, diagnostics);
        if (find is null || replacement is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        bool shouldPreserveRuns = preserveRuns ?? true;

        if (DocxTargetId.TryParse(target, out DocxTargetId parsedTargetId) &&
            parsedTargetId.Kind is DocxTargetKind.Cell or DocxTargetKind.MergeGroup)
        {
            return ExecuteReplaceTextInCell(package, operation, target, find, replacement, expected, shouldPreserveRuns, occurrence, replaceAll, options, generatedRevisionIds, cancellationToken);
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
            return
            [
                Diagnostic(
                    DocxSeverity.Error,
                    "E3201",
                    $"Guard failed for {target}. Expected text does not match current text.",
                    operation,
                    target, fieldName: "expect-text")
            ];
        }

        IReadOnlyList<TextRange> matches = FindTextMatches(current, find, replaceAll ? null : occurrence);
        if (matches.Count == 0)
        {
            return [Diagnostic(DocxSeverity.Error, "E4203", $"Find text was not found in {target}.", operation, target, fieldName: "find")];
        }

        if (!replaceAll && occurrence is null && matches.Count > 1)
        {
            return [Diagnostic(DocxSeverity.Error, "E1202", $"Find text matched {matches.Count} occurrences in {target}. Specify occurrence N to select one match or occurrence all to replace every match.", operation, target) with { MatchCount = matches.Count }];
        }

        if (string.Equals(ApplyTextReplacement(current, matches, replacement), current, StringComparison.Ordinal))
        {
            return NoOpResult(operation, target, "Replace-text for " + target + " leaves the text unchanged; nothing was written and no revisions were generated.");
        }

        // D05: run-preserving edits check the actual match spans against
        // protected boundaries. Paragraph rewrites keep the whole-paragraph gate
        // because they rebuild the container; tracked output preserves
        // surrounding markup in place around the edited spans.
        bool planSpans = shouldPreserveRuns;
        SpanEditPlan spanPlan = planSpans
            ? PlanSpanEdit(paragraphTarget.Paragraph, current, matches)
            : SpanEditPlan.Legacy;
        string? protectedFeature = spanPlan.UseLegacyGate
            ? TryGetStoryProtectedTextEditFeature(paragraphTarget.Paragraph, out string legacyFeature) ? legacyFeature : null
            : spanPlan.ProtectedFeature;
        if (protectedFeature is not null)
        {
            if (options.TrackChanges == TrackChangesMode.Require)
            {
                TrackUnsupportedShape(options, operation, target, "paragraph contains protected OOXML boundary " + Quote(protectedFeature), diagnostics);
                return diagnostics;
            }

            return [Diagnostic(DocxSeverity.Error, "E4305", "Text edit for " + target + " crosses protected OOXML boundary " + Quote(protectedFeature) + ".", operation, target)];
        }

        bool useTrackedChanges = options.TrackChanges is TrackChangesMode.Require or TrackChangesMode.Suggest;
        bool canUseTrackedChanges = true;
        string? trackedUnsupportedReason = null;
        List<VisibleCharEntry>? spanMap = spanPlan.Map;
        bool useSpanTrackedPath = useTrackedChanges && !spanPlan.UseLegacyGate && spanPlan.ProtectedFeature is null && spanMap is not null;
        if (useTrackedChanges)
        {
            if (useSpanTrackedPath)
            {
                canUseTrackedChanges = TryValidateTrackedSpanReplacement(spanMap!, current, matches, replacement, out trackedUnsupportedReason);
            }
            else
            {
                canUseTrackedChanges = TryValidateTrackedTextReplacement(paragraphTarget.Paragraph, current, matches, replacement, out trackedUnsupportedReason);
            }

            if (!canUseTrackedChanges && trackedUnsupportedReason is not null &&
                !TrackUnsupportedShape(options, operation, target, trackedUnsupportedReason, diagnostics))
            {
                return diagnostics;
            }
        }


        if (useTrackedChanges && canUseTrackedChanges)
        {
            if (useSpanTrackedPath)
            {
                if (!ReplaceParagraphTextWithTrackedSpans(package, paragraphTarget.Paragraph, current, matches, replacement, options, generatedRevisionIds, cancellationToken, out trackedUnsupportedReason))
                {
                    if (!TrackUnsupportedShape(options, operation, target, trackedUnsupportedReason ?? "tracked span replacement is not supported for this target shape", diagnostics))
                    {
                        return diagnostics;
                    }

                    canUseTrackedChanges = false;
                }
            }
            else
            {
                ReplaceParagraphTextWithTrackedChanges(package, paragraphTarget.Paragraph, current, matches, replacement, options, generatedRevisionIds, cancellationToken);
            }

            if (canUseTrackedChanges)
            {
                SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
                return [];
            }
        }

        if (shouldPreserveRuns)
        {
            string? unsupportedReason;
            bool replaced = !spanPlan.UseLegacyGate && spanPlan.Positions is not null && spanPlan.DirectMatches is not null
                ? ReplaceDirectTextRanges(paragraphTarget.Paragraph, spanPlan.Positions, spanPlan.DirectMatches, replacement, out unsupportedReason)
                : TryReplaceParagraphTextPreservingRuns(paragraphTarget.Paragraph, matches, replacement, out unsupportedReason);
            if (!replaced)
            {
                return [Diagnostic(DocxSeverity.Error, "E4306", "Run-preserving replacement is not supported for " + target + ": " + unsupportedReason + ". Use preserve-runs false to allow paragraph-level rewriting.", operation, target)];
            }
        }
        else
        {
            string edited = ApplyTextReplacement(current, matches, replacement);
            ReplaceParagraphText(paragraphTarget.Paragraph, edited);
        }

        SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
        return diagnostics;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteReplaceTextInCell(
        OoxmlPackage package,
        DocxPatchOperation operation,
        string target,
        string find,
        string replacement,
        string? expected,
        bool shouldPreserveRuns,
        int? occurrence,
        bool replaceAll,
        DocxEditOptions options,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        CellTarget? cellTarget = ResolveCellTarget(package, target, cancellationToken);
        if (cellTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (IsVerticalMergeContinuation(cellTarget.Cell))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Unsupported merged-cell target '{target}'. Target the vertical-merge root cell instead.", operation, target)];
        }



        string guard = ReadVisibleText(cellTarget.Cell);
        if (expected is not null && !string.Equals(guard, expected, StringComparison.Ordinal))
        {
            return
            [
                Diagnostic(
                    DocxSeverity.Error,
                    "E3201",
                    $"Guard failed for {target}. Expected text does not match current text.",
                    operation,
                    target, fieldName: "expect-text")
            ];
        }

        var paragraphs = new List<(XElement Paragraph, string Current, IReadOnlyList<TextRange> Matches)>();
        foreach (XElement paragraph in cellTarget.Cell.Elements(OoxmlNs.W + "p"))
        {
            string current = ReadVisibleText(paragraph);
            IReadOnlyList<TextRange> matches = FindTextMatches(current, find, null);
            if (matches.Count != 0)
            {
                paragraphs.Add((paragraph, current, matches));
            }
        }

        int total = paragraphs.Sum(static hit => hit.Matches.Count);
        if (total == 0)
        {
            return [Diagnostic(DocxSeverity.Error, "E4203", $"Find text was not found in {target}.", operation, target, fieldName: "find")];
        }

        List<(XElement Paragraph, string Current, IReadOnlyList<TextRange> Matches)> selected;
        if (replaceAll)
        {
            selected = paragraphs;
        }
        else if (occurrence is not null)
        {
            int remaining = occurrence.Value;
            (XElement Paragraph, string Current, IReadOnlyList<TextRange> Matches)? pick = null;
            foreach (var hit in paragraphs)
            {
                if (remaining <= hit.Matches.Count)
                {
                    pick = (hit.Paragraph, hit.Current, new[] { hit.Matches[remaining - 1] });
                    break;
                }

                remaining -= hit.Matches.Count;
            }

            if (pick is null)
            {
                return [Diagnostic(DocxSeverity.Error, "E4203", $"Find text was not found in {target}.", operation, target, fieldName: "find")];
            }

            selected = [pick.Value];
        }
        else if (total > 1)
        {
            return [Diagnostic(DocxSeverity.Error, "E1202", $"Find text matched {total} occurrences in {target}. Specify occurrence N to select one match or occurrence all to replace every match.", operation, target) with { MatchCount = total }];
        }
        else
        {
            selected = paragraphs;
        }

        bool changed = false;
        foreach (var hit in selected)
        {
            if (!string.Equals(ApplyTextReplacement(hit.Current, hit.Matches, replacement), hit.Current, StringComparison.Ordinal))
            {
                changed = true;
                break;
            }
        }

        if (!changed)
        {
            return NoOpResult(operation, target, "Replace-text for " + target + " leaves the text unchanged; nothing was written and no revisions were generated.");
        }

        var diagnostics = new List<DocxDiagnostic>();
        bool useTrackedChanges = options.TrackChanges is TrackChangesMode.Require or TrackChangesMode.Suggest;
        bool canUseTrackedChanges = true;
        var plans = new List<(XElement Paragraph, string Current, IReadOnlyList<TextRange> Matches, SpanEditPlan Plan, List<VisibleCharEntry>? Map, bool SpanPath)>();
        foreach (var hit in selected)
        {
            SpanEditPlan spanPlan = shouldPreserveRuns
                ? PlanSpanEdit(hit.Paragraph, hit.Current, hit.Matches)
                : SpanEditPlan.Legacy;
            string? protectedFeature = spanPlan.UseLegacyGate
                ? TryGetStoryProtectedTextEditFeature(hit.Paragraph, out string legacyFeature) ? legacyFeature : null
                : spanPlan.ProtectedFeature;
            if (protectedFeature is not null)
            {
                if (options.TrackChanges == TrackChangesMode.Require)
                {
                    TrackUnsupportedShape(options, operation, target, "paragraph contains protected OOXML boundary " + Quote(protectedFeature), diagnostics);
                    return diagnostics;
                }

                return [Diagnostic(DocxSeverity.Error, "E4305", "Text edit for " + target + " crosses protected OOXML boundary " + Quote(protectedFeature) + ".", operation, target)];
            }

            List<VisibleCharEntry>? spanMap = spanPlan.Map;
            bool spanPath = useTrackedChanges && !spanPlan.UseLegacyGate && spanPlan.ProtectedFeature is null && spanMap is not null;
            plans.Add((hit.Paragraph, hit.Current, hit.Matches, spanPlan, spanMap, spanPath));
        }

        if (useTrackedChanges)
        {
            foreach (var plan in plans)
            {
                bool validated;
                string? trackedUnsupportedReason;
                if (plan.SpanPath)
                {
                    validated = TryValidateTrackedSpanReplacement(plan.Map!, plan.Current, plan.Matches, replacement, out trackedUnsupportedReason);
                }
                else
                {
                    validated = TryValidateTrackedTextReplacement(plan.Paragraph, plan.Current, plan.Matches, replacement, out trackedUnsupportedReason);
                }

                if (!validated)
                {
                    canUseTrackedChanges = false;
                    if (trackedUnsupportedReason is not null &&
                        !TrackUnsupportedShape(options, operation, target, trackedUnsupportedReason, diagnostics))
                    {
                        return diagnostics;
                    }
                }
            }
        }

        if (useTrackedChanges && canUseTrackedChanges)
        {
            foreach (var plan in plans)
            {
                if (plan.SpanPath)
                {
                    if (!ReplaceParagraphTextWithTrackedSpans(package, plan.Paragraph, plan.Current, plan.Matches, replacement, options, generatedRevisionIds, cancellationToken, out string? applyReason))
                    {
                        if (!TrackUnsupportedShape(options, operation, target, applyReason ?? "tracked span replacement is not supported for this target shape", diagnostics))
                        {
                            return diagnostics;
                        }

                        canUseTrackedChanges = false;
                        break;
                    }
                }
                else
                {
                    ReplaceParagraphTextWithTrackedChanges(package, plan.Paragraph, plan.Current, plan.Matches, replacement, options, generatedRevisionIds, cancellationToken);
                }
            }

            if (canUseTrackedChanges)
            {
                SaveDocumentPart(package, cellTarget.PartName, cellTarget.Document);
                return [];
            }
        }

        if (shouldPreserveRuns)
        {
            foreach (var plan in plans)
            {
                string? unsupportedReason;
                bool replaced = !plan.Plan.UseLegacyGate && plan.Plan.Positions is not null && plan.Plan.DirectMatches is not null
                    ? ReplaceDirectTextRanges(plan.Paragraph, plan.Plan.Positions, plan.Plan.DirectMatches, replacement, out unsupportedReason)
                    : TryReplaceParagraphTextPreservingRuns(plan.Paragraph, plan.Matches, replacement, out unsupportedReason);
                if (!replaced)
                {
                    return [Diagnostic(DocxSeverity.Error, "E4306", "Run-preserving replacement is not supported for " + target + ": " + unsupportedReason + ". Use preserve-runs false to allow paragraph-level rewriting.", operation, target)];
                }
            }
        }
        else
        {
            foreach (var plan in plans)
            {
                string edited = ApplyTextReplacement(plan.Current, plan.Matches, replacement);
                ReplaceParagraphText(plan.Paragraph, edited);
            }
        }

        SaveDocumentPart(package, cellTarget.PartName, cellTarget.Document);
        return diagnostics;
    }
    private static IReadOnlyList<DocxDiagnostic> ExecuteReplaceParagraph(
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
        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        string? style = operation.Fields.GetValueOrDefault("style");
        if (text is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
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
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected text does not match current text.", operation, target, fieldName: "expect-text")];
        }

        string? styleId = null;
        if (style is not null)
        {
            if (!TryResolveStyleId(package, style, "paragraph", cancellationToken, out styleId, out DocxDiagnostic? styleDiagnostic, operation, target))
            {
                return [styleDiagnostic];
            }
        }

        string? currentStyle = (string?)paragraphTarget.Paragraph
            .Element(OoxmlNs.W + "pPr")
            ?.Element(OoxmlNs.W + "pStyle")
            ?.Attribute(OoxmlNs.W + "val");
        if (string.Equals(text, current, StringComparison.Ordinal) &&
            (styleId is null || string.Equals(styleId, currentStyle, StringComparison.Ordinal)))
        {
            return NoOpResult(operation, target, "Replace-paragraph for " + target + " leaves the paragraph unchanged; nothing was written and no revisions were generated.");
        }

        if (TryGetStoryProtectedTextEditFeature(paragraphTarget.Paragraph, out string protectedFeature))
        {
            if (options.TrackChanges == TrackChangesMode.Require)
            {
                TrackUnsupportedShape(options, operation, target, $"paragraph contains protected OOXML boundary '{protectedFeature}'", diagnostics);
                return diagnostics;
            }

            return [Diagnostic(DocxSeverity.Error, "E4305", $"Paragraph replacement for {target} would remove protected OOXML boundary '{protectedFeature}'.", operation, target)];
        }

        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges &&
            !TryValidateTrackedWholeParagraphReplacement(paragraphTarget.Paragraph, current, text, out string? trackedUnsupportedReason))
        {
            if (!TryFallbackToDirectEdit(options, operation, target, trackedUnsupportedReason, diagnostics, ref useTrackedChanges))
            {
                return diagnostics;
            }
        }


        if (useTrackedChanges)
        {
            ReplaceWholeParagraphTextWithTrackedChanges(package, paragraphTarget.Paragraph, current, text, options, generatedRevisionIds, cancellationToken);
            if (styleId is not null)
            {
                SetParagraphStyleWithTrackedChange(package, paragraphTarget.Paragraph, styleId, options, generatedRevisionIds, cancellationToken);
            }

            SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
            return diagnostics;
        }

        ReplaceParagraphText(paragraphTarget.Paragraph, text);
        if (styleId is not null)
        {
            SetParagraphStyle(paragraphTarget.Paragraph, styleId);
        }

        SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
        return diagnostics;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteInsertBlock(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool insertAfter,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        bool copyParagraphProperties = ReadBooleanField(operation, "copy-paragraph-properties", diagnostics) ?? false;
        if (text is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        string[] texts = operation.FieldValues.Where(static field => field.Name == "text").Select(static field => field.Value).ToArray();
        string[] styles = operation.FieldValues.Where(static field => field.Name == "style").Select(static field => field.Value).ToArray();
        if (styles.Length > 1 && styles.Length != texts.Length)
        {
            return [Diagnostic(DocxSeverity.Error, "E4205", $"Operation '{operation.OperationName}' accepts at most one 'style' field or one per 'text' field.", operation, target)];
        }

        BlockTarget? blockTarget = ResolveBlockTarget(package, operation, target, cancellationToken, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
        if (selectorDiagnostics.Count != 0)
        {
            return selectorDiagnostics;
        }

        if (blockTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (copyParagraphProperties && blockTarget.Block.Name != OoxmlNs.W + "p")
        {
            return [Diagnostic(DocxSeverity.Error, "E4307", $"Field 'copy-paragraph-properties' requires paragraph target '{target}'.", operation, target)];
        }

        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges && texts.Any(TextContainsTrackedUnsupportedCharacters))
        {
            if (!TryFallbackToDirectEdit(options, operation, target, "inserted paragraph text contains tabs or line breaks", diagnostics, ref useTrackedChanges))
            {
                return diagnostics;
            }
        }


        string?[] styleIds = new string?[styles.Length];
        for (int index = 0; index < styles.Length; index++)
        {
            if (!TryResolveStyleId(package, styles[index], "paragraph", cancellationToken, out styleIds[index], out DocxDiagnostic? insertStyleDiagnostic, operation, target))
            {
                return [insertStyleDiagnostic];
            }
        }

        XElement? baseParagraphProperties = copyParagraphProperties
            ? CloneParagraphPropertiesForInsertion(blockTarget.Block.Element(OoxmlNs.W + "pPr"))
            : null;
        XElement sibling = blockTarget.Block;
        for (int index = 0; index < texts.Length; index++)
        {
            XElement? paragraphProperties = baseParagraphProperties is null ? null : new XElement(baseParagraphProperties);
            string? paragraphStyleId = styles.Length <= 1 ? styleIds.FirstOrDefault() : styleIds[index];
            XElement paragraph = useTrackedChanges
                ? CreateTrackedInsertedParagraph(package, texts[index], paragraphStyleId, paragraphProperties, options, generatedRevisionIds, cancellationToken)
                : CreateSimpleParagraph(texts[index], paragraphStyleId, paragraphProperties);
            if (insertAfter)
            {
                sibling.AddAfterSelf(paragraph);
                sibling = paragraph;
            }
            else
            {
                blockTarget.Block.AddBeforeSelf(paragraph);
            }
        }

        SaveDocumentPart(package, blockTarget.PartName, blockTarget.Document);
        return [];
    }

    private static XElement? CloneParagraphPropertiesForInsertion(XElement? paragraphProperties)
    {
        if (paragraphProperties is null)
        {
            return null;
        }

        var clone = new XElement(paragraphProperties);
        clone.Elements(OoxmlNs.W + "pPrChange").Remove();
        clone.Elements(OoxmlNs.W + "sectPr").Remove();
        return clone.HasElements || clone.HasAttributes ? clone : null;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteDeleteBlock(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        BlockTarget? blockTarget = ResolveBlockTarget(package, operation, target, cancellationToken, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
        if (selectorDiagnostics.Count != 0)
        {
            return selectorDiagnostics;
        }

        if (blockTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        string current = ReadVisibleText(blockTarget.Block);
        if (expected is not null)
        {
            if (!string.Equals(current, expected, StringComparison.Ordinal))
            {
                return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected text does not match current text.", operation, target, fieldName: "expect-text")];
            }
        }

        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges)
        {
            if (blockTarget.Block.Name != OoxmlNs.W + "p")
            {
                if (!TryFallbackToDirectEdit(options, operation, target, "tracked block deletion is supported only for paragraph targets", diagnostics, ref useTrackedChanges))
                {
                    return diagnostics;
                }
            }
            else if (ParagraphHasSectionProperties(blockTarget.Block))
            {
                if (!TryFallbackToDirectEdit(options, operation, target, "paragraph contains section properties", diagnostics, ref useTrackedChanges))
                {
                    return diagnostics;
                }
            }
            else if (TryGetStoryProtectedTextEditFeature(blockTarget.Block, out string protectedFeature))
            {
                if (!TryFallbackToDirectEdit(options, operation, target, $"paragraph contains protected OOXML boundary '{protectedFeature}'", diagnostics, ref useTrackedChanges))
                {
                    return diagnostics;
                }
            }
            else if (TextContainsTrackedUnsupportedCharacters(current))
            {
                if (!TryFallbackToDirectEdit(options, operation, target, "deleted paragraph text contains tabs or line breaks", diagnostics, ref useTrackedChanges))
                {
                    return diagnostics;
                }
            }
        }


        if (useTrackedChanges)
        {
            ReplaceWholeParagraphTextWithTrackedChanges(package, blockTarget.Block, current, string.Empty, options, generatedRevisionIds, cancellationToken);
            SaveDocumentPart(package, blockTarget.PartName, blockTarget.Document);
            return diagnostics;
        }

        if (TryGetStoryEnteringProtectedFeature(blockTarget.Block, out string enteringFeature))
        {
            return [Diagnostic(DocxSeverity.Error, "E4305", "Delete for " + target + " crosses protected OOXML boundary " + Quote(enteringFeature) + ".", operation, target)];
        }
        if (FindOrphanedRangeBoundary(blockTarget.Document, blockTarget.Block) is { } orphaned)
        {
            return [Diagnostic(DocxSeverity.Error, "E4305", $"Delete for {target} would orphan {orphaned} outside the deleted element. Delete the range first or choose another target.", operation, target)];
        }

        blockTarget.Block.Remove();
        SaveDocumentPart(package, blockTarget.PartName, blockTarget.Document);
        return diagnostics;
    }

    private static bool ParagraphHasSectionProperties(XElement paragraph)
    {
        return paragraph
            .Element(OoxmlNs.W + "pPr")
            ?.Element(OoxmlNs.W + "sectPr") is not null;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetStyle(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? style = ReadRequiredField(operation, "style", diagnostics);
        string? expectedStyle = operation.Fields.GetValueOrDefault("expect-style");
        if (style is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
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

        string? currentStyle = (string?)paragraphTarget.Paragraph
            .Element(OoxmlNs.W + "pPr")
            ?.Element(OoxmlNs.W + "pStyle")
            ?.Attribute(OoxmlNs.W + "val");
        if (expectedStyle is not null && !string.Equals(currentStyle, expectedStyle, StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected paragraph style " + "\u0027" + expectedStyle + "\u0027" + ", found " + "\u0027" + (currentStyle ?? "none") + "\u0027" + ".", operation, target, fieldName: "expect-style")];
        }

        if (!TryResolveStyleId(package, style, "paragraph", cancellationToken, out string? styleId, out DocxDiagnostic? styleDiagnostic, operation, target))
        {
            return [styleDiagnostic];
        }

        if (string.Equals(styleId, currentStyle, StringComparison.Ordinal))
        {
            return NoOpResult(operation, target, "Set-style for " + target + " leaves the style unchanged; nothing was written and no revisions were generated.");
        }

        if (IsTrackedMode(options))
        {
            SetParagraphStyleWithTrackedChange(package, paragraphTarget.Paragraph, styleId, options, generatedRevisionIds, cancellationToken);
        }
        else
        {
            SetParagraphStyle(paragraphTarget.Paragraph, styleId);
        }

        SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
        return [];
    }
}

