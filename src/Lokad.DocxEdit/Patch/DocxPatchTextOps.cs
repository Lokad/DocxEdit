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
        int? occurrence = ReadPositiveOccurrence(operation, diagnostics);
        if (find is null || replacement is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        bool shouldPreserveRuns = preserveRuns ?? true;
        ParagraphTarget? paragraphTarget = ResolveParagraphTarget(package, operation, target, cancellationToken, out IReadOnlyList<DocxDiagnostic> selectorDiagnostics);
        if (selectorDiagnostics.Count != 0)
        {
            return selectorDiagnostics;
        }

        if (paragraphTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (find.Length == 0)
        {
            return [Diagnostic(DocxSeverity.Error, "E4205", "Field 'find' must not be empty.", operation, target)];
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
                    target)
            ];
        }

        if (TryGetProtectedTextEditFeature(paragraphTarget.Paragraph, out string protectedFeature))
        {
            if (options.TrackChanges == TrackChangesMode.Require)
            {
                TrackUnsupportedShape(options, operation, target, $"paragraph contains protected OOXML boundary '{protectedFeature}'", diagnostics);
                return diagnostics;
            }

            return [Diagnostic(DocxSeverity.Error, "E4305", $"Text edit for {target} crosses protected OOXML boundary '{protectedFeature}'.", operation, target)];
        }

        IReadOnlyList<TextRange> matches = FindTextMatches(current, find, occurrence);
        if (matches.Count == 0)
        {
            return [Diagnostic(DocxSeverity.Error, "E4203", $"Find text was not found in {target}.", operation, target)];
        }

        bool useTrackedChanges = options.TrackChanges is TrackChangesMode.Require or TrackChangesMode.Suggest;
        bool canUseTrackedChanges = true;
        string? trackedUnsupportedReason = null;
        if (useTrackedChanges)
        {
            canUseTrackedChanges = TryValidateTrackedTextReplacement(paragraphTarget.Paragraph, current, matches, replacement, out trackedUnsupportedReason);
            if (!canUseTrackedChanges && trackedUnsupportedReason is not null &&
                !TrackUnsupportedShape(options, operation, target, trackedUnsupportedReason, diagnostics))
            {
                return diagnostics;
            }
        }

        if (!apply)
        {
            return diagnostics;
        }

        if (useTrackedChanges && canUseTrackedChanges)
        {
            ReplaceParagraphTextWithTrackedChanges(package, paragraphTarget.Paragraph, current, matches, replacement, options, generatedRevisionIds, cancellationToken);
            SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
            return [];
        }

        if (shouldPreserveRuns)
        {
            if (!TryReplaceParagraphTextPreservingRuns(paragraphTarget.Paragraph, matches, replacement, out string? unsupportedReason))
            {
                return [Diagnostic(DocxSeverity.Error, "E4306", $"Run-preserving replacement is not supported for {target}: {unsupportedReason}. Use preserve-runs false to allow paragraph-level rewriting.", operation, target)];
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
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected text does not match current text.", operation, target)];
        }

        if (TryGetProtectedTextEditFeature(paragraphTarget.Paragraph, out string protectedFeature))
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
            !TryValidateTrackedWholeParagraphReplacement(paragraphTarget.Paragraph, current, text, style, out string? trackedUnsupportedReason))
        {
            if (!TrackUnsupportedShape(options, operation, target, trackedUnsupportedReason, diagnostics))
            {
                return diagnostics;
            }

            useTrackedChanges = false;
        }

        if (!apply)
        {
            return diagnostics;
        }

        if (useTrackedChanges)
        {
            ReplaceWholeParagraphTextWithTrackedChanges(package, paragraphTarget.Paragraph, current, text, options, generatedRevisionIds, cancellationToken);
            if (style is not null)
            {
                SetParagraphStyleWithTrackedChange(package, paragraphTarget.Paragraph, style, options, generatedRevisionIds, cancellationToken);
            }

            SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
            return diagnostics;
        }

        ReplaceParagraphText(paragraphTarget.Paragraph, text);
        if (style is not null)
        {
            SetParagraphStyle(paragraphTarget.Paragraph, style);
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
        string? style = operation.Fields.GetValueOrDefault("style");
        bool copyParagraphProperties = ReadBooleanField(operation, "copy-paragraph-properties", diagnostics) ?? false;
        if (text is null || target is null || diagnostics.Count != 0)
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

        if (copyParagraphProperties && blockTarget.Block.Name != OoxmlNs.W + "p")
        {
            return [Diagnostic(DocxSeverity.Error, "E4307", $"Field 'copy-paragraph-properties' requires paragraph target '{target}'.", operation, target)];
        }

        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges && TextContainsTrackedUnsupportedCharacters(text))
        {
            if (!TrackUnsupportedShape(options, operation, target, "inserted paragraph text contains tabs or line breaks", diagnostics))
            {
                return diagnostics;
            }

            useTrackedChanges = false;
        }

        if (!apply)
        {
            return diagnostics;
        }

        XElement? paragraphProperties = copyParagraphProperties
            ? CloneParagraphPropertiesForInsertion(blockTarget.Block.Element(OoxmlNs.W + "pPr"))
            : null;
        XElement paragraph = useTrackedChanges
            ? CreateTrackedInsertedParagraph(package, text, style, paragraphProperties, options, generatedRevisionIds, cancellationToken)
            : CreateSimpleParagraph(text, style, paragraphProperties);
        if (insertAfter)
        {
            blockTarget.Block.AddAfterSelf(paragraph);
        }
        else
        {
            blockTarget.Block.AddBeforeSelf(paragraph);
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
                return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected text does not match current text.", operation, target)];
            }
        }

        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges)
        {
            if (blockTarget.Block.Name != OoxmlNs.W + "p")
            {
                if (!TrackUnsupportedShape(options, operation, target, "tracked block deletion is supported only for paragraph targets", diagnostics))
                {
                    return diagnostics;
                }

                useTrackedChanges = false;
            }
            else if (ParagraphHasSectionProperties(blockTarget.Block))
            {
                if (!TrackUnsupportedShape(options, operation, target, "paragraph contains section properties", diagnostics))
                {
                    return diagnostics;
                }

                useTrackedChanges = false;
            }
            else if (TryGetProtectedTextEditFeature(blockTarget.Block, out string protectedFeature))
            {
                if (!TrackUnsupportedShape(options, operation, target, $"paragraph contains protected OOXML boundary '{protectedFeature}'", diagnostics))
                {
                    return diagnostics;
                }

                useTrackedChanges = false;
            }
            else if (TextContainsTrackedUnsupportedCharacters(current))
            {
                if (!TrackUnsupportedShape(options, operation, target, "deleted paragraph text contains tabs or line breaks", diagnostics))
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

        if (useTrackedChanges)
        {
            ReplaceWholeParagraphTextWithTrackedChanges(package, blockTarget.Block, current, string.Empty, options, generatedRevisionIds, cancellationToken);
            SaveDocumentPart(package, blockTarget.PartName, blockTarget.Document);
            return diagnostics;
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

        if (!TryResolveStyleId(package, style, "paragraph", cancellationToken, out string? styleId, out DocxDiagnostic? styleDiagnostic, operation, target))
        {
            return [styleDiagnostic!];
        }

        if (!apply)
        {
            return [];
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
