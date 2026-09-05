using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

internal static partial class DocxPatchEngine
{
    private static IReadOnlyList<DocxDiagnostic> ExecuteSetHyperlinkTarget(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        if (!TryReadHyperlinkDestination(operation, diagnostics, out string? uri, out string? anchor))
        {
            return diagnostics;
        }

        bool? history = ReadBooleanField(operation, "history", diagnostics);

        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        HyperlinkTarget? hyperlinkTarget = ResolveHyperlinkTarget(package, target, cancellationToken);
        if (hyperlinkTarget is null && !IsSupportedHyperlinkTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported hyperlink target '{target}'. Expected a hyperlink ID such as M.L0001 or H001.L0001.", operation, target)];
        }

        if (hyperlinkTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        if (uri is not null)
        {
            SetExternalHyperlinkTarget(package, hyperlinkTarget, uri, cancellationToken);
        }
        else if (anchor is not null)
        {
            SetInternalHyperlinkAnchor(package, hyperlinkTarget, anchor, cancellationToken);
        }

        if (operation.Fields.TryGetValue("tooltip", out string? tooltip))
        {
            hyperlinkTarget.Hyperlink.SetAttributeValue(OoxmlNs.W + "tooltip", tooltip);
        }

        if (operation.Fields.TryGetValue("target-frame", out string? targetFrame))
        {
            hyperlinkTarget.Hyperlink.SetAttributeValue(OoxmlNs.W + "tgtFrame", targetFrame);
        }

        if (history is not null)
        {
            hyperlinkTarget.Hyperlink.SetAttributeValue(OoxmlNs.W + "history", history.Value ? "true" : "false");
        }

        SaveDocumentPart(package, hyperlinkTarget.PartName, hyperlinkTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetHyperlinkText(
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

        HyperlinkTarget? hyperlinkTarget = ResolveHyperlinkTarget(package, target, cancellationToken);
        if (hyperlinkTarget is null && !IsSupportedHyperlinkTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported hyperlink target '{target}'. Expected a hyperlink ID such as M.L0001 or H001.L0001.", operation, target)];
        }

        if (hyperlinkTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        string current = ReadVisibleText(hyperlinkTarget.Hyperlink);
        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges)
        {
            if (TryGetProtectedTextEditFeature(hyperlinkTarget.Hyperlink, out string protectedFeature))
            {
                if (!TryFallbackToDirectEdit(options, operation, target, $"hyperlink contains protected OOXML boundary '{protectedFeature}'", diagnostics, ref useTrackedChanges))
                {
                    return diagnostics;
                }
            }
            else if (!TryValidateTrackedWholeParagraphReplacement(hyperlinkTarget.Hyperlink, current, text, style: null, out string? trackedUnsupportedReason))
            {
                if (!TryFallbackToDirectEdit(options, operation, target, trackedUnsupportedReason, diagnostics, ref useTrackedChanges))
                {
                    return diagnostics;
                }
            }
        }

        if (!apply)
        {
            return diagnostics;
        }

        if (useTrackedChanges)
        {
            ReplaceWholeParagraphTextWithTrackedChanges(package, hyperlinkTarget.Hyperlink, current, text, options, generatedRevisionIds, cancellationToken);
            SaveDocumentPart(package, hyperlinkTarget.PartName, hyperlinkTarget.Document);
            return diagnostics;
        }

        ReplaceHyperlinkText(hyperlinkTarget.Hyperlink, text);
        SaveDocumentPart(package, hyperlinkTarget.PartName, hyperlinkTarget.Document);
        return diagnostics;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteInsertHyperlinkAfter(
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
        if (!TryReadHyperlinkDestination(operation, diagnostics, out string? uri, out string? anchor))
        {
            return diagnostics;
        }

        bool? history = ReadBooleanField(operation, "history", diagnostics);

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

        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges && TextContainsTrackedUnsupportedCharacters(text))
        {
            if (!TryFallbackToDirectEdit(options, operation, target, "inserted hyperlink text contains tabs or line breaks", diagnostics, ref useTrackedChanges))
            {
                return diagnostics;
            }
        }

        if (!apply)
        {
            return diagnostics;
        }

        string? relationshipId = null;
        if (uri is not null)
        {
            relationshipId = OoxmlIds.AllocateRelationshipId(package.GetRelationships(blockTarget.PartName, cancellationToken).Select(relationship => relationship.Id));
            package.AddRelationship(blockTarget.PartName, relationshipId, OoxmlRelTypes.Hyperlink, uri, "External", cancellationToken);
        }

        XElement paragraph = CreateHyperlinkParagraph(
            useTrackedChanges ? string.Empty : text,
            relationshipId,
            anchor,
            operation.Fields.GetValueOrDefault("tooltip"),
            operation.Fields.GetValueOrDefault("target-frame"),
            history);
        if (useTrackedChanges)
        {
            XElement hyperlink = paragraph.Element(OoxmlNs.W + "hyperlink")
                ?? throw new InvalidDataException("Hyperlink paragraph did not contain a hyperlink element.");
            hyperlink.RemoveNodes();
            ReplaceWholeParagraphTextWithTrackedChanges(package, hyperlink, string.Empty, text, options, generatedRevisionIds, cancellationToken);
        }

        blockTarget.Block.AddAfterSelf(paragraph);
        SaveDocumentPart(package, blockTarget.PartName, blockTarget.Document);
        return diagnostics;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteRemoveHyperlink(
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

        HyperlinkTarget? hyperlinkTarget = ResolveHyperlinkTarget(package, target, cancellationToken);
        if (hyperlinkTarget is null && !IsSupportedHyperlinkTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported hyperlink target '{target}'. Expected a hyperlink ID such as M.L0001 or H001.L0001.", operation, target)];
        }

        if (hyperlinkTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        string? relationshipId = (string?)hyperlinkTarget.Hyperlink.Attribute(OoxmlNs.R + "id");
        hyperlinkTarget.Hyperlink.ReplaceWith(hyperlinkTarget.Hyperlink.Nodes().ToArray());
        if (!string.IsNullOrWhiteSpace(relationshipId) &&
            !UsesHyperlinkRelationship(hyperlinkTarget.Document, relationshipId))
        {
            package.RemoveRelationship(hyperlinkTarget.PartName, relationshipId, cancellationToken);
        }

        SaveDocumentPart(package, hyperlinkTarget.PartName, hyperlinkTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteUnsupportedCommentThreadOperation(DocxPatchOperation operation)
    {
        string? target = operation.Fields.GetValueOrDefault("target");
        return
        [
            Diagnostic(
                DocxSeverity.Error,
                "E4314",
                $"Operation '{operation.OperationName}' is not supported because threaded comment reply edits require comment-thread metadata that DocxEdit does not safely modify yet.",
                operation,
                target)
        ];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteUnsupportedRepeatingSectionOperation(DocxPatchOperation operation)
    {
        string? target = operation.Fields.GetValueOrDefault("target");
        return
        [
            Diagnostic(
                DocxSeverity.Error,
                "E4315",
                $"Operation '{operation.OperationName}' is not supported because repeating-section item edits require cloning or deleting structured document tag subtrees while preserving IDs, bindings, and section boundaries.",
                operation,
                target)
        ];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteUnsupportedColumnOperation(DocxPatchOperation operation)
    {
        string? target = operation.Fields.GetValueOrDefault("target");
        return
        [
            Diagnostic(
                DocxSeverity.Error,
                "E4316",
                $"Operation '{operation.OperationName}' is not supported because column edits require rebuilding table grids, horizontal spans, omitted cells, and vertical merge state.",
                operation,
                target)
        ];
    }

    private static bool TryReadHyperlinkDestination(
        DocxPatchOperation operation,
        List<DocxDiagnostic> diagnostics,
        out string? uri,
        out string? anchor)
    {
        uri = operation.Fields.GetValueOrDefault("uri");
        anchor = operation.Fields.GetValueOrDefault("anchor");
        string? target = operation.Fields.GetValueOrDefault("target");
        if (string.IsNullOrWhiteSpace(uri) == string.IsNullOrWhiteSpace(anchor))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Exactly one of 'uri' or 'anchor' is required for hyperlink destination operations.", operation, target));
            return false;
        }

        if (uri is not null &&
            !TryValidateExternalHyperlinkUri(uri, out string? uriDiagnostic))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", uriDiagnostic, operation, target));
            return false;
        }

        if (anchor is not null && (anchor.Length == 0 || anchor.Any(char.IsWhiteSpace)))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Field 'anchor' must be a non-empty bookmark anchor without whitespace.", operation, target));
            return false;
        }

        return true;
    }

    private static bool TryValidateExternalHyperlinkUri(string uri, [NotNullWhen(false)] out string? diagnostic)
    {
        diagnostic = null;
        if (!Uri.TryCreate(uri, UriKind.RelativeOrAbsolute, out Uri? parsed))
        {
            diagnostic = $"Field 'uri' must be a well-formed absolute http, https, or mailto URI; received malformed URI '{uri}'.";
            return false;
        }

        if (!parsed.IsAbsoluteUri)
        {
            diagnostic = $"Field 'uri' must be an absolute http, https, or mailto URI; relative hyperlink targets are not supported: {uri}.";
            return false;
        }

        if (parsed.Scheme is not ("http" or "https" or "mailto"))
        {
            diagnostic = $"Unsupported hyperlink URI scheme '{parsed.Scheme}'. Allowed schemes are http, https, and mailto.";
            return false;
        }

        return true;
    }

    private static void SetExternalHyperlinkTarget(
        OoxmlPackage package,
        HyperlinkTarget hyperlinkTarget,
        string uri,
        CancellationToken cancellationToken)
    {
        string? oldRelationshipId = (string?)hyperlinkTarget.Hyperlink.Attribute(OoxmlNs.R + "id");
        bool canReuseRelationship = !string.IsNullOrWhiteSpace(oldRelationshipId) &&
            CountHyperlinkRelationshipUses(hyperlinkTarget.Document, oldRelationshipId) == 1;
        string relationshipId = canReuseRelationship && oldRelationshipId is not null
            ? oldRelationshipId
            : OoxmlIds.AllocateRelationshipId(package.GetRelationships(hyperlinkTarget.PartName, cancellationToken).Select(relationship => relationship.Id));
        if (canReuseRelationship)
        {
            package.RemoveRelationship(hyperlinkTarget.PartName, relationshipId, cancellationToken);
        }

        package.AddRelationship(hyperlinkTarget.PartName, relationshipId, OoxmlRelTypes.Hyperlink, uri, "External", cancellationToken);
        hyperlinkTarget.Hyperlink.SetAttributeValue(OoxmlNs.R + "id", relationshipId);
        hyperlinkTarget.Hyperlink.SetAttributeValue(OoxmlNs.W + "anchor", null);
    }

    private static void SetInternalHyperlinkAnchor(
        OoxmlPackage package,
        HyperlinkTarget hyperlinkTarget,
        string anchor,
        CancellationToken cancellationToken)
    {
        string? oldRelationshipId = (string?)hyperlinkTarget.Hyperlink.Attribute(OoxmlNs.R + "id");
        hyperlinkTarget.Hyperlink.SetAttributeValue(OoxmlNs.R + "id", null);
        hyperlinkTarget.Hyperlink.SetAttributeValue(OoxmlNs.W + "anchor", anchor);
        if (!string.IsNullOrWhiteSpace(oldRelationshipId) &&
            !UsesHyperlinkRelationship(hyperlinkTarget.Document, oldRelationshipId))
        {
            package.RemoveRelationship(hyperlinkTarget.PartName, oldRelationshipId, cancellationToken);
        }
    }

    private static void ReplaceHyperlinkText(XElement hyperlink, string text)
    {
        var run = new XElement(OoxmlNs.W + "r");
        foreach (XNode node in CreateTextNodes(text))
        {
            run.Add(node);
        }

        hyperlink.RemoveNodes();
        hyperlink.Add(run);
    }

    private static XElement CreateHyperlinkParagraph(
        string text,
        string? relationshipId,
        string? anchor,
        string? tooltip,
        string? targetFrame,
        bool? history)
    {
        var hyperlink = new XElement(OoxmlNs.W + "hyperlink");
        if (relationshipId is not null)
        {
            hyperlink.SetAttributeValue(OoxmlNs.R + "id", relationshipId);
        }

        if (anchor is not null)
        {
            hyperlink.SetAttributeValue(OoxmlNs.W + "anchor", anchor);
        }

        if (!string.IsNullOrWhiteSpace(tooltip))
        {
            hyperlink.SetAttributeValue(OoxmlNs.W + "tooltip", tooltip);
        }

        if (!string.IsNullOrWhiteSpace(targetFrame))
        {
            hyperlink.SetAttributeValue(OoxmlNs.W + "tgtFrame", targetFrame);
        }

        if (history is not null)
        {
            hyperlink.SetAttributeValue(OoxmlNs.W + "history", history.Value ? "true" : "false");
        }

        ReplaceHyperlinkText(hyperlink, text);
        return new XElement(OoxmlNs.W + "p", hyperlink);
    }

    private static bool UsesHyperlinkRelationship(XDocument document, string relationshipId)
    {
        return CountHyperlinkRelationshipUses(document, relationshipId) > 0;
    }

    private static int CountHyperlinkRelationshipUses(XDocument document, string relationshipId)
    {
        return document
            .Descendants(OoxmlNs.W + "hyperlink")
            .Count(hyperlink => string.Equals((string?)hyperlink.Attribute(OoxmlNs.R + "id"), relationshipId, StringComparison.Ordinal));
    }
}
