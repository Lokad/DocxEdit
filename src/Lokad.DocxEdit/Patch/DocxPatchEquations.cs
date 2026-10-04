using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

internal static partial class DocxPatchEngine
{
    private sealed record EquationTarget(string PartName, XDocument Document, XElement Math);

    private static DocxDiagnostic? EquationRewriteProtection(OoxmlPackage package, DocxPatchOperation operation, CancellationToken token)
    {
        string? target = operation.Fields.GetValueOrDefault("target");
        if (target is null) return null;
        XElement? container = operation.OperationName switch
        {
            "set-cell" => ResolveCellTarget(package, target, token)?.Cell,
            "set-hyperlink-text" => ResolveHyperlinkTarget(package, target, token)?.Hyperlink,
            "set-field-result" or "refresh-field-result" => ResolveFieldTarget(package, target, token)?.Element,
            "set-comment-text" => ResolveCommentTarget(package, target, operation, token, out _)?.Comment,
            "set-content-control-text" or "set-content-control-checkbox" or "set-content-control-choice" or "set-content-control-date" =>
                ResolveContentControlTarget(package, operation, target, token, out _)?.ContentControl,
            _ => null
        };
        return container?.Descendants(OoxmlNs.M + "oMath").Any() == true
            ? Diagnostic(DocxSeverity.Error, "E4305", "Text replacement would remove a native equation. Use whole-equation operations or a text span outside the equation.", operation, target)
            : null;
    }

    private static EquationTarget? ResolveEquationTarget(OoxmlPackage package, string target, CancellationToken token, bool live = false)
    {
        if (IsAliasReference(target))
        {
            var found = FindMarkedStoryElement(package, SnapshotAliasName, AliasReferenceName(target), OoxmlNs.M + "oMath", token);
            return found is null ? null : new(found.Value.Story.PartName, found.Value.Document, found.Value.Element);
        }
        if (!DocxTargetId.TryParse(target, out DocxTargetId id) || id.Kind != DocxTargetKind.Equation) return null;
        StoryDocument? story = TryResolveStoryDocument(package, id, token);
        if (story is null) return null;
        XElement? math = FindSnapshotElement(story.Document, OoxmlNs.M + "oMath", target);
        if (math is null && live) math = story.Document.Descendants(OoxmlNs.M + "oMath").ElementAtOrDefault(id.Primary - 1);
        return math is null ? null : new(story.PartName, story.Document, math);
    }

    private static bool CanEditEquation(XElement math)
    {
        XElement? paragraph = math.Parent?.Name == OoxmlNs.M + "oMathPara" ? math.Parent.Parent : math.Parent;
        if (paragraph?.Name != OoxmlNs.W + "p" || !EquationXml.IsVisible(math, DocxTextView.Final)) return false;
        GetEnteringRangeCounts(math, out int fieldDepth);
        if (fieldDepth != 0) return false;
        // Whole-equation replacement must not silently remove revisions, range
        // anchors, embedded objects or controls carried inside a native equation.
        if (TryGetProtectedTextEditFeature(math, out _)) return false;
        if (math.Descendants().Any(e => e.Name.Namespace == OoxmlNs.W && e.Name.LocalName.EndsWith("Change", StringComparison.Ordinal))) return false;
        return !math.Ancestors().Any(e => e.Name == OoxmlNs.W + "ins" || e.Name == OoxmlNs.W + "del" ||
            e.Name == OoxmlNs.W + "moveFrom" || e.Name == OoxmlNs.W + "moveTo" || e.Name == OoxmlNs.W + "sdtContent");
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteEquation(
        OoxmlPackage package, DocxPatchOperation operation, DocxEditOptions options, bool apply,
        List<string> revisions, CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        if (target is null) return diagnostics;
        XElement? replacement = null;
        if (operation.OperationName != "delete-equation")
        {
            string? latex = ReadRequiredField(operation, "latex", diagnostics);
            if (latex is null) return diagnostics;
            try { replacement = LatexMath.Parse(latex); }
            catch (FormatException exception)
            {
                return [Diagnostic(DocxSeverity.Error, "E4205", exception.Message, operation, target, fieldName: "latex")];
            }
        }
        if (operation.OperationName == "insert-equation")
        {
            string placement = operation.Fields.GetValueOrDefault("placement") ?? "after";
            if (placement is not ("after" or "inline"))
                return [Diagnostic(DocxSeverity.Error, "E4205", "Equation placement must be after or inline.", operation, target, fieldName: "placement")];
            BlockTarget? anchor = ResolveBlockTarget(package, operation, target, cancellationToken, out var selectorDiagnostics);
            if (selectorDiagnostics.Count != 0) return selectorDiagnostics;
            if (anchor is null) return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
            if (placement == "inline" && anchor.Block.Name != OoxmlNs.W + "p")
                return [Diagnostic(DocxSeverity.Error, "E4305", "Inline equations require a paragraph target.", operation, target)];
            if (anchor.Block.Ancestors().Any(e => e.Name == OoxmlNs.W + "ins" || e.Name == OoxmlNs.W + "del" ||
                e.Name == OoxmlNs.W + "moveFrom" || e.Name == OoxmlNs.W + "moveTo" || e.Name == OoxmlNs.W + "sdtContent"))
                return [Diagnostic(DocxSeverity.Error, "E4305", "Equation insertion inside a revision or content control is not supported.", operation, target)];
            XElement math = replacement!;
            math.SetAttributeValue(SnapshotCreatedName, CreatedMarkValue(operation, 0));
            if (operation.Fields.TryGetValue("as", out string? alias)) math.SetAttributeValue(SnapshotAliasName, alias);
            if (placement == "inline") anchor.Block.Add(math);
            else anchor.Block.AddAfterSelf(new XElement(OoxmlNs.W + "p", new XElement(OoxmlNs.M + "oMathPara", math)));
            SaveDocumentPart(package, anchor.PartName, anchor.Document);
            return [];
        }

        EquationTarget? equation = ResolveEquationTarget(package, target, cancellationToken);
        if (equation is null) return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 equation targets: {target}.", operation, target)];
        string? expected = operation.Fields.GetValueOrDefault("expect-hash");
        if (expected is not null && !string.Equals(expected, EquationXml.Hash(equation.Math), StringComparison.Ordinal))
            return [Diagnostic(DocxSeverity.Error, "E3201", "Equation content hash does not match expect-hash.", operation, target, fieldName: "expect-hash")];
        if (!CanEditEquation(equation.Math))
            return [Diagnostic(DocxSeverity.Error, "E4305", "Equation contains protected markup or is inside an unsupported wrapper. Preserve it or edit its containing structure explicitly.", operation, target)];
        if (replacement is not null)
        {
            if (EquationXml.Hash(equation.Math) == EquationXml.Hash(replacement))
                return NoOpResult(operation, target, "The equation is unchanged.");
            equation.Math.ReplaceNodes(replacement.Nodes());
        }
        else
        {
            XElement? display = equation.Math.Parent;
            equation.Math.Remove();
            if (display?.Name == OoxmlNs.M + "oMathPara" && !display.Elements(OoxmlNs.M + "oMath").Any()) display.Remove();
        }
        SaveDocumentPart(package, equation.PartName, equation.Document);
        return [];
    }

    private static IReadOnlyList<string> CreatedEquationIds(DocxPatchOperation operation, OoxmlPackage package, CancellationToken token)
    {
        var found = FindMarkedStoryElement(package, SnapshotCreatedName, CreatedMarkValue(operation, 0), OoxmlNs.M + "oMath", token);
        if (found is null) return [];
        (char story, int part) = DocxTargetId.ParseStoryPrefix(found.Value.Story.Prefix);
        int ordinal = found.Value.Document.Descendants(OoxmlNs.M + "oMath").TakeWhile(e => !ReferenceEquals(e, found.Value.Element)).Count() + 1;
        return [new DocxTargetId(story, part, DocxTargetKind.Equation, ordinal, 0, 0).ToWireValue()];
    }

    private static ParagraphCapabilitiesOutcome GetEquationCapabilities(OoxmlPackage package, string target, TrackChangesMode mode, CancellationToken token)
    {
        EquationTarget? equation = ResolveEquationTarget(package, target, token, live: true);
        if (equation is null) return ParagraphCapabilitiesNotFound(target);
        bool supported = CanEditEquation(equation.Math) && mode != TrackChangesMode.Require;
        string reason = !CanEditEquation(equation.Math) ? "Protected equation markup or wrapper prevents whole-equation editing (E4305)." :
            mode == TrackChangesMode.Require ? "Equation revisions are not implemented; Require fails with E6001. Use Off or Suggest." :
            "Whole-equation operation; expect-hash guards native content. Suggest applies directly with W4001. Check remains authoritative.";
        return new ParagraphCapabilitiesOutcome
        {
            Capabilities = new DocxTargetCapabilities(target, "equation", target.StartsWith("M.", StringComparison.Ordinal) ? "main" : target.StartsWith("H", StringComparison.Ordinal) ? "header" : "footer",
                new[] { "replace-equation", "delete-equation" }.Select(name => new DocxOperationCapability(name, supported ? "supported" : "unsupported", reason, name, null)).ToArray())
        };
    }

    private static void ProtectEquationCapabilities(XElement container, List<DocxOperationCapability> operations)
    {
        if (!container.Descendants(OoxmlNs.M + "oMath").Any()) return;
        for (int i = 0; i < operations.Count; i++)
        {
            if (operations[i].Operation is "set-cell" or "set-hyperlink-text" or "set-field-result" or "refresh-field-result" or
                "set-content-control-text" or "set-content-control-checkbox" or "set-content-control-choice" or "set-content-control-date")
                operations[i] = operations[i] with { Support = "unsupported", Reason = "Text replacement would remove a native equation (E4305). Use whole-equation operations or a text span outside the equation.", Alternative = "replace-equation" };
        }
    }
}
