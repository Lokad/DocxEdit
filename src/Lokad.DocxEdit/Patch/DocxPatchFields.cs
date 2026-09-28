using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

internal static partial class DocxPatchEngine
{
    private static IReadOnlyList<DocxDiagnostic> ExecuteSetFieldFlag(
        OoxmlPackage package,
        DocxPatchOperation operation,
        string fieldName,
        string attributeName,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        _ = ReadRequiredField(operation, fieldName, diagnostics);
        bool? value = ReadBooleanField(operation, fieldName, diagnostics);
        if (target is null || value is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (IsAllFieldsTarget(target))
        {
            IReadOnlyList<FieldTarget> fieldTargets = ResolveAllFieldTargets(package, cancellationToken);
            if (fieldTargets.Count == 0)
            {
                return [Diagnostic(DocxSeverity.Error, "E1201", "Selector matched 0 fields: all.", operation, target)];
            }


            foreach (FieldTarget targetField in fieldTargets)
            {
                targetField.Element.SetAttributeValue(OoxmlNs.W + attributeName, value.Value ? "true" : "false");
            }

            foreach (IGrouping<string, FieldTarget> partGroup in fieldTargets.GroupBy(fieldTarget => fieldTarget.PartName, StringComparer.Ordinal))
            {
                SaveDocumentPart(package, partGroup.Key, partGroup.First().Document);
            }

            return [];
        }

        FieldTarget? fieldTarget = ResolveFieldTarget(package, target, cancellationToken);
        if (fieldTarget is null && !IsSupportedFieldTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported field target '{target}'. Expected 'all' or a field ID such as M.F0001 or H001.F0001.", operation, target)];
        }

        if (fieldTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 fields: {target}.", operation, target)];
        }


        fieldTarget.Element.SetAttributeValue(OoxmlNs.W + attributeName, value.Value ? "true" : "false");
        SaveDocumentPart(package, fieldTarget.PartName, fieldTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetFieldCode(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? code = ReadRequiredField(operation, "code", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-code");
        if (code is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        FieldTarget? fieldTarget = ResolveFieldTarget(package, target, cancellationToken);
        if (fieldTarget is null && !IsSupportedFieldTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported field target '{target}'. Expected a field ID such as M.F0001 or H001.F0001.", operation, target)];
        }

        if (fieldTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 fields: {target}.", operation, target)];
        }

        if (fieldTarget.Element.Name != OoxmlNs.W + "fldSimple")
        {
            return [Diagnostic(DocxSeverity.Error, "E4313", $"Field code replacement for {target} currently supports only simple w:fldSimple fields.", operation, target)];
        }

        string current = NormalizeFieldCodeForGuard((string?)fieldTarget.Element.Attribute(OoxmlNs.W + "instr") ?? string.Empty);
        if (expected is not null && !string.Equals(current, NormalizeFieldCodeForGuard(expected), StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected field code does not match current code.", operation, target, fieldName: "expect-code")];
        }


        fieldTarget.Element.SetAttributeValue(OoxmlNs.W + "instr", code);
        fieldTarget.Element.SetAttributeValue(OoxmlNs.W + "dirty", "true");
        SaveDocumentPart(package, fieldTarget.PartName, fieldTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetFieldResult(
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
        string? expected = operation.Fields.GetValueOrDefault("expect-result");
        if (text is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        FieldTarget? fieldTarget = ResolveFieldTarget(package, target, cancellationToken);
        if (fieldTarget is null && !IsSupportedFieldTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported field target '{target}'. Expected a field ID such as M.F0001 or H001.F0001.", operation, target)];
        }

        if (fieldTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 fields: {target}.", operation, target)];
        }

        bool simpleField = fieldTarget.Element.Name == OoxmlNs.W + "fldSimple";
        XElement? complexSeparateRun = null;
        XElement[] complexResultRuns = [];
        if (!simpleField)
        {
            if (!TryGetSimpleComplexFieldResultRuns(
                fieldTarget.Element,
                out complexSeparateRun,
                out complexResultRuns,
                out string? complexUnsupportedReason))
            {
                return [Diagnostic(DocxSeverity.Error, "E4313", $"Cached-result replacement for {target} is not safe: {complexUnsupportedReason}.", operation, target)];
            }
        }

        string current = simpleField
            ? ReadVisibleText(fieldTarget.Element)
            : ReadVisibleText(new XElement(OoxmlNs.W + "p", complexResultRuns));
        if (expected is not null && !string.Equals(current, expected, StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected field result does not match current result.", operation, target, fieldName: "expect-result")];
        }

        bool useTrackedChanges = IsTrackedMode(options);
        if (useTrackedChanges)
        {
            if (!simpleField)
            {
                if (!TryFallbackToDirectEdit(options, operation, target, "tracked complex-field result replacement is not modeled", diagnostics, ref useTrackedChanges))
                {
                    return diagnostics;
                }
            }
            else if (TryGetProtectedTextEditFeature(fieldTarget.Element, out string protectedFeature))
            {
                if (!TryFallbackToDirectEdit(options, operation, target, $"field result contains protected OOXML boundary '{protectedFeature}'", diagnostics, ref useTrackedChanges))
                {
                    return diagnostics;
                }
            }
            else if (!TryValidateTrackedWholeParagraphReplacement(fieldTarget.Element, current, text, out string? trackedUnsupportedReason))
            {
                if (!TryFallbackToDirectEdit(options, operation, target, trackedUnsupportedReason, diagnostics, ref useTrackedChanges))
                {
                    return diagnostics;
                }
            }
        }


        if (useTrackedChanges)
        {
            ReplaceWholeParagraphTextWithTrackedChanges(package, fieldTarget.Element, current, text, options, generatedRevisionIds, cancellationToken);
            SaveDocumentPart(package, fieldTarget.PartName, fieldTarget.Document);
            return diagnostics;
        }

        if (simpleField)
        {
            ReplaceSimpleFieldResult(fieldTarget.Element, text);
        }
        else if (complexSeparateRun is not null)
        {
            ReplaceComplexFieldResult(complexSeparateRun, complexResultRuns, text);
        }

        SaveDocumentPart(package, fieldTarget.PartName, fieldTarget.Document);
        return diagnostics;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteRefreshFieldResult(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? expectedCode = operation.Fields.GetValueOrDefault("expect-code");
        string? expectedResult = operation.Fields.GetValueOrDefault("expect-result");
        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        FieldTarget? fieldTarget = ResolveFieldTarget(package, target, cancellationToken);
        if (fieldTarget is null && !IsSupportedFieldTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported field target '{target}'. Expected a field ID such as M.F0001 or H001.F0001.", operation, target)];
        }

        if (fieldTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 fields: {target}.", operation, target)];
        }

        if (fieldTarget.Element.Name != OoxmlNs.W + "fldSimple")
        {
            return [Diagnostic(DocxSeverity.Error, "E4313", $"Field refresh for {target} currently supports only simple w:fldSimple REF-style fields.", operation, target)];
        }

        string currentCode = NormalizeFieldCodeForGuard((string?)fieldTarget.Element.Attribute(OoxmlNs.W + "instr") ?? string.Empty);
        if (expectedCode is not null && !string.Equals(currentCode, NormalizeFieldCodeForGuard(expectedCode), StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected field code does not match current code.", operation, target, fieldName: "expect-code")];
        }

        string currentResult = ReadVisibleText(fieldTarget.Element);
        if (expectedResult is not null && !string.Equals(currentResult, expectedResult, StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected field result does not match current result.", operation, target, fieldName: "expect-result")];
        }

        if (TryReadQuoteFieldText(currentCode, out string? quoteText))
        {

            ReplaceSimpleFieldResult(fieldTarget.Element, quoteText);
            fieldTarget.Element.SetAttributeValue(OoxmlNs.W + "dirty", null);
            SaveDocumentPart(package, fieldTarget.PartName, fieldTarget.Document);
            return [];
        }

        if (!TryReadRefFieldBookmarkName(currentCode, out string? bookmarkName))
        {
            return [UnsupportedFieldRefreshDiagnostic(operation, target, currentCode)];
        }

        if (!TryReadSimpleBookmarkText(fieldTarget.Document, bookmarkName, out string? bookmarkText, out string? unsupportedReason))
        {
            return [Diagnostic(DocxSeverity.Error, "E4313", $"Field refresh for {target} cannot resolve bookmark '{bookmarkName}': {unsupportedReason}.", operation, target)];
        }


        ReplaceSimpleFieldResult(fieldTarget.Element, bookmarkText);
        fieldTarget.Element.SetAttributeValue(OoxmlNs.W + "dirty", null);
        SaveDocumentPart(package, fieldTarget.PartName, fieldTarget.Document);
        return [];
    }

    private static string NormalizeFieldCodeForGuard(string code)
    {
        return string.Join(
            " ",
            code.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));
    }

    private static bool TryReadRefFieldBookmarkName(string code, [NotNullWhen(true)] out string? bookmarkName)
    {
        bookmarkName = null;
        string[] tokens = TokenizeFieldCodeForPatch(code);
        if (tokens.Length < 2)
        {
            return false;
        }

        string fieldType = NormalizeFieldTypeForPatch(tokens[0]);
        if (fieldType is not ("REF" or "PAGEREF" or "NOTEREF"))
        {
            return false;
        }

        foreach (string token in tokens.Skip(1))
        {
            if (token.StartsWith('\\'))
            {
                continue;
            }

            bookmarkName = token;
            return !string.IsNullOrWhiteSpace(bookmarkName);
        }

        return false;
    }

    private static bool TryReadQuoteFieldText(string code, [NotNullWhen(true)] out string? text)
    {
        text = null;
        string[] tokens = TokenizeFieldCodeForPatch(code);
        if (tokens.Length < 2 || NormalizeFieldTypeForPatch(tokens[0]) != "QUOTE")
        {
            return false;
        }

        string[] arguments = tokens
            .Skip(1)
            .TakeWhile(token => !token.StartsWith('\\'))
            .ToArray();
        if (arguments.Length == 0)
        {
            return false;
        }

        text = string.Join(" ", arguments);
        return true;
    }

    private static DocxDiagnostic UnsupportedFieldRefreshDiagnostic(
        DocxPatchOperation operation,
        string target,
        string code)
    {
        string[] tokens = TokenizeFieldCodeForPatch(code);
        string fieldType = tokens.Length == 0 ? "unknown" : NormalizeFieldTypeForPatch(tokens[0]);
        string reason = fieldType switch
        {
            "TOC" or "PAGE" or "NUMPAGES" or "SECTIONPAGES" => "requires Word layout or pagination state",
            "DOCPROPERTY" or "DOCVARIABLE" or "AUTHOR" or "TITLE" or "SUBJECT" or "KEYWORDS" => "requires document property state",
            "MERGEFIELD" or "MERGEREC" or "MERGESEQ" or "NEXT" or "NEXTIF" or "SKIPIF" => "requires mail merge data or mail merge state",
            "FORMULA" => "requires Word formula evaluation",
            "IF" => "requires Word conditional field evaluation",
            "DATE" or "TIME" or "CREATEDATE" or "SAVEDATE" or "PRINTDATE" => "requires Word date/time evaluation",
            "HYPERLINK" or "INCLUDETEXT" or "INCLUDEPICTURE" or "LINK" => "requires hyperlink or external target state",
            _ => "is not modeled for deterministic refresh"
        };
        return Diagnostic(
            DocxSeverity.Error,
            "E4313",
            $"Field refresh for {target} does not support field type '{fieldType}': {reason}. Supported deterministic refresh fields are REF, PAGEREF, NOTEREF, and QUOTE.",
            operation,
            target);
    }

    private static string NormalizeFieldTypeForPatch(string token)
    {
        return token.StartsWith('=') ? "FORMULA" : token.ToUpperInvariant();
    }

    private static string[] TokenizeFieldCodeForPatch(string code)
    {
        var tokens = new List<string>();
        var current = new StringBuilder();
        bool inQuote = false;
        foreach (char character in code)
        {
            if (character == '"')
            {
                inQuote = !inQuote;
                continue;
            }

            if (char.IsWhiteSpace(character) && !inQuote)
            {
                AddFieldCodeToken(tokens, current);
                continue;
            }

            current.Append(character);
        }

        AddFieldCodeToken(tokens, current);
        return tokens.ToArray();
    }

    private static void AddFieldCodeToken(List<string> tokens, StringBuilder current)
    {
        if (current.Length == 0)
        {
            return;
        }

        tokens.Add(current.ToString());
        current.Clear();
    }

    private static bool TryReadSimpleBookmarkText(
        XDocument document,
        string bookmarkName,
        [NotNullWhen(true)] out string? text,
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        text = null;
        unsupportedReason = null;
        XElement[] starts = document
            .Descendants(OoxmlNs.W + "bookmarkStart")
            .Where(bookmark => string.Equals((string?)bookmark.Attribute(OoxmlNs.W + "name"), bookmarkName, StringComparison.Ordinal))
            .ToArray();
        if (starts.Length == 0)
        {
            unsupportedReason = "bookmark was not found";
            return false;
        }

        if (starts.Length > 1)
        {
            unsupportedReason = "bookmark name is ambiguous";
            return false;
        }

        XElement start = starts[0];
        string? ooxmlId = (string?)start.Attribute(OoxmlNs.W + "id");
        XElement? end = ooxmlId is null
            ? null
            : document
                .Descendants(OoxmlNs.W + "bookmarkEnd")
                .FirstOrDefault(element => string.Equals((string?)element.Attribute(OoxmlNs.W + "id"), ooxmlId, StringComparison.Ordinal));
        if (end is null)
        {
            unsupportedReason = "bookmark end marker was not found";
            return false;
        }

        if (start.Parent is null || start.Parent != end.Parent || start.Parent.Name != OoxmlNs.W + "p")
        {
            unsupportedReason = "bookmark range is not a simple same-paragraph range";
            return false;
        }

        XNode[] nodes = start.NodesAfterSelf()
            .TakeWhile(node => node != end)
            .ToArray();
        if (ContainsProtectedBookmarkReplacementNode(nodes, out string? protectedFeature))
        {
            unsupportedReason = $"bookmark range contains protected OOXML boundary '{protectedFeature}'";
            return false;
        }

        text = ReadVisibleText(new XElement(OoxmlNs.W + "p", nodes));
        return true;
    }

    private static void ReplaceSimpleFieldResult(XElement field, string text)
    {
        XElement? firstRunProperties = field
            .Elements(OoxmlNs.W + "r")
            .Elements(OoxmlNs.W + "rPr")
            .FirstOrDefault();
        field.RemoveNodes();
        var run = new XElement(OoxmlNs.W + "r");
        if (firstRunProperties is not null)
        {
            run.Add(new XElement(firstRunProperties));
        }

        foreach (XNode node in CreateTextNodes(text))
        {
            run.Add(node);
        }

        field.Add(run);
    }

    private static bool TryGetSimpleComplexFieldResultRuns(
        XElement beginFieldChar,
        out XElement? separateRun,
        out XElement[] resultRuns,
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        separateRun = null;
        resultRuns = [];
        unsupportedReason = null;
        if (beginFieldChar.Name != OoxmlNs.W + "fldChar" ||
            !string.Equals((string?)beginFieldChar.Attribute(OoxmlNs.W + "fldCharType"), "begin", StringComparison.Ordinal))
        {
            unsupportedReason = "target is not a complex field begin marker";
            return false;
        }

        XElement? beginRun = beginFieldChar.Parent;
        XElement? paragraph = beginRun?.Parent;
        if (beginRun?.Name != OoxmlNs.W + "r" || paragraph?.Name != OoxmlNs.W + "p")
        {
            unsupportedReason = "complex field markers must be direct run children in one paragraph";
            return false;
        }

        XElement[] siblings = paragraph.Elements().ToArray();
        int beginIndex = Array.IndexOf(siblings, beginRun);
        if (beginIndex < 0)
        {
            unsupportedReason = "complex field begin marker was not found in its paragraph";
            return false;
        }

        var results = new List<XElement>();
        bool foundSeparate = false;
        for (int i = beginIndex + 1; i < siblings.Length; i++)
        {
            XElement sibling = siblings[i];
            if (sibling.Name != OoxmlNs.W + "r")
            {
                unsupportedReason = foundSeparate
                    ? $"complex field result contains non-run content '{sibling.Name.LocalName}'"
                    : $"complex field instruction contains non-run content '{sibling.Name.LocalName}'";
                return false;
            }

            XElement[] fieldChars = sibling.Descendants(OoxmlNs.W + "fldChar").ToArray();
            if (!foundSeparate)
            {
                foreach (XElement fieldChar in fieldChars)
                {
                    string? type = (string?)fieldChar.Attribute(OoxmlNs.W + "fldCharType");
                    if (string.Equals(type, "begin", StringComparison.Ordinal))
                    {
                        unsupportedReason = "complex field contains nested field topology before the result";
                        return false;
                    }

                    if (string.Equals(type, "separate", StringComparison.Ordinal))
                    {
                        separateRun = sibling;
                        foundSeparate = true;
                        break;
                    }

                    if (string.Equals(type, "end", StringComparison.Ordinal))
                    {
                        unsupportedReason = "complex field has an end marker before its separate marker";
                        return false;
                    }
                }

                continue;
            }

            foreach (XElement fieldChar in fieldChars)
            {
                string? type = (string?)fieldChar.Attribute(OoxmlNs.W + "fldCharType");
                if (string.Equals(type, "begin", StringComparison.Ordinal))
                {
                    unsupportedReason = "complex field result contains nested complex field markup";
                    return false;
                }

                if (string.Equals(type, "separate", StringComparison.Ordinal))
                {
                    unsupportedReason = "complex field has duplicate separate markers";
                    return false;
                }

                if (string.Equals(type, "end", StringComparison.Ordinal))
                {
                    resultRuns = results.ToArray();
                    return true;
                }
            }

            if (ContainsProtectedBookmarkReplacementNode([sibling], out string? protectedFeature))
            {
                unsupportedReason = $"complex field result contains protected OOXML boundary '{protectedFeature}'";
                return false;
            }

            results.Add(sibling);
        }

        unsupportedReason = foundSeparate
            ? "matching complex field end marker was not found in the same paragraph"
            : "complex field separate marker was not found in the same paragraph";
        return false;
    }

    private static void ReplaceComplexFieldResult(XElement separateRun, IReadOnlyList<XElement> resultRuns, string text)
    {
        XElement? firstRunProperties = resultRuns
            .Elements(OoxmlNs.W + "rPr")
            .FirstOrDefault();
        foreach (XElement run in resultRuns)
        {
            run.Remove();
        }

        var replacementRun = new XElement(OoxmlNs.W + "r");
        if (firstRunProperties is not null)
        {
            replacementRun.Add(new XElement(firstRunProperties));
        }

        foreach (XNode node in CreateTextNodes(text))
        {
            replacementRun.Add(node);
        }

        separateRun.AddAfterSelf(replacementRun);
    }
}
