using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

internal static partial class DocxPatchEngine
{
    private static IReadOnlyList<DocxDiagnostic> ExecuteSetContentControlText(
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
        if (text is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ContentControlTarget? controlTarget = ResolveContentControlTarget(package, target, cancellationToken);
        if (controlTarget is null && !IsSupportedContentControlTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported content-control target '{target}'. Expected a content control ID such as M.CC0001 or H001.CC0001.", operation, target)];
        }

        if (controlTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        bool isPlainText = IsPlainTextContentControl(controlTarget.ContentControl);
        bool isRichText = IsRichTextContentControl(controlTarget.ContentControl);
        if (!isPlainText && !isRichText)
        {
            return [UnsupportedContentControlKindDiagnostic(controlTarget.ContentControl, operation, target, "a plain-text or rich-text content control")];
        }

        DocxDiagnostic? lockDiagnostic = ValidateContentControlUnlocked(controlTarget.ContentControl, operation, target);
        if (lockDiagnostic is not null)
        {
            return [lockDiagnostic];
        }

        XElement? content = controlTarget.ContentControl.Element(OoxmlNs.W + "sdtContent");
        if (content is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E4310", $"Content control '{target}' has no editable content container.", operation, target)];
        }

        string current = ReadVisibleText(content);
        if (expected is not null && !string.Equals(current, expected, StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected content-control text does not match current text.", operation, target)];
        }

        if (!isPlainText)
        {
            if (expected is null)
            {
                return [Diagnostic(DocxSeverity.Error, "E4205", $"Rich-text content control '{target}' requires expect-text before replacement.", operation, target)];
            }

            if (content.Elements().Any(element => element.Name != OoxmlNs.W + "p"))
            {
                return [Diagnostic(DocxSeverity.Error, "E4310", $"Rich-text content control '{target}' contains non-paragraph content.", operation, target)];
            }

            if (TryGetProtectedTextEditFeature(content, out string protectedFeature))
            {
                return [Diagnostic(DocxSeverity.Error, "E4310", $"Rich-text content control '{target}' contains protected OOXML boundary '{protectedFeature}'.", operation, target)];
            }
        }

        bool useTrackedChanges = IsTrackedMode(options);
        XElement? trackedContainer = null;
        XElement[]? trackedParagraphs = null;
        string trackedCurrent = current;
        if (useTrackedChanges)
        {
            if (!isPlainText)
            {
                if (!TryGetTrackedRichTextContentControlParagraphs(content, text, out trackedParagraphs, out string? trackedUnsupportedReason))
                {
                    if (!TrackUnsupportedShape(options, operation, target, trackedUnsupportedReason, diagnostics))
                    {
                        return diagnostics;
                    }

                    useTrackedChanges = false;
                }
            }
            else if (!TryGetTrackedContentControlTextContainer(content, text, out trackedContainer, out trackedCurrent, out string? trackedUnsupportedReason))
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

        if (useTrackedChanges)
        {
            if (trackedParagraphs is not null)
            {
                ReplaceCellParagraphTextWithTrackedChanges(package, trackedParagraphs, text, options, generatedRevisionIds, cancellationToken);
            }
            else if (trackedContainer is not null)
            {
                ReplaceWholeParagraphTextWithTrackedChanges(package, trackedContainer, trackedCurrent, text, options, generatedRevisionIds, cancellationToken);
            }

            SaveDocumentPart(package, controlTarget.PartName, controlTarget.Document);
            return diagnostics;
        }

        ReplaceContentControlText(content, text);
        SaveDocumentPart(package, controlTarget.PartName, controlTarget.Document);
        return diagnostics;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetContentControlCheckbox(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        _ = ReadRequiredField(operation, "checked", diagnostics);
        bool? checkedValue = ReadBooleanField(operation, "checked", diagnostics);
        if (target is null || checkedValue is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ContentControlTarget? controlTarget = ResolveContentControlTarget(package, target, cancellationToken);
        if (controlTarget is null && !IsSupportedContentControlTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported content-control target '{target}'. Expected a content control ID such as M.CC0001 or H001.CC0001.", operation, target)];
        }

        if (controlTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        XElement? checkBox = controlTarget.ContentControl
            .Element(OoxmlNs.W + "sdtPr")
            ?.Element(OoxmlNs.W + "checkBox");
        if (checkBox is null)
        {
            return [UnsupportedContentControlKindDiagnostic(controlTarget.ContentControl, operation, target, "a checkbox content control")];
        }

        DocxDiagnostic? lockDiagnostic = ValidateContentControlUnlocked(controlTarget.ContentControl, operation, target);
        if (lockDiagnostic is not null)
        {
            return [lockDiagnostic];
        }

        XElement? content = controlTarget.ContentControl.Element(OoxmlNs.W + "sdtContent");
        if (content is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E4310", $"Content control '{target}' has no editable content container.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        SetCheckboxChecked(checkBox, checkedValue.Value);
        ReplaceContentControlText(content, GetCheckboxDisplaySymbol(checkBox, checkedValue.Value));
        SaveDocumentPart(package, controlTarget.PartName, controlTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetContentControlChoice(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? value = operation.Fields.GetValueOrDefault("value");
        string? displayText = operation.Fields.GetValueOrDefault("display-text");
        if ((value is null) == (displayText is null))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Exactly one of 'value' or 'display-text' is required for set-content-control-choice.", operation, target));
        }

        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ContentControlTarget? controlTarget = ResolveContentControlTarget(package, target, cancellationToken);
        if (controlTarget is null && !IsSupportedContentControlTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported content-control target '{target}'. Expected a content control ID such as M.CC0001 or H001.CC0001.", operation, target)];
        }

        if (controlTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        XElement? list = controlTarget.ContentControl
            .Element(OoxmlNs.W + "sdtPr")
            ?.Elements()
            .FirstOrDefault(element => element.Name == OoxmlNs.W + "dropDownList" || element.Name == OoxmlNs.W + "comboBox");
        if (list is null)
        {
            return [UnsupportedContentControlKindDiagnostic(controlTarget.ContentControl, operation, target, "a dropdown or combo box content control")];
        }

        DocxDiagnostic? lockDiagnostic = ValidateContentControlUnlocked(controlTarget.ContentControl, operation, target);
        if (lockDiagnostic is not null)
        {
            return [lockDiagnostic];
        }

        ContentControlChoice? choice = ResolveContentControlChoice(list, value, displayText);
        if (choice is null)
        {
            string selector = value is not null ? $"value '{value}'" : $"display-text '{displayText}'";
            return [Diagnostic(DocxSeverity.Error, "E4205", $"Content control '{target}' has no list item with {selector}.", operation, target)];
        }

        XElement? content = controlTarget.ContentControl.Element(OoxmlNs.W + "sdtContent");
        if (content is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E4310", $"Content control '{target}' has no editable content container.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        ReplaceContentControlText(content, choice.DisplayText);
        SaveDocumentPart(package, controlTarget.PartName, controlTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetContentControlDate(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? value = ReadRequiredField(operation, "value", diagnostics);
        string? displayText = operation.Fields.GetValueOrDefault("display-text");
        if (value is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ContentControlTarget? controlTarget = ResolveContentControlTarget(package, target, cancellationToken);
        if (controlTarget is null && !IsSupportedContentControlTargetShape(target))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported content-control target '{target}'. Expected a content control ID such as M.CC0001 or H001.CC0001.", operation, target)];
        }

        if (controlTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        XElement? date = controlTarget.ContentControl
            .Element(OoxmlNs.W + "sdtPr")
            ?.Element(OoxmlNs.W + "date");
        if (date is null)
        {
            return [UnsupportedContentControlKindDiagnostic(controlTarget.ContentControl, operation, target, "a date content control")];
        }

        DocxDiagnostic? lockDiagnostic = ValidateContentControlUnlocked(controlTarget.ContentControl, operation, target);
        if (lockDiagnostic is not null)
        {
            return [lockDiagnostic];
        }

        XElement? content = controlTarget.ContentControl.Element(OoxmlNs.W + "sdtContent");
        if (content is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E4310", $"Content control '{target}' has no editable content container.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        SetContentControlDateValue(date, value);
        ReplaceContentControlText(content, displayText ?? value);
        SaveDocumentPart(package, controlTarget.PartName, controlTarget.Document);
        return [];
    }
    private static bool IsPlainTextContentControl(XElement contentControl)
    {
        return contentControl
            .Element(OoxmlNs.W + "sdtPr")
            ?.Element(OoxmlNs.W + "text") is not null;
    }

    private static bool IsRichTextContentControl(XElement contentControl)
    {
        XElement? properties = contentControl.Element(OoxmlNs.W + "sdtPr");
        if (properties is null)
        {
            return true;
        }

        if (properties.Element(OoxmlNs.W + "richText") is not null)
        {
            return true;
        }

        return !properties.Elements().Any(element => element.Name.LocalName is
            "text" or
            "checkBox" or
            "dropDownList" or
            "comboBox" or
            "date" or
            "picture" or
            "group" or
            "repeatingSection" or
            "repeatingSectionItem");
    }

    private static string ReadContentControlKind(XElement contentControl)
    {
        XElement? properties = contentControl.Element(OoxmlNs.W + "sdtPr");
        if (properties is null)
        {
            return "rich-text";
        }

        string? kind = properties.Elements()
            .Select(element => element.Name.LocalName)
            .FirstOrDefault(name => name is "text" or "richText" or "checkBox" or "dropDownList" or "comboBox" or "date" or "picture" or "group" or "repeatingSection" or "repeatingSectionItem");
        return kind switch
        {
            "text" => "plain-text",
            "richText" => "rich-text",
            "checkBox" => "checkbox",
            "dropDownList" => "dropdown-list",
            "comboBox" => "combo-box",
            "date" => "date",
            "picture" => "picture",
            "group" => "group",
            "repeatingSection" => "repeating-section",
            "repeatingSectionItem" => "repeating-section-item",
            _ => "rich-text"
        };
    }

    private static DocxDiagnostic UnsupportedContentControlKindDiagnostic(
        XElement contentControl,
        DocxPatchOperation operation,
        string target,
        string expected)
    {
        string kind = ReadContentControlKind(contentControl);
        string guidance = kind switch
        {
            "picture" => "Picture content controls preserve a picture container; use read/media to inspect the contained image and target image operations when applicable.",
            "group" => "Group content controls protect a container; target an editable child content control instead.",
            "repeating-section" or "repeating-section-item" => "Repeating-section subtree edits require cloning or deleting structured document tag subtrees, which DocxEdit currently rejects with E4315.",
            "checkbox" => "Use set-content-control-checkbox for checkbox state edits.",
            "dropdown-list" or "combo-box" => "Use set-content-control-choice for dropdown or combo-box selections.",
            "date" => "Use set-content-control-date for date values.",
            _ => "Choose an operation that matches the content-control kind."
        };
        return Diagnostic(
            DocxSeverity.Error,
            "E4310",
            $"Content control '{target}' is kind '{kind}', not {expected}. {guidance}",
            operation,
            target);
    }

    private static DocxDiagnostic? ValidateContentControlUnlocked(
        XElement contentControl,
        DocxPatchOperation operation,
        string target)
    {
        XElement? lockElement = contentControl.Element(OoxmlNs.W + "sdtPr")?.Element(OoxmlNs.W + "lock");
        if (lockElement is null)
        {
            return null;
        }

        string lockValue = (string?)lockElement.Attribute(OoxmlNs.W + "val") ?? "locked";
        if (string.Equals(lockValue, "unlocked", StringComparison.Ordinal))
        {
            return null;
        }

        return Diagnostic(
            DocxSeverity.Error,
            "E4310",
            $"Content control '{target}' is locked by w:lock='{lockValue}'.",
            operation,
            target);
    }

    private static void SetCheckboxChecked(XElement checkBox, bool checkedValue)
    {
        XElement? checkedElement = checkBox.Element(OoxmlNs.W + "checked");
        if (checkedElement is null)
        {
            checkedElement = new XElement(OoxmlNs.W + "checked");
            checkBox.Add(checkedElement);
        }

        checkedElement.SetAttributeValue(OoxmlNs.W + "val", checkedValue ? "1" : "0");
    }

    private static string GetCheckboxDisplaySymbol(XElement checkBox, bool checkedValue)
    {
        string stateElementName = checkedValue ? "checkedState" : "uncheckedState";
        string? stateValue = (string?)checkBox
            .Element(OoxmlNs.W + stateElementName)
            ?.Attribute(OoxmlNs.W + "val");
        return TryDecodeStateSymbol(stateValue, out string? symbol)
            ? symbol
            : char.ConvertFromUtf32(checkedValue ? 0x2612 : 0x2610);
    }

    private static bool TryDecodeStateSymbol(string? value, [NotNullWhen(true)] out string? symbol)
    {
        symbol = null;
        if (string.IsNullOrEmpty(value))
        {
            return false;
        }

        if (value.Length == 1)
        {
            symbol = value;
            return true;
        }

        if (value.All(Uri.IsHexDigit) &&
            int.TryParse(value, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out int codePoint) &&
            codePoint > 0)
        {
            symbol = char.ConvertFromUtf32(codePoint);
            return true;
        }

        return false;
    }

    private static ContentControlChoice? ResolveContentControlChoice(XElement list, string? value, string? displayText)
    {
        foreach (XElement item in list.Elements(OoxmlNs.W + "listItem"))
        {
            string? itemValue = (string?)item.Attribute(OoxmlNs.W + "value");
            string? itemDisplayText = (string?)item.Attribute(OoxmlNs.W + "displayText") ?? itemValue;
            if (value is not null && string.Equals(itemValue, value, StringComparison.Ordinal))
            {
                return new ContentControlChoice(itemDisplayText ?? string.Empty);
            }

            if (displayText is not null && string.Equals(itemDisplayText, displayText, StringComparison.Ordinal))
            {
                return new ContentControlChoice(itemDisplayText ?? string.Empty);
            }
        }

        return null;
    }

    private static void SetContentControlDateValue(XElement date, string value)
    {
        XElement? fullDate = date.Element(OoxmlNs.W + "fullDate");
        if (fullDate is null)
        {
            fullDate = new XElement(OoxmlNs.W + "fullDate");
            date.Add(fullDate);
        }

        fullDate.SetAttributeValue(OoxmlNs.W + "val", value);
    }

    private static void ReplaceContentControlText(XElement content, string text)
    {
        bool blockLevel = content.Elements(OoxmlNs.W + "p").Any();
        content.RemoveNodes();
        if (blockLevel)
        {
            content.Add(CreateSimpleParagraph(text, style: null, paragraphProperties: null));
        }
        else
        {
            content.Add(CreateSimpleRun(text));
        }
    }
}
