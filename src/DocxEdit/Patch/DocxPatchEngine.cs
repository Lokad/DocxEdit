using System.Text;
using System.Xml.Linq;
using DocxEdit.Ooxml;

namespace DocxEdit;

internal static class DocxPatchEngine
{
    public static PatchExecutionResult Check(
        OoxmlPackage package,
        DocxPatch patch,
        DocxEditOptions options,
        CancellationToken cancellationToken = default)
    {
        return Execute(package, patch, options, apply: false, cancellationToken);
    }

    public static PatchExecutionResult Apply(
        OoxmlPackage package,
        DocxPatch patch,
        DocxEditOptions options,
        CancellationToken cancellationToken = default)
    {
        return Execute(package, patch, options, apply: true, cancellationToken);
    }

    private static PatchExecutionResult Execute(
        OoxmlPackage package,
        DocxPatch patch,
        DocxEditOptions options,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        var reports = new List<DocxPatchOperationReport>();
        foreach (DocxPatchOperation operation in patch.Operations)
        {
            cancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<DocxDiagnostic> operationDiagnostics = operation.OperationName switch
            {
                "replace-text" => ExecuteReplaceText(package, operation, apply, cancellationToken),
                "replace-paragraph" => ExecuteReplaceParagraph(package, operation, apply, cancellationToken),
                "insert-before" => ExecuteInsertBlock(package, operation, insertAfter: false, apply, cancellationToken),
                "insert-after" => ExecuteInsertBlock(package, operation, insertAfter: true, apply, cancellationToken),
                "delete-block" => ExecuteDeleteBlock(package, operation, apply, cancellationToken),
                "set-cell" => ExecuteSetCell(package, operation, apply, cancellationToken),
                "append-row" => ExecuteAppendRow(package, operation, apply, cancellationToken),
                "insert-row-before" => ExecuteInsertRow(package, operation, insertAfter: false, apply, cancellationToken),
                "insert-row-after" => ExecuteInsertRow(package, operation, insertAfter: true, apply, cancellationToken),
                "delete-row" => ExecuteDeleteRow(package, operation, apply, cancellationToken),
                "replace-image" => ExecuteReplaceImage(package, operation, options, apply, cancellationToken),
                "set-image-alt" => ExecuteSetImageAlt(package, operation, apply, cancellationToken),
                "delete-image" => ExecuteDeleteImage(package, operation, apply, cancellationToken),
                _ => [Diagnostic(DocxSeverity.Error, "E4201", $"Unsupported operation '{operation.OperationName}'.", operation)]
            };
            bool operationSuccess = operationDiagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error);
            diagnostics.AddRange(operationDiagnostics);
            reports.Add(new DocxPatchOperationReport(
                operation.Index,
                operation.OperationName,
                operation.Fields.GetValueOrDefault("target"),
                operationSuccess,
                operationDiagnostics));
        }

        return new PatchExecutionResult(diagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error), diagnostics, reports);
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteReplaceText(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? find = ReadRequiredField(operation, "find", diagnostics);
        string? replacement = ReadRequiredField(operation, "with", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ParagraphTarget? paragraphTarget = ResolveParagraphTarget(package, target!, cancellationToken);
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
                    target)
            ];
        }

        if (!current.Contains(find!, StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E4203", $"Find text was not found in {target}.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        string edited = current.Replace(find!, replacement!, StringComparison.Ordinal);
        ReplaceParagraphText(paragraphTarget.Paragraph, edited);
        SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteReplaceParagraph(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        string? style = operation.Fields.GetValueOrDefault("style");
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        ParagraphTarget? paragraphTarget = ResolveParagraphTarget(package, target!, cancellationToken);
        if (paragraphTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        string current = ReadVisibleText(paragraphTarget.Paragraph);
        if (expected is not null && !string.Equals(current, expected, StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected text does not match current text.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        ReplaceParagraphText(paragraphTarget.Paragraph, text!);
        if (style is not null)
        {
            SetParagraphStyle(paragraphTarget.Paragraph, style);
        }

        SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteInsertBlock(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool insertAfter,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        string? style = operation.Fields.GetValueOrDefault("style");
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
        XElement? targetBlock = ResolveMainBlock(body, target!);
        if (targetBlock is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        XElement paragraph = CreateSimpleParagraph(text!, style);
        if (insertAfter)
        {
            targetBlock.AddAfterSelf(paragraph);
        }
        else
        {
            targetBlock.AddBeforeSelf(paragraph);
        }

        SaveMainDocument(package, document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteDeleteBlock(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
        XElement? targetBlock = ResolveMainBlock(body, target!);
        if (targetBlock is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (expected is not null)
        {
            string current = ReadVisibleText(targetBlock);
            if (!string.Equals(current, expected, StringComparison.Ordinal))
            {
                return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected text does not match current text.", operation, target)];
            }
        }

        if (!apply)
        {
            return [];
        }

        targetBlock.Remove();
        SaveMainDocument(package, document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteReplaceImage(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? asset = ReadRequiredField(operation, "asset", diagnostics);
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!TryParseMainImageTarget(target!, out int imageOrdinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported replace-image target '{target}'. Expected an image ID such as M.I0001.", operation, target)];
        }

        ImageTarget? imageTarget = FindMainImageTarget(package, imageOrdinal, cancellationToken);
        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (!TryReadAsset(options.AssetProvider, asset!, cancellationToken, out byte[] bytes, out string? contentType, out DocxDiagnostic? assetDiagnostic, operation, target))
        {
            return [assetDiagnostic!];
        }

        if (imageTarget.Part.ContentType is not null &&
            !string.Equals(imageTarget.Part.ContentType, contentType, StringComparison.OrdinalIgnoreCase))
        {
            return [Diagnostic(DocxSeverity.Error, "E5204", $"Replacing image content type '{imageTarget.Part.ContentType}' with '{contentType}' is not supported for existing media part {imageTarget.Part.Name}.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        package.ReplacePartBytes(imageTarget.Part.Name, bytes);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetImageAlt(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? alt = ReadRequiredField(operation, "alt", diagnostics);
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!TryParseMainImageTarget(target!, out int imageOrdinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-image-alt target '{target}'. Expected an image ID such as M.I0001.", operation, target)];
        }

        ImageBlipTarget? imageTarget = FindMainImageBlipTarget(package, imageOrdinal, cancellationToken);
        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        XElement? drawing = imageTarget.Blip.Ancestors(OoxmlNs.W + "drawing").FirstOrDefault();
        XElement? container = drawing?.Descendants(OoxmlNs.Wp + "inline").FirstOrDefault()
            ?? drawing?.Descendants(OoxmlNs.Wp + "anchor").FirstOrDefault();
        if (container is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E5205", $"Image '{target}' does not have editable DrawingML properties.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        XElement? docPr = container.Element(OoxmlNs.Wp + "docPr");
        if (docPr is null)
        {
            docPr = new XElement(OoxmlNs.Wp + "docPr");
            docPr.SetAttributeValue("id", "1");
            docPr.SetAttributeValue("name", target);
            container.AddFirst(docPr);
        }

        docPr.SetAttributeValue("descr", alt);
        SaveMainDocument(package, imageTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteDeleteImage(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!TryParseMainImageTarget(target!, out int imageOrdinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported delete-image target '{target}'. Expected an image ID such as M.I0001.", operation, target)];
        }

        ImageBlipTarget? imageTarget = FindMainImageBlipTarget(package, imageOrdinal, cancellationToken);
        if (imageTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        XElement? drawing = imageTarget.Blip.Ancestors(OoxmlNs.W + "drawing").FirstOrDefault();
        if (drawing is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E5205", $"Image '{target}' does not have an editable DrawingML object.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        drawing.Remove();
        SaveMainDocument(package, imageTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetCell(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        bool force = ReadBooleanField(operation, "force", diagnostics) ?? false;
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!TryParseMainCellTarget(target!, out int tableOrdinal, out int rowOrdinal, out int cellOrdinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported set-cell target '{target}'. Expected a table cell ID such as M.T0001.R02.C03.", operation, target)];
        }

        XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
        XElement? table = FindTable(body, tableOrdinal);
        XElement? row = table?.Elements(OoxmlNs.W + "tr").ElementAtOrDefault(rowOrdinal - 1);
        XElement? cell = row?.Elements(OoxmlNs.W + "tc").ElementAtOrDefault(cellOrdinal - 1);
        if (cell is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        string current = ReadVisibleText(cell);
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

        if (IsVerticalMergeContinuation(cell))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Unsupported merged-cell target '{target}'.", operation, target)];
        }

        if (!force && !IsSimpleEditableCell(cell))
        {
            return [Diagnostic(DocxSeverity.Error, "E4302", $"Cell '{target}' contains unsupported content. Use force true only when replacing all cell content is intended.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        ReplaceCellText(cell, text!);
        SaveMainDocument(package, document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteAppendRow(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!TryParseMainTableTarget(target!, out int tableOrdinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported append-row target '{target}'. Expected a table ID such as M.T0001.", operation, target)];
        }

        string[] cellTexts = operation.FieldValues
            .Where(field => field.Name == "cell")
            .Select(field => field.Value)
            .ToArray();
        if (cellTexts.Length == 0)
        {
            return [Diagnostic(DocxSeverity.Error, "E4202", "Operation 'append-row' is missing required field 'cell'.", operation, target)];
        }

        XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
        XElement? table = FindTable(body, tableOrdinal);
        if (table is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        XElement[] rows = table.Elements(OoxmlNs.W + "tr").ToArray();
        if (rows.Length == 0)
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' has no rows to clone.", operation, target)];
        }

        if (ContainsVerticalMerges(table))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' contains vertical merges that are not supported by append-row.", operation, target)];
        }

        if (!IsRectangular(table, out int columnCount))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' is not rectangular and cannot be appended safely.", operation, target)];
        }

        if (cellTexts.Length != columnCount)
        {
            return [Diagnostic(DocxSeverity.Error, "E4303", $"append-row expected {columnCount} cell field(s), but received {cellTexts.Length}.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        XElement newRow = CreateRowFromTemplate(rows[^1], cellTexts);
        rows[^1].AddAfterSelf(newRow);
        SaveMainDocument(package, document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteInsertRow(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool insertAfter,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        bool force = ReadBooleanField(operation, "force", diagnostics) ?? false;
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!TryParseMainRowTarget(target!, out int tableOrdinal, out int rowOrdinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported {operation.OperationName} target '{target}'. Expected a table row ID such as M.T0001.R02.", operation, target)];
        }

        string[] cellTexts = operation.FieldValues
            .Where(field => field.Name == "cell")
            .Select(field => field.Value)
            .ToArray();
        if (cellTexts.Length == 0)
        {
            return [Diagnostic(DocxSeverity.Error, "E4202", $"Operation '{operation.OperationName}' is missing required field 'cell'.", operation, target)];
        }

        XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
        XElement? table = FindTable(body, tableOrdinal);
        XElement[] rows = table?.Elements(OoxmlNs.W + "tr").ToArray() ?? [];
        XElement? templateRow = rowOrdinal < 1 || rowOrdinal > rows.Length ? null : rows[rowOrdinal - 1];
        if (templateRow is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (ContainsVerticalMerges(table!))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' contains vertical merges that are not supported by {operation.OperationName}.", operation, target)];
        }

        int expectedCellCount;
        if (IsRectangular(table!, out int columnCount))
        {
            expectedCellCount = columnCount;
        }
        else
        {
            if (!force)
            {
                return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' is not rectangular and cannot be edited safely without force true.", operation, target)];
            }

            expectedCellCount = templateRow.Elements(OoxmlNs.W + "tc").Count();
        }

        if (expectedCellCount == 0)
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Row '{target}' has no cells to clone.", operation, target)];
        }

        if (cellTexts.Length != expectedCellCount)
        {
            return [Diagnostic(DocxSeverity.Error, "E4303", $"{operation.OperationName} expected {expectedCellCount} cell field(s), but received {cellTexts.Length}.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        XElement newRow = CreateRowFromTemplate(templateRow, cellTexts);
        if (insertAfter)
        {
            templateRow.AddAfterSelf(newRow);
        }
        else
        {
            templateRow.AddBeforeSelf(newRow);
        }

        SaveMainDocument(package, document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteDeleteRow(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        bool force = ReadBooleanField(operation, "force", diagnostics) ?? false;
        if (diagnostics.Count != 0)
        {
            return diagnostics;
        }

        if (!TryParseMainRowTarget(target!, out int tableOrdinal, out int rowOrdinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Unsupported delete-row target '{target}'. Expected a table row ID such as M.T0001.R02.", operation, target)];
        }

        XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
        XElement? table = FindTable(body, tableOrdinal);
        XElement[] rows = table?.Elements(OoxmlNs.W + "tr").ToArray() ?? [];
        XElement? row = rowOrdinal < 1 || rowOrdinal > rows.Length ? null : rows[rowOrdinal - 1];
        if (row is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 targets: {target}.", operation, target)];
        }

        if (rows.Length == 1)
        {
            return [Diagnostic(DocxSeverity.Error, "E4304", $"Cannot delete the last row of table '{target}'.", operation, target)];
        }

        if (ContainsVerticalMerges(table!))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' contains vertical merges that are not supported by delete-row.", operation, target)];
        }

        if (!force && !IsRectangular(table!, out _))
        {
            return [Diagnostic(DocxSeverity.Error, "E4301", $"Table '{target}' is not rectangular and cannot be edited safely without force true.", operation, target)];
        }

        if (!apply)
        {
            return [];
        }

        row.Remove();
        SaveMainDocument(package, document);
        return [];
    }

    private static string? ReadRequiredField(
        DocxPatchOperation operation,
        string fieldName,
        List<DocxDiagnostic> diagnostics)
    {
        if (operation.Fields.TryGetValue(fieldName, out string? value) && value.Length != 0)
        {
            return value;
        }

        diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4202", $"Operation '{operation.OperationName}' is missing required field '{fieldName}'.", operation));
        return null;
    }

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

    private static XElement? ResolveMainBlock(XElement body, string target)
    {
        if (TryParseMainParagraphTarget(target, out int paragraphOrdinal))
        {
            return FindParagraph(body, paragraphOrdinal);
        }

        if (TryParseMainTableTarget(target, out int tableOrdinal))
        {
            return FindTable(body, tableOrdinal);
        }

        return null;
    }

    private static ParagraphTarget? ResolveParagraphTarget(
        OoxmlPackage package,
        string target,
        CancellationToken cancellationToken)
    {
        if (TryParseMainParagraphTarget(target, out int mainParagraphOrdinal))
        {
            XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
            XElement? paragraph = FindParagraph(body, mainParagraphOrdinal);
            return paragraph is null || package.MainDocumentPartName is null
                ? null
                : new ParagraphTarget(package.MainDocumentPartName, document, paragraph);
        }

        if (TryParseStoryParagraphTarget(target, 'H', out int headerOrdinal, out int headerParagraphOrdinal))
        {
            return ResolveRelatedStoryParagraphTarget(package, OoxmlRelTypes.Header, headerOrdinal, headerParagraphOrdinal, cancellationToken);
        }

        if (TryParseStoryParagraphTarget(target, 'F', out int footerOrdinal, out int footerParagraphOrdinal))
        {
            return ResolveRelatedStoryParagraphTarget(package, OoxmlRelTypes.Footer, footerOrdinal, footerParagraphOrdinal, cancellationToken);
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
        if (package.MainDocumentPartName is null || storyOrdinal < 1)
        {
            return null;
        }

        OoxmlRelationship? relationship = package
            .GetRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => !relationship.IsExternal && relationship.Type == relationshipType && relationship.ResolvedTarget is not null)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .ElementAtOrDefault(storyOrdinal - 1);
        if (relationship?.ResolvedTarget is null)
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

    private static ImageTarget? FindMainImageTarget(
        OoxmlPackage package,
        int imageOrdinal,
        CancellationToken cancellationToken)
    {
        ImageBlipTarget? target = FindMainImageBlipTarget(package, imageOrdinal, cancellationToken);
        return target is null ? null : new ImageTarget(target.RelationshipId, target.Part);
    }

    private static ImageBlipTarget? FindMainImageBlipTarget(
        OoxmlPackage package,
        int imageOrdinal,
        CancellationToken cancellationToken)
    {
        if (imageOrdinal < 1 || package.MainDocumentPartName is null)
        {
            return null;
        }

        XDocument document = LoadMainDocument(package, cancellationToken, out _);
        IReadOnlyDictionary<string, OoxmlRelationship> relationships = package
            .GetRelationships(package.MainDocumentPartName, cancellationToken)
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
                return new ImageBlipTarget(document, blip, relationshipId, part);
            }
        }

        return null;
    }

    private static bool TryReadAsset(
        IDocxAssetProvider? assetProvider,
        string asset,
        CancellationToken cancellationToken,
        out byte[] bytes,
        out string? contentType,
        out DocxDiagnostic? diagnostic,
        DocxPatchOperation operation,
        string? target)
    {
        bytes = [];
        contentType = null;
        diagnostic = null;
        if (assetProvider is null)
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E5201", "No asset provider is configured for image operation assets.", operation, target);
            return false;
        }

        if (!assetProvider.TryOpen(asset, out Stream stream, out string? contentTypeHint, out string? fileNameHint))
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E5202", $"Asset '{asset}' could not be resolved.", operation, target);
            return false;
        }

        using (stream)
        using (var memory = new MemoryStream())
        {
            CopyTo(stream, memory, cancellationToken);
            bytes = memory.ToArray();
        }

        contentType = DetectImageContentType(bytes, contentTypeHint, fileNameHint ?? asset);
        if (contentType is null)
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E5203", $"Asset '{asset}' is not a supported PNG or JPEG image.", operation, target);
            return false;
        }

        return true;
    }

    private static string? DetectImageContentType(byte[] bytes, string? contentTypeHint, string? fileNameHint)
    {
        string? normalizedHint = NormalizeImageContentType(contentTypeHint);
        if (normalizedHint is not null)
        {
            return normalizedHint;
        }

        string extension = Path.GetExtension(fileNameHint ?? string.Empty).ToLowerInvariant();
        if (extension is ".png")
        {
            return "image/png";
        }

        if (extension is ".jpg" or ".jpeg")
        {
            return "image/jpeg";
        }

        if (bytes.Length >= 8 &&
            bytes[0] == 0x89 &&
            bytes[1] == 0x50 &&
            bytes[2] == 0x4E &&
            bytes[3] == 0x47 &&
            bytes[4] == 0x0D &&
            bytes[5] == 0x0A &&
            bytes[6] == 0x1A &&
            bytes[7] == 0x0A)
        {
            return "image/png";
        }

        if (bytes.Length >= 3 &&
            bytes[0] == 0xFF &&
            bytes[1] == 0xD8 &&
            bytes[2] == 0xFF)
        {
            return "image/jpeg";
        }

        return null;
    }

    private static string? NormalizeImageContentType(string? contentType)
    {
        return contentType?.ToLowerInvariant() switch
        {
            "image/png" => "image/png",
            "image/jpeg" => "image/jpeg",
            "image/jpg" => "image/jpeg",
            _ => null
        };
    }

    private static void CopyTo(Stream source, Stream destination, CancellationToken cancellationToken)
    {
        byte[] buffer = new byte[81920];
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int read = source.Read(buffer, 0, buffer.Length);
            if (read == 0)
            {
                return;
            }

            destination.Write(buffer, 0, read);
        }
    }

    private static string ReadVisibleText(XElement container)
    {
        var builder = new StringBuilder();
        foreach (XElement element in container.Descendants())
        {
            if (element.Ancestors(OoxmlNs.W + "del").Any() ||
                element.Ancestors(OoxmlNs.W + "moveFrom").Any())
            {
                continue;
            }

            if (element.Name == OoxmlNs.W + "t")
            {
                builder.Append(element.Value);
            }
            else if (element.Name == OoxmlNs.W + "tab")
            {
                builder.Append('\t');
            }
            else if (element.Name == OoxmlNs.W + "br")
            {
                builder.Append('\n');
            }
        }

        return builder.ToString();
    }

    private static void ReplaceParagraphText(XElement paragraph, string text)
    {
        XElement? paragraphProperties = paragraph.Element(OoxmlNs.W + "pPr");
        XElement? firstRunProperties = paragraph
            .Elements(OoxmlNs.W + "r")
            .Elements(OoxmlNs.W + "rPr")
            .FirstOrDefault();

        paragraph.RemoveNodes();
        if (paragraphProperties is not null)
        {
            paragraph.Add(new XElement(paragraphProperties));
        }

        var run = new XElement(OoxmlNs.W + "r");
        if (firstRunProperties is not null)
        {
            run.Add(new XElement(firstRunProperties));
        }

        var textElement = new XElement(OoxmlNs.W + "t", text);
        if (RequiresPreserveSpace(text))
        {
            textElement.SetAttributeValue(OoxmlNs.Xml + "space", "preserve");
        }

        run.Add(textElement);
        paragraph.Add(run);
    }

    private static void ReplaceCellText(XElement cell, string text)
    {
        XElement? cellProperties = cell.Element(OoxmlNs.W + "tcPr");
        cell.RemoveNodes();
        if (cellProperties is not null)
        {
            cell.Add(new XElement(cellProperties));
        }

        cell.Add(CreateSimpleParagraph(text));
    }

    private static XElement CreateRowFromTemplate(XElement templateRow, IReadOnlyList<string> cellTexts)
    {
        var row = new XElement(OoxmlNs.W + "tr");
        XElement? rowProperties = templateRow.Element(OoxmlNs.W + "trPr");
        if (rowProperties is not null)
        {
            row.Add(new XElement(rowProperties));
        }

        XElement[] templateCells = templateRow.Elements(OoxmlNs.W + "tc").ToArray();
        for (int i = 0; i < cellTexts.Count; i++)
        {
            var cell = new XElement(OoxmlNs.W + "tc");
            XElement? cellProperties = templateCells[i].Element(OoxmlNs.W + "tcPr");
            if (cellProperties is not null)
            {
                cell.Add(new XElement(cellProperties));
            }

            cell.Add(CreateSimpleParagraph(cellTexts[i]));
            row.Add(cell);
        }

        return row;
    }

    private static XElement CreateSimpleParagraph(string text, string? style = null)
    {
        var paragraph = new XElement(OoxmlNs.W + "p");
        if (style is not null)
        {
            SetParagraphStyle(paragraph, style);
        }

        var run = new XElement(OoxmlNs.W + "r");
        foreach (XNode node in CreateTextNodes(text))
        {
            run.Add(node);
        }

        paragraph.Add(run);
        return paragraph;
    }

    private static void SetParagraphStyle(XElement paragraph, string style)
    {
        XElement? paragraphProperties = paragraph.Element(OoxmlNs.W + "pPr");
        if (paragraphProperties is null)
        {
            paragraphProperties = new XElement(OoxmlNs.W + "pPr");
            paragraph.AddFirst(paragraphProperties);
        }

        XElement? paragraphStyle = paragraphProperties.Element(OoxmlNs.W + "pStyle");
        if (paragraphStyle is null)
        {
            paragraphStyle = new XElement(OoxmlNs.W + "pStyle");
            paragraphProperties.AddFirst(paragraphStyle);
        }

        paragraphStyle.SetAttributeValue(OoxmlNs.W + "val", style);
    }

    private static IEnumerable<XNode> CreateTextNodes(string text)
    {
        var buffer = new StringBuilder();
        foreach (char ch in text)
        {
            if (ch == '\t' || ch == '\n')
            {
                if (buffer.Length != 0)
                {
                    yield return CreateTextElement(buffer.ToString());
                    buffer.Clear();
                }

                yield return ch == '\t'
                    ? new XElement(OoxmlNs.W + "tab")
                    : new XElement(OoxmlNs.W + "br");
                continue;
            }

            buffer.Append(ch);
        }

        yield return CreateTextElement(buffer.ToString());
    }

    private static XElement CreateTextElement(string text)
    {
        var textElement = new XElement(OoxmlNs.W + "t", text);
        if (RequiresPreserveSpace(text))
        {
            textElement.SetAttributeValue(OoxmlNs.Xml + "space", "preserve");
        }

        return textElement;
    }

    private static bool IsSimpleEditableCell(XElement cell)
    {
        XElement[] paragraphs = cell.Elements(OoxmlNs.W + "p").ToArray();
        return paragraphs.Length == 1 &&
            !cell.Elements(OoxmlNs.W + "tbl").Any() &&
            !cell.Descendants(OoxmlNs.W + "drawing").Any() &&
            !cell.Descendants(OoxmlNs.W + "fldSimple").Any() &&
            !cell.Descendants(OoxmlNs.W + "fldChar").Any() &&
            !cell.Descendants(OoxmlNs.W + "instrText").Any();
    }

    private static bool ContainsVerticalMerges(XElement table)
    {
        return table.Descendants(OoxmlNs.W + "vMerge").Any();
    }

    private static bool IsRectangular(XElement table, out int columnCount)
    {
        XElement[] rows = table.Elements(OoxmlNs.W + "tr").ToArray();
        columnCount = rows.Length == 0 ? 0 : rows[0].Elements(OoxmlNs.W + "tc").Count();
        int expectedColumnCount = columnCount;
        return expectedColumnCount != 0 && rows.All(row => row.Elements(OoxmlNs.W + "tc").Count() == expectedColumnCount);
    }

    private static bool IsVerticalMergeContinuation(XElement cell)
    {
        XElement? verticalMerge = cell.Element(OoxmlNs.W + "tcPr")?.Element(OoxmlNs.W + "vMerge");
        if (verticalMerge is null)
        {
            return false;
        }

        string? value = (string?)verticalMerge.Attribute(OoxmlNs.W + "val");
        return !string.Equals(value, "restart", StringComparison.OrdinalIgnoreCase);
    }

    private static bool? ReadBooleanField(DocxPatchOperation operation, string fieldName, List<DocxDiagnostic> diagnostics)
    {
        if (!operation.Fields.TryGetValue(fieldName, out string? value))
        {
            return null;
        }

        if (string.Equals(value, "true", StringComparison.Ordinal))
        {
            return true;
        }

        if (string.Equals(value, "false", StringComparison.Ordinal))
        {
            return false;
        }

        diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4204", $"Field '{fieldName}' must be true or false.", operation));
        return null;
    }

    private static XDocument LoadMainDocument(
        OoxmlPackage package,
        CancellationToken cancellationToken,
        out XElement body)
    {
        XDocument document = LoadDocumentPart(package, package.MainDocumentPartName!, cancellationToken, out XElement root);
        body = root.Element(OoxmlNs.W + "body")
            ?? throw new InvalidDataException("Main document part is missing w:body.");
        return document;
    }

    private static XDocument LoadDocumentPart(
        OoxmlPackage package,
        string partName,
        CancellationToken cancellationToken,
        out XElement root)
    {
        OoxmlPart documentPart = package.GetPart(partName)
            ?? throw new InvalidDataException($"Document part '{partName}' does not exist.");
        using Stream stream = documentPart.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        root = document.Root
            ?? throw new InvalidDataException($"Document part '{partName}' has no XML root.");
        return document;
    }

    private static void SaveMainDocument(OoxmlPackage package, XDocument document)
    {
        SaveDocumentPart(package, package.MainDocumentPartName!, document);
    }

    private static void SaveDocumentPart(OoxmlPackage package, string partName, XDocument document)
    {
        using var output = new MemoryStream();
        document.Save(output, SaveOptions.DisableFormatting);
        package.ReplacePartBytes(partName, output.ToArray());
    }

    private static bool RequiresPreserveSpace(string text)
    {
        return text.Length != 0 &&
            (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[^1]) || text.Contains("  ", StringComparison.Ordinal));
    }

    private static DocxDiagnostic Diagnostic(
        DocxSeverity severity,
        string code,
        string message,
        DocxPatchOperation operation,
        string? targetId = null)
    {
        return new DocxDiagnostic(severity, code, message, TargetId: targetId, OperationIndex: operation.Index);
    }
}

internal sealed record PatchExecutionResult(
    bool Success,
    IReadOnlyList<DocxDiagnostic> Diagnostics,
    IReadOnlyList<DocxPatchOperationReport> Reports);

internal sealed record ImageTarget(string RelationshipId, OoxmlPart Part);

internal sealed record ImageBlipTarget(XDocument Document, XElement Blip, string RelationshipId, OoxmlPart Part);

internal sealed record ParagraphTarget(string PartName, XDocument Document, XElement Paragraph);
