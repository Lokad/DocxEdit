using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

internal static partial class DocxPatchEngine
{
    private const string SettingsContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
    private const string CommentsContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
    private const string CommentsExtendedContentType = "application/vnd.ms-word.commentsExtended+xml";
    private const string CommentsIdsContentType = "application/vnd.ms-word.commentsIds+xml";

    private static readonly IReadOnlyDictionary<XName, string> ProtectedTextEditElements = new Dictionary<XName, string>
    {
        [OoxmlNs.W + "hyperlink"] = "hyperlink",
        [OoxmlNs.W + "fldSimple"] = "field",
        [OoxmlNs.W + "fldChar"] = "field",
        [OoxmlNs.W + "instrText"] = "field",
        [OoxmlNs.W + "commentRangeStart"] = "comment",
        [OoxmlNs.W + "commentRangeEnd"] = "comment",
        [OoxmlNs.W + "commentReference"] = "comment",
        [OoxmlNs.W + "bookmarkStart"] = "bookmark",
        [OoxmlNs.W + "bookmarkEnd"] = "bookmark",
        [OoxmlNs.W + "sdt"] = "content-control",
        [OoxmlNs.W + "ins"] = "tracked-insertion",
        [OoxmlNs.W + "del"] = "tracked-deletion",
        [OoxmlNs.W + "moveFrom"] = "tracked-move-from",
        [OoxmlNs.W + "moveTo"] = "tracked-move-to",
        [OoxmlNs.W + "moveFromRangeStart"] = "tracked-move-from-range",
        [OoxmlNs.W + "moveFromRangeEnd"] = "tracked-move-from-range",
        [OoxmlNs.W + "moveToRangeStart"] = "tracked-move-to-range",
        [OoxmlNs.W + "moveToRangeEnd"] = "tracked-move-to-range",
        [OoxmlNs.W + "customXmlInsRangeStart"] = "tracked-custom-xml-insertion",
        [OoxmlNs.W + "customXmlInsRangeEnd"] = "tracked-custom-xml-insertion",
        [OoxmlNs.W + "customXmlDelRangeStart"] = "tracked-custom-xml-deletion",
        [OoxmlNs.W + "customXmlDelRangeEnd"] = "tracked-custom-xml-deletion",
        [OoxmlNs.W + "customXmlMoveFromRangeStart"] = "tracked-custom-xml-move-from",
        [OoxmlNs.W + "customXmlMoveFromRangeEnd"] = "tracked-custom-xml-move-from",
        [OoxmlNs.W + "customXmlMoveToRangeStart"] = "tracked-custom-xml-move-to",
        [OoxmlNs.W + "customXmlMoveToRangeEnd"] = "tracked-custom-xml-move-to",
        [OoxmlNs.W + "drawing"] = "drawing",
        [OoxmlNs.W + "pict"] = "picture",
        [OoxmlNs.W + "object"] = "object"
    };

    public static PatchExecutionResult Check(
        OoxmlPackage package,
        DocxPatch patch,
        DocxEditOptions options,
        CancellationToken cancellationToken)
    {
        return Execute(package, patch, options, apply: false, cancellationToken);
    }

    public static PatchExecutionResult Apply(
        OoxmlPackage package,
        DocxPatch patch,
        DocxEditOptions options,
        CancellationToken cancellationToken)
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
        IReadOnlyList<DocxDiagnostic> trackChangeOptionDiagnostics = ValidateTrackChangeOptions(options);
        if (trackChangeOptionDiagnostics.Count != 0)
        {
            return new PatchExecutionResult(false, trackChangeOptionDiagnostics, []);
        }

        var reports = new List<DocxPatchOperationReport>();
        foreach (DocxPatchOperation operation in patch.Operations)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var operationDiagnostics = new List<DocxDiagnostic>();
            var generatedRevisionIds = new List<string>();
            TableOperationSnapshot? tableBefore = CaptureTableOperationSnapshot(package, operation, cancellationToken);
            bool supportsTrackedChanges = SupportsTrackedChangeOutput(operation.OperationName);
            if (options.TrackChanges == TrackChangesMode.Require && !supportsTrackedChanges)
            {
                operationDiagnostics.Add(Diagnostic(
                    DocxSeverity.Error,
                    "E6001",
                    BuildUnsupportedTrackedOperationMessage(options.TrackChanges, operation),
                    operation,
                    operation.Fields.GetValueOrDefault("target"),
                    "track-changes-no-revision-representation",
                    "require-failed"));
            }
            else
            {
                if (options.TrackChanges == TrackChangesMode.Suggest && !supportsTrackedChanges)
                {
                    operationDiagnostics.Add(Diagnostic(
                        DocxSeverity.Warning,
                        "W4001",
                        BuildUnsupportedTrackedOperationMessage(options.TrackChanges, operation),
                        operation,
                        operation.Fields.GetValueOrDefault("target"),
                        "track-changes-no-revision-representation",
                        "direct-edit-preserve-existing-revisions"));
                }

                IReadOnlyList<DocxDiagnostic>? executed = TryExecuteOperation(package, operation, options, apply, generatedRevisionIds, cancellationToken);
                if (executed is not null)
                {
                    operationDiagnostics.AddRange(executed);
                }
                else
                {
                    operationDiagnostics.Add(Diagnostic(DocxSeverity.Error, "E4201", $"Unsupported operation '{operation.OperationName}'.", operation));
                }
            }
            bool operationSuccess = operationDiagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error);
            diagnostics.AddRange(operationDiagnostics);
            reports.Add(new DocxPatchOperationReport(
                operation.Index,
                operation.OperationName,
                operation.Fields.GetValueOrDefault("target"),
                operationSuccess,
                operationDiagnostics)
            {
                AffectedTargets = operationSuccess ? BuildAffectedTargets(operation, tableBefore) : [],
                GeneratedRevisionIds = operationSuccess ? generatedRevisionIds.ToArray() : []
            });
        }

        bool shouldMarkFieldsDirty = apply &&
            options.MarkFieldsDirtyWhenEditing &&
            patch.Operations.Count != 0 &&
            patch.Operations.Any(operation => MarksFieldsDirtyAfterEdit(operation.OperationName)) &&
            diagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error);
        bool containsFieldsBeforeRefresh = shouldMarkFieldsDirty && PackageContainsFieldMarkup(package, cancellationToken);
        if (shouldMarkFieldsDirty)
        {
            diagnostics.AddRange(MarkFieldsDirty(package, cancellationToken));
            if (containsFieldsBeforeRefresh && diagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error))
            {
                diagnostics.Add(new DocxDiagnostic(
                    DocxSeverity.Warning,
                    "W5103",
                    "Document contains fields and was marked for Word-side field refresh; DocxEdit does not recalculate field results.",
                    Feature: "field",
                    Fallback: "word-refresh-required"));
            }
        }

        if (apply && diagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error))
        {
            diagnostics.AddRange(ValidateEditedPackage(package, cancellationToken));
        }

        return new PatchExecutionResult(diagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error), diagnostics, reports);
    }

    private static string BuildUnsupportedTrackedOperationMessage(TrackChangesMode mode, DocxPatchOperation operation)
    {
        string operationName = operation.OperationName;
        string support = TrackChangesSupportValue(operationName);

        return mode switch
        {
            TrackChangesMode.Require =>
                $"TrackChangesMode.Require cannot apply operation '{operationName}' as tracked output because its catalog support is '{support}'. Existing tracked-change markup is preserved, but this operation does not generate new revision markup.",
            TrackChangesMode.Suggest =>
                $"TrackChangesMode.Suggest will apply operation '{operationName}' directly because its catalog support is '{support}'. Existing tracked-change markup is preserved, but this operation does not generate new revision markup.",
            _ =>
                $"TrackChangesMode.{mode} does not generate tracked output for operation '{operationName}' because its catalog support is '{support}'."
        };
    }


    private static bool PackageContainsFieldMarkup(OoxmlPackage package, CancellationToken cancellationToken)
    {
        foreach (OoxmlPart part in package.Parts.Values
            .Where(part => part.Name.StartsWith("/word/", StringComparison.OrdinalIgnoreCase) &&
                part.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                !part.Name.Contains("/_rels/", StringComparison.OrdinalIgnoreCase))
            .OrderBy(part => part.Name, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            if (document.Descendants().Any(element =>
                element.Name == OoxmlNs.W + "fldSimple" ||
                element.Name == OoxmlNs.W + "fldChar" ||
                element.Name == OoxmlNs.W + "instrText"))
            {
                return true;
            }
        }

        return false;
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
    private static IReadOnlyList<string> GetEditableStoryPartNames(
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {

        var partNames = new List<string> { package.MainDocumentPartName };
        IReadOnlyList<ResolvedOoxmlRelationship> relationships = package.GetResolvedRelationships(package.MainDocumentPartName, cancellationToken);
        partNames.AddRange(relationships
            .Where(relationship => relationship.Type == OoxmlRelTypes.Header)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .Select(relationship => relationship.ResolvedTarget)
            .Where(partName => package.GetPart(partName) is not null));
        partNames.AddRange(relationships
            .Where(relationship => relationship.Type == OoxmlRelTypes.Footer)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .Select(relationship => relationship.ResolvedTarget)
            .Where(partName => package.GetPart(partName) is not null));
        return partNames;
    }

    private static bool TryReadAsset(
        IDocxAssetProvider? assetProvider,
        string asset,
        CancellationToken cancellationToken,
        out byte[] bytes,
        [NotNullWhen(true)] out string? contentType,
        [NotNullWhen(false)] out DocxDiagnostic? diagnostic,
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

    private static bool TryResolveStyleId(
        OoxmlPackage package,
        string requestedStyle,
        string styleType,
        CancellationToken cancellationToken,
        [NotNullWhen(true)] out string? styleId,
        out DocxDiagnostic? diagnostic,
        DocxPatchOperation operation,
        string? target)
    {
        styleId = null;
        diagnostic = null;
        IReadOnlyList<DocxStyleInfo> styles = DocxStyleScanner.Scan(package, cancellationToken)
            .Where(style => style.Type == styleType)
            .ToArray();
        DocxStyleInfo? byId = styles.FirstOrDefault(style => string.Equals(style.StyleId, requestedStyle, StringComparison.Ordinal));
        if (byId is not null)
        {
            styleId = byId.StyleId;
            return true;
        }

        DocxStyleInfo[] byName = styles
            .Where(style => string.Equals(style.Name, requestedStyle, StringComparison.Ordinal))
            .ToArray();
        if (byName.Length == 1)
        {
            styleId = byName[0].StyleId;
            return true;
        }

        diagnostic = byName.Length > 1
            ? Diagnostic(DocxSeverity.Error, "E7102", $"Style name '{requestedStyle}' is ambiguous.", operation, target)
            : Diagnostic(DocxSeverity.Error, "E7101", $"Style '{requestedStyle}' was not found.", operation, target);
        return false;
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

    private static bool TryGetProtectedTextEditFeature(XElement paragraph, out string feature)
    {
        foreach (XElement element in paragraph.Descendants())
        {
            if (ProtectedTextEditElements.TryGetValue(element.Name, out string? found) && found is not null)
            {
                feature = found;
                return true;
            }
        }

        feature = string.Empty;
        return false;
    }

    private static int? ReadPositiveOccurrence(DocxPatchOperation operation, List<DocxDiagnostic> diagnostics)
    {
        if (!operation.Fields.TryGetValue("occurrence", out string? value))
        {
            return null;
        }

        if (!int.TryParse(value, out int occurrence) || occurrence <= 0)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Field 'occurrence' must be greater than 0.", operation, operation.Fields.GetValueOrDefault("target")));
            return null;
        }

        return occurrence;
    }

    private static IReadOnlyList<TextRange> FindTextMatches(string text, string find, int? occurrence)
    {
        var matches = new List<TextRange>();
        int index = 0;
        int seen = 0;
        while (index <= text.Length)
        {
            int found = text.IndexOf(find, index, StringComparison.Ordinal);
            if (found < 0)
            {
                break;
            }

            seen++;
            if (occurrence is null || seen == occurrence)
            {
                matches.Add(new TextRange(found, find.Length));
                if (occurrence is not null)
                {
                    break;
                }
            }

            index = found + find.Length;
        }

        return matches;
    }

    private static string ApplyTextReplacement(string text, IReadOnlyList<TextRange> matches, string replacement)
    {
        var builder = new StringBuilder(text);
        for (int i = matches.Count - 1; i >= 0; i--)
        {
            TextRange match = matches[i];
            builder.Remove(match.Start, match.Length);
            builder.Insert(match.Start, replacement);
        }

        return builder.ToString();
    }

    private static bool TryReplaceParagraphTextPreservingRuns(
        XElement paragraph,
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

        TextPosition?[] positions = BuildTextPositions(paragraph);
        foreach (TextRange match in matches)
        {
            for (int i = match.Start; i < match.Start + match.Length; i++)
            {
                if (i < 0 || i >= positions.Length || positions[i] is null)
                {
                    unsupportedReason = "match includes tabs, line breaks, or non-text run content";
                    return false;
                }
            }
        }

        for (int i = matches.Count - 1; i >= 0; i--)
        {
            TextRange match = matches[i];
            TextPosition start = positions[match.Start]!;
            TextPosition end = positions[match.Start + match.Length - 1]!;
            if (ReferenceEquals(start.TextElement, end.TextElement))
            {
                string value = start.TextElement.Value;
                string edited = value.Remove(start.Offset, match.Length).Insert(start.Offset, replacement);
                SetTextElementValue(start.TextElement, edited);
                continue;
            }

            string startValue = start.TextElement.Value;
            string endValue = end.TextElement.Value;
            SetTextElementValue(start.TextElement, startValue[..start.Offset] + replacement);
            SetTextElementValue(end.TextElement, endValue[(end.Offset + 1)..]);

            bool insideRange = false;
            foreach (XElement textElement in paragraph.Elements(OoxmlNs.W + "r").Elements(OoxmlNs.W + "t"))
            {
                if (ReferenceEquals(textElement, start.TextElement))
                {
                    insideRange = true;
                    continue;
                }

                if (ReferenceEquals(textElement, end.TextElement))
                {
                    break;
                }

                if (insideRange)
                {
                    SetTextElementValue(textElement, string.Empty);
                }
            }
        }

        return true;
    }

    private static TextPosition?[] BuildTextPositions(XElement paragraph)
    {
        var positions = new List<TextPosition?>();
        foreach (XElement run in paragraph.Elements(OoxmlNs.W + "r"))
        {
            foreach (XElement child in run.Elements())
            {
                if (child.Name == OoxmlNs.W + "t")
                {
                    for (int i = 0; i < child.Value.Length; i++)
                    {
                        positions.Add(new TextPosition(child, i));
                    }
                }
                else if (child.Name == OoxmlNs.W + "tab" || child.Name == OoxmlNs.W + "br")
                {
                    positions.Add(null);
                }
            }
        }

        return positions.ToArray();
    }

    private static void SetTextElementValue(XElement textElement, string text)
    {
        textElement.Value = text;
        if (RequiresPreserveSpace(text))
        {
            textElement.SetAttributeValue(OoxmlNs.Xml + "space", "preserve");
        }
        else
        {
            textElement.SetAttributeValue(OoxmlNs.Xml + "space", null);
        }
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
        XElement? paragraphProperties = cell
            .Elements(OoxmlNs.W + "p")
            .Elements(OoxmlNs.W + "pPr")
            .FirstOrDefault();
        cell.RemoveNodes();
        if (cellProperties is not null)
        {
            cell.Add(new XElement(cellProperties));
        }

        cell.Add(CreateSimpleParagraph(text, style: null, paragraphProperties: paragraphProperties));
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

            XElement? paragraphProperties = templateCells[i]
                .Elements(OoxmlNs.W + "p")
                .Elements(OoxmlNs.W + "pPr")
                .FirstOrDefault();
            cell.Add(CreateSimpleParagraph(cellTexts[i], style: null, paragraphProperties: paragraphProperties));
            row.Add(cell);
        }

        return row;
    }

    private static XElement CreateSimpleParagraph(string text, string? style, XElement? paragraphProperties)
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

        var run = new XElement(OoxmlNs.W + "r");
        foreach (XNode node in CreateTextNodes(text))
        {
            run.Add(node);
        }

        paragraph.Add(run);
        return paragraph;
    }

    private static XElement CreateSimpleRun(string text)
    {
        var run = new XElement(OoxmlNs.W + "r");
        foreach (XNode node in CreateTextNodes(text))
        {
            run.Add(node);
        }

        return run;
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

    private static XElement CreateDeletedTextElement(string text)
    {
        var textElement = new XElement(OoxmlNs.W + "delText", text);
        if (RequiresPreserveSpace(text))
        {
            textElement.SetAttributeValue(OoxmlNs.Xml + "space", "preserve");
        }

        return textElement;
    }

    private static string[] AllocateRevisionIds(
        OoxmlPackage package,
        int count,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        int nextId = FindNextRevisionId(package, cancellationToken);
        foreach (string generatedRevisionId in generatedRevisionIds)
        {
            if (int.TryParse(generatedRevisionId, out int id) && id >= nextId)
            {
                nextId = id + 1;
            }
        }

        string[] ids = BuildRevisionIds(nextId, count);
        generatedRevisionIds.AddRange(ids);
        return ids;
    }

    private static string[] AllocateRevisionIds(
        OoxmlPackage package,
        int count,
        CancellationToken cancellationToken)
    {
        return BuildRevisionIds(FindNextRevisionId(package, cancellationToken), count);
    }

    private static int FindNextRevisionId(
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        int nextId = 1;
        foreach (OoxmlPart part in package.Parts.Values
            .Where(part => part.Name.StartsWith("/word/", StringComparison.OrdinalIgnoreCase) && part.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(part => part.Name, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            foreach (XAttribute idAttribute in document.Descendants().Attributes(OoxmlNs.W + "id"))
            {
                if (int.TryParse(idAttribute.Value, out int id) && id >= nextId)
                {
                    nextId = id + 1;
                }
            }
        }

        return nextId;
    }

    private static string[] BuildRevisionIds(int nextId, int count)
    {
        string[] ids = new string[count];
        for (int i = 0; i < ids.Length; i++)
        {
            ids[i] = (nextId++).ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        return ids;
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

    private static IReadOnlyList<DocxDiagnostic> MarkFieldsDirty(OoxmlPackage package, CancellationToken cancellationToken)
    {
        string settingsPartName = ResolveOrCreateSettingsPart(package, cancellationToken);
        OoxmlPart settingsPart = package.GetPart(settingsPartName)
            ?? throw new InvalidDataException($"Settings part '{settingsPartName}' does not exist.");
        using Stream stream = settingsPart.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        XElement settings = document.Root
            ?? throw new InvalidDataException($"Settings part '{settingsPartName}' has no XML root.");
        if (settings.Name != OoxmlNs.W + "settings")
        {
            return [new DocxDiagnostic(DocxSeverity.Error, "E9001", $"Post-edit validation failed for {settingsPartName}: Expected root element 'settings', found '{settings.Name.LocalName}'.", PartName: settingsPartName)];
        }

        XElement? updateFields = settings.Element(OoxmlNs.W + "updateFields");
        if (updateFields is null)
        {
            updateFields = new XElement(OoxmlNs.W + "updateFields");
            settings.Add(updateFields);
        }

        updateFields.SetAttributeValue(OoxmlNs.W + "val", "true");
        SaveDocumentPart(package, settingsPartName, document);
        return [];
    }

    private static string ResolveOrCreateSettingsPart(OoxmlPackage package, CancellationToken cancellationToken)
    {
        ResolvedOoxmlRelationship? relationship = package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .FirstOrDefault(relationship => relationship.Type == OoxmlRelTypes.Settings);
        if (relationship is not null)
        {
            return relationship.ResolvedTarget;
        }

        const string settingsPartName = "/word/settings.xml";
        if (package.GetPart(settingsPartName) is null)
        {
            byte[] bytes = Encoding.UTF8.GetBytes("""
                <?xml version="1.0" encoding="utf-8"?>
                <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />
                """);
            package.AddPart(settingsPartName, SettingsContentType, bytes, cancellationToken);
        }

        string relationshipId = OoxmlIds.AllocateRelationshipId(package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .Select(relationship => relationship.Id));
        package.AddRelationship(package.MainDocumentPartName, relationshipId, OoxmlRelTypes.Settings, GetRelativeRelationshipTarget(package.MainDocumentPartName, settingsPartName), targetMode: null, cancellationToken);
        return settingsPartName;
    }

    private static IReadOnlyList<DocxDiagnostic> ValidateEditedPackage(OoxmlPackage package, CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        foreach (string partName in package.TouchedPartNames.Order(StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                diagnostics.Add(PostEditValidationDiagnostic(partName, $"Touched part '{partName}' is missing."));
                continue;
            }

            if (!ShouldValidateTouchedPart(part))
            {
                ValidateTouchedBinaryPart(part, diagnostics);
                continue;
            }

            try
            {
                using Stream stream = part.OpenRead();
                XDocument document = SafeXml.Load(stream, cancellationToken);
                ValidateTouchedPartRoot(package, part, document, diagnostics, cancellationToken);
            }
            catch (Exception ex) when (ex is InvalidDataException or IOException or XmlException)
            {
                diagnostics.Add(PostEditValidationDiagnostic(part.Name, ex.Message));
            }
        }

        return diagnostics;
    }

    private static bool ShouldValidateTouchedPart(OoxmlPart part)
    {
        return string.Equals(part.Name, "/[Content_Types].xml", StringComparison.OrdinalIgnoreCase) ||
            part.Name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) ||
            part.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(part.ContentType, OoxmlContentTypeNames.Xml, StringComparison.OrdinalIgnoreCase) ||
            part.ContentType?.EndsWith("+xml", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static void ValidateTouchedBinaryPart(OoxmlPart part, List<DocxDiagnostic> diagnostics)
    {
        if (part.Bytes.Length != 0)
        {
            return;
        }

        if (string.Equals(part.ContentType, "image/png", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(part.ContentType, "image/jpeg", StringComparison.OrdinalIgnoreCase))
        {
            diagnostics.Add(PostEditValidationDiagnostic(part.Name, "Image part is empty."));
        }
    }

    private static void ValidateTouchedPartRoot(
        OoxmlPackage package,
        OoxmlPart part,
        XDocument document,
        List<DocxDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        XElement root = document.Root
            ?? throw new InvalidDataException("XML part has no root element.");

        if (string.Equals(part.Name, "/[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
        {
            RequireRoot(part, root, OoxmlNs.Ct + "Types", diagnostics);
            return;
        }

        if (part.Name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
        {
            RequireRoot(part, root, OoxmlNs.Rel + "Relationships", diagnostics);
            ValidateRelationshipTargets(package, part, diagnostics, cancellationToken);
            return;
        }

        if (string.Equals(part.Name, package.MainDocumentPartName, StringComparison.OrdinalIgnoreCase))
        {
            RequireRoot(part, root, OoxmlNs.W + "document", diagnostics);
            if (root.Element(OoxmlNs.W + "body") is null)
            {
                diagnostics.Add(PostEditValidationDiagnostic(part.Name, "Main document part is missing w:body."));
            }

            DocxPackageValidator.ValidateRevisionMarkup(document, part.Name, diagnostics);
            return;
        }

        if (string.Equals(part.ContentType, "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml", StringComparison.OrdinalIgnoreCase))
        {
            RequireRoot(part, root, OoxmlNs.W + "hdr", diagnostics);
            DocxPackageValidator.ValidateRevisionMarkup(document, part.Name, diagnostics);
            return;
        }

        if (string.Equals(part.ContentType, "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml", StringComparison.OrdinalIgnoreCase))
        {
            RequireRoot(part, root, OoxmlNs.W + "ftr", diagnostics);
            DocxPackageValidator.ValidateRevisionMarkup(document, part.Name, diagnostics);
            return;
        }

        if (string.Equals(part.ContentType, CommentsContentType, StringComparison.OrdinalIgnoreCase))
        {
            RequireRoot(part, root, OoxmlNs.W + "comments", diagnostics);
            DocxPackageValidator.ValidateRevisionMarkup(document, part.Name, diagnostics);
            return;
        }

        if (string.Equals(part.ContentType, SettingsContentType, StringComparison.OrdinalIgnoreCase))
        {
            RequireRoot(part, root, OoxmlNs.W + "settings", diagnostics);
        }
    }

    private static void RequireRoot(OoxmlPart part, XElement root, XName expectedRoot, List<DocxDiagnostic> diagnostics)
    {
        if (root.Name != expectedRoot)
        {
            diagnostics.Add(PostEditValidationDiagnostic(part.Name, $"Expected root element '{expectedRoot.LocalName}', found '{root.Name.LocalName}'."));
        }
    }

    private static void ValidateRelationshipTargets(
        OoxmlPackage package,
        OoxmlPart relationshipPart,
        List<DocxDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        string sourcePartName = GetSourcePartNameFromRelationshipPartName(relationshipPart.Name);
        using Stream stream = relationshipPart.OpenRead();
        foreach (OoxmlRelationship relationship in OoxmlPackage.ParseRelationships(stream, sourcePartName, cancellationToken))
        {
            if (!relationship.IsExternal &&
                relationship.ResolvedTarget is not null &&
                package.GetPart(relationship.ResolvedTarget) is null)
            {
                diagnostics.Add(PostEditValidationDiagnostic(relationshipPart.Name, $"Relationship '{relationship.Id}' targets missing part '{relationship.ResolvedTarget}'."));
            }
        }
    }

    private static string GetSourcePartNameFromRelationshipPartName(string relationshipPartName)
    {
        string normalized = OoxmlPath.NormalizePartName(relationshipPartName);
        if (normalized == "/_rels/.rels")
        {
            return "/";
        }

        const string relationshipMarker = "/_rels/";
        int markerIndex = normalized.LastIndexOf(relationshipMarker, StringComparison.Ordinal);
        if (markerIndex < 0 || !normalized.EndsWith(".rels", StringComparison.Ordinal))
        {
            throw new InvalidDataException($"Invalid relationship part name '{relationshipPartName}'.");
        }

        string directory = normalized[..markerIndex];
        string fileName = normalized[(markerIndex + relationshipMarker.Length)..^".rels".Length];
        return OoxmlPath.NormalizePartName($"{directory}/{fileName}");
    }

    private static DocxDiagnostic PostEditValidationDiagnostic(string partName, string message)
    {
        return new DocxDiagnostic(DocxSeverity.Error, "E9001", $"Post-edit validation failed for {partName}: {message}", PartName: partName);
    }

    private static XDocument LoadMainDocument(
        OoxmlPackage package,
        CancellationToken cancellationToken,
        out XElement body)
    {
        XDocument document = LoadDocumentPart(package, package.MainDocumentPartName, cancellationToken, out XElement root);
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
        SaveDocumentPart(package, package.MainDocumentPartName, document);
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
        DocxPatchOperation operation)
    {
        return Diagnostic(severity, code, message, operation, targetId: null);
    }

    private static DocxDiagnostic Diagnostic(
        DocxSeverity severity,
        string code,
        string message,
        DocxPatchOperation operation,
        string? targetId)
    {
        return Diagnostic(severity, code, message, operation, targetId, feature: null, fallback: null);
    }

    private static DocxDiagnostic Diagnostic(
        DocxSeverity severity,
        string code,
        string message,
        DocxPatchOperation operation,
        string? targetId,
        string? feature,
        string? fallback)
    {
        return new DocxDiagnostic(
            severity,
            code,
            message,
            TargetId: targetId,
            Feature: feature,
            Fallback: fallback,
            OperationIndex: operation.Index);
    }
}
