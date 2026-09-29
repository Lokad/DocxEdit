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

        Dictionary<string, int> rangeBaseline = CaptureRangeStructureBaseline(package, cancellationToken);
        // D01: bind explicit paragraph/table/row/cell/section IDs to the input
        // snapshot for this patch; semantic selectors keep resolving live.
        // Untouched parts are restored byte-identical afterwards, so snapshot
        // bookkeeping never rewrites output the patch did not edit.
        Dictionary<string, byte[]> snapshotOriginals = RecordStoryPartBytes(package, cancellationToken);
        CaptureTargetSnapshot(package, cancellationToken);
        UnmarkStoryParts(package, cancellationToken);
        // One shared revision-ID allocator per execution; per-operation reports slice
        // their own ID ranges out of it below.
        var revisionIds = new RevisionIdTracker();
        bool anyMutation = false;
        bool priorFailure = false;
        var reports = new List<DocxPatchOperationReport>();
        foreach (DocxPatchOperation operation in patch.Operations)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (priorFailure)
            {
                // D14: operations after the first failure are skipped, never simulated
                // against the divergent disposable package. The overall failure plus
                // the first error already tell the repair story.
                var skippedDiagnostics = new List<DocxDiagnostic>
                {
                    Diagnostic(DocxSeverity.Info, SkippedDiagnosticCode, "Operation skipped after an earlier operation failed; nothing was attempted.", operation, operation.Fields.GetValueOrDefault("target"))
                };
                diagnostics.AddRange(skippedDiagnostics);
                reports.Add(new DocxPatchOperationReport(
                    operation.Index,
                    operation.OperationName,
                    operation.Fields.GetValueOrDefault("target"),
                    false,
                    skippedDiagnostics)
                {
                    AffectedTargets = [],
                    GeneratedRevisionIds = []
                });
                continue;
            }
            var operationDiagnostics = new List<DocxDiagnostic>();
            int revisionMark = revisionIds.Count;
            TableOperationSnapshot? tableBefore = CaptureTableOperationSnapshot(package, operation, cancellationToken);
            DocxTargetId? resolvedBefore = CaptureResolvedTargetSnapshot(package, operation, cancellationToken);
            PreviewSnapshot? previewBefore = CapturePreviewBefore(package, operation, options, cancellationToken);
            HashSet<string>? commentsBefore = operation.OperationName == "add-comment" ? ReadCommentIds(package, cancellationToken) : null;
            bool supportsTrackedChanges = SupportsTrackedChangeOutput(operation.OperationName);
            // D15: intrinsic review/annotation operations are permitted under Require:
            // a comment is already review markup, so requiring generated text
            // revisions for it is not the same policy as tracking content changes.
            if (options.TrackChanges == TrackChangesMode.Require && !supportsTrackedChanges && !IsAnnotationOperation(operation.OperationName))
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

                IReadOnlyList<DocxDiagnostic>? executed = TryExecuteOperation(package, operation, options, apply, revisionIds, cancellationToken);
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
            bool operationMutated = operationSuccess && operationDiagnostics.All(diagnostic => diagnostic.Code != NoOpDiagnosticCode);
            var (previewBeforeText, previewAfterText, previewTruncated) = FinalizePreview(package, operation, options, previewBefore, operationSuccess, cancellationToken);
            anyMutation |= operationMutated;
            priorFailure |= !operationSuccess;
            diagnostics.AddRange(operationDiagnostics);
            reports.Add(new DocxPatchOperationReport(
                operation.Index,
                operation.OperationName,
                operation.Fields.GetValueOrDefault("target"),
                operationSuccess,
                operationDiagnostics)
            {
                AffectedTargets = operationSuccess && operationMutated ? BuildAffectedTargets(operation, tableBefore, resolvedBefore) : [],
                CreatedTargetIds = operationSuccess && operationMutated ? BuildCreatedTargetIds(operation, package, commentsBefore, cancellationToken) : [],
                PreviewBefore = previewBeforeText,
                PreviewAfter = previewAfterText,
                PreviewTruncated = previewTruncated,
                GeneratedRevisionIds = apply && operationSuccess ? revisionIds.Skip(revisionMark).ToArray() : []
            });
            // Comment and bookmark operations allocate w:id values outside revision
            // allocation; later revision allocations must rescan to stay above them.
            if (WordIdAllocatingOperations.Contains(operation.OperationName))
            {
                revisionIds.Invalidate();
            }
        }

        // D01: strip snapshot marks before field refresh, validation, and
        // publication so transient IDs never reach validators or output.
        StripTargetSnapshot(package, snapshotOriginals, cancellationToken);

        bool shouldMarkFieldsDirty = options.MarkFieldsDirtyWhenEditing &&
            anyMutation &&
            patch.Operations.Count != 0 &&
            patch.Operations.Any(operation => MarksFieldsDirtyAfterEdit(operation.OperationName)) &&
            diagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error);
        bool containsFieldsBeforeRefresh = shouldMarkFieldsDirty && PackageContainsFieldMarkup(package, cancellationToken);
        if (shouldMarkFieldsDirty)
        {
            diagnostics.AddRange(MarkFieldsDirty(package, cancellationToken));
            if (containsFieldsBeforeRefresh && diagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error))
            {
                diagnostics.Add(new DocxDiagnostic(DocxSeverity.Warning, "W5103", "Document contains fields and was marked for Word-side field refresh; DocxEdit does not recalculate field results.") with
                {
                    Feature = "field",
                    Fallback = "word-refresh-required"
                });
            }
        }

        if (diagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error))
        {
            diagnostics.AddRange(ValidateEditedPackage(package, cancellationToken));
        }

        if (diagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error))
        {
            diagnostics.AddRange(ValidateNewRangeViolations(package, rangeBaseline, cancellationToken));
        }

        if (diagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error))
        {
            long editedTotalBytes = 0;
            foreach (OoxmlPart part in package.Parts.Values)
            {
                cancellationToken.ThrowIfCancellationRequested();
                checked
                {
                    editedTotalBytes += part.Bytes.Length;
                }
            }

            if (editedTotalBytes > options.Quotas.MaxUncompressedBytes)
            {
                diagnostics.Add(new DocxDiagnostic(DocxSeverity.Error, "E0001", "Edited package exceeds the maximum supported uncompressed size."));
            }
        }

        bool success = diagnostics.All(diagnostic => diagnostic.Severity != DocxSeverity.Error);
        if (!success)
        {
            // D14: revision IDs from simulated edits were never published when the
            // patch fails, so reports must not carry them as if they were committed.
            reports = reports.Select(report => report with { GeneratedRevisionIds = [] }).ToList();
        }

        return new PatchExecutionResult(success, diagnostics, reports);
    }

    private static Dictionary<string, int> CaptureRangeStructureBaseline(
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        var baseline = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (string partName in RangeCheckedStoryParts(package, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            foreach (DocxDiagnostic diagnostic in DocxPackageValidator.ValidateRangeStructure(document, partName))
            {
                baseline.TryGetValue(RangeDiagnosticKey(diagnostic), out int count);
                baseline[RangeDiagnosticKey(diagnostic)] = count + 1;
            }
        }

        return baseline;
    }

    private static IReadOnlyList<DocxDiagnostic> ValidateNewRangeViolations(
        OoxmlPackage package,
        Dictionary<string, int> baseline,
        CancellationToken cancellationToken)
    {
        var remaining = new Dictionary<string, int>(baseline, StringComparer.Ordinal);
        var fresh = new List<DocxDiagnostic>();
        foreach (string partName in package.TouchedPartNames.Order(StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!IsRangeCheckedStoryPart(package, partName, cancellationToken))
            {
                continue;
            }

            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            try
            {
                using Stream stream = part.OpenRead();
                XDocument document = SafeXml.Load(stream, cancellationToken);
                foreach (DocxDiagnostic diagnostic in DocxPackageValidator.ValidateRangeStructure(document, partName))
                {
                    string key = RangeDiagnosticKey(diagnostic);
                    if (remaining.TryGetValue(key, out int count) && count > 0)
                    {
                        remaining[key] = count - 1;
                        continue;
                    }

                    fresh.Add(diagnostic);
                }
            }
            catch (Exception ex) when (ex is InvalidDataException or IOException or XmlException)
            {
                fresh.Add(PostEditValidationDiagnostic(part.Name, ex.Message));
            }
        }

        return fresh;
    }

    private static IEnumerable<string> RangeCheckedStoryParts(OoxmlPackage package, CancellationToken cancellationToken)
    {
        foreach (StoryPartRef story in DocxPartRoles.GetOrderedStories(package, includeHeadersFooters: true, cancellationToken))
        {
            yield return story.PartName;
        }

        foreach (string? partName in new[]
        {
            DocxPartRoles.FindCommentsPartName(package, cancellationToken),
            DocxPartRoles.FindFootnotesPartName(package, cancellationToken),
            DocxPartRoles.FindEndnotesPartName(package, cancellationToken)
        })
        {
            if (partName is not null && package.GetPart(partName) is not null)
            {
                yield return partName;
            }
        }
    }

    private static bool IsRangeCheckedStoryPart(OoxmlPackage package, string partName, CancellationToken cancellationToken)
    {
        return RangeCheckedStoryParts(package, cancellationToken)
            .Contains(partName, StringComparer.OrdinalIgnoreCase);
    }

    private static string RangeDiagnosticKey(DocxDiagnostic diagnostic)
    {
        return string.Concat(diagnostic.Code, "|", diagnostic.PartName, "|", diagnostic.Message);
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
        foreach (string wordPartName in DocxPartRoles.GetWordProcessingParts(package, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            OoxmlPart? fieldPart = package.GetPart(wordPartName);
            if (fieldPart is null)
            {
                continue;
            }

            using Stream stream = fieldPart.OpenRead();
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
        // D10: empty-value policy lives in the registry alongside the other
        // field contracts; handlers stay unaware of per-field allowances.
        bool allowEmpty = false;
        if (OperationsByName.TryGetValue(operation.OperationName, out OperationRegistration? registration))
        {
            OperationFieldDefinition? definition = registration.Fields.FirstOrDefault(field => string.Equals(field.Name, fieldName, StringComparison.Ordinal));
            allowEmpty = definition?.AllowEmpty ?? false;
        }

        if (operation.Fields.TryGetValue(fieldName, out string? value) && (value.Length != 0 || allowEmpty))
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

        return DocxPartRoles.GetOrderedStories(package, includeHeadersFooters: true, cancellationToken)
            .Select(story => story.PartName)
            .ToArray();
    }

    private static bool TryReadAsset(
        IDocxAssetProvider? assetProvider,
        string asset,
        long maxAssetBytes,
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
            try
            {
                CopyTo(stream, memory, maxAssetBytes, cancellationToken);
            }
            catch (InvalidDataException)
            {
                bytes = [];
                contentType = null;
                diagnostic = Diagnostic(DocxSeverity.Error, "E5207", $"Image asset '{asset}' exceeds the maximum supported size of {maxAssetBytes} bytes.", operation, target);
                return false;
            }

            bytes = memory.ToArray();
        }

        // Content is authoritative: magic bytes decide the type, provider hints
        // and file names are only hints, and the structure is validated boundedly.
        string? magicType = DetectImageMagicContentType(bytes);
        if (magicType is null)
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E5203", $"Asset '{asset}' is not a supported PNG or JPEG image.", operation, target);
            return false;
        }

        string? normalizedHint = NormalizeImageContentType(contentTypeHint);
        if (normalizedHint is not null && !string.Equals(normalizedHint, magicType, StringComparison.Ordinal))
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E5203", $"Asset '{asset}' content-type hint suggests '{normalizedHint}' but the content is '{magicType}'.", operation, target);
            return false;
        }

        if (contentTypeHint is null &&
            DetectImageExtensionContentType(fileNameHint ?? asset) is { } extensionType &&
            !string.Equals(extensionType, magicType, StringComparison.Ordinal))
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E5203", $"Asset '{asset}' file extension suggests '{extensionType}' but the content is '{magicType}'.", operation, target);
            return false;
        }

        if (!TryValidateImageStructure(bytes, magicType, out string? structuralReason))
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E5203", $"Asset '{asset}' is not a structurally valid {magicType} image ({structuralReason}).", operation, target);
            return false;
        }

        contentType = magicType;
        return true;
    }

    internal static bool TryValidateImageStructure(byte[] bytes, string contentType, out string? reason)
    {
        switch (contentType)
        {
            case "image/png":
                return TryValidatePngStructure(bytes, out reason);
            case "image/jpeg":
                return TryValidateJpegStructure(bytes, out reason);
            default:
                reason = "unsupported content type";
                return false;
        }
    }

    private static bool TryValidatePngStructure(byte[] bytes, out string? reason)
    {
        reason = null;
        if (DetectImageMagicContentType(bytes) != "image/png" || bytes.Length < 33)
        {
            reason = "missing signature or header";
            return false;
        }

        // Walk chunks: 4-byte big-endian length, 4-byte type, data, 4-byte CRC.
        // Requires IHDR first (length 13, sane dimensions and encoding), at
        // least one IDAT chunk, and IEND terminating the file exactly.
        int position = 8;
        bool seenHeader = false;
        bool seenData = false;
        while (true)
        {
            if (position + 8 > bytes.Length)
            {
                reason = "truncated chunk header";
                return false;
            }

            long length = ((long)bytes[position] << 24) | ((long)bytes[position + 1] << 16) | ((long)bytes[position + 2] << 8) | bytes[position + 3];
            if (length < 0 || length > 1_000_000_000 || position + 8 + length + 4 > bytes.Length)
            {
                reason = "truncated chunk data";
                return false;
            }

            string type = System.Text.Encoding.ASCII.GetString(bytes, position + 4, 4);
            if (!seenHeader)
            {
                if (type != "IHDR" || length != 13)
                {
                    reason = "missing IHDR header";
                    return false;
                }

                long width = ((long)bytes[position + 8] << 24) | ((long)bytes[position + 9] << 16) | ((long)bytes[position + 10] << 8) | bytes[position + 11];
                long height = ((long)bytes[position + 12] << 24) | ((long)bytes[position + 13] << 16) | ((long)bytes[position + 14] << 8) | bytes[position + 15];
                int bitDepth = bytes[position + 16];
                int colorType = bytes[position + 17];
                int compression = bytes[position + 18];
                int filter = bytes[position + 19];
                int interlace = bytes[position + 20];
                if (width < 1 || height < 1 || width > 1_000_000 || height > 1_000_000)
                {
                    reason = "invalid dimensions";
                    return false;
                }

                bool validEncoding = (bitDepth, colorType) switch
                {
                    (1, 0 or 3) => true,
                    (2, 0 or 3) => true,
                    (4, 0 or 3) => true,
                    (8, 0 or 2 or 3 or 4 or 6) => true,
                    (16, 0 or 2 or 4 or 6) => true,
                    _ => false
                };
                if (!validEncoding || compression != 0 || filter != 0 || interlace is not (0 or 1))
                {
                    reason = "unsupported encoding";
                    return false;
                }

                seenHeader = true;
            }
            else if (type == "IDAT")
            {
                seenData = true;
            }
            else if (type == "IEND")
            {
                if (!seenData)
                {
                    reason = "missing image data";
                    return false;
                }

                if (position + 12 != bytes.Length)
                {
                    reason = "trailing bytes after image end";
                    return false;
                }

                return true;
            }

            position += 8 + (int)length + 4;
        }
    }

    private static bool TryValidateJpegStructure(byte[] bytes, out string? reason)
    {
        reason = null;
        if (bytes.Length < 4 || bytes[0] != 0xFF || bytes[1] != 0xD8)
        {
            reason = "missing start marker";
            return false;
        }

        // Walk segments to the first frame header (SOF), then require scan data
        // closed by an end marker: truncated files without either are rejected.
        // Dimensions follow the JPEG specification (1..65500).
        int position = 2;
        bool seenFrame = false;
        while (position + 1 < bytes.Length)
        {
            if (bytes[position] != 0xFF)
            {
                position++;
                continue;
            }

            while (position < bytes.Length && bytes[position] == 0xFF)
            {
                position++;
            }

            if (position >= bytes.Length)
            {
                reason = "truncated marker";
                return false;
            }

            byte marker = bytes[position++];
            if (marker == 0x00 || marker is >= 0xD0 and <= 0xD7)
            {
                continue;
            }

            if (marker == 0xD9)
            {
                if (!seenFrame)
                {
                    reason = "missing frame header";
                    return false;
                }

                return true;
            }

            if (marker == 0xDA)
            {
                if (!seenFrame)
                {
                    reason = "missing frame header";
                    return false;
                }

                // Scan data runs to the end marker; only stuffed (FF00) and
                // restart markers may appear inside it.
                while (position + 1 < bytes.Length)
                {
                    if (bytes[position] != 0xFF)
                    {
                        position++;
                        continue;
                    }

                    byte following = bytes[position + 1];
                    if (following == 0x00 || following is >= 0xD0 and <= 0xD7)
                    {
                        position += 2;
                        continue;
                    }

                    if (following == 0xD9)
                    {
                        return true;
                    }

                    reason = "truncated scan data";
                    return false;
                }

                reason = "truncated scan data";
                return false;
            }

            if (marker is 0x01 || marker == 0xD8)
            {
                continue;
            }

            if (IsJpegFrameMarker(marker))
            {
                if (position + 7 > bytes.Length)
                {
                    reason = "truncated frame header";
                    return false;
                }

                int length = (bytes[position] << 8) | bytes[position + 1];
                if (length < 8)
                {
                    reason = "invalid frame header";
                    return false;
                }

                int height = (bytes[position + 3] << 8) | bytes[position + 4];
                int width = (bytes[position + 5] << 8) | bytes[position + 6];
                if (width < 1 || height < 1 || width > 65500 || height > 65500)
                {
                    reason = "invalid dimensions";
                    return false;
                }

                seenFrame = true;
            }

            if (position + 1 >= bytes.Length)
            {
                reason = "truncated segment";
                return false;
            }

            int segmentLength = (bytes[position] << 8) | bytes[position + 1];
            if (segmentLength < 2 || position + segmentLength > bytes.Length)
            {
                reason = "truncated segment";
                return false;
            }

            position += segmentLength;
        }

        reason = seenFrame ? "truncated scan data" : "missing frame header";
        return false;
    }

    private static bool IsJpegFrameMarker(byte marker)
    {
        return marker is >= 0xC0 and <= 0xCF
            && marker is not (0xC4 or 0xC8 or 0xCC);
    }

    private static bool TryResolveStyleId(
        OoxmlPackage package,
        string requestedStyle,
        string styleType,
        CancellationToken cancellationToken,
        [NotNullWhen(true)] out string? styleId,
        [NotNullWhen(false)] out DocxDiagnostic? diagnostic,
        DocxPatchOperation operation,
        string? target)
    {
        styleId = null;
        diagnostic = null;
        IReadOnlyList<DocxStyleInfo> allStyles = DocxStyleScanner.Scan(package, cancellationToken);
        IReadOnlyList<DocxStyleInfo> styles = allStyles
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

        if (byName.Length > 1)
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E7102", $"Style name '{requestedStyle}' is ambiguous.", operation, target);
            return false;
        }

        DocxStyleInfo? idOfAnotherKind = allStyles.FirstOrDefault(style => style.Type != styleType && string.Equals(style.StyleId, requestedStyle, StringComparison.Ordinal));
        if (idOfAnotherKind is not null)
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E7103", $"Style '{requestedStyle}' is a {idOfAnotherKind.Type} style, not a {styleType} style.", operation, target);
            return false;
        }

        DocxStyleInfo[] namesOfAnotherKind = allStyles
            .Where(style => style.Type != styleType && string.Equals(style.Name, requestedStyle, StringComparison.Ordinal))
            .ToArray();
        if (namesOfAnotherKind.Length == 1)
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E7103", $"Style name '{requestedStyle}' identifies the {namesOfAnotherKind[0].Type} style '{namesOfAnotherKind[0].StyleId}', not a {styleType} style.", operation, target);
            return false;
        }

        diagnostic = Diagnostic(DocxSeverity.Error, "E7101", $"Style '{requestedStyle}' was not found.", operation, target);
        return false;
    }

    private static string? DetectImageExtensionContentType(string? fileNameHint)
    {
        return Path.GetExtension(fileNameHint ?? string.Empty).ToLowerInvariant() switch
        {
            ".png" => "image/png",
            ".jpg" or ".jpeg" => "image/jpeg",
            _ => null
        };
    }

    private static string? DetectImageMagicContentType(byte[] bytes)
    {
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

    private static void CopyTo(Stream source, Stream destination, long maxBytes, CancellationToken cancellationToken)
    {
        byte[] buffer = new byte[81920];
        long total = 0;
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int read = source.Read(buffer, 0, buffer.Length);
            if (read == 0)
            {
                return;
            }

            total += read;
            if (total > maxBytes)
            {
                throw new InvalidDataException("Asset stream exceeds the maximum supported size.");
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

    /// <summary>
    /// Finds the first paired range the removal of <paramref name="removed"/> would newly orphan:
    /// a bookmark/comment-range marker inside the subtree whose counterpart lies outside it
    /// in an otherwise healthy pair, or complex-field markers whose removal unbalances the part.
    /// Returns a human-readable range description, or null when removal keeps pairing intact.
    /// </summary>
    private static string? FindOrphanedRangeBoundary(XDocument document, XElement removed)
    {
        foreach ((XName startName, XName endName, bool named, string label) in new[]
        {
            (OoxmlNs.W + "bookmarkStart", OoxmlNs.W + "bookmarkEnd", true, "bookmark"),
            (OoxmlNs.W + "commentRangeStart", OoxmlNs.W + "commentRangeEnd", false, "comment range")
        })
        {
            foreach (XElement marker in removed.Descendants(startName))
            {
                string? id = (string?)marker.Attribute(OoxmlNs.W + "id");
                if (string.IsNullOrWhiteSpace(id) || !IsHealthyPair(document, startName, endName, id))
                {
                    continue;
                }

                if (!HasCounterpartInside(removed, endName, id))
                {
                    return DescribeRange(label, marker, id);
                }
            }

            foreach (XElement marker in removed.Descendants(endName))
            {
                string? id = (string?)marker.Attribute(OoxmlNs.W + "id");
                if (string.IsNullOrWhiteSpace(id) || !IsHealthyPair(document, startName, endName, id))
                {
                    continue;
                }

                if (!HasCounterpartInside(removed, startName, id))
                {
                    return DescribeRange(label, marker, id);
                }
            }
        }

        if (WouldUnbalanceFieldMarkers(document, removed))
        {
            return "complex field boundary";
        }

        return null;
    }

    private static bool IsHealthyPair(XDocument document, XName startName, XName endName, string id)
    {
        return document.Descendants(startName).Count(element => string.Equals((string?)element.Attribute(OoxmlNs.W + "id"), id, StringComparison.Ordinal)) == 1 &&
            document.Descendants(endName).Count(element => string.Equals((string?)element.Attribute(OoxmlNs.W + "id"), id, StringComparison.Ordinal)) == 1;
    }

    private static bool HasCounterpartInside(XElement removed, XName counterpartName, string id)
    {
        return removed.Descendants(counterpartName)
            .Any(element => string.Equals((string?)element.Attribute(OoxmlNs.W + "id"), id, StringComparison.Ordinal));
    }

    private static string DescribeRange(string label, XElement marker, string id)
    {
        string? name = (string?)marker.Attribute(OoxmlNs.W + "name");
        return string.IsNullOrWhiteSpace(name)
            ? $"{label} (id '{id}')"
            : $"{label} '{name}' (id '{id}')";
    }

    private static bool WouldUnbalanceFieldMarkers(XDocument document, XElement removed)
    {
        int documentBegins = 0;
        int documentEnds = 0;
        foreach (XElement field in document.Descendants(OoxmlNs.W + "fldChar"))
        {
            string? type = (string?)field.Attribute(OoxmlNs.W + "fldCharType");
            if (string.Equals(type, "begin", StringComparison.Ordinal))
            {
                documentBegins++;
            }
            else if (string.Equals(type, "end", StringComparison.Ordinal))
            {
                documentEnds++;
            }
        }

        if (documentBegins != documentEnds)
        {
            return false;
        }

        int removedBegins = 0;
        int removedEnds = 0;
        foreach (XElement field in removed.Descendants(OoxmlNs.W + "fldChar"))
        {
            string? type = (string?)field.Attribute(OoxmlNs.W + "fldCharType");
            if (string.Equals(type, "begin", StringComparison.Ordinal))
            {
                removedBegins++;
            }
            else if (string.Equals(type, "end", StringComparison.Ordinal))
            {
                removedEnds++;
            }
        }

        return removedBegins != removedEnds;
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

        return ReplaceDirectTextRanges(paragraph, BuildTextPositions(paragraph), matches, replacement, out unsupportedReason);
    }

    private static bool ReplaceDirectTextRanges(
        XElement paragraph,
        TextPosition?[] positions,
        IReadOnlyList<TextRange> matches,
        string replacement,
        [NotNullWhen(false)] out string? unsupportedReason)
    {
        unsupportedReason = null;
        // 9 and 10 encode tab and line feed; quote bytes stay out of this file.
        if (replacement.Contains((char)9) || replacement.Contains((char)10))
        {
            unsupportedReason = "replacement contains tabs or line breaks";
            return false;
        }

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
            TextPosition? start = positions[match.Start];
            TextPosition? end = positions[match.Start + match.Length - 1];
            if (start is null || end is null)
            {
                unsupportedReason = "match includes tabs, line breaks, or non-text run content";
                return false;
            }
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

        foreach (XNode node in CreateTextNodes(text))
        {
            run.Add(node);
        }
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

    // Execution-scoped revision-ID state. The per-operation generated-ID lists flow
    // through one shared instance, so repeated allocations reuse a single package scan.
    // The counter only moves forward; package deletions can only lower the stored
    // maximum, and cloned nodes copy existing IDs. Comment and bookmark allocations
    // write w:id values outside this tracker, so the engine invalidates it after the
    // operations that allocate them (see WordIdAllocatingOperations); the next
    // allocation then rescans. Per-operation reports still slice their own ID ranges
    // out of the shared list.
    private sealed class RevisionIdTracker : List<string>
    {
        public int NextId { get; set; } = 1;

        public bool NeedsScan { get; set; } = true;

        public void Invalidate()
        {
            NeedsScan = true;
        }
    }

    private static string[] AllocateRevisionIds(
        OoxmlPackage package,
        int count,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        RevisionIdTracker? tracker = generatedRevisionIds as RevisionIdTracker;
        int nextId = tracker is not null && !tracker.NeedsScan ? tracker.NextId : FindNextRevisionId(package, cancellationToken);
        foreach (string generatedRevisionId in generatedRevisionIds)
        {
            if (int.TryParse(generatedRevisionId, out int id) && id >= nextId)
            {
                nextId = id + 1;
            }
        }

        string[] ids = BuildRevisionIds(nextId, count);
        generatedRevisionIds.AddRange(ids);
        if (tracker is not null)
        {
            tracker.NextId = nextId + count;
            tracker.NeedsScan = false;
        }

        return ids;
    }
    private static int FindNextRevisionId(
        OoxmlPackage package,
        CancellationToken cancellationToken)
    {
        int nextId = 1;
        foreach (string wordPartName in DocxPartRoles.GetWordProcessingParts(package, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            OoxmlPart? part = package.GetPart(wordPartName);
            if (part is null)
            {
                continue;
            }

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
            return [new DocxDiagnostic(DocxSeverity.Error, "E9001", $"Post-edit validation failed for {settingsPartName}: Expected root element 'settings', found '{settings.Name.LocalName}'.") with { PartName = settingsPartName }];
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
            .GetRelationships(package.MainDocumentPartName, cancellationToken)
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
        string? contentType = part.ContentType;
        bool isSupportedImage = string.Equals(contentType, "image/png", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(contentType, "image/jpeg", StringComparison.OrdinalIgnoreCase);
        if (!isSupportedImage || contentType is null)
        {
            return;
        }

        if (part.Bytes.Length == 0)
        {
            diagnostics.Add(PostEditValidationDiagnostic(part.Name, "Image part is empty."));
            return;
        }

        // Defense in depth: assets were structurally validated before embedding,
        // so this only trips on corruption introduced after that gate.
        if (!TryValidateImageStructure(part.Bytes, contentType.ToLowerInvariant(), out _))
        {
            diagnostics.Add(PostEditValidationDiagnostic(part.Name, $"Image part is not a structurally valid {contentType} image."));
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
        string sourcePartName = OoxmlPath.GetSourcePartNameFromRelationshipPartName(relationshipPart.Name);
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

    private static DocxDiagnostic PostEditValidationDiagnostic(string partName, string message)
    {
        return new DocxDiagnostic(DocxSeverity.Error, "E9001", $"Post-edit validation failed for {partName}: {message}") with { PartName = partName };
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

    // D16: semantic no-ops succeed with this informational code instead of
    // writing unchanged bytes or fabricating revision history. Execution uses
    // it to tell mutating operations apart for field-refresh decisions and
    // affected-target reporting.
    internal const string NoOpDiagnosticCode = "I0001";
    internal const string SkippedDiagnosticCode = "I0002";

    private static IReadOnlyList<DocxDiagnostic> NoOpResult(DocxPatchOperation operation, string? target, string message)
    {
        return [Diagnostic(DocxSeverity.Info, NoOpDiagnosticCode, message, operation, target)];
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
        string? targetId,
        string? fieldName = null)
    {
        return Diagnostic(severity, code, message, operation, targetId, feature: null, fallback: null, fieldName: fieldName);
    }

    private static DocxDiagnostic Diagnostic(
        DocxSeverity severity,
        string code,
        string message,
        DocxPatchOperation operation,
        string? targetId,
        string? feature,
        string? fallback,
        string? fieldName = null)
    {
        string? locationField = fieldName ?? (targetId is null ? null : "target");
        DocxPatchField? field = locationField is null
            ? null
            : operation.FieldValues.FirstOrDefault(candidate => string.Equals(candidate.Name, locationField, StringComparison.Ordinal));
        return new DocxDiagnostic(severity, code, message) with
        {
            TargetId = targetId,
            Feature = feature,
            Fallback = fallback,
            HelpTopic = DocxHelp.TryGetPatchOperation(operation.OperationName, out _) ? operation.OperationName : null,
            OperationIndex = operation.Index,
            Line = field?.Line,
            Column = field?.Column
        };
    }
}
