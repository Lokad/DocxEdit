using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

internal static partial class DocxPatchEngine
{
    private static IReadOnlyList<DocxDiagnostic> ExecuteAddComment(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        string? anchorText = operation.Fields.GetValueOrDefault("anchor-text");
        int? occurrence = ReadPositiveOccurrence(operation, diagnostics);
        string author = operation.Fields.GetValueOrDefault("author") ?? options.Author;
        string? initials = operation.Fields.GetValueOrDefault("initials");
        DateTimeOffset timestampUtc = options.TimestampUtc.ToUniversalTime();
        if (operation.Fields.TryGetValue("date", out string? date))
        {
            if (!DateTimeOffset.TryParse(date, null, System.Globalization.DateTimeStyles.RoundtripKind, out DateTimeOffset parsedDate))
            {
                diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", $"Field 'date' must be an ISO-8601 timestamp: {date}.", operation, target));
            }
            else
            {
                timestampUtc = parsedDate.ToUniversalTime();
            }
        }

        if (string.IsNullOrWhiteSpace(author))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Field 'author' must not be empty.", operation, target));
        }

        if (operation.Fields.ContainsKey("occurrence") && anchorText is null)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Field 'occurrence' requires field 'anchor-text'.", operation, target));
        }

        if (anchorText is not null && anchorText.Length == 0)
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Field 'anchor-text' must not be empty.", operation, target));
        }

        if (target is null || text is null || diagnostics.Count != 0)
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
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 paragraph targets: {target}.", operation, target)];
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

        TextRange? anchorRange = null;
        if (anchorText is not null)
        {
            IReadOnlyList<TextRange> matches = FindTextMatches(current, anchorText, occurrence);
            if (matches.Count == 0)
            {
                return [Diagnostic(DocxSeverity.Error, "E4203", $"Anchor text was not found in {target}.", operation, target, fieldName: "anchor-text")];
            }

            if (occurrence is null && matches.Count > 1)
            {
                return
                [
                    Diagnostic(
                        DocxSeverity.Error,
                        "E1202",
                        $"Anchor text matched {matches.Count} ranges in {target}. Specify occurrence to select one range.",
                        operation,
                        target) with { MatchCount = matches.Count }
                ];
            }

            if (TryGetProtectedTextEditFeature(paragraphTarget.Paragraph, out string protectedFeature))
            {
                return [Diagnostic(DocxSeverity.Error, "E4305", $"Selected comment range for {target} crosses protected OOXML boundary '{protectedFeature}'.", operation, target)];
            }

            anchorRange = matches[0];
            if (!CanAddSelectedCommentAnchor(paragraphTarget.Paragraph, anchorRange.Value, out string? unsupportedReason))
            {
                return [Diagnostic(DocxSeverity.Error, "E4317", $"Selected comment range is not supported for {target}: {unsupportedReason}.", operation, target)];
            }
        }

        DocxDiagnostic? commentsPartDiagnostic = ValidateExistingCommentsPart(package, cancellationToken);
        if (commentsPartDiagnostic is not null)
        {
            return [commentsPartDiagnostic];
        }


        CommentsPartTarget commentsPart = ResolveOrCreateCommentsPart(package, cancellationToken);
        string commentId = AllocateCommentId(package, cancellationToken);
        XElement createdComment = CreateComment(commentId, text, author, initials, timestampUtc);
        createdComment.SetAttributeValue(SnapshotCreatedName, CreatedMarkValue(operation, 0));
        commentsPart.Root.Add(createdComment);
        if (anchorRange is TextRange selectedRange)
        {
            AddSelectedCommentAnchor(paragraphTarget.Paragraph, commentId, selectedRange);
        }
        else
        {
            AddCommentAnchor(paragraphTarget.Paragraph, commentId);
        }

        SaveDocumentPart(package, commentsPart.PartName, commentsPart.Document);
        SaveDocumentPart(package, paragraphTarget.PartName, paragraphTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetCommentText(
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

        CommentTarget? commentTarget = ResolveCommentTarget(package, target, operation, cancellationToken, out DocxDiagnostic? diagnostic);
        if (diagnostic is not null)
        {
            return [diagnostic];
        }

        if (commentTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 comments: {target}.", operation, target)];
        }

        bool useTrackedChanges = IsTrackedMode(options);
        XElement[]? trackedParagraphs = null;

        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        string current = ReadVisibleText(commentTarget.Comment);
        if (expected is not null && !string.Equals(current, expected, StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected text does not match current text.", operation, target, fieldName: "expect-text")];
        }
        if (useTrackedChanges &&
            !TryGetTrackedCommentParagraphs(commentTarget.Comment, text, out trackedParagraphs, out string? trackedUnsupportedReason))
        {
            if (!TryFallbackToDirectEdit(options, operation, target, trackedUnsupportedReason, diagnostics, ref useTrackedChanges))
            {
                return diagnostics;
            }
        }


        if (useTrackedChanges && trackedParagraphs is not null)
        {
            ReplaceCellParagraphTextWithTrackedChanges(package, trackedParagraphs, text, options, generatedRevisionIds, cancellationToken);
        }
        else
        {
            ReplaceCommentText(commentTarget.Comment, text);
        }

        SaveDocumentPart(package, commentTarget.PartName, commentTarget.Document);
        return diagnostics;
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteSetCommentResolved(
        OoxmlPackage package,
        DocxPatchOperation operation,
        bool resolved,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        if (target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        CommentTarget? commentTarget = ResolveCommentTarget(package, target, operation, cancellationToken, out DocxDiagnostic? diagnostic);
        if (diagnostic is not null)
        {
            return [diagnostic];
        }

        if (commentTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 comments: {target}.", operation, target)];
        }

        DocxDiagnostic? commentsExtendedDiagnostic = ValidateExistingCommentsExtendedPart(package, cancellationToken);
        if (commentsExtendedDiagnostic is not null)
        {
            return [commentsExtendedDiagnostic];
        }


        CommentExtensionTarget? extensionTarget = ResolveOrCreateCommentExtensionTarget(
            package,
            commentTarget,
            operation,
            target,
            cancellationToken,
            out diagnostic,
            out bool commentDocumentChanged);
        if (diagnostic is not null)
        {
            return [diagnostic];
        }

        if (extensionTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E4312", $"Comment '{target}' has no commentsExtended resolution metadata.", operation, target)];
        }

        extensionTarget.CommentExtension.SetAttributeValue(OoxmlNs.W15 + "done", resolved ? "1" : "0");
        if (commentDocumentChanged)
        {
            SaveDocumentPart(package, commentTarget.PartName, commentTarget.Document);
        }

        SaveDocumentPart(package, extensionTarget.PartName, extensionTarget.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteDeleteComment(
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

        CommentTarget? commentTarget = ResolveCommentTarget(package, target, operation, cancellationToken, out DocxDiagnostic? diagnostic);
        if (diagnostic is not null)
        {
            return [diagnostic];
        }

        if (commentTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 comments: {target}.", operation, target)];
        }

        string? expected = operation.Fields.GetValueOrDefault("expect-text");
        string current = ReadVisibleText(commentTarget.Comment);
        if (expected is not null && !string.Equals(current, expected, StringComparison.Ordinal))
        {
            return [Diagnostic(DocxSeverity.Error, "E3201", $"Guard failed for {target}. Expected text does not match current text.", operation, target, fieldName: "expect-text")];
        }

        string commentId = (string?)commentTarget.Comment.Attribute(OoxmlNs.W + "id") ?? string.Empty;
        RemoveCommentExtensionRecords(package, commentTarget.Comment, cancellationToken);
        RemoveCommentIdRecords(package, commentTarget.Comment, cancellationToken);
        commentTarget.Comment.Remove();
        SaveDocumentPart(package, commentTarget.PartName, commentTarget.Document);
        RemoveCommentAnchors(package, commentId, cancellationToken);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteAddCommentReply(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxDiagnostic>();
        string? target = ReadRequiredField(operation, "target", diagnostics);
        string? text = ReadRequiredField(operation, "text", diagnostics);
        string author = operation.Fields.GetValueOrDefault("author") ?? options.Author;
        string? initials = operation.Fields.GetValueOrDefault("initials");
        DateTimeOffset timestampUtc = options.TimestampUtc.ToUniversalTime();
        if (operation.Fields.TryGetValue("date", out string? date))
        {
            if (!DateTimeOffset.TryParse(date, null, System.Globalization.DateTimeStyles.RoundtripKind, out DateTimeOffset parsedDate))
            {
                diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", $"Field 'date' must be an ISO-8601 timestamp: {date}.", operation, target));
            }
            else
            {
                timestampUtc = parsedDate.ToUniversalTime();
            }
        }

        if (string.IsNullOrWhiteSpace(author))
        {
            diagnostics.Add(Diagnostic(DocxSeverity.Error, "E4205", "Field 'author' must not be empty.", operation, target));
        }

        if (text is null || target is null || diagnostics.Count != 0)
        {
            return diagnostics;
        }

        CommentTarget? parentTarget = ResolveCommentTarget(package, target, operation, cancellationToken, out DocxDiagnostic? diagnostic);
        if (diagnostic is not null)
        {
            return [diagnostic];
        }

        if (parentTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 comments: {target}.", operation, target)];
        }

        if (parentTarget.Comment.Elements(OoxmlNs.W + "p").FirstOrDefault() is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E4314", $"Comment '{target}' has no body paragraph for reply metadata.", operation, target)];
        }

        DocxDiagnostic? commentsPartDiagnostic = ValidateExistingCommentsPart(package, cancellationToken);
        if (commentsPartDiagnostic is not null)
        {
            return [commentsPartDiagnostic];
        }

        DocxDiagnostic? commentsExtendedDiagnostic = ValidateExistingCommentsExtendedPart(package, cancellationToken);
        if (commentsExtendedDiagnostic is not null)
        {
            return [commentsExtendedDiagnostic];
        }

        DocxDiagnostic? commentsIdsDiagnostic = ValidateExistingCommentsIdsPart(package, cancellationToken);
        if (commentsIdsDiagnostic is not null)
        {
            return [commentsIdsDiagnostic];
        }


        CommentExtensionTarget? parentExtensionTarget = ResolveOrCreateCommentExtensionTarget(
            package,
            parentTarget,
            operation,
            target,
            cancellationToken,
            out diagnostic,
            out _);
        if (diagnostic is not null)
        {
            return [diagnostic];
        }

        if (parentExtensionTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E4314", $"Comment '{target}' cannot host reply metadata.", operation, target)];
        }

        string parentParaId = ReadCommentParaId(parentTarget.Comment)
            ?? throw new InvalidDataException("Parent comment paraId was not created.");
        string replyParaId = AllocateCommentParaId(package, [parentParaId], cancellationToken);
        string replyCommentId = AllocateCommentId(package, cancellationToken);
        EnsureNamespaceDeclaration(parentTarget.Document.Root, "w15", OoxmlNs.W15);
        XElement createdReply = CreateComment(replyCommentId, text, author, initials, timestampUtc, replyParaId);
        createdReply.SetAttributeValue(SnapshotCreatedName, CreatedMarkValue(operation, 0));
        parentTarget.Comment.AddAfterSelf(createdReply);

        XElement parentExtensionRoot = parentExtensionTarget?.Document.Root
            ?? throw new InvalidDataException("commentsExtended document has no XML root.");
        EnsureNamespaceDeclaration(parentExtensionRoot, "w15", OoxmlNs.W15);
        parentExtensionRoot.Add(new XElement(
            OoxmlNs.W15 + "commentEx",
            new XAttribute(OoxmlNs.W15 + "paraId", replyParaId),
            new XAttribute(OoxmlNs.W15 + "paraIdParent", parentParaId),
            new XAttribute(OoxmlNs.W15 + "done", "0")));

        CommentsIdsPartTarget commentsIdsPart = ResolveOrCreateCommentsIdsPart(package, cancellationToken);
        EnsureNamespaceDeclaration(commentsIdsPart.Root, "w16cid", OoxmlNs.W16Cid);
        commentsIdsPart.Root.Add(new XElement(
            OoxmlNs.W16Cid + "commentId",
            new XAttribute(OoxmlNs.W16Cid + "paraId", replyParaId),
            new XAttribute(OoxmlNs.W16Cid + "durableId", AllocateCommentDurableId(package, cancellationToken))));

        SaveDocumentPart(package, parentTarget.PartName, parentTarget.Document);
        SaveDocumentPart(package, parentExtensionTarget.PartName, parentExtensionTarget.Document);
        SaveDocumentPart(package, commentsIdsPart.PartName, commentsIdsPart.Document);
        return [];
    }

    private static IReadOnlyList<DocxDiagnostic> ExecuteDeleteCommentReply(
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

        CommentReplyTarget? replyTarget = ResolveCommentReplyTarget(package, target, operation, cancellationToken, out DocxDiagnostic? diagnostic);
        if (diagnostic is not null)
        {
            return [diagnostic];
        }

        if (replyTarget is null)
        {
            return [Diagnostic(DocxSeverity.Error, "E1201", $"Selector matched 0 comment replies: {target}.", operation, target)];
        }

        if (HasChildCommentReplies(package, replyTarget.ParaId, cancellationToken))
        {
            return [Diagnostic(DocxSeverity.Error, "E4314", $"Comment reply '{target}' has child replies and cannot be deleted without changing thread topology.", operation, target)];
        }


        string replyCommentId = (string?)replyTarget.Comment.Attribute(OoxmlNs.W + "id") ?? string.Empty;
        RemoveCommentExtensionRecords(package, replyTarget.Comment, cancellationToken);
        RemoveCommentIdRecords(package, replyTarget.Comment, cancellationToken);
        replyTarget.Comment.Remove();
        SaveDocumentPart(package, replyTarget.PartName, replyTarget.Document);
        RemoveCommentAnchors(package, replyCommentId, cancellationToken);
        return [];

        static bool HasChildCommentReplies(
            OoxmlPackage package,
            string paraId,
            CancellationToken cancellationToken)
        {
            foreach (string partName in GetCommentsExtendedPartNames(package, cancellationToken))
            {
                OoxmlPart? part = package.GetPart(partName);
                if (part is null)
                {
                    continue;
                }

                using Stream stream = part.OpenRead();
                XDocument document = SafeXml.Load(stream, cancellationToken);
                if (document
                    .Descendants(OoxmlNs.W15 + "commentEx")
                    .Any(element => string.Equals((string?)element.Attribute(OoxmlNs.W15 + "paraIdParent"), paraId, StringComparison.Ordinal)))
                {
                    return true;
                }
            }

            return false;
        }
    }
    private static CommentTarget? ResolveCommentTarget(
        OoxmlPackage package,
        string target,
        DocxPatchOperation operation,
        CancellationToken cancellationToken,
        out DocxDiagnostic? diagnostic)
    {
        diagnostic = null;
        if (IsAliasReference(target))
        {
            return ResolveAliasCommentTarget(package, operation, target, cancellationToken, out diagnostic);
        }

        if (target.StartsWith("comment:", StringComparison.Ordinal))
        {
            string commentId = target["comment:".Length..].Trim();
            if (commentId.Length == 0)
            {
                diagnostic = Diagnostic(DocxSeverity.Error, "E1203", "Comment target must be comment:<id> or C001.C0001.", operation, target);
                return null;
            }

            return FindCommentTargetById(package, commentId, cancellationToken);
        }

        if (TryParseCommentBodyTarget(target, out int storyOrdinal, out int commentOrdinal))
        {
            string? partName = GetCommentsPartNames(package, cancellationToken).ElementAtOrDefault(storyOrdinal - 1);
            return partName is null ? null : FindCommentTargetByOrdinal(package, partName, storyOrdinal, commentOrdinal, cancellationToken);
        }

        diagnostic = Diagnostic(DocxSeverity.Error, "E1203", "Comment target must be comment:<id> or C001.C0001.", operation, target);
        return null;
    }

    private static CommentReplyTarget? ResolveCommentReplyTarget(
        OoxmlPackage package,
        string target,
        DocxPatchOperation operation,
        CancellationToken cancellationToken,
        out DocxDiagnostic? diagnostic)
    {
        diagnostic = null;
        if (TryParseCommentReplyTarget(target, out string? parentCommentId, out int replyOrdinal))
        {
            CommentTarget? parentTarget = FindCommentTargetById(package, parentCommentId, cancellationToken);
            if (parentTarget is null)
            {
                return null;
            }

            string? parentReplyLookupParaId = ReadCommentParaId(parentTarget.Comment);
            if (string.IsNullOrWhiteSpace(parentReplyLookupParaId))
            {
                diagnostic = Diagnostic(DocxSeverity.Error, "E4314", $"Comment '{target}' parent has no w15:paraId for threaded reply lookup.", operation, target);
                return null;
            }

            return FindCommentReplyTargetByOrdinal(package, parentTarget.PartName, parentReplyLookupParaId, replyOrdinal, cancellationToken);
        }

        CommentTarget? commentTarget = ResolveCommentTarget(package, target, operation, cancellationToken, out diagnostic);
        if (diagnostic is not null || commentTarget is null)
        {
            return null;
        }

        string? paraId = ReadCommentParaId(commentTarget.Comment);
        if (string.IsNullOrWhiteSpace(paraId))
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E4314", $"Comment target '{target}' is not a threaded reply because it has no w15:paraId.", operation, target);
            return null;
        }

        string? parentParaId = ReadCommentParentParaId(package, paraId, cancellationToken);
        if (string.IsNullOrWhiteSpace(parentParaId))
        {
            diagnostic = Diagnostic(DocxSeverity.Error, "E4314", $"Comment target '{target}' is not a threaded reply.", operation, target);
            return null;
        }

        return new CommentReplyTarget(commentTarget.PartName, commentTarget.Document, commentTarget.Comment, paraId, parentParaId);
    }

    private static bool TryParseCommentReplyTarget(string target, out string parentCommentId, out int replyOrdinal)
    {
        parentCommentId = string.Empty;
        replyOrdinal = 0;
        if (!target.StartsWith("comment:", StringComparison.Ordinal))
        {
            return false;
        }

        int markerIndex = target.IndexOf(".reply:", StringComparison.Ordinal);
        if (markerIndex < 0)
        {
            return false;
        }

        parentCommentId = target["comment:".Length..markerIndex];
        string ordinalText = target[(markerIndex + ".reply:".Length)..];
        return parentCommentId.Length != 0 &&
            int.TryParse(ordinalText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out replyOrdinal) &&
            replyOrdinal > 0;
    }

    private static CommentReplyTarget? FindCommentReplyTargetByOrdinal(
        OoxmlPackage package,
        string partName,
        string parentParaId,
        int replyOrdinal,
        CancellationToken cancellationToken)
    {
        XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
        XElement? reply = document
            .Descendants(OoxmlNs.W + "comment")
            .Where(comment =>
            {
                string? paraId = ReadCommentParaId(comment);
                return paraId is not null &&
                    string.Equals(ReadCommentParentParaId(package, paraId, cancellationToken), parentParaId, StringComparison.Ordinal);
            })
            .ElementAtOrDefault(replyOrdinal - 1);
        if (reply is null)
        {
            return null;
        }

        string replyParaId = ReadCommentParaId(reply)
            ?? throw new InvalidDataException("Reply comment target has no paraId.");
        return new CommentReplyTarget(partName, document, reply, replyParaId, parentParaId);
    }

    private static string? ReadCommentParentParaId(
        OoxmlPackage package,
        string paraId,
        CancellationToken cancellationToken)
    {
        foreach (string partName in GetCommentsExtendedPartNames(package, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            XElement? extension = document
                .Descendants(OoxmlNs.W15 + "commentEx")
                .FirstOrDefault(element => string.Equals((string?)element.Attribute(OoxmlNs.W15 + "paraId"), paraId, StringComparison.Ordinal));
            if (extension is not null)
            {
                return (string?)extension.Attribute(OoxmlNs.W15 + "paraIdParent");
            }
        }

        return null;
    }

    private static bool TryParseCommentBodyTarget(string target, out int storyOrdinal, out int commentOrdinal)
    {
        storyOrdinal = 0;
        commentOrdinal = 0;
        if (target.Length != 10 ||
            target[0] != 'C' ||
            target[4..6] != ".C")
        {
            return false;
        }

        return int.TryParse(target[1..4], out storyOrdinal) &&
            int.TryParse(target[6..], out commentOrdinal) &&
            storyOrdinal > 0 &&
            commentOrdinal > 0;
    }

    private static CommentTarget? FindCommentTargetById(
        OoxmlPackage package,
        string commentId,
        CancellationToken cancellationToken)
    {
        foreach (string partName in GetCommentsPartNames(package, cancellationToken))
        {
            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
            XElement? comment = document
                .Descendants(OoxmlNs.W + "comment")
                .FirstOrDefault(comment => string.Equals((string?)comment.Attribute(OoxmlNs.W + "id"), commentId, StringComparison.Ordinal));
            if (comment is not null)
            {
                return new CommentTarget(partName, document, comment);
            }
        }

        return null;
    }

    private static CommentTarget? FindCommentTargetByOrdinal(
        OoxmlPackage package,
        string partName,
        int storyOrdinal,
        int commentOrdinal,
        CancellationToken cancellationToken)
    {
        if (commentOrdinal < 1)
        {
            return null;
        }

        XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
        string wireId = "C" + storyOrdinal.ToString("D3", System.Globalization.CultureInfo.InvariantCulture) + ".C" + commentOrdinal.ToString("D4", System.Globalization.CultureInfo.InvariantCulture);
        XElement? comment = FindSnapshotElement(document, OoxmlNs.W + "comment", wireId);
        return comment is null ? null : new CommentTarget(partName, document, comment);
    }

    private static IReadOnlyList<string> GetCommentsPartNames(OoxmlPackage package, CancellationToken cancellationToken)
    {

        return package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => relationship.Type == OoxmlRelTypes.Comments)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .Select(relationship => relationship.ResolvedTarget)
            .ToArray();
    }

    private static IReadOnlyList<string> GetCommentsExtendedPartNames(OoxmlPackage package, CancellationToken cancellationToken)
    {

        var partNames = package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => relationship.Type == OoxmlRelTypes.CommentsExtended)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .Select(relationship => relationship.ResolvedTarget)
            .ToList();
        return partNames;
    }

    private static IReadOnlyList<string> GetCommentsIdsPartNames(OoxmlPackage package, CancellationToken cancellationToken)
    {

        var partNames = package
            .GetResolvedRelationships(package.MainDocumentPartName, cancellationToken)
            .Where(relationship => relationship.Type == OoxmlRelTypes.CommentsIds)
            .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
            .Select(relationship => relationship.ResolvedTarget)
            .ToList();
        return partNames;
    }

    private static DocxDiagnostic? ValidateExistingCommentsPart(OoxmlPackage package, CancellationToken cancellationToken)
    {
        foreach (string partName in GetCommentsPartNames(package, cancellationToken))
        {
            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out XElement root);
            if (root.Name != OoxmlNs.W + "comments")
            {
                return new DocxDiagnostic(DocxSeverity.Error, "E9001", $"Comments part '{partName}' has root '{root.Name.LocalName}', expected 'comments'.") with { PartName = partName };
            }
        }

        return null;
    }

    private static DocxDiagnostic? ValidateExistingCommentsExtendedPart(OoxmlPackage package, CancellationToken cancellationToken)
    {
        foreach (string partName in GetCommentsExtendedPartNames(package, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out XElement root);
            if (root.Name != OoxmlNs.W15 + "commentsEx")
            {
                return new DocxDiagnostic(DocxSeverity.Error, "E9001", $"commentsExtended part '{partName}' has root '{root.Name.LocalName}', expected 'commentsEx'.") with { PartName = partName };
            }
        }

        return null;
    }

    private static DocxDiagnostic? ValidateExistingCommentsIdsPart(OoxmlPackage package, CancellationToken cancellationToken)
    {
        foreach (string partName in GetCommentsIdsPartNames(package, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out XElement root);
            if (root.Name != OoxmlNs.W16Cid + "commentsIds")
            {
                return new DocxDiagnostic(DocxSeverity.Error, "E9001", $"commentsIds part '{partName}' has root '{root.Name.LocalName}', expected 'commentsIds'.") with { PartName = partName };
            }
        }

        return null;
    }

    private static CommentsPartTarget ResolveOrCreateCommentsPart(OoxmlPackage package, CancellationToken cancellationToken)
    {
        string? commentsPartName = GetCommentsPartNames(package, cancellationToken).FirstOrDefault();
        if (commentsPartName is null)
        {
            commentsPartName = "/word/comments.xml";
            if (package.GetPart(commentsPartName) is null)
            {
                byte[] bytes = Encoding.UTF8.GetBytes("""
                    <?xml version="1.0" encoding="utf-8"?>
                    <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />
                    """);
                package.AddPart(commentsPartName, CommentsContentType, bytes, cancellationToken);
            }

            string relationshipId = OoxmlIds.AllocateRelationshipId(package
                .GetRelationships(package.MainDocumentPartName, cancellationToken)
                .Select(relationship => relationship.Id));
            package.AddRelationship(package.MainDocumentPartName, relationshipId, OoxmlRelTypes.Comments, GetRelativeRelationshipTarget(package.MainDocumentPartName, commentsPartName), targetMode: null, cancellationToken);
        }
        else if (package.GetPart(commentsPartName) is null)
        {
            byte[] bytes = Encoding.UTF8.GetBytes("""
                <?xml version="1.0" encoding="utf-8"?>
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />
                """);
            package.AddPart(commentsPartName, CommentsContentType, bytes, cancellationToken);
        }

        XDocument document = LoadDocumentPart(package, commentsPartName, cancellationToken, out XElement root);
        if (root.Name != OoxmlNs.W + "comments")
        {
            throw new InvalidDataException($"Comments part '{commentsPartName}' has root '{root.Name.LocalName}', expected 'comments'.");
        }

        return new CommentsPartTarget(commentsPartName, document, root);
    }

    // Allocates a comment w:id, which shares the scanned w:id space with revisions. Callers live in
    // WordIdAllocatingOperations so later revision allocations rescan past fresh comment IDs.
    private static string AllocateCommentId(OoxmlPackage package, CancellationToken cancellationToken)
    {
        int maxId = -1;
        foreach (string partName in GetCommentsPartNames(package, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            foreach (string id in document.Descendants(OoxmlNs.W + "comment").Select(comment => (string?)comment.Attribute(OoxmlNs.W + "id")).OfType<string>())
            {
                maxId = Math.Max(maxId, ParseNonNegativeIdOrDefault(id, -1));
            }
        }

        foreach (OoxmlPart part in package.Parts.Values
            .Where(part => part.Name.StartsWith("/word/", StringComparison.OrdinalIgnoreCase) &&
                part.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(part => part.Name, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            foreach (string id in document
                .Descendants()
                .Where(element => element.Name == OoxmlNs.W + "commentRangeStart" ||
                    element.Name == OoxmlNs.W + "commentRangeEnd" ||
                    element.Name == OoxmlNs.W + "commentReference")
                .Select(element => (string?)element.Attribute(OoxmlNs.W + "id"))
                .OfType<string>())
            {
                maxId = Math.Max(maxId, ParseNonNegativeIdOrDefault(id, -1));
            }
        }

        return (maxId + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    private static int ParseNonNegativeIdOrDefault(string value, int fallback)
    {
        return int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int id) && id >= 0
            ? id
            : fallback;
    }

    private static CommentExtensionTarget? ResolveOrCreateCommentExtensionTarget(
        OoxmlPackage package,
        CommentTarget commentTarget,
        DocxPatchOperation operation,
        string target,
        CancellationToken cancellationToken,
        out DocxDiagnostic? diagnostic,
        out bool commentDocumentChanged)
    {
        diagnostic = null;
        commentDocumentChanged = false;
        string? paraId = ReadCommentParaId(commentTarget.Comment);
        if (string.IsNullOrWhiteSpace(paraId))
        {
            XElement? firstParagraph = commentTarget.Comment.Elements(OoxmlNs.W + "p").FirstOrDefault();
            if (firstParagraph is null)
            {
                diagnostic = Diagnostic(DocxSeverity.Error, "E4312", $"Comment '{target}' has no body paragraph for resolution metadata.", operation, target);
                return null;
            }

            paraId = AllocateCommentParaId(package, cancellationToken);
            EnsureNamespaceDeclaration(commentTarget.Document.Root, "w15", OoxmlNs.W15);
            firstParagraph.SetAttributeValue(OoxmlNs.W15 + "paraId", paraId);
            commentDocumentChanged = true;
        }

        foreach (string partName in GetCommentsExtendedPartNames(package, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
            XElement? commentExtension = document
                .Descendants(OoxmlNs.W15 + "commentEx")
                .FirstOrDefault(element => string.Equals((string?)element.Attribute(OoxmlNs.W15 + "paraId"), paraId, StringComparison.Ordinal));
            if (commentExtension is not null)
            {
                return new CommentExtensionTarget(partName, document, commentExtension);
            }
        }

        CommentsExtendedPartTarget extensionPart = ResolveOrCreateCommentsExtendedPart(package, cancellationToken);
        var extension = new XElement(OoxmlNs.W15 + "commentEx", new XAttribute(OoxmlNs.W15 + "paraId", paraId));
        extensionPart.Root.Add(extension);
        return new CommentExtensionTarget(extensionPart.PartName, extensionPart.Document, extension);
    }

    private static CommentsExtendedPartTarget ResolveOrCreateCommentsExtendedPart(OoxmlPackage package, CancellationToken cancellationToken)
    {
        string? commentsExtendedPartName = GetCommentsExtendedPartNames(package, cancellationToken).FirstOrDefault();
        if (commentsExtendedPartName is null)
        {
            commentsExtendedPartName = "/word/commentsExtended.xml";
            if (package.GetPart(commentsExtendedPartName) is null)
            {
                byte[] bytes = Encoding.UTF8.GetBytes("""
                    <?xml version="1.0" encoding="utf-8"?>
                    <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" />
                    """);
                package.AddPart(commentsExtendedPartName, CommentsExtendedContentType, bytes, cancellationToken);
            }

            string relationshipId = OoxmlIds.AllocateRelationshipId(package
                .GetRelationships(package.MainDocumentPartName, cancellationToken)
                .Select(relationship => relationship.Id));
            package.AddRelationship(package.MainDocumentPartName, relationshipId, OoxmlRelTypes.CommentsExtended, GetRelativeRelationshipTarget(package.MainDocumentPartName, commentsExtendedPartName), targetMode: null, cancellationToken);
        }
        else if (package.GetPart(commentsExtendedPartName) is null)
        {
            byte[] bytes = Encoding.UTF8.GetBytes("""
                <?xml version="1.0" encoding="utf-8"?>
                <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" />
                """);
            package.AddPart(commentsExtendedPartName, CommentsExtendedContentType, bytes, cancellationToken);
        }

        XDocument document = LoadDocumentPart(package, commentsExtendedPartName, cancellationToken, out XElement root);
        if (root.Name != OoxmlNs.W15 + "commentsEx")
        {
            throw new InvalidDataException($"commentsExtended part '{commentsExtendedPartName}' has root '{root.Name.LocalName}', expected 'commentsEx'.");
        }

        EnsureNamespaceDeclaration(root, "w15", OoxmlNs.W15);
        return new CommentsExtendedPartTarget(commentsExtendedPartName, document, root);
    }

    private static CommentsIdsPartTarget ResolveOrCreateCommentsIdsPart(OoxmlPackage package, CancellationToken cancellationToken)
    {
        string? commentsIdsPartName = GetCommentsIdsPartNames(package, cancellationToken).FirstOrDefault();
        if (commentsIdsPartName is null)
        {
            commentsIdsPartName = "/word/commentsIds.xml";
            if (package.GetPart(commentsIdsPartName) is null)
            {
                byte[] bytes = Encoding.UTF8.GetBytes("""
                    <?xml version="1.0" encoding="utf-8"?>
                    <w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" />
                    """);
                package.AddPart(commentsIdsPartName, CommentsIdsContentType, bytes, cancellationToken);
            }

            string relationshipId = OoxmlIds.AllocateRelationshipId(package
                .GetRelationships(package.MainDocumentPartName, cancellationToken)
                .Select(relationship => relationship.Id));
            package.AddRelationship(package.MainDocumentPartName, relationshipId, OoxmlRelTypes.CommentsIds, GetRelativeRelationshipTarget(package.MainDocumentPartName, commentsIdsPartName), targetMode: null, cancellationToken);
        }
        else if (package.GetPart(commentsIdsPartName) is null)
        {
            byte[] bytes = Encoding.UTF8.GetBytes("""
                <?xml version="1.0" encoding="utf-8"?>
                <w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" />
                """);
            package.AddPart(commentsIdsPartName, CommentsIdsContentType, bytes, cancellationToken);
        }

        XDocument document = LoadDocumentPart(package, commentsIdsPartName, cancellationToken, out XElement root);
        if (root.Name != OoxmlNs.W16Cid + "commentsIds")
        {
            throw new InvalidDataException($"commentsIds part '{commentsIdsPartName}' has root '{root.Name.LocalName}', expected 'commentsIds'.");
        }

        EnsureNamespaceDeclaration(root, "w16cid", OoxmlNs.W16Cid);
        return new CommentsIdsPartTarget(commentsIdsPartName, document, root);
    }

    private static string AllocateCommentDurableId(OoxmlPackage package, CancellationToken cancellationToken)
    {
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string partName in GetCommentsIdsPartNames(package, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            foreach (string durableId in document
                .Descendants(OoxmlNs.W16Cid + "commentId")
                .Select(element => (string?)element.Attribute(OoxmlNs.W16Cid + "durableId"))
                .Where(durableId => !string.IsNullOrWhiteSpace(durableId))
                .OfType<string>())
            {
                used.Add(durableId);
            }
        }

        for (uint id = 1; id < uint.MaxValue; id++)
        {
            string candidate = id.ToString("X8", System.Globalization.CultureInfo.InvariantCulture);
            if (!used.Contains(candidate))
            {
                return candidate;
            }
        }

        throw new InvalidDataException("Unable to allocate a unique comment durableId.");
    }

    private static string AllocateCommentParaId(OoxmlPackage package, CancellationToken cancellationToken)
    {
        return AllocateCommentParaId(package, [], cancellationToken);
    }

    private static string AllocateCommentParaId(
        OoxmlPackage package,
        IEnumerable<string> reservedParaIds,
        CancellationToken cancellationToken)
    {
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string reservedParaId in reservedParaIds)
        {
            if (!string.IsNullOrWhiteSpace(reservedParaId))
            {
                used.Add(reservedParaId);
            }
        }

        foreach (string partName in GetCommentsPartNames(package, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            foreach (string paraId in document
                .Descendants(OoxmlNs.W + "p")
                .Select(paragraph => (string?)paragraph.Attribute(OoxmlNs.W15 + "paraId"))
                .Where(paraId => !string.IsNullOrWhiteSpace(paraId))
                .OfType<string>())
            {
                used.Add(paraId);
            }
        }

        foreach (string partName in GetCommentsExtendedPartNames(package, cancellationToken))
        {
            OoxmlPart? part = package.GetPart(partName);
            if (part is null)
            {
                continue;
            }

            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            foreach (XElement extension in document.Descendants(OoxmlNs.W15 + "commentEx"))
            {
                foreach (string paraId in new[] { (string?)extension.Attribute(OoxmlNs.W15 + "paraId"), (string?)extension.Attribute(OoxmlNs.W15 + "paraIdParent") }
                    .Where(paraId => !string.IsNullOrWhiteSpace(paraId))
                    .OfType<string>())
                {
                    used.Add(paraId);
                }
            }
        }

        for (uint id = 1; id < uint.MaxValue; id++)
        {
            string candidate = id.ToString("X8", System.Globalization.CultureInfo.InvariantCulture);
            if (!used.Contains(candidate))
            {
                return candidate;
            }
        }

        throw new InvalidDataException("Unable to allocate a unique comment paraId.");
    }

    private static void EnsureNamespaceDeclaration(XElement? root, string prefix, XNamespace ns)
    {
        if (root is not null && root.GetNamespaceOfPrefix(prefix) != ns)
        {
            root.SetAttributeValue(XNamespace.Xmlns + prefix, ns.NamespaceName);
        }
    }

    private static string? ReadCommentParaId(XElement comment)
    {
        return (string?)comment
            .Elements(OoxmlNs.W + "p")
            .FirstOrDefault()
            ?.Attribute(OoxmlNs.W15 + "paraId");
    }

    private static void ReplaceCommentText(XElement comment, string text)
    {
        comment.RemoveNodes();
        comment.Add(CreateSimpleParagraph(text, style: null, paragraphProperties: null));
    }

    private static XElement CreateComment(
        string commentId,
        string text,
        string author,
        string? initials,
        DateTimeOffset timestampUtc)
    {
        return CreateComment(commentId, text, author, initials, timestampUtc, paraId: null);
    }

    private static XElement CreateComment(
        string commentId,
        string text,
        string author,
        string? initials,
        DateTimeOffset timestampUtc,
        string? paraId)
    {
        var comment = new XElement(
            OoxmlNs.W + "comment",
            new XAttribute(OoxmlNs.W + "id", commentId),
            new XAttribute(OoxmlNs.W + "author", author),
            new XAttribute(OoxmlNs.W + "date", timestampUtc.ToUniversalTime().ToString("O")));
        if (!string.IsNullOrWhiteSpace(initials))
        {
            comment.SetAttributeValue(OoxmlNs.W + "initials", initials);
        }

        XElement paragraph = CreateSimpleParagraph(text, style: null, paragraphProperties: null);
        if (!string.IsNullOrWhiteSpace(paraId))
        {
            paragraph.SetAttributeValue(OoxmlNs.W15 + "paraId", paraId);
        }

        comment.Add(paragraph);
        return comment;
    }

    private static void AddCommentAnchor(XElement paragraph, string commentId)
    {
        var start = new XElement(OoxmlNs.W + "commentRangeStart", new XAttribute(OoxmlNs.W + "id", commentId));
        XElement? paragraphProperties = paragraph.Element(OoxmlNs.W + "pPr");
        if (paragraphProperties is null)
        {
            paragraph.AddFirst(start);
        }
        else
        {
            paragraphProperties.AddAfterSelf(start);
        }

        paragraph.Add(new XElement(OoxmlNs.W + "commentRangeEnd", new XAttribute(OoxmlNs.W + "id", commentId)));
        paragraph.Add(new XElement(
            OoxmlNs.W + "r",
            new XElement(OoxmlNs.W + "commentReference", new XAttribute(OoxmlNs.W + "id", commentId))));
    }

    private static bool CanAddSelectedCommentAnchor(XElement paragraph, TextRange range, out string? unsupportedReason)
    {
        unsupportedReason = null;
        TextPosition?[] positions = BuildTextPositions(paragraph);
        if (positions.Length != ReadVisibleText(paragraph).Length)
        {
            unsupportedReason = "selected ranges are limited to direct paragraph runs";
            return false;
        }

        if (range.Length <= 0 || range.Start < 0 || range.Start + range.Length > positions.Length)
        {
            unsupportedReason = "selected range is outside paragraph text";
            return false;
        }

        for (int i = range.Start; i < range.Start + range.Length; i++)
        {
            if (positions[i] is null)
            {
                unsupportedReason = "selected range includes tabs, line breaks, or non-text run content";
                return false;
            }
        }

        return true;
    }

    private static void AddSelectedCommentAnchor(XElement paragraph, string commentId, TextRange range)
    {
        InsertCommentBoundaryAtTextOffset(
            paragraph,
            range.Start + range.Length,
            new XElement(OoxmlNs.W + "commentRangeEnd", new XAttribute(OoxmlNs.W + "id", commentId)));
        InsertCommentBoundaryAtTextOffset(
            paragraph,
            range.Start,
            new XElement(OoxmlNs.W + "commentRangeStart", new XAttribute(OoxmlNs.W + "id", commentId)));
        paragraph.Add(new XElement(
            OoxmlNs.W + "r",
            new XElement(OoxmlNs.W + "commentReference", new XAttribute(OoxmlNs.W + "id", commentId))));
    }

    private static void InsertCommentBoundaryAtTextOffset(XElement paragraph, int textOffset, XElement boundary)
    {
        TextPosition?[] positions = BuildTextPositions(paragraph);
        if (textOffset == 0)
        {
            XElement? paragraphProperties = paragraph.Element(OoxmlNs.W + "pPr");
            if (paragraphProperties is null)
            {
                paragraph.AddFirst(boundary);
            }
            else
            {
                paragraphProperties.AddAfterSelf(boundary);
            }

            return;
        }

        if (textOffset == positions.Length)
        {
            paragraph.Add(boundary);
            return;
        }

        TextPosition position = positions[textOffset]
            ?? throw new InvalidDataException("Selected comment range boundary cannot be placed on non-text run content.");
        XElement afterRun = SplitRunAtTextPosition(position);
        afterRun.AddBeforeSelf(boundary);
    }

    private static XElement SplitRunAtTextPosition(TextPosition position)
    {
        XElement textElement = position.TextElement;
        XElement run = textElement.Parent
            ?? throw new InvalidDataException("Text element has no parent run.");
        if (run.Name != OoxmlNs.W + "r" || run.Parent?.Name != OoxmlNs.W + "p")
        {
            throw new InvalidDataException("Selected comment range boundary is not inside a direct paragraph run.");
        }

        var beforeRun = new XElement(OoxmlNs.W + "r");
        var afterRun = new XElement(OoxmlNs.W + "r");
        XElement? runProperties = run.Element(OoxmlNs.W + "rPr");
        if (runProperties is not null)
        {
            beforeRun.Add(new XElement(runProperties));
            afterRun.Add(new XElement(runProperties));
        }

        bool reachedTextElement = false;
        foreach (XElement child in run.Elements().Where(child => child.Name != OoxmlNs.W + "rPr"))
        {
            if (!ReferenceEquals(child, textElement))
            {
                (reachedTextElement ? afterRun : beforeRun).Add(new XElement(child));
                continue;
            }

            reachedTextElement = true;
            string value = textElement.Value;
            string beforeText = value[..position.Offset];
            string afterText = value[position.Offset..];
            if (beforeText.Length != 0)
            {
                XElement beforeTextElement = new XElement(textElement);
                SetTextElementValue(beforeTextElement, beforeText);
                beforeRun.Add(beforeTextElement);
            }

            if (afterText.Length != 0)
            {
                XElement afterTextElement = new XElement(textElement);
                SetTextElementValue(afterTextElement, afterText);
                afterRun.Add(afterTextElement);
            }
        }

        bool hasBeforeContent = beforeRun.Elements().Any(element => element.Name != OoxmlNs.W + "rPr");
        bool hasAfterContent = afterRun.Elements().Any(element => element.Name != OoxmlNs.W + "rPr");
        if (hasBeforeContent)
        {
            run.AddBeforeSelf(beforeRun);
        }

        if (hasAfterContent)
        {
            run.AddBeforeSelf(afterRun);
        }

        run.Remove();
        return hasAfterContent
            ? afterRun
            : beforeRun;
    }

    private static void RemoveCommentExtensionRecords(OoxmlPackage package, XElement comment, CancellationToken cancellationToken)
    {
        string? paraId = ReadCommentParaId(comment);
        if (string.IsNullOrWhiteSpace(paraId))
        {
            return;
        }

        foreach (string partName in GetCommentsExtendedPartNames(package, cancellationToken))
        {
            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
            XElement[] extensionRecords = document
                .Descendants(OoxmlNs.W15 + "commentEx")
                .Where(element => string.Equals((string?)element.Attribute(OoxmlNs.W15 + "paraId"), paraId, StringComparison.Ordinal))
                .ToArray();
            if (extensionRecords.Length == 0)
            {
                continue;
            }

            foreach (XElement extensionRecord in extensionRecords)
            {
                extensionRecord.Remove();
            }

            SaveDocumentPart(package, partName, document);
        }
    }

    private static void RemoveCommentIdRecords(OoxmlPackage package, XElement comment, CancellationToken cancellationToken)
    {
        string? paraId = ReadCommentParaId(comment);
        if (string.IsNullOrWhiteSpace(paraId))
        {
            return;
        }

        foreach (string partName in GetCommentsIdsPartNames(package, cancellationToken))
        {
            XDocument document = LoadDocumentPart(package, partName, cancellationToken, out _);
            XElement[] commentIdRecords = document
                .Descendants(OoxmlNs.W16Cid + "commentId")
                .Where(element => string.Equals((string?)element.Attribute(OoxmlNs.W16Cid + "paraId"), paraId, StringComparison.Ordinal))
                .ToArray();
            if (commentIdRecords.Length == 0)
            {
                continue;
            }

            foreach (XElement commentIdRecord in commentIdRecords)
            {
                commentIdRecord.Remove();
            }

            SaveDocumentPart(package, partName, document);
        }
    }

    private static void RemoveCommentAnchors(OoxmlPackage package, string commentId, CancellationToken cancellationToken)
    {
        string? commentsPartName = DocxPartRoles.FindCommentsPartName(package, cancellationToken);
        foreach (string wordPartName in DocxPartRoles.GetWordProcessingParts(package, cancellationToken))
        {
            if (string.Equals(wordPartName, commentsPartName, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            OoxmlPart? part = package.GetPart(wordPartName);
            if (part is null)
            {
                continue;
            }

            cancellationToken.ThrowIfCancellationRequested();
            using Stream stream = part.OpenRead();
            XDocument document = SafeXml.Load(stream, cancellationToken);
            bool changed = false;
            foreach (XElement marker in document
                .Descendants()
                .Where(element => element.Name == OoxmlNs.W + "commentRangeStart" || element.Name == OoxmlNs.W + "commentRangeEnd")
                .Where(element => string.Equals((string?)element.Attribute(OoxmlNs.W + "id"), commentId, StringComparison.Ordinal))
                .ToArray())
            {
                marker.Remove();
                changed = true;
            }

            foreach (XElement reference in document
                .Descendants(OoxmlNs.W + "commentReference")
                .Where(element => string.Equals((string?)element.Attribute(OoxmlNs.W + "id"), commentId, StringComparison.Ordinal))
                .ToArray())
            {
                XElement? run = reference.Ancestors(OoxmlNs.W + "r").FirstOrDefault();
                reference.Remove();
                if (run is not null && !run.Elements().Any() && string.IsNullOrEmpty(run.Value))
                {
                    run.Remove();
                }

                changed = true;
            }

            if (changed)
            {
                SaveDocumentPart(package, part.Name, document);
            }
        }
    }

}
