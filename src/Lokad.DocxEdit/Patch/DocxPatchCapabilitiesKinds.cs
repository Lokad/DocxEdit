using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

// D17: capabilities for content-control, cell, merge-group, bookmark, table, row, section, hyperlink, field, and image targets.
// Like the paragraph surface, every verdict reuses the execution predicates,
// so guidance cannot disagree with check. Patch-dependent details stay
// conditional until the exact patch is checked.
internal static partial class DocxPatchEngine
{
    private sealed record StoryDocument(StoryPartRef Story, string PartName, XDocument Document);

    private sealed record ContentControlCapabilityFacts(
        string Kind,
        bool IsPlainText,
        bool IsRichText,
        string? LockValue,
        bool HasContainer,
        bool HasNonParagraphContent,
        bool HasProtected,
        string ProtectedFeature,
        bool TrackedShapeOk,
        string? TrackedShapeReason);

    private sealed record CellCapabilityFacts(
        bool IsContinuation,
        string? RootCellId,
        bool IsSimple,
        bool TrackedShapeOk,
        string? TrackedShapeReason);

    internal static ParagraphCapabilitiesOutcome GetNonParagraphCapabilities(
        OoxmlPackage package,
        DocxTargetId parsed,
        string requestedTargetId,
        TrackChangesMode mode,
        CancellationToken cancellationToken)
    {
        if (parsed.Kind == DocxTargetKind.ContentControl)
        {
            return GetContentControlCapabilities(package, parsed, requestedTargetId, mode, cancellationToken);
        }

        if (parsed.Kind == DocxTargetKind.Cell)
        {
            return GetCellCapabilities(package, parsed, requestedTargetId, mode, isMergeGroup: false, cancellationToken);
        }

        if (parsed.Kind == DocxTargetKind.MergeGroup)
        {
            return GetCellCapabilities(package, parsed, requestedTargetId, mode, isMergeGroup: true, cancellationToken);
        }

        if (parsed.Kind == DocxTargetKind.Image)
        {
            return GetImageCapabilities(package, parsed, requestedTargetId, mode, cancellationToken);
        }

        if (parsed.Kind == DocxTargetKind.Field)
        {
            return GetFieldCapabilities(package, parsed, requestedTargetId, mode, cancellationToken);
        }

        if (parsed.Kind == DocxTargetKind.Hyperlink)
        {
            return GetHyperlinkCapabilities(package, parsed, requestedTargetId, mode, cancellationToken);
        }

        if (parsed.Kind == DocxTargetKind.Section)
        {
            return GetSectionCapabilities(package, parsed, requestedTargetId, mode, cancellationToken);
        }

        if (parsed.Kind == DocxTargetKind.Table)
        {
            return GetTableCapabilities(package, parsed, requestedTargetId, mode, cancellationToken);
        }

        if (parsed.Kind == DocxTargetKind.Row)
        {
            return GetRowCapabilities(package, parsed, requestedTargetId, mode, cancellationToken);
        }

        if (parsed.Kind == DocxTargetKind.Bookmark)
        {
            return GetBookmarkCapabilities(package, parsed, requestedTargetId, mode, cancellationToken);
        }

        return new ParagraphCapabilitiesOutcome
        {
            Error = new DocxDiagnostic(
                DocxSeverity.Error,
                "E1201",
                "Target " + Quote(requestedTargetId) + " was not found. Capabilities currently cover explicit paragraph, content-control, cell, merge-group, bookmark, table, row, section, hyperlink, field, and image IDs such as M.P0001.") with
            {
                TargetId = requestedTargetId
            }
        };
    }

    private static StoryDocument? TryResolveStoryDocument(
        OoxmlPackage package,
        DocxTargetId parsed,
        CancellationToken cancellationToken)
    {
        // 77 encodes M. Letter codes keep this file free of quote bytes,
        // which the authoring transport corrupts.
        string wantedPrefix = parsed.Story == (char)77
            ? "M"
            : parsed.Story.ToString() + parsed.StoryPart.ToString("000");
        StoryPartRef? story = null;
        foreach (StoryPartRef candidate in DocxPartRoles.GetOrderedStories(package, includeHeadersFooters: true, cancellationToken))
        {
            if (string.Equals(candidate.Prefix, wantedPrefix, StringComparison.Ordinal))
            {
                story = candidate;
                break;
            }
        }

        if (story is null)
        {
            return null;
        }

        OoxmlPart? part = package.GetPart(story.PartName);
        if (part is null)
        {
            return null;
        }

        using Stream stream = part.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        return document.Root is null ? null : new StoryDocument(story, story.PartName, document);
    }    private static ParagraphCapabilitiesOutcome GetContentControlCapabilities(
        OoxmlPackage package,
        DocxTargetId parsed,
        string requestedTargetId,
        TrackChangesMode mode,
        CancellationToken cancellationToken)
    {
        StoryDocument? storyDocument = TryResolveStoryDocument(package, parsed, cancellationToken);
        if (storyDocument is null)
        {
            return ParagraphCapabilitiesNotFound(requestedTargetId);
        }

        XElement? control = storyDocument.Document.Descendants(OoxmlNs.W + "sdt").ElementAtOrDefault(parsed.Primary - 1);
        if (control is null)
        {
            return ParagraphCapabilitiesNotFound(requestedTargetId);
        }

        string kind = ReadContentControlKind(control);
        bool isPlainText = IsPlainTextContentControl(control);
        bool isRichText = IsRichTextContentControl(control);
        XElement? content = control.Element(OoxmlNs.W + "sdtContent");
        bool hasNonParagraphContent = false;
        bool hasProtected = false;
        string protectedFeature = string.Empty;
        if (content is not null && isRichText)
        {
            hasNonParagraphContent = content.Elements().Any(element => element.Name != OoxmlNs.W + "p");
            hasProtected = TryGetProtectedTextEditFeature(content, out protectedFeature);
        }

        bool trackedShapeOk = true;
        string? trackedShapeReason = null;
        bool isTracked = mode is TrackChangesMode.Require or TrackChangesMode.Suggest;
        if (isTracked && content is not null && (isPlainText || isRichText))
        {
            if (isRichText)
            {
                trackedShapeOk = TryGetTrackedRichTextContentControlParagraphs(content, string.Empty, out _, out trackedShapeReason);
            }
            else
            {
                trackedShapeOk = TryGetTrackedContentControlTextContainer(content, string.Empty, out _, out _, out trackedShapeReason);
            }
        }

        var facts = new ContentControlCapabilityFacts(
            kind,
            isPlainText,
            isRichText,
            ReadContentControlLock(control),
            content is not null,
            hasNonParagraphContent,
            hasProtected,
            protectedFeature,
            trackedShapeOk,
            trackedShapeReason);
        bool isRequire = mode == TrackChangesMode.Require;
        var operations = new List<DocxOperationCapability>
        {
            ContentControlTextCapability(facts, isTracked, isRequire),
            ContentControlCheckboxCapability(facts, mode),
            ContentControlChoiceCapability(facts, mode),
            ContentControlDateCapability(facts, mode),
        };

        return new ParagraphCapabilitiesOutcome
        {
            Capabilities = new DocxTargetCapabilities(
                parsed.ToWireValue(),
                "content-control",
                storyDocument.Story.StoryLabel,
                operations)
        };
    }

    // Mirrors the lock rule in ValidateContentControlUnlocked: a missing lock
    // element or an explicit unlocked value edits freely; any other value,
    // including a bare lock element, locks the control.
    private static string? ReadContentControlLock(XElement control)
    {
        XElement? lockElement = control.Element(OoxmlNs.W + "sdtPr")?.Element(OoxmlNs.W + "lock");
        if (lockElement is null)
        {
            return null;
        }

        string lockValue = (string?)lockElement.Attribute(OoxmlNs.W + "val") ?? "locked";
        return string.Equals(lockValue, "unlocked", StringComparison.Ordinal) ? null : lockValue;
    }

    private static string ContentControlKindGuidance(string kind)
    {
        if (string.Equals(kind, "picture", StringComparison.Ordinal))
        {
            return "Picture content controls preserve a picture container; use read/media to inspect the contained image and target image operations when applicable.";
        }

        if (string.Equals(kind, "group", StringComparison.Ordinal))
        {
            return "Group content controls protect a container; target an editable child content control instead.";
        }

        if (string.Equals(kind, "repeating-section", StringComparison.Ordinal) || string.Equals(kind, "repeating-section-item", StringComparison.Ordinal))
        {
            return "Repeating-section subtree edits require cloning or deleting structured document tag subtrees, which currently fail with E4315.";
        }

        if (string.Equals(kind, "checkbox", StringComparison.Ordinal))
        {
            return "Use set-content-control-checkbox for checkbox state edits.";
        }

        if (string.Equals(kind, "dropdown-list", StringComparison.Ordinal) || string.Equals(kind, "combo-box", StringComparison.Ordinal))
        {
            return "Use set-content-control-choice for dropdown or combo-box selections.";
        }

        if (string.Equals(kind, "date", StringComparison.Ordinal))
        {
            return "Use set-content-control-date for date values.";
        }

        return "Choose an operation that matches the content-control kind.";
    }
    private static string? ContentControlKindAlternative(string kind)
    {
        if (string.Equals(kind, "checkbox", StringComparison.Ordinal))
        {
            return "set-content-control-checkbox";
        }

        if (string.Equals(kind, "dropdown-list", StringComparison.Ordinal) || string.Equals(kind, "combo-box", StringComparison.Ordinal))
        {
            return "set-content-control-choice";
        }

        if (string.Equals(kind, "date", StringComparison.Ordinal))
        {
            return "set-content-control-date";
        }

        if (string.Equals(kind, "plain-text", StringComparison.Ordinal) || string.Equals(kind, "rich-text", StringComparison.Ordinal))
        {
            return "set-content-control-text";
        }

        return null;
    }

    private static DocxOperationCapability? ContentControlLockCapability(string operation, string? lockValue)
    {
        if (lockValue is null)
        {
            return null;
        }

        return new DocxOperationCapability(
            operation,
            "unsupported",
            "Content control is locked by w:lock=" + Quote(lockValue) + ". Lock values other than unlocked reject every content-control edit with E4310.",
            operation,
            null);
    }

    private static DocxOperationCapability ContentControlTextCapability(ContentControlCapabilityFacts facts, bool isTracked, bool isRequire)
    {
        const string operation = "set-content-control-text";
        if (!facts.IsPlainText && !facts.IsRichText)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Content control is kind " + Quote(facts.Kind) + ", not a plain-text or rich-text content control. " + ContentControlKindGuidance(facts.Kind),
                operation,
                ContentControlKindAlternative(facts.Kind));
        }

        DocxOperationCapability? locked = ContentControlLockCapability(operation, facts.LockValue);
        if (locked is not null)
        {
            return locked;
        }

        if (!facts.HasContainer)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Content control has no editable content container, so every content-control edit fails with E4310.",
                operation,
                null);
        }

        if (facts.IsRichText && facts.HasNonParagraphContent)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Rich-text content control contains non-paragraph content, so text replacement fails with E4310.",
                operation,
                null);
        }

        if (facts.IsRichText && facts.HasProtected)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Rich-text content control contains protected OOXML boundary " + Quote(facts.ProtectedFeature) + ", so text replacement fails with E4310.",
                operation,
                null);
        }

        if (isTracked && !facts.TrackedShapeOk)
        {
            if (isRequire)
            {
                return new DocxOperationCapability(
                    operation,
                    "unsupported",
                    "Content-control shape needs a direct rewrite (" + facts.TrackedShapeReason + "), so Require fails with E6002. Suggest falls back to a direct rewrite with W4002 and Off rewrites directly. Check remains authoritative for the exact replacement.",
                    operation,
                    null);
            }

            return new DocxOperationCapability(
                operation,
                "conditional",
                "Content-control shape needs a direct rewrite (" + facts.TrackedShapeReason + "), so tracked replacement falls back to a direct rewrite with W4002. The exact replacement decides whether tracked markup or a direct rewrite applies. Check remains authoritative.",
                operation,
                null);
        }

        if (isTracked)
        {
            return new DocxOperationCapability(
                operation,
                "supported",
                isRequire
                    ? "Simple replacements emit tracked delete and insert markup while preserving wrappers, bindings, locks, and paragraph containers; complex shapes fail with E6002 under Require. Check remains authoritative."
                    : "Simple replacements emit tracked delete and insert markup while preserving wrappers, bindings, locks, and paragraph containers; complex shapes fall back to a direct rewrite with W4002. Check remains authoritative.",
                operation,
                null);
        }

        if (facts.IsRichText)
        {
            return new DocxOperationCapability(
                operation,
                "supported",
                "Direct rewrite replaces the control text and preserves wrappers, bindings, and locks. Rich-text replacement requires expect-text that matches the current content.",
                operation,
                null);
        }

        return new DocxOperationCapability(
            operation,
            "supported",
            "Direct rewrite replaces the control text and preserves wrappers, bindings, and locks.",
            operation,
            null);
    }
    private static DocxOperationCapability ContentControlCheckboxCapability(ContentControlCapabilityFacts facts, TrackChangesMode mode)
    {
        const string operation = "set-content-control-checkbox";
        if (!string.Equals(facts.Kind, "checkbox", StringComparison.Ordinal))
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Content control is kind " + Quote(facts.Kind) + ", not a checkbox content control. " + ContentControlKindGuidance(facts.Kind),
                operation,
                ContentControlKindAlternative(facts.Kind));
        }

        DocxOperationCapability? locked = ContentControlLockCapability(operation, facts.LockValue);
        if (locked is not null)
        {
            return locked;
        }

        if (!facts.HasContainer)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Content control has no editable content container, so every content-control edit fails with E4310.",
                operation,
                null);
        }

        if (!SupportsTrackedChangeOutput(operation) && mode == TrackChangesMode.Require && !IsAnnotationOperation(operation))
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Checkbox state updates metadata without a tracked revision representation, so Require fails with E6001 before editing. Use Suggest, which warns with W4001 and applies directly, or Off for a direct edit. Check remains authoritative.",
                operation,
                null);
        }

        return new DocxOperationCapability(
            operation,
            "supported",
            "Checkbox state and display symbol update together; the checked field is boolean.",
            operation,
            null);
    }

    private static DocxOperationCapability ContentControlChoiceCapability(ContentControlCapabilityFacts facts, TrackChangesMode mode)
    {
        const string operation = "set-content-control-choice";
        if (!string.Equals(facts.Kind, "dropdown-list", StringComparison.Ordinal) && !string.Equals(facts.Kind, "combo-box", StringComparison.Ordinal))
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Content control is kind " + Quote(facts.Kind) + ", not a dropdown or combo-box content control. " + ContentControlKindGuidance(facts.Kind),
                operation,
                ContentControlKindAlternative(facts.Kind));
        }

        DocxOperationCapability? locked = ContentControlLockCapability(operation, facts.LockValue);
        if (locked is not null)
        {
            return locked;
        }

        if (!facts.HasContainer)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Content control has no editable content container, so every content-control edit fails with E4310.",
                operation,
                null);
        }

        if (!SupportsTrackedChangeOutput(operation) && mode == TrackChangesMode.Require && !IsAnnotationOperation(operation))
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Dropdown and combo-box selections update list metadata without a tracked revision representation, so Require fails with E6001 before editing. Use Suggest, which warns with W4001 and applies directly, or Off for a direct edit. Check remains authoritative.",
                operation,
                null);
        }

        return new DocxOperationCapability(
            operation,
            "supported",
            "Exactly one of value or display-text selects a list item, and the item must exist or check fails. The item display text replaces the control content.",
            operation,
            null);
    }

    private static DocxOperationCapability ContentControlDateCapability(ContentControlCapabilityFacts facts, TrackChangesMode mode)
    {
        const string operation = "set-content-control-date";
        if (!string.Equals(facts.Kind, "date", StringComparison.Ordinal))
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Content control is kind " + Quote(facts.Kind) + ", not a date content control. " + ContentControlKindGuidance(facts.Kind),
                operation,
                ContentControlKindAlternative(facts.Kind));
        }

        DocxOperationCapability? locked = ContentControlLockCapability(operation, facts.LockValue);
        if (locked is not null)
        {
            return locked;
        }

        if (!facts.HasContainer)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Content control has no editable content container, so every content-control edit fails with E4310.",
                operation,
                null);
        }

        if (!SupportsTrackedChangeOutput(operation) && mode == TrackChangesMode.Require && !IsAnnotationOperation(operation))
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Date value updates metadata without a tracked revision representation, so Require fails with E6001 before editing. Use Suggest, which warns with W4001 and applies directly, or Off for a direct edit. Check remains authoritative.",
                operation,
                null);
        }

        return new DocxOperationCapability(
            operation,
            "supported",
            "Value sets the date metadata while display-text, or value when display-text is absent, replaces the control content.",
            operation,
            null);
    }
    // D17: section capabilities reuse the guard and tracked-revision predicates from the section engine.
    private sealed record SectionCapabilityFacts(
        int CurrentColumns,
        string CurrentOrientation,
        bool HasTrackedRevision);
    internal static ParagraphCapabilitiesOutcome GetSectionCapabilities(
        OoxmlPackage package,
        DocxTargetId parsed,
        string requestedTargetId,
        TrackChangesMode mode,
        CancellationToken cancellationToken)
    {
        if (parsed.Story != (char)77 || parsed.Primary < 1)
        {
            return ParagraphCapabilitiesNotFound(requestedTargetId);
        }
        XDocument document = LoadMainDocument(package, cancellationToken, out XElement body);
        XElement? sectionProperties = null;
        int sectionIndex = 0;
        foreach (DocxStoryBlocks.StoryBlock entry in DocxStoryBlocks.EnumeratePhysicalBlocks(body))
        {
            XElement block = entry.Block;
            if (block.Name == OoxmlNs.W + "p")
            {
                XElement? paragraphSectPr = block.Element(OoxmlNs.W + "pPr")?.Element(OoxmlNs.W + "sectPr");
                if (paragraphSectPr is not null)
                {
                    sectionIndex++;
                    if (sectionIndex == parsed.Primary)
                    {
                        sectionProperties = paragraphSectPr;
                        break;
                    }
                }
            }
            else if (block.Name == OoxmlNs.W + "sectPr")
            {
                sectionIndex++;
                if (sectionIndex == parsed.Primary)
                {
                    sectionProperties = block;
                    break;
                }
            }
        }
        if (sectionProperties is null)
        {
            return ParagraphCapabilitiesNotFound(requestedTargetId);
        }
        var facts = new SectionCapabilityFacts(
            ReadSectionColumnCount(sectionProperties),
            ReadSectionOrientation(sectionProperties).ToWireValue(),
            sectionProperties.Elements(OoxmlNs.W + "sectPrChange").Any());
        bool isTracked = mode is TrackChangesMode.Require or TrackChangesMode.Suggest;
        var operations = new List<DocxOperationCapability>
        {
            SectionColumnsCapability(facts, isTracked),
            SectionOrientationCapability(facts, isTracked),
        };
        return new ParagraphCapabilitiesOutcome
        {
            Capabilities = new DocxTargetCapabilities(
                parsed.ToWireValue(),
                "section",
                "main",
                operations)
        };
    }
    private static DocxOperationCapability SectionColumnsCapability(SectionCapabilityFacts facts, bool isTracked)
    {
        const string operation = "set-section-columns";
        if (facts.HasTrackedRevision && isTracked)
        {
            return new DocxOperationCapability(
                operation,
                "conditional",
                "Section already contains tracked section property revision markup: Suggest falls back to a direct edit with W4002 and Require fails with E6002. The count value must still be 1 through 4, otherwise check fails with E6201. Check remains authoritative.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "conditional",
            "Section currently uses " + facts.CurrentColumns + " column(s) in " + facts.CurrentOrientation + ". The count value must be 1 through 4, otherwise check fails with E6201; expect-columns and expect-orientation can guard current values. Tracked modes emit section property revisions with w:sectPrChange while preserving previous section properties and references. Check remains authoritative for the exact count.",
            operation,
            null);
    }
    private static DocxOperationCapability SectionOrientationCapability(SectionCapabilityFacts facts, bool isTracked)
    {
        const string operation = "set-section-orientation";
        if (facts.HasTrackedRevision && isTracked)
        {
            return new DocxOperationCapability(
                operation,
                "conditional",
                "Section already contains tracked section property revision markup: Suggest falls back to a direct edit with W4002 and Require fails with E6002. The orientation value must still be portrait or landscape. Check remains authoritative.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "conditional",
            "Section currently uses " + facts.CurrentOrientation + " with " + facts.CurrentColumns + " column(s). The orientation value must be portrait or landscape, otherwise check fails with E6202; expect-columns and expect-orientation can guard current values. Tracked modes emit section property revisions with w:sectPrChange while preserving page size, section properties, and references. Check remains authoritative for the exact orientation.",
            operation,
            null);
    }
    // D17: hyperlink capabilities reuse the protection and track-support predicates from the hyperlink engine.
    private sealed record HyperlinkCapabilityFacts(
        bool HasProtected,
        string ProtectedFeature,
        string DestinationKind);
    internal static ParagraphCapabilitiesOutcome GetHyperlinkCapabilities(
        OoxmlPackage package,
        DocxTargetId parsed,
        string requestedTargetId,
        TrackChangesMode mode,
        CancellationToken cancellationToken)
    {
        StoryDocument? storyDocument = TryResolveStoryDocument(package, parsed, cancellationToken);
        HyperlinkTarget? hyperlinkTarget = storyDocument is null
            ? null
            : FindHyperlinkTarget(package, storyDocument.PartName, parsed.Story, parsed.StoryPart, parsed.Primary, cancellationToken, allowLiveFallback: true);
        if (storyDocument is null || hyperlinkTarget is null)
        {
            return ParagraphCapabilitiesNotFound(requestedTargetId);
        }
        bool hasProtected = TryGetProtectedTextEditFeature(hyperlinkTarget.Hyperlink, out string protectedFeature);
        string destinationKind = (string?)hyperlinkTarget.Hyperlink.Attribute(OoxmlNs.W + "anchor") is not null
            ? "an internal anchor"
            : (string?)hyperlinkTarget.Hyperlink.Attribute(OoxmlNs.R + "id") is not null
                ? "an external URI"
                : "no destination";
        var facts = new HyperlinkCapabilityFacts(
            hasProtected,
            protectedFeature,
            destinationKind);
        bool isTracked = mode is TrackChangesMode.Require or TrackChangesMode.Suggest;
        bool isRequire = mode == TrackChangesMode.Require;
        var operations = new List<DocxOperationCapability>
        {
            HyperlinkTargetCapability(facts, isRequire),
            HyperlinkTextCapability(facts, isTracked, isRequire),
            HyperlinkRemoveCapability(isRequire),
        };
        return new ParagraphCapabilitiesOutcome
        {
            Capabilities = new DocxTargetCapabilities(
                parsed.ToWireValue(),
                "hyperlink",
                storyDocument.Story.StoryLabel,
                operations)
        };
    }
    private static DocxOperationCapability HyperlinkTargetCapability(HyperlinkCapabilityFacts facts, bool isRequire)
    {
        const string operation = "set-hyperlink-target";
        if (isRequire)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Hyperlink target updates modify relationship or anchor metadata without a tracked revision representation, so Require fails with E6001 before editing. Use Suggest, which warns with W4001 and applies directly, or Off for a direct edit. Check remains authoritative.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "supported",
            "Link currently points to " + facts.DestinationKind + ". The destination must be an absolute http, https, or mailto URI or an internal anchor name; tooltip, target-frame, and history are optional.",
            operation,
            null);
    }
    private static DocxOperationCapability HyperlinkTextCapability(HyperlinkCapabilityFacts facts, bool isTracked, bool isRequire)
    {
        const string operation = "set-hyperlink-text";
        if (facts.HasProtected)
        {
            if (isRequire)
            {
                return new DocxOperationCapability(
                    operation,
                    "unsupported",
                    "Hyperlink contains protected OOXML boundary " + Quote(facts.ProtectedFeature) + ", so tracked text replacement fails with E6002 under Require. Suggest falls back to a direct rewrite with W4002 and Off rewrites directly. Check remains authoritative.",
                    operation,
                    null);
            }
            if (isTracked)
            {
                return new DocxOperationCapability(
                    operation,
                    "conditional",
                    "Hyperlink contains protected OOXML boundary " + Quote(facts.ProtectedFeature) + ", so tracked text replacement falls back to a direct rewrite with W4002. The exact replacement decides. Check remains authoritative.",
                    operation,
                    null);
            }
            return new DocxOperationCapability(
                operation,
                "supported",
                "Direct rewrite replaces the display text and preserves the relationship or anchor. A guarded replacement needs expect-text that matches the current text; identical text is a no-op.",
                operation,
                null);
        }
        if (isTracked)
        {
            return new DocxOperationCapability(
                operation,
                "conditional",
                "Simple display text emits tracked delete and insert markup while preserving the relationship or anchor; complex shapes fall back to a direct rewrite with W4002 under Suggest and fail with E6002 under Require. The exact replacement decides. Check remains authoritative.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "supported",
            "Direct rewrite replaces the display text and preserves the relationship or anchor. A guarded replacement needs expect-text that matches the current text; identical text is a no-op.",
            operation,
            null);
    }
    private static DocxOperationCapability HyperlinkRemoveCapability(bool isRequire)
    {
        const string operation = "remove-hyperlink";
        if (isRequire)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Hyperlink removal changes wrapper and relationship metadata without a tracked revision representation, so Require fails with E6001 before editing. Use Suggest, which warns with W4001 and applies directly, or Off for a direct edit. Check remains authoritative.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "supported",
            "Removal unwraps the hyperlink and preserves display runs; the relationship is dropped when nothing else uses it. A guarded removal needs expect-text that matches the current display text.",
            operation,
            null);
    }
    // D17: field capabilities reuse the shape, guard-support, and track-support predicates from the field engine.
    private sealed record FieldCapabilityFacts(
        bool IsSimple,
        bool ComplexSafe,
        string ComplexReason,
        bool HasProtected,
        string ProtectedFeature,
        string FieldType,
        string RefreshKind,
        string RefreshDetail);
    internal static ParagraphCapabilitiesOutcome GetFieldCapabilities(
        OoxmlPackage package,
        DocxTargetId parsed,
        string requestedTargetId,
        TrackChangesMode mode,
        CancellationToken cancellationToken)
    {
        StoryDocument? storyDocument = TryResolveStoryDocument(package, parsed, cancellationToken);
        FieldTarget? fieldTarget = storyDocument is null
            ? null
            : FindFieldTarget(package, storyDocument.PartName, parsed.Story, parsed.StoryPart, parsed.Primary, cancellationToken, allowLiveFallback: true);
        if (storyDocument is null || fieldTarget is null)
        {
            return ParagraphCapabilitiesNotFound(requestedTargetId);
        }
        bool simple = fieldTarget.Element.Name == OoxmlNs.W + "fldSimple";
        bool complexSafe = false;
        string complexReason = string.Empty;
        if (!simple)
        {
            complexSafe = TryGetSimpleComplexFieldResultRuns(
                fieldTarget.Element,
                out _,
                out _,
                out string? reason);
            complexReason = reason ?? string.Empty;
        }
        bool hasProtected = TryGetProtectedTextEditFeature(fieldTarget.Element, out string protectedFeature);
        string fieldType = "complex";
        string refreshKind = "complex";
        string refreshDetail = string.Empty;
        if (simple)
        {
            string code = NormalizeFieldCodeForGuard((string?)fieldTarget.Element.Attribute(OoxmlNs.W + "instr") ?? string.Empty);
            string[] tokens = TokenizeFieldCodeForPatch(code);
            fieldType = tokens.Length == 0 ? "unknown" : NormalizeFieldTypeForPatch(tokens[0]);
            if (TryReadQuoteFieldText(code, out _))
            {
                refreshKind = "quote";
            }
            else if (TryReadRefFieldBookmarkName(code, out string? bookmarkName))
            {
                if (TryReadSimpleBookmarkText(fieldTarget.Document, bookmarkName, out _, out string? bookmarkReason))
                {
                    refreshKind = "ref";
                }
                else
                {
                    refreshKind = "ref-unresolved";
                    refreshDetail = bookmarkReason ?? string.Empty;
                }
            }
            else
            {
                refreshKind = "unsupported";
            }
        }
        var facts = new FieldCapabilityFacts(
            simple,
            complexSafe,
            complexReason,
            hasProtected,
            protectedFeature,
            fieldType,
            refreshKind,
            refreshDetail);
        bool isTracked = mode is TrackChangesMode.Require or TrackChangesMode.Suggest;
        bool isRequire = mode == TrackChangesMode.Require;
        var operations = new List<DocxOperationCapability>
        {
            FieldFlagCapability("set-field-dirty", isRequire),
            FieldFlagCapability("set-field-lock", isRequire),
            FieldCodeCapability(facts, isRequire),
            FieldResultCapability(facts, isTracked, isRequire),
            FieldRefreshCapability(facts, isRequire),
        };
        return new ParagraphCapabilitiesOutcome
        {
            Capabilities = new DocxTargetCapabilities(
                parsed.ToWireValue(),
                "field",
                storyDocument.Story.StoryLabel,
                operations)
        };
    }
    private static DocxOperationCapability FieldFlagCapability(string operation, bool isRequire)
    {
        if (isRequire)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Field flag updates are field metadata without a tracked revision representation, so Require fails with E6001 before editing. Use Suggest, which warns with W4001 and applies directly, or Off for a direct edit. Check remains authoritative.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "supported",
            "Sets the field flag directly with no shape checks.",
            operation,
            null);
    }
    private static DocxOperationCapability FieldCodeCapability(FieldCapabilityFacts facts, bool isRequire)
    {
        const string operation = "set-field-code";
        if (!facts.IsSimple)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Field code replacement currently supports only simple w:fldSimple fields, so this complex field fails with E4313.",
                operation,
                null);
        }
        if (isRequire)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Field codes are instruction metadata without a tracked revision representation, so Require fails with E6001 before editing. Use Suggest, which warns with W4001 and applies directly, or Off for a direct edit. Check remains authoritative.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "supported",
            "Replaces the field instruction and marks the field dirty. The new code is patch-supplied; expect-code can guard the current code.",
            operation,
            null);
    }
    private static DocxOperationCapability FieldResultCapability(FieldCapabilityFacts facts, bool isTracked, bool isRequire)
    {
        const string operation = "set-field-result";
        if (!facts.IsSimple && !facts.ComplexSafe)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Cached-result replacement for this complex field is not safe: " + facts.ComplexReason + ", so set-field-result fails with E4313.",
                operation,
                null);
        }
        if (!facts.IsSimple)
        {
            if (isRequire)
            {
                return new DocxOperationCapability(
                    operation,
                    "unsupported",
                    "Tracked complex-field result replacement is not modeled, so Require fails with E6002. Suggest falls back to a direct rewrite with W4002 and Off rewrites directly. Check remains authoritative.",
                    operation,
                    null);
            }
            if (isTracked)
            {
                return new DocxOperationCapability(
                    operation,
                    "conditional",
                    "Tracked complex-field result replacement is not modeled, so Suggest falls back to a direct rewrite with W4002. Check remains authoritative.",
                    operation,
                    null);
            }
            return new DocxOperationCapability(
                operation,
                "supported",
                "Direct rewrite replaces the cached result of this simple complex field. Guard expect-result can assert the current result.",
                operation,
                null);
        }
        if (facts.HasProtected)
        {
            if (isRequire)
            {
                return new DocxOperationCapability(
                    operation,
                    "unsupported",
                    "Field result contains protected OOXML boundary " + Quote(facts.ProtectedFeature) + ", so tracked result replacement fails with E6002 under Require. Suggest falls back to a direct rewrite with W4002 and Off rewrites directly. Check remains authoritative.",
                    operation,
                    null);
            }
            if (isTracked)
            {
                return new DocxOperationCapability(
                    operation,
                    "conditional",
                    "Field result contains protected OOXML boundary " + Quote(facts.ProtectedFeature) + ", so tracked result replacement falls back to a direct rewrite with W4002. The exact replacement decides. Check remains authoritative.",
                    operation,
                    null);
            }
        }
        if (isTracked)
        {
            return new DocxOperationCapability(
                operation,
                "conditional",
                "Simple cached results emit tracked delete and insert markup while preserving the field instruction; complex shapes fall back to a direct rewrite with W4002 under Suggest and fail with E6002 under Require. The exact replacement decides. Check remains authoritative.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "supported",
            "Direct rewrite replaces the cached result and preserves the field instruction. Guard expect-result can assert the current result.",
            operation,
            null);
    }
    private static DocxOperationCapability FieldRefreshCapability(FieldCapabilityFacts facts, bool isRequire)
    {
        const string operation = "refresh-field-result";
        if (isRequire)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Field refresh updates cached result text without a tracked revision representation, so Require fails with E6001 before editing. Use Suggest, which warns with W4001 and applies directly, or Off for a direct refresh. Check remains authoritative.",
                operation,
                null);
        }
        if (facts.RefreshKind == "quote" || facts.RefreshKind == "ref")
        {
            return new DocxOperationCapability(
                operation,
                "supported",
                "Deterministic refresh recomputes this " + facts.FieldType + " field from modeled document state and clears its dirty flag. Guards expect-code and expect-result can assert current values.",
                operation,
                null);
        }
        if (facts.RefreshKind == "ref-unresolved")
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Field refresh cannot resolve its bookmark: " + facts.RefreshDetail + ", so refresh-field-result fails with E4313.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "unsupported",
            "Field refresh does not support field type " + Quote(facts.FieldType) + ": deterministic refresh covers REF, PAGEREF, NOTEREF, and QUOTE, so refresh-field-result fails with E4313.",
            operation,
            null);
    }
    // D17: image capabilities reuse the drawing-shape predicates from the image engine.
    private sealed record ImageCapabilityFacts(
        bool HasDrawing,
        bool HasContainer,
        bool IsAnchored,
        string? ContentType);
    internal static ParagraphCapabilitiesOutcome GetImageCapabilities(
        OoxmlPackage package,
        DocxTargetId parsed,
        string requestedTargetId,
        TrackChangesMode mode,
        CancellationToken cancellationToken)
    {
        StoryDocument? storyDocument = TryResolveStoryDocument(package, parsed, cancellationToken);
        ImageBlipTarget? imageTarget = storyDocument is null
            ? null
            : FindImageBlipTarget(package, storyDocument.PartName, parsed.Story, parsed.StoryPart, parsed.Primary, cancellationToken, allowLiveFallback: true);
        if (storyDocument is null || imageTarget is null)
        {
            return ParagraphCapabilitiesNotFound(requestedTargetId);
        }
        XElement? drawing = imageTarget.Blip.Ancestors(OoxmlNs.W + "drawing").FirstOrDefault();
        XElement? container = drawing?.Descendants(OoxmlNs.Wp + "inline").FirstOrDefault()
            ?? drawing?.Descendants(OoxmlNs.Wp + "anchor").FirstOrDefault();
        var facts = new ImageCapabilityFacts(
            drawing is not null,
            container is not null,
            container is not null && container.Name == OoxmlNs.Wp + "anchor",
            imageTarget.Part.ContentType);
        bool isRequire = mode == TrackChangesMode.Require;
        var operations = new List<DocxOperationCapability>
        {
            ImageReplaceCapability(isRequire),
            ImageDrawingCapability("set-image-alt", facts, isRequire, needsAnchor: false),
            ImageDrawingCapability("set-image-metadata", facts, isRequire, needsAnchor: false),
            ImageDrawingCapability("set-image-size", facts, isRequire, needsAnchor: false),
            ImageDrawingCapability("set-image-wrap", facts, isRequire, needsAnchor: true),
            ImageDrawingCapability("set-image-position", facts, isRequire, needsAnchor: true),
            ImageDrawingCapability("set-image-crop", facts, isRequire, needsAnchor: false),
            ImageDeleteCapability(facts, isRequire),
        };
        return new ParagraphCapabilitiesOutcome
        {
            Capabilities = new DocxTargetCapabilities(
                parsed.ToWireValue(),
                "image",
                storyDocument.Story.StoryLabel,
                operations)
        };
    }
    private static DocxOperationCapability ImageReplaceCapability(bool isRequire)
    {
        const string operation = "replace-image";
        if (isRequire)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Image replacement updates DrawingML and package media without a tracked revision representation, so Require fails with E6001 before editing. Use Suggest, which warns with W4001 and applies directly, or Off for a direct edit. Check remains authoritative.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "conditional",
            "The asset must be readable and its content type must match the existing media part, otherwise check fails with E5204; an optional alt update needs editable DrawingML properties. Guards expect-content-type can assert the current media type. Check remains authoritative for the exact asset.",
            operation,
            null);
    }
    private static DocxOperationCapability ImageDrawingCapability(string operation, ImageCapabilityFacts facts, bool isRequire, bool needsAnchor)
    {
        if (isRequire)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Image DrawingML updates carry no tracked revision representation, so Require fails with E6001 before editing. Use Suggest, which warns with W4001 and applies directly, or Off for a direct edit. Check remains authoritative.",
                operation,
                null);
        }
        if (!facts.HasContainer)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Image does not have editable DrawingML properties, so " + operation + " fails with E5205.",
                operation,
                null);
        }
        if (needsAnchor && !facts.IsAnchored)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Image is inline; this metadata is only editable on anchored images, so " + operation + " fails with E5205.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "supported",
            "Direct edit updates the DrawingML metadata; value fields and the expect-content-type guard are decided by check.",
            operation,
            null);
    }
    private static DocxOperationCapability ImageDeleteCapability(ImageCapabilityFacts facts, bool isRequire)
    {
        const string operation = "delete-image";
        if (isRequire)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Image deletion removes DrawingML and package media without a tracked revision representation, so Require fails with E6001 before editing. Use Suggest, which warns with W4001 and applies directly, or Off for a direct edit. Check remains authoritative.",
                operation,
                null);
        }
        if (!facts.HasDrawing)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Image does not have an editable DrawingML object, so delete-image fails with E5205.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "supported",
            "Deletion removes the DrawingML object and drops the media relationship and part when nothing else uses them.",
            operation,
            null);
    }
    // D17: row capabilities reuse the insertion-boundary, deletion, grid-shape, orphan, and track-support predicates from the table and text engines.
    private sealed record RowCapabilityFacts(
        bool InsertBeforeOk,
        string InsertBeforeReason,
        bool InsertAfterOk,
        string InsertAfterReason,
        bool IsOnlyRow,
        bool CanDelete,
        string DeleteReason,
        bool GridConsistent,
        bool RowHasCells,
        string? OrphanDescription);
    internal static ParagraphCapabilitiesOutcome GetRowCapabilities(
        OoxmlPackage package,
        DocxTargetId parsed,
        string requestedTargetId,
        TrackChangesMode mode,
        CancellationToken cancellationToken)
    {
        StoryDocument? storyDocument = TryResolveStoryDocument(package, parsed, cancellationToken);
        XElement? table = storyDocument is null
            ? null
            : DocxStoryBlocks.FindTableByPhysicalOrdinal(
                storyDocument.Document.Root?.Element(OoxmlNs.W + "body") ?? storyDocument.Document.Root!,
                parsed.Primary);
        XElement[] rows = table?.Elements(OoxmlNs.W + "tr").ToArray() ?? [];
        XElement? row = parsed.Secondary >= 1 ? rows.ElementAtOrDefault(parsed.Secondary - 1) : null;
        if (storyDocument is null || table is null || row is null)
        {
            return ParagraphCapabilitiesNotFound(requestedTargetId);
        }
        bool insertBeforeOk = CanInsertRowWithVerticalMerges(table, row, insertAfter: false, out string? beforeReason);
        bool insertAfterOk = CanInsertRowWithVerticalMerges(table, row, insertAfter: true, out string? afterReason);
        bool canDelete = CanDeleteRowWithVerticalMerges(table, row, out string? deleteReason);
        var facts = new RowCapabilityFacts(
            insertBeforeOk,
            beforeReason ?? string.Empty,
            insertAfterOk,
            afterReason ?? string.Empty,
            rows.Length == 1,
            canDelete,
            deleteReason ?? string.Empty,
            IsRectangular(table, out _) || TryGetConsistentVisualColumnCount(table, out _),
            row.Elements(OoxmlNs.W + "tc").Any(),
            FindOrphanedRangeBoundary(storyDocument.Document, row));
        bool isTracked = mode is TrackChangesMode.Require or TrackChangesMode.Suggest;
        bool isRequire = mode == TrackChangesMode.Require;
        var operations = new List<DocxOperationCapability>
        {
            RowInsertCapability("insert-row-before", facts.InsertBeforeOk, facts.InsertBeforeReason),
            RowInsertCapability("insert-row-after", facts.InsertAfterOk, facts.InsertAfterReason),
            RowDeleteCapability(facts, isTracked, isRequire),
            RowHeaderCapability(),
        };
        return new ParagraphCapabilitiesOutcome
        {
            Capabilities = new DocxTargetCapabilities(
                parsed.ToWireValue(),
                "row",
                storyDocument.Story.StoryLabel,
                operations)
        };
    }
    private static DocxOperationCapability RowInsertCapability(string operation, bool boundaryOk, string boundaryReason)
    {
        if (!boundaryOk)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Table cannot be edited safely by " + operation + ": " + boundaryReason + ", so " + operation + " fails with E4301.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "conditional",
            "The patch must supply one cell field per cloned cell, and table guards still apply; ragged grids need force true, which clones the raw cell count and may misalign merged or spanned columns. Simple rectangular tables emit row insertion revisions under tracked modes while complex shapes fall back with W4002 under Suggest and fail with E6002 under Require. Check remains authoritative for the exact cells.",
            operation,
            null);
    }
    private static DocxOperationCapability RowDeleteCapability(RowCapabilityFacts facts, bool isTracked, bool isRequire)
    {
        const string operation = "delete-row";
        if (facts.IsOnlyRow)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Cannot delete the last row of a table, so delete-row fails with E4304.",
                operation,
                null);
        }
        if (!facts.CanDelete)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Table cannot be edited safely by delete-row: " + facts.DeleteReason + ", so delete-row fails with E4301.",
                operation,
                null);
        }
        if (facts.OrphanDescription is not null && !isTracked)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Deleting this row would orphan " + facts.OrphanDescription + " outside the deleted element, so direct delete-row fails with E4305. Tracked modes mark the deletion instead; delete the range first or choose another target.",
                operation,
                null);
        }
        if (!facts.GridConsistent)
        {
            return new DocxOperationCapability(
                operation,
                "conditional",
                "Table does not have a consistent visual grid: without force true delete-row fails with E4301, and with force true the row is deleted anyway, which may leave the remaining grid ragged. Tracked shape is decided by check. Check remains authoritative.",
                operation,
                null);
        }
        if (isTracked)
        {
            return new DocxOperationCapability(
                operation,
                "conditional",
                "Simple rectangular tables emit row deletion revisions; deleting the row promotes the next vertical-merge continuation when a merge root goes away. Complex shapes fall back with W4002 under Suggest and fail with E6002 under Require. Check remains authoritative.",
                operation,
                null);
        }
        if (facts.OrphanDescription is not null)
        {
            return new DocxOperationCapability(
                operation,
                "conditional",
                "Deleting this row would orphan " + facts.OrphanDescription + " outside the deleted element: direct delete fails with E4305 while tracked modes mark the deletion instead. Check remains authoritative.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "supported",
            "Direct delete removes the row and promotes the next vertical-merge continuation when a merge root goes away. Guard expect-contains can assert the row text.",
            operation,
            null);
    }
    private static DocxOperationCapability RowHeaderCapability()
    {
        const string operation = "set-row-header";
        return new DocxOperationCapability(
            operation,
            "supported",
            "Sets or clears the repeating-header flag on the row. Guard expect-header can assert the current flag. Tracked modes emit row property revisions with w:trPrChange while preserving previous row properties.",
            operation,
            null);
    }
    // D17: table capabilities reuse the grid-shape, orphan, and track-support predicates from the table and text engines.
    private sealed record TableCapabilityFacts(
        int RowCount,
        bool GridConsistent,
        bool LastRowAppendable,
        string LastRowReason,
        bool TemplateHasCells,
        string? OrphanDescription);
    internal static ParagraphCapabilitiesOutcome GetTableCapabilities(
        OoxmlPackage package,
        DocxTargetId parsed,
        string requestedTargetId,
        TrackChangesMode mode,
        CancellationToken cancellationToken)
    {
        StoryDocument? storyDocument = TryResolveStoryDocument(package, parsed, cancellationToken);
        XElement? table = storyDocument is null
            ? null
            : DocxStoryBlocks.FindTableByPhysicalOrdinal(
                storyDocument.Document.Root?.Element(OoxmlNs.W + "body") ?? storyDocument.Document.Root!,
                parsed.Primary);
        if (storyDocument is null || table is null)
        {
            return ParagraphCapabilitiesNotFound(requestedTargetId);
        }
        XElement[] rows = table.Elements(OoxmlNs.W + "tr").ToArray();
        bool gridConsistent = TryGetConsistentVisualColumnCount(table, out _);
        bool lastRowAppendable = false;
        string lastRowReason = string.Empty;
        bool templateHasCells = false;
        if (rows.Length != 0)
        {
            lastRowAppendable = CanAppendRowWithVerticalMerges(table, rows[^1], out string? reason);
            lastRowReason = reason ?? string.Empty;
            templateHasCells = rows[^1].Elements(OoxmlNs.W + "tc").Any();
        }
        var facts = new TableCapabilityFacts(
            rows.Length,
            gridConsistent,
            lastRowAppendable,
            lastRowReason,
            templateHasCells,
            FindOrphanedRangeBoundary(storyDocument.Document, table));
        bool isTracked = mode is TrackChangesMode.Require or TrackChangesMode.Suggest;
        bool isRequire = mode == TrackChangesMode.Require;
        var operations = new List<DocxOperationCapability>
        {
            TableInsertCapability("insert-before", isTracked),
            TableInsertCapability("insert-after", isTracked),
            TableHyperlinkInsertCapability(isTracked),
            TableDeleteCapability(facts, isTracked, isRequire),
            TableStyleCapability(),
            TableMetadataCapability(isRequire),
            TableAppendRowCapability(facts),
        };
        return new ParagraphCapabilitiesOutcome
        {
            Capabilities = new DocxTargetCapabilities(
                parsed.ToWireValue(),
                "table",
                storyDocument.Story.StoryLabel,
                operations)
        };
    }
    private static DocxOperationCapability TableInsertCapability(string operation, bool isTracked)
    {
        if (isTracked)
        {
            return new DocxOperationCapability(
                operation,
                "supported",
                "Structural insert next to this table; the anchor table is unchanged. Inserted text without tabs or line breaks is recorded as tracked insertion markup; other inserted text falls back to a direct insert with W4002 under Suggest and fails with E6002 under Require. Check remains authoritative for the exact text.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "supported",
            "Structural insert next to this table; the anchor table is unchanged.",
            operation,
            null);
    }
    private static DocxOperationCapability TableHyperlinkInsertCapability(bool isTracked)
    {
        const string operation = "insert-hyperlink-after";
        if (isTracked)
        {
            return new DocxOperationCapability(
                operation,
                "supported",
                "Hyperlink paragraph insert after this table; the anchor table is unchanged. Simple display text is recorded as tracked insertion markup; text with tabs or line breaks falls back to a direct insert with W4002 under Suggest and fails with E6002 under Require. Check remains authoritative for the exact text.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "supported",
            "Hyperlink paragraph insert after this table; the anchor table is unchanged.",
            operation,
            null);
    }
    private static DocxOperationCapability TableDeleteCapability(TableCapabilityFacts facts, bool isTracked, bool isRequire)
    {
        const string operation = "delete-block";
        if (facts.OrphanDescription is not null)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Deleting this table would orphan " + facts.OrphanDescription + " outside the deleted element, so delete-block fails with E4305 in every mode. Delete the range first or choose another target.",
                operation,
                null);
        }
        if (isRequire)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Tracked block deletion is modeled only for paragraph targets, so Require fails with E6002. Suggest falls back to a direct delete with W4002 and Off deletes directly, removing the table. Check remains authoritative.",
                operation,
                null);
        }
        if (isTracked)
        {
            return new DocxOperationCapability(
                operation,
                "supported",
                "Tracked block deletion is modeled only for paragraph targets, so Suggest falls back to a direct delete with W4002 while Off deletes directly; either removes the table. Check remains authoritative.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "supported",
            "Direct delete removes the table.",
            operation,
            null);
    }
    private static DocxOperationCapability TableStyleCapability()
    {
        const string operation = "set-table-style";
        return new DocxOperationCapability(
            operation,
            "conditional",
            "The style value must resolve to a table style by ID or by unique name, otherwise check fails; expect-style can guard the current style. Tracked modes emit table property revisions with w:tblPrChange while preserving previous table properties. Check remains authoritative for the exact style.",
            operation,
            null);
    }
    private static DocxOperationCapability TableMetadataCapability(bool isRequire)
    {
        const string operation = "set-table-metadata";
        if (isRequire)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Table caption and description updates are table metadata without a tracked revision representation, so Require fails with E6001 before editing. Use Suggest, which warns with W4001 and applies directly, or Off for a direct edit. Check remains authoritative.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "supported",
            "Sets or clears table caption and description metadata; at least one of them is required. Guards expect-caption and expect-description can assert current values.",
            operation,
            null);
    }
    private static DocxOperationCapability TableAppendRowCapability(TableCapabilityFacts facts)
    {
        const string operation = "append-row";
        if (facts.RowCount == 0)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Table has no rows to clone, so append-row fails with E4301.",
                operation,
                null);
        }
        if (!facts.GridConsistent)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Table does not have a consistent visual grid, so append-row fails with E4301.",
                operation,
                null);
        }
        if (!facts.LastRowAppendable)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Table cannot be appended safely: " + facts.LastRowReason + ", so append-row fails with E4301.",
                operation,
                null);
        }
        if (!facts.TemplateHasCells)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Table last row has no cells to clone, so append-row fails with E4301.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "conditional",
            "The patch must supply one cell field per cloned cell, and table guards still apply; simple rectangular tables emit row insertion revisions under tracked modes while complex shapes fall back with W4002 under Suggest and fail with E6002 under Require. Check remains authoritative for the exact cells.",
            operation,
            null);
    }
    // D17: bookmark capabilities reuse the range-shape, protection, hyperlink-reference, and track-support predicates from the bookmark engine.
    private sealed record BookmarkCapabilityFacts(
        string? Name,
        bool HasEnd,
        bool ParagraphBounded,
        bool SameParagraph,
        bool HasProtected,
        string ProtectedFeature,
        bool ReferencedByHyperlink,
        bool RenameAmbiguous);
    internal static ParagraphCapabilitiesOutcome GetBookmarkCapabilities(
        OoxmlPackage package,
        DocxTargetId parsed,
        string requestedTargetId,
        TrackChangesMode mode,
        CancellationToken cancellationToken)
    {
        StoryDocument? storyDocument = TryResolveStoryDocument(package, parsed, cancellationToken);
        BookmarkTarget? bookmarkTarget = storyDocument is null
            ? null
            : FindBookmarkTarget(package, storyDocument.PartName, parsed.Story, parsed.StoryPart, parsed.Primary, cancellationToken, allowLiveFallback: true);
        if (storyDocument is null || bookmarkTarget is null)
        {
            return ParagraphCapabilitiesNotFound(requestedTargetId);
        }
        string? name = (string?)bookmarkTarget.Start.Attribute(OoxmlNs.W + "name");
        XElement? end = bookmarkTarget.End;
        bool paragraphBounded = end is not null &&
            bookmarkTarget.Start.Parent is not null &&
            bookmarkTarget.Start.Parent.Name == OoxmlNs.W + "p" &&
            end.Parent is not null &&
            end.Parent.Name == OoxmlNs.W + "p";
        bool sameParagraph = paragraphBounded &&
            end is not null &&
            bookmarkTarget.Start.Parent == end.Parent;
        bool hasProtected = false;
        string protectedFeature = string.Empty;
        if (sameParagraph)
        {
            XNode[] nodes = bookmarkTarget.Start.NodesAfterSelf().TakeWhile(node => node != bookmarkTarget.End).ToArray();
            hasProtected = ContainsProtectedBookmarkReplacementNode(nodes, out string? feature);
            protectedFeature = feature ?? string.Empty;
        }
        string currentName = name ?? string.Empty;
        bool referenced = !string.IsNullOrWhiteSpace(currentName) && HasInternalHyperlinkAnchor(bookmarkTarget.Document, currentName);
        bool renameAmbiguous = !string.IsNullOrWhiteSpace(currentName) &&
            CountBookmarkName(bookmarkTarget.Document, currentName) > 1 &&
            HasInternalHyperlinkAnchor(bookmarkTarget.Document, currentName);
        var facts = new BookmarkCapabilityFacts(
            name,
            bookmarkTarget.End is not null,
            paragraphBounded,
            sameParagraph,
            hasProtected,
            protectedFeature,
            referenced,
            renameAmbiguous);
        bool isTracked = mode is TrackChangesMode.Require or TrackChangesMode.Suggest;
        bool isRequire = mode == TrackChangesMode.Require;
        var operations = new List<DocxOperationCapability>
        {
            BookmarkReplaceCapability(facts, isTracked, isRequire),
            BookmarkRenameCapability(facts, mode),
            BookmarkDeleteCapability(facts, mode),
        };
        return new ParagraphCapabilitiesOutcome
        {
            Capabilities = new DocxTargetCapabilities(
                parsed.ToWireValue(),
                "bookmark",
                storyDocument.Story.StoryLabel,
                operations)
        };
    }
    private static DocxOperationCapability BookmarkReplaceCapability(BookmarkCapabilityFacts facts, bool isTracked, bool isRequire)
    {
        const string operation = "replace-bookmark-text";
        if (!facts.ParagraphBounded)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Bookmark range is not a complete paragraph-bounded range, so text replacement fails with E4311.",
                operation,
                null);
        }
        if (facts.HasProtected)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Bookmark range contains protected OOXML boundary " + Quote(facts.ProtectedFeature) + ", so text replacement fails with E4311.",
                operation,
                null);
        }
        if (!facts.SameParagraph)
        {
            return new DocxOperationCapability(
                operation,
                "conditional",
                "Multi-paragraph ranges rewrite directly with guarded replacements, and table-spanning ranges need one replacement line per visible text slot; tracked multi-paragraph output falls back to a direct rewrite with W4002 under Suggest and fails with E6002 under Require. Check remains authoritative for the exact replacement.",
                operation,
                null);
        }
        if (isTracked)
        {
            return new DocxOperationCapability(
                operation,
                "conditional",
                isRequire
                    ? "Simple run-only ranges emit tracked delete and insert markup while preserving markers; complex shapes fail with E6002 under Require. The exact replacement decides. Check remains authoritative."
                    : "Simple run-only ranges emit tracked delete and insert markup while preserving markers; complex shapes fall back to a direct rewrite with W4002. The exact replacement decides. Check remains authoritative.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "supported",
            "Direct rewrite replaces the range text and preserves bookmark markers. A guarded replacement needs expect-text that matches the current content.",
            operation,
            null);
    }
    private static DocxOperationCapability BookmarkRenameCapability(BookmarkCapabilityFacts facts, TrackChangesMode mode)
    {
        const string operation = "rename-bookmark";
        string currentName = facts.Name ?? string.Empty;
        if (string.IsNullOrWhiteSpace(currentName))
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Bookmark has no current name, so rename fails with E4311.",
                operation,
                null);
        }
        if (facts.RenameAmbiguous)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Bookmark name " + Quote(currentName) + " is duplicated with same-part hyperlink anchors, so rename fails with E4311 as ambiguous.",
                operation,
                null);
        }
        if (!SupportsTrackedChangeOutput(operation) && mode == TrackChangesMode.Require && !IsAnnotationOperation(operation))
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Bookmark rename changes anchor metadata without a tracked revision representation, so Require fails with E6001 before editing. Use Suggest, which warns with W4001 and applies directly, or Off for a direct edit. Check remains authoritative.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "supported",
            "Rename updates markers and same-story internal hyperlink anchors when unambiguous. The new name must be non-empty without whitespace and unique in the part; expect-name can guard the current name.",
            operation,
            null);
    }
    private static DocxOperationCapability BookmarkDeleteCapability(BookmarkCapabilityFacts facts, TrackChangesMode mode)
    {
        const string operation = "delete-bookmark";
        if (!facts.HasEnd)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Bookmark is incomplete and cannot be deleted safely, so deletion fails with E4311.",
                operation,
                null);
        }
        if (facts.ReferencedByHyperlink)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Bookmark is referenced by same-part hyperlink anchors, so deletion fails with E4311. Update or remove those hyperlinks first.",
                operation,
                null);
        }
        if (!SupportsTrackedChangeOutput(operation) && mode == TrackChangesMode.Require && !IsAnnotationOperation(operation))
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Bookmark deletion removes anchor metadata without a tracked revision representation, so Require fails with E6001 before editing. Use Suggest, which warns with W4001 and applies directly, or Off for a direct edit. Check remains authoritative.",
                operation,
                null);
        }
        return new DocxOperationCapability(
            operation,
            "supported",
            "Deletion removes complete unreferenced markers and preserves content; expect-name can guard the current name.",
            operation,
            null);
    }
    private static ParagraphCapabilitiesOutcome GetCellCapabilities(
        OoxmlPackage package,
        DocxTargetId parsed,
        string requestedTargetId,
        TrackChangesMode mode,
        bool isMergeGroup,
        CancellationToken cancellationToken)
    {
        StoryDocument? storyDocument = TryResolveStoryDocument(package, parsed, cancellationToken);
        if (storyDocument is null)
        {
            return ParagraphCapabilitiesNotFound(requestedTargetId);
        }

        XElement? root = storyDocument.Document.Root;
        XElement container = root?.Element(OoxmlNs.W + "body") ?? root!;
        XElement? table = DocxStoryBlocks.FindTableByPhysicalOrdinal(container, parsed.Primary);
        if (table is null)
        {
            return ParagraphCapabilitiesNotFound(requestedTargetId);
        }

        XElement? row = null;
        XElement? cell = null;
        int rowOrdinal = parsed.Secondary;
        int visualColumn = parsed.Tertiary;
        if (isMergeGroup)
        {
            if (!TryFindMergeGroupRoot(table, parsed.Secondary, out XElement? groupRow, out XElement? groupCell, out int groupColumn, out int groupRowOrdinal))
            {
                return ParagraphCapabilitiesNotFound(requestedTargetId);
            }

            row = groupRow;
            cell = groupCell;
            rowOrdinal = groupRowOrdinal;
            visualColumn = groupColumn;
        }
        else
        {
            row = table.Elements(OoxmlNs.W + "tr").ElementAtOrDefault(parsed.Secondary - 1);
            cell = row is null ? null : FindCellByVisualColumn(row, parsed.Tertiary);
            if (row is null || cell is null)
            {
                return ParagraphCapabilitiesNotFound(requestedTargetId);
            }
        }

        if (row is null || cell is null)
        {
            return ParagraphCapabilitiesNotFound(requestedTargetId);
        }

        bool isContinuation = IsVerticalMergeContinuation(cell);
        string? rootCellId = null;
        if (isContinuation)
        {
            rootCellId = FindMergeRootCellId(table, parsed, rowOrdinal, visualColumn);
        }

        bool trackedShapeOk = true;
        string? trackedShapeReason = null;
        bool isTracked = mode is TrackChangesMode.Require or TrackChangesMode.Suggest;
        if (isTracked)
        {
            trackedShapeOk = TryGetTrackedSetCellParagraphs(cell, string.Empty, out _, out trackedShapeReason);
        }

        var facts = new CellCapabilityFacts(
            isContinuation,
            rootCellId,
            IsSimpleEditableCell(cell),
            trackedShapeOk,
            trackedShapeReason);
        bool isRequire = mode == TrackChangesMode.Require;
        var operations = new List<DocxOperationCapability>
        {
            SetCellCapability(facts, parsed.ToWireValue(), isTracked, isRequire),
            SetCellShadingCapability(facts, parsed.ToWireValue()),
            ReplaceTextCellCapability(facts, parsed.ToWireValue(), isTracked),
        };

        return new ParagraphCapabilitiesOutcome
        {
            Capabilities = new DocxTargetCapabilities(
                parsed.ToWireValue(),
                isMergeGroup ? "merge-group" : "cell",
                storyDocument.Story.StoryLabel,
                operations)
        };
    }
    // Live copy of the merge-group walk in ResolveMergeGroupCellTarget: finds
    // the root cell of the Nth merge group in document order.
    private static bool TryFindMergeGroupRoot(
        XElement table,
        int mergeGroupOrdinal,
        out XElement? row,
        out XElement? cell,
        out int visualColumn,
        out int rowOrdinal)
    {
        row = null;
        cell = null;
        visualColumn = 0;
        rowOrdinal = 0;
        if (mergeGroupOrdinal < 1)
        {
            return false;
        }

        int mergeGroupIndex = 1;
        var activeVerticalMerges = new Dictionary<int, MergeGroupRootState>();
        int currentRowOrdinal = 0;
        foreach (XElement currentRow in table.Elements(OoxmlNs.W + "tr"))
        {
            currentRowOrdinal++;
            int gridBefore = ReadTableRowGridOffset(currentRow, "gridBefore");
            RemoveActiveMergeGroups(activeVerticalMerges, 1, gridBefore);
            int columnIndex = 1 + gridBefore;
            foreach (XElement currentCell in currentRow.Elements(OoxmlNs.W + "tc"))
            {
                int columnSpan = ReadTableCellColumnSpan(currentCell);
                DocxVerticalMerge? verticalMerge = ReadTableCellVerticalMerge(currentCell);
                if (verticalMerge == DocxVerticalMerge.Restart)
                {
                    int currentMergeGroup = mergeGroupIndex++;
                    SetActiveMergeGroup(activeVerticalMerges, columnIndex, columnSpan, new MergeGroupRootState(currentRow, currentCell, columnIndex));
                    if (currentMergeGroup == mergeGroupOrdinal)
                    {
                        row = currentRow;
                        cell = currentCell;
                        visualColumn = columnIndex;
                        rowOrdinal = currentRowOrdinal;
                        return true;
                    }
                }
                else if (verticalMerge is not null)
                {
                    MergeGroupRootState? root = FindActiveMergeGroup(activeVerticalMerges, columnIndex, columnSpan);
                    if (root is null)
                    {
                        int currentMergeGroup = mergeGroupIndex++;
                        if (currentMergeGroup == mergeGroupOrdinal)
                        {
                            row = currentRow;
                            cell = currentCell;
                            visualColumn = columnIndex;
                            rowOrdinal = currentRowOrdinal;
                            return true;
                        }

                        SetActiveMergeGroup(activeVerticalMerges, columnIndex, columnSpan, new MergeGroupRootState(currentRow, currentCell, columnIndex));
                    }
                }
                else
                {
                    RemoveActiveMergeGroups(activeVerticalMerges, columnIndex, columnSpan);
                    if (columnSpan > 1)
                    {
                        int currentMergeGroup = mergeGroupIndex++;
                        if (currentMergeGroup == mergeGroupOrdinal)
                        {
                            row = currentRow;
                            cell = currentCell;
                            visualColumn = columnIndex;
                            rowOrdinal = currentRowOrdinal;
                            return true;
                        }
                    }
                }

                columnIndex += columnSpan;
            }

            int gridAfter = ReadTableRowGridOffset(currentRow, "gridAfter");
            RemoveActiveMergeGroups(activeVerticalMerges, columnIndex, gridAfter);
        }

        return false;
    }

    // Finds the vertical-merge root cell above a continuation cell in the same
    // visual column, so the E4301 refusal can point at the editable root.
    private static string? FindMergeRootCellId(
        XElement table,
        DocxTargetId parsed,
        int rowOrdinal,
        int visualColumn)
    {
        XElement[] rows = table.Elements(OoxmlNs.W + "tr").ToArray();
        for (int index = rowOrdinal - 2; index >= 0; index--)
        {
            XElement? covering = FindCellByVisualColumn(rows[index], visualColumn);
            if (covering is null || IsVerticalMergeContinuation(covering))
            {
                continue;
            }

            int startColumn = 1 + ReadTableRowGridOffset(rows[index], "gridBefore");
            foreach (XElement candidate in rows[index].Elements(OoxmlNs.W + "tc"))
            {
                if (ReferenceEquals(candidate, covering))
                {
                    break;
                }

                startColumn += ReadTableCellColumnSpan(candidate);
            }

            return new DocxTargetId(parsed.Story, parsed.StoryPart, DocxTargetKind.Cell, parsed.Primary, index + 1, startColumn).ToWireValue();
        }

        return null;
    }
    private static DocxOperationCapability SetCellCapability(CellCapabilityFacts facts, string targetId, bool isTracked, bool isRequire)
    {
        const string operation = "set-cell";
        if (facts.IsContinuation)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Unsupported merged-cell target " + Quote(targetId) + ". Only the vertical-merge root holds the merged value, so editing a continuation fails with E4301. Target the vertical-merge root cell instead.",
                operation,
                facts.RootCellId);
        }

        if (isTracked && !facts.TrackedShapeOk)
        {
            if (isRequire)
            {
                return new DocxOperationCapability(
                    operation,
                    "unsupported",
                    "Cell shape needs a direct rewrite (" + facts.TrackedShapeReason + "), so Require fails with E6002; tracked set-cell also rejects force true under Require. Suggest falls back to a direct rewrite with W4002 and Off rewrites directly, with force true when the cell holds complex content. Check remains authoritative for the exact replacement.",
                    operation,
                    null);
            }

            return new DocxOperationCapability(
                operation,
                "conditional",
                "Cell shape needs a direct rewrite (" + facts.TrackedShapeReason + "): use force true to replace all cell content, otherwise check fails with E4302. Suggest warns with W4002 on fallback. Check remains authoritative for the exact replacement.",
                operation,
                null);
        }

        if (!isTracked && !facts.IsSimple)
        {
            return new DocxOperationCapability(
                operation,
                "conditional",
                "Cell holds complex content such as multiple paragraphs, a nested table, a drawing, or a field, so direct replacement needs force true; otherwise check fails with E4302. Check remains authoritative for the exact replacement.",
                operation,
                null);
        }

        if (isTracked)
        {
            return new DocxOperationCapability(
                operation,
                "supported",
                "Simple cell shapes emit paragraph-level tracked markup; replacements with tabs or line breaks, mixed formatting, drawings, fields, or protected boundaries fall back with W4002 under Suggest and fail with E6002 under Require. Check remains authoritative.",
                operation,
                null);
        }

        return new DocxOperationCapability(
            operation,
            "supported",
            "Direct rewrite replaces the cell text. Table guards such as expect-text, expect-row-count, expect-column-count, and expect-cell-count still apply to the patch.",
            operation,
            null);
    }

    private static DocxOperationCapability ReplaceTextCellCapability(CellCapabilityFacts facts, string targetId, bool isTracked)
    {
        const string operation = "replace-text";
        if (facts.IsContinuation)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Unsupported merged-cell target " + Quote(targetId) + ". Only the vertical-merge root holds the merged value, so editing a continuation fails with E4301. Target the vertical-merge root cell instead.",
                operation,
                facts.RootCellId);
        }

        if (isTracked)
        {
            return new DocxOperationCapability(
                operation,
                "conditional",
                "Substring replacement applies to cell paragraphs in document order with occurrence counted across those paragraphs; tracked output follows the paragraph rules per paragraph. Check remains authoritative for the exact find span and replacement.",
                operation,
                null);
        }

        return new DocxOperationCapability(
            operation,
            "conditional",
            "Substring replacement applies to cell paragraphs in document order with occurrence counted across those paragraphs; expect-text guards the whole cell text. Check remains authoritative for the exact find span.",
            operation,
            null);
    }

    private static DocxOperationCapability SetCellShadingCapability(CellCapabilityFacts facts, string targetId)
    {
        const string operation = "set-cell-shading";
        if (facts.IsContinuation)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Unsupported merged-cell target " + Quote(targetId) + ". Only the vertical-merge root holds the merged value, so editing a continuation fails with E4301. Target the vertical-merge root cell instead.",
                operation,
                facts.RootCellId);
        }

        return new DocxOperationCapability(
            operation,
            "supported",
            "Cell shading updates the cell properties without touching cell text. Fill accepts a 6-digit hexadecimal color or auto, clear true removes shading, and expect-fill guards the current fill. Tracked modes record shading revision markup.",
            operation,
            null);
    }
}
