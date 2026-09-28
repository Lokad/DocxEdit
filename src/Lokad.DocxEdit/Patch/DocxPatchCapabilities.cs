using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

// D17: target-specific editing capabilities for paragraph targets. Every
// verdict below reuses the execution predicates from the patch engine, so
// guidance cannot disagree with check: protected boundaries, tracked shape
// checks, section properties, and orphaned ranges come from the same helpers
// that gate the Execute methods. Patch-dependent details such as the exact
// find span, replacement text, style name, or anchor-text span stay
// conditional until the exact patch is checked.
internal static partial class DocxPatchEngine
{
    internal sealed class ParagraphCapabilitiesOutcome
    {
        public DocxTargetCapabilities? Capabilities { get; init; }
        public DocxDiagnostic? Error { get; init; }
    }

    private sealed record ParagraphCapabilityFacts(
        bool HasProtected,
        string ProtectedFeature,
        bool HasUnsupportedRunContent,
        string UnsupportedRunContent,
        bool HasMixedRunProperties,
        bool HasSectionProperties,
        string? OrphanDescription,
        bool CurrentHasTrackedUnsupportedCharacters);

    internal static ParagraphCapabilitiesOutcome GetParagraphCapabilities(
        OoxmlPackage package,
        string requestedTargetId,
        TrackChangesMode mode,
        CancellationToken cancellationToken)
    {
        if (!DocxTargetId.TryParse(requestedTargetId, out DocxTargetId parsed))
        {
            return ParagraphCapabilitiesNotFound(requestedTargetId);
        }

        if (parsed.Kind != DocxTargetKind.Paragraph)
        {
            return GetNonParagraphCapabilities(package, parsed, requestedTargetId, mode, cancellationToken);
        }

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
            return ParagraphCapabilitiesNotFound(requestedTargetId);
        }

        OoxmlPart? part = package.GetPart(story.PartName);
        if (part is null)
        {
            return ParagraphCapabilitiesNotFound(requestedTargetId);
        }

        using Stream stream = part.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        XElement? root = document.Root;
        if (root is null)
        {
            return ParagraphCapabilitiesNotFound(requestedTargetId);
        }

        XElement container = root.Element(OoxmlNs.W + "body") ?? root;
        XElement? paragraph = DocxStoryBlocks.FindParagraphByPhysicalOrdinal(container, parsed.Primary);
        if (paragraph is null)
        {
            return ParagraphCapabilitiesNotFound(requestedTargetId);
        }

        bool hasProtected = TryGetProtectedTextEditFeature(paragraph, out string protectedFeature);
        bool hasUnsupportedRunContent = TryGetUnsupportedTrackedRunContent(paragraph, out string unsupportedRunContent);
        var facts = new ParagraphCapabilityFacts(
            hasProtected,
            protectedFeature,
            hasUnsupportedRunContent,
            unsupportedRunContent,
            HasMixedDirectTextRunProperties(paragraph),
            ParagraphHasSectionProperties(paragraph),
            FindOrphanedRangeBoundary(document, paragraph),
            TextContainsTrackedUnsupportedCharacters(ReadVisibleText(paragraph)));

        bool isTracked = mode is TrackChangesMode.Require or TrackChangesMode.Suggest;
        bool isRequire = mode == TrackChangesMode.Require;
        var operations = new List<DocxOperationCapability>
        {
            ReplaceTextCapability(facts, isTracked, isRequire),
            ReplaceParagraphCapability(facts, isTracked, isRequire),
            InsertCapability("insert-before", isTracked),
            InsertCapability("insert-after", isTracked),
            DeleteBlockCapability(facts, isTracked, isRequire),
            SetStyleCapability(isTracked),
            AddBookmarkCapability(facts, mode),
            AddCommentCapability(facts),
        };

        return new ParagraphCapabilitiesOutcome
        {
            Capabilities = new DocxTargetCapabilities(
                parsed.ToWireValue(),
                "paragraph",
                story.StoryLabel,
                operations)
        };
    }

    private static ParagraphCapabilitiesOutcome ParagraphCapabilitiesNotFound(string requestedTargetId)
    {
        return new ParagraphCapabilitiesOutcome
        {
            Error = new DocxDiagnostic(
                DocxSeverity.Error,
                "E1201",
                "Target " + Quote(requestedTargetId) + " was not found.") with
            {
                TargetId = requestedTargetId
            }
        };
    }

    private static string Quote(string value)
    {
        return ((char)39).ToString() + value + ((char)39).ToString();
    }

    private static DocxOperationCapability ReplaceTextCapability(ParagraphCapabilityFacts facts, bool isTracked, bool isRequire)
    {
        const string operation = "replace-text";
        if (facts.HasProtected)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Paragraph contains protected OOXML boundary " + Quote(facts.ProtectedFeature) + ". Replace-text for this target fails with E4305 in direct modes and E6002 under Require, because the edit would cross that boundary. Use insert-before or insert-after to add content without editing the protected span. Check remains authoritative for the exact find span.",
                operation,
                "insert-after");
        }

        if (isTracked && (facts.HasUnsupportedRunContent || facts.HasMixedRunProperties))
        {
            string detail = facts.HasUnsupportedRunContent
                ? "unsupported run content " + Quote(facts.UnsupportedRunContent)
                : "mixed direct run formatting";
            if (isRequire)
            {
                return new DocxOperationCapability(
                    operation,
                    "unsupported",
                    "Paragraph has " + detail + ", so tracked text replacement is not modeled and Require fails with E6002. Suggest falls back to a direct edit with W4002, and Off applies directly. Check remains authoritative for the exact find and replacement.",
                    operation,
                    null);
            }

            return new DocxOperationCapability(
                operation,
                "conditional",
                "Paragraph has " + detail + ", so tracked text replacement falls back to a direct edit with W4002. The exact find span and replacement decide success: spans that include tabs, line breaks, or non-text run content need preserve-runs false. Check remains authoritative.",
                operation,
                null);
        }

        if (facts.CurrentHasTrackedUnsupportedCharacters)
        {
            return new DocxOperationCapability(
                operation,
                "conditional",
                "Paragraph text contains tabs or line breaks. Replace-text for a span that avoids them succeeds; a span that includes them needs preserve-runs false, and a replacement with tabs or line breaks needs preserve-runs false. Check remains authoritative for the exact find span.",
                operation,
                null);
        }

        if (isTracked)
        {
            return new DocxOperationCapability(
                operation,
                "supported",
                isRequire
                    ? "Plain paragraph text. Simple find and replacement pairs emit tracked delete and insert markup; complex shapes fail with E6002 under Require. Check remains authoritative for the exact pair."
                    : "Plain paragraph text. Simple find and replacement pairs emit tracked delete and insert markup; replacements with tabs or line breaks fall back to a direct edit with W4002. Check remains authoritative for the exact pair.",
                operation,
                null);
        }

        return new DocxOperationCapability(
            operation,
            "supported",
            "Plain paragraph text. Replace-text with an unambiguous find span succeeds; use occurrence N for one match or occurrence all for every match. Style and unrelated content are preserved.",
            operation,
            null);
    }

    private static DocxOperationCapability ReplaceParagraphCapability(ParagraphCapabilityFacts facts, bool isTracked, bool isRequire)
    {
        const string operation = "replace-paragraph";
        if (facts.HasProtected)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Paragraph contains protected OOXML boundary " + Quote(facts.ProtectedFeature) + ". Whole-paragraph replacement would remove that boundary, so it fails with E4305 in direct modes and E6002 under Require. Use insert-before or insert-after to add a new paragraph instead. Check remains authoritative.",
                operation,
                "insert-after");
        }

        if (isTracked && (facts.HasUnsupportedRunContent || facts.HasMixedRunProperties || facts.CurrentHasTrackedUnsupportedCharacters))
        {
            string detail = facts.HasUnsupportedRunContent
                ? "unsupported run content " + Quote(facts.UnsupportedRunContent)
                : facts.HasMixedRunProperties
                    ? "mixed direct run formatting"
                    : "tabs or line breaks in the current text";
            if (isRequire)
            {
                return new DocxOperationCapability(
                    operation,
                    "unsupported",
                    "Paragraph has " + detail + ", so tracked whole-paragraph replacement is not modeled and Require fails with E6002. Suggest falls back to a direct rewrite with W4002, and Off rewrites directly. Check remains authoritative for the exact replacement.",
                    operation,
                    null);
            }

            return new DocxOperationCapability(
                operation,
                "conditional",
                "Paragraph has " + detail + ", so tracked whole-paragraph replacement falls back to a direct rewrite with W4002. The exact replacement decides whether tracked markup or a direct rewrite applies. Check remains authoritative.",
                operation,
                null);
        }

        if (isTracked)
        {
            return new DocxOperationCapability(
                operation,
                "supported",
                isRequire
                    ? "Plain paragraph text. Simple replacements emit whole-paragraph tracked markup, with paragraph property revision markup when a compatible style change is included; complex shapes fail with E6002 under Require. Check remains authoritative."
                    : "Plain paragraph text. Simple replacements emit whole-paragraph tracked markup, with paragraph property revision markup when a compatible style change is included; complex shapes fall back to a direct rewrite with W4002. Check remains authoritative.",
                operation,
                null);
        }

        return new DocxOperationCapability(
            operation,
            "supported",
            "Direct rewrite replaces the paragraph text and preserves paragraph properties; an optional style field sets the paragraph style in the same edit.",
            operation,
            null);
    }

    private static DocxOperationCapability InsertCapability(string operation, bool isTracked)
    {
        if (isTracked)
        {
            return new DocxOperationCapability(
                operation,
                "supported",
                "Structural insert next to this paragraph; the anchor paragraph is unchanged. Inserted text without tabs or line breaks is recorded as tracked insertion markup; other inserted text falls back to a direct insert with W4002 under Suggest and fails with E6002 under Require. Check remains authoritative for the exact text.",
                operation,
                null);
        }

        return new DocxOperationCapability(
            operation,
            "supported",
            "Structural insert next to this paragraph; the anchor paragraph is unchanged. Use style for the paragraph style of the new block and copy-paragraph-properties to inherit the anchor formatting.",
            operation,
            null);
    }

    private static DocxOperationCapability DeleteBlockCapability(ParagraphCapabilityFacts facts, bool isTracked, bool isRequire)
    {
        const string operation = "delete-block";
        if (facts.OrphanDescription is not null)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Deleting this paragraph would orphan " + facts.OrphanDescription + " outside the deleted element, so delete-block fails with E4305 in every mode. Delete the range first or choose another target.",
                operation,
                null);
        }

        string? blocker = null;
        if (facts.HasSectionProperties)
        {
            blocker = "carries section properties";
        }
        else if (facts.HasProtected)
        {
            blocker = "contains protected OOXML boundary " + Quote(facts.ProtectedFeature);
        }
        else if (facts.CurrentHasTrackedUnsupportedCharacters)
        {
            blocker = "contains tabs or line breaks";
        }

        if (blocker is not null && isRequire)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Paragraph " + blocker + ", so tracked deletion is not modeled and Require fails with E6002. Suggest falls back to a direct delete with W4002 and Off deletes directly, removing the paragraph together with the blocking content. Check remains authoritative.",
                operation,
                null);
        }

        if (blocker is not null)
        {
            return new DocxOperationCapability(
                operation,
                "supported",
                "Paragraph " + blocker + ", so Suggest falls back to a direct delete with W4002 while Off deletes directly; either removes the paragraph together with the blocking content. Check remains authoritative.",
                operation,
                null);
        }

        if (isTracked)
        {
            return new DocxOperationCapability(
                operation,
                "supported",
                "Simple paragraph text is recorded as tracked deletion markup; existing revisions are preserved.",
                operation,
                null);
        }

        string loss = facts.HasSectionProperties
            ? " The paragraph carries section properties, which are removed with it."
            : string.Empty;
        return new DocxOperationCapability(
            operation,
            "supported",
            "Direct delete removes the paragraph." + loss,
            operation,
            null);
    }

    private static DocxOperationCapability SetStyleCapability(bool isTracked)
    {
        const string operation = "set-style";
        if (isTracked)
        {
            return new DocxOperationCapability(
                operation,
                "supported",
                "Set-style updates the paragraph style reference and preserves paragraph text. The style value must resolve to a paragraph style by ID or by unique name, otherwise check fails with E7101, E7102, or E7103. The change is recorded with paragraph property revision markup.",
                operation,
                null);
        }

        return new DocxOperationCapability(
            operation,
            "supported",
            "Set-style updates the paragraph style reference and preserves paragraph text. The style value must resolve to a paragraph style by ID or by unique name, otherwise check fails with E7101, E7102, or E7103.",
            operation,
            null);
    }

    private static DocxOperationCapability AddBookmarkCapability(ParagraphCapabilityFacts facts, TrackChangesMode mode)
    {
        const string operation = "add-bookmark";
        if (!SupportsTrackedChangeOutput(operation) && mode == TrackChangesMode.Require && !IsAnnotationOperation(operation))
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Bookmark creation adds anchor metadata without a tracked revision representation, so Require fails with E6001 before editing. Use Suggest, which warns with W4001 and inserts directly, or Off for a direct insert. Check remains authoritative.",
                operation,
                null);
        }

        if (facts.HasProtected)
        {
            return new DocxOperationCapability(
                operation,
                "unsupported",
                "Paragraph contains protected OOXML boundary " + Quote(facts.ProtectedFeature) + ", so bookmark creation would cross it and fails with E4311. Bookmark names must also be unique within the part. Use add-comment to annotate this paragraph without bookmark markers.",
                operation,
                "add-comment");
        }

        if (mode == TrackChangesMode.Suggest)
        {
            return new DocxOperationCapability(
                operation,
                "supported",
                "Bookmark creation inserts a guarded paragraph bookmark with a unique name; names with whitespace fail with E4205 and duplicate names fail with E4311. Markers are metadata without tracked revision output, so Suggest warns with W4001 and inserts directly.",
                operation,
                null);
        }

        return new DocxOperationCapability(
            operation,
            "supported",
            "Bookmark creation inserts a guarded paragraph bookmark with a unique name; names with whitespace fail with E4205 and duplicate names fail with E4311. Markers are metadata without tracked revision output.",
            operation,
            null);
    }

    private static DocxOperationCapability AddCommentCapability(ParagraphCapabilityFacts facts)
    {
        const string operation = "add-comment";
        if (facts.HasProtected)
        {
            return new DocxOperationCapability(
                operation,
                "conditional",
                "Paragraph contains protected OOXML boundary " + Quote(facts.ProtectedFeature) + ". A whole-paragraph comment succeeds, but an anchored comment with anchor-text fails with E4305 when the selected range crosses that boundary. Comments are review markup and stay permitted under Require. Check remains authoritative for the exact anchor-text span.",
                operation,
                null);
        }

        if (facts.CurrentHasTrackedUnsupportedCharacters)
        {
            return new DocxOperationCapability(
                operation,
                "conditional",
                "Paragraph text contains tabs or line breaks. A whole-paragraph comment succeeds; an anchored comment span must use direct text only, so spans that include tabs or line breaks fail with E4317. Use occurrence to disambiguate repeated anchor text. Check remains authoritative.",
                operation,
                null);
        }

        return new DocxOperationCapability(
            operation,
            "supported",
            "A new comment anchors to this paragraph, or to one direct text span selected with anchor-text; use occurrence when the anchor text repeats. Comments are review markup and stay permitted under Require.",
            operation,
            null);
    }
}
