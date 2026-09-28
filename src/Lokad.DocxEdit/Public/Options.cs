namespace Lokad.DocxEdit;

/// <summary>
/// Shared package-load quotas. Every per-operation options class composes this
/// value as <c>Quotas</c> so the limits stay single-sourced; stream ownership
/// stays an explicit boolean on each options class.
/// </summary>
public sealed record DocxPackageLimits(
    int MaxZipEntries,
    long MaxUncompressedBytes,
    long MaxSinglePartBytes)
{
    /// <summary>Default quotas: 10,000 ZIP entries, 512MB total, 128MB per part.</summary>
    public static DocxPackageLimits Default { get; } = new(10_000, 512L * 1024L * 1024L, 128L * 1024L * 1024L);
}

/// <summary>Options for reading a document.</summary>
public sealed class DocxReadOptions
{
    /// <summary>Whether to include header/footer stories.</summary>
    public bool IncludeHeadersFooters { get; init; }
    /// <summary>Which text view to read.</summary>
    public DocxTextView TextView { get; init; } = DocxTextView.Final;
    /// <summary>Maximum body-text characters per item in text and structured/JSON output; 0 drops body text. IDs, counts, style IDs/names, bookmark names, content-control tags/aliases, field codes/kinds, hyperlink URIs/anchors, authors, and revision IDs are retained as structural metadata. Raising it increases peak output size.</summary>
    public int MaxText { get; init; } = 4_000;
    /// <summary>Package quotas: ZIP entry count, total uncompressed bytes, and single-part bytes.</summary>
    public DocxPackageLimits Quotas { get; init; } = DocxPackageLimits.Default;
    /// <summary>Whether to leave the input stream open; when false the input is disposed on every exit (success, failure, or cancellation).</summary>
    public bool LeaveInputOpen { get; init; } = true;
}

/// <summary>Options for outlining document structure.</summary>
public sealed class DocxOutlineOptions
{
    /// <summary>Whether to include header/footer stories.</summary>
    public bool IncludeHeadersFooters { get; init; }
    /// <summary>Which text view to read.</summary>
    public DocxTextView TextView { get; init; } = DocxTextView.Final;
    /// <summary>Maximum body-text characters per item in text and structured/JSON output; 0 drops body text. IDs, counts, style IDs/names, bookmark names, content-control tags/aliases, field codes/kinds, hyperlink URIs/anchors, authors, and revision IDs are retained as structural metadata. Raising it increases peak output size.</summary>
    public int MaxText { get; init; } = 4_000;
    /// <summary>Package quotas: ZIP entry count, total uncompressed bytes, and single-part bytes.</summary>
    public DocxPackageLimits Quotas { get; init; } = DocxPackageLimits.Default;
    /// <summary>Whether to leave the input stream open; when false the input is disposed on every exit (success, failure, or cancellation).</summary>
    public bool LeaveInputOpen { get; init; } = true;
}

/// <summary>Options for text search.</summary>
public sealed class DocxFindOptions
{
    /// <summary>Whether to include header/footer stories.</summary>
    public bool IncludeHeadersFooters { get; init; }
    /// <summary>Which text view to read.</summary>
    public DocxTextView TextView { get; init; } = DocxTextView.Final;
    /// <summary>Maximum body-text characters per item in text and structured/JSON output; 0 drops body text. IDs, counts, style IDs/names, bookmark names, content-control tags/aliases, field codes/kinds, hyperlink URIs/anchors, authors, and revision IDs are retained as structural metadata. Raising it increases peak output size.</summary>
    public int MaxText { get; init; } = 4_000;
    /// <summary>Package quotas: ZIP entry count, total uncompressed bytes, and single-part bytes.</summary>
    public DocxPackageLimits Quotas { get; init; } = DocxPackageLimits.Default;
    /// <summary>Whether to leave the input stream open; when false the input is disposed on every exit (success, failure, or cancellation).</summary>
    public bool LeaveInputOpen { get; init; } = true;
}

/// <summary>Options for dumping one target.</summary>
public sealed class DocxDumpOptions
{
    /// <summary>Whether to include run-level markup detail.</summary>
    public bool IncludeRuns { get; init; }
    /// <summary>Which text view to read.</summary>
    public DocxTextView TextView { get; init; } = DocxTextView.Final;
    /// <summary>Maximum body-text characters per item in text and structured/JSON output; 0 drops body text. IDs, counts, style IDs/names, bookmark names, content-control tags/aliases, field codes/kinds, hyperlink URIs/anchors, authors, and revision IDs are retained as structural metadata. Raising it increases peak output size.</summary>
    public int MaxText { get; init; } = 4_000;
    /// <summary>Package quotas: ZIP entry count, total uncompressed bytes, and single-part bytes.</summary>
    public DocxPackageLimits Quotas { get; init; } = DocxPackageLimits.Default;
    /// <summary>Whether to leave the input stream open; when false the input is disposed on every exit (success, failure, or cancellation).</summary>
    public bool LeaveInputOpen { get; init; } = true;
}

/// <summary>Options for target context.</summary>
public sealed class DocxContextOptions
{
    /// <summary>Whether to include header/footer stories.</summary>
    public bool IncludeHeadersFooters { get; init; }
    /// <summary>Which text view to read.</summary>
    public DocxTextView TextView { get; init; } = DocxTextView.Final;
    /// <summary>Neighboring targets shown on each side (clamped at 0). Raising it widens context output.</summary>
    public int Radius { get; init; } = 1;
    /// <summary>Maximum body-text characters per item in text and structured/JSON output; 0 drops body text. IDs, counts, style IDs/names, bookmark names, content-control tags/aliases, field codes/kinds, hyperlink URIs/anchors, authors, and revision IDs are retained as structural metadata. Raising it increases peak output size.</summary>
    public int MaxText { get; init; }
    /// <summary>Package quotas: ZIP entry count, total uncompressed bytes, and single-part bytes.</summary>
    public DocxPackageLimits Quotas { get; init; } = DocxPackageLimits.Default;
    /// <summary>Whether to leave the input stream open; when false the input is disposed on every exit (success, failure, or cancellation).</summary>
    public bool LeaveInputOpen { get; init; } = true;
}

/// <summary>Options for the style inventory.</summary>
public sealed class DocxStylesOptions
{
    /// <summary>Package quotas: ZIP entry count, total uncompressed bytes, and single-part bytes.</summary>
    public DocxPackageLimits Quotas { get; init; } = DocxPackageLimits.Default;
    /// <summary>Whether to leave the input stream open; when false the input is disposed on every exit (success, failure, or cancellation).</summary>
    public bool LeaveInputOpen { get; init; } = true;
}

/// <summary>Options for the image inventory.</summary>
public sealed class DocxMediaOptions
{
    /// <summary>Package quotas: ZIP entry count, total uncompressed bytes, and single-part bytes.</summary>
    public DocxPackageLimits Quotas { get; init; } = DocxPackageLimits.Default;
    /// <summary>Whether to leave the input stream open; when false the input is disposed on every exit (success, failure, or cancellation).</summary>
    public bool LeaveInputOpen { get; init; } = true;
}

/// <summary>Options for package validation.</summary>
public sealed class DocxValidateOptions
{
    /// <summary>Validation profile.</summary>
    public DocxValidationProfile Profile { get; init; } = DocxValidationProfile.Structural;
    /// <summary>Maximum diagnostics returned; lowering it hides findings behind the cap markers sooner.</summary>
    public int MaxDiagnostics { get; init; } = 500;
    /// <summary>Package quotas: ZIP entry count, total uncompressed bytes, and single-part bytes.</summary>
    public DocxPackageLimits Quotas { get; init; } = DocxPackageLimits.Default;
    /// <summary>Whether to leave the input stream open; when false the input is disposed on every exit (success, failure, or cancellation).</summary>
    public bool LeaveInputOpen { get; init; } = true;
}

/// <summary>Options for target editing capabilities.</summary>
public sealed class DocxCapabilitiesOptions
{
    /// <summary>Effective track-change policy the capabilities describe.</summary>
    public TrackChangesMode TrackChanges { get; init; } = TrackChangesMode.Off;
    /// <summary>Package quotas: ZIP entry count, total uncompressed bytes, and single-part bytes.</summary>
    public DocxPackageLimits Quotas { get; init; } = DocxPackageLimits.Default;
    /// <summary>Whether to leave the input stream open; when false the input is disposed on every exit (success, failure, or cancellation).</summary>
    public bool LeaveInputOpen { get; init; } = true;
}

/// <summary>Options for a guarded patch template.</summary>
public sealed class DocxTemplateOptions
{
    /// <summary>Effective track-change policy the template is generated under; the active block passes check under this policy.</summary>
    public TrackChangesMode TrackChanges { get; init; } = TrackChangesMode.Off;
    /// <summary>Package quotas: ZIP entry count, total uncompressed bytes, and single-part bytes.</summary>
    public DocxPackageLimits Quotas { get; init; } = DocxPackageLimits.Default;
    /// <summary>Whether to leave the input stream open; when false the input is disposed on every exit (success, failure, or cancellation).</summary>
    public bool LeaveInputOpen { get; init; } = true;
}

/// <summary>Options for the change inventory.</summary>
public sealed class DocxChangesOptions
{
    /// <summary>Whether to include comment text.</summary>
    public bool IncludeCommentText { get; init; }
    /// <summary>Maximum comment text characters; snippets truncate beyond it.</summary>
    public int MaxCommentText { get; init; } = 240;
    /// <summary>Prior operation reports used to annotate changes. They must come from the check/apply run that produced the scanned document: annotation joins on revision IDs alone, which do not establish document provenance.</summary>
    public IReadOnlyList<DocxPatchOperationReport> OperationReports { get; init; } = [];
    /// <summary>Package quotas: ZIP entry count, total uncompressed bytes, and single-part bytes.</summary>
    public DocxPackageLimits Quotas { get; init; } = DocxPackageLimits.Default;
    /// <summary>Whether to leave the input stream open; when false the input is disposed on every exit (success, failure, or cancellation).</summary>
    public bool LeaveInputOpen { get; init; } = true;
}

/// <summary>Options for patch check and apply.</summary>
public sealed class DocxEditOptions
{
    /// <summary>Track-change output mode.</summary>
    public TrackChangesMode TrackChanges { get; init; } = TrackChangesMode.Off;
    /// <summary>Requested author for generated revisions. Leading and trailing whitespace is trimmed before use; check/apply results carry the trimmed (effective) value.</summary>
    public string Author { get; init; } = "docxedit";
    /// <summary>Requested timestamp for generated revisions. Converted to UTC before use; check/apply results carry the UTC (effective) value.</summary>
    public DateTimeOffset TimestampUtc { get; init; } = DateTimeOffset.UtcNow;

    /// <summary>Resolves image asset references for image operations.</summary>
    public IDocxAssetProvider? AssetProvider { get; init; }

    /// <summary>Package quotas: ZIP entry count, total uncompressed bytes, and single-part bytes.</summary>
    public DocxPackageLimits Quotas { get; init; } = DocxPackageLimits.Default;
    /// <summary>Whether to leave the input stream open; when false the input is disposed on every exit (success, failure, or cancellation).</summary>
    public bool LeaveInputOpen { get; init; } = true;
    /// <summary>Whether to leave the output stream open; when false the output is disposed on every exit (success, failure, or cancellation).</summary>
    public bool LeaveOutputOpen { get; init; } = true;

    /// <summary>Maximum patch text characters accepted; the patch read is bounded before allocation and parsing.</summary>
    public int MaxPatchChars { get; init; } = 4_000_000;

    /// <summary>Whether macro-enabled documents load instead of failing.</summary>
    public bool AllowMacroEnabledDocuments { get; init; }
    /// <summary>Whether successful edits mark fields dirty.</summary>
    public bool MarkFieldsDirtyWhenEditing { get; init; } = true;

    /// <summary>Maximum before/after preview characters per side in operation reports; 0 disables previews and keeps reports metadata-only. Raising it exposes document text in reports.</summary>
    public int MaxPreviewChars { get; init; }

    /// <summary>Copies these options with revision metadata normalized the way generated revisions record it: trimmed author and UTC timestamp.</summary>
    internal DocxEditOptions WithNormalizedRevisionMetadata()
    {
        return new DocxEditOptions
        {
            TrackChanges = TrackChanges,
            Author = Author.Trim(),
            TimestampUtc = TimestampUtc.ToUniversalTime(),
            AssetProvider = AssetProvider,
            LeaveInputOpen = LeaveInputOpen,
            LeaveOutputOpen = LeaveOutputOpen,
            Quotas = Quotas,
            MaxPatchChars = MaxPatchChars,
            AllowMacroEnabledDocuments = AllowMacroEnabledDocuments,
            MarkFieldsDirtyWhenEditing = MarkFieldsDirtyWhenEditing,
            MaxPreviewChars = MaxPreviewChars
        };
    }
}

/// <summary>How edits interact with tracked-change markup.</summary>
public enum TrackChangesMode
{
    /// <summary>Apply direct edits.</summary>
    Off,
    /// <summary>Apply direct edits while preserving existing tracked-change markup where possible.</summary>
    Preserve,
    /// <summary>Generate new revision markup for supported operations; warn and apply directly otherwise.</summary>
    Suggest,
    /// <summary>Require generated revision markup; fail unsupported operations.</summary>
    Require
}

/// <summary>Which text to read.</summary>
public enum DocxTextView
{
    /// <summary>Accepted visible text (default).</summary>
    Final,
    /// <summary>Text before tracked insertions/deletions.</summary>
    Original,
    /// <summary>Text with lightweight insertion/deletion markers.</summary>
    Markup
}

/// <summary>How much to validate.</summary>
public enum DocxValidationProfile
{
    /// <summary>Package checks plus structural invariants (default).</summary>
    Structural,
    /// <summary>Package checks only.</summary>
    Package
}
