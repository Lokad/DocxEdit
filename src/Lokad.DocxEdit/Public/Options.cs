namespace Lokad.DocxEdit;

/// <summary>
/// Shared package-load quotas. The per-operation options classes below default
/// their quota properties from <see cref="Default"/> so the limits stay
/// single-sourced; the editor maps them to <c>OoxmlPackageOptions</c> through
/// one mapper.
/// </summary>
public sealed record DocxPackageLimits(
    int MaxZipEntries,
    long MaxUncompressedBytes,
    long MaxSinglePartBytes,
    bool LeaveInputOpen)
{
    /// <summary>Default quotas: 10,000 ZIP entries, 512MB total, 128MB per part, input left open.</summary>
    public static DocxPackageLimits Default { get; } = new(10_000, 512L * 1024L * 1024L, 128L * 1024L * 1024L, true);
}

/// <summary>
/// Options classes carrying package-load quotas. Internal mapping contract:
/// lets the single package-options mapper read the quotas without one overload
/// per options type.
/// </summary>
internal interface IHasPackageLimits
{
    int MaxZipEntries { get; }
    long MaxUncompressedBytes { get; }
    long MaxSinglePartBytes { get; }
    bool LeaveInputOpen { get; }
}

/// <summary>Options for reading a document.</summary>
public sealed class DocxReadOptions : IHasPackageLimits
{
    /// <summary>Whether to include header/footer stories.</summary>
    public bool IncludeHeadersFooters { get; init; }
    /// <summary>Whether to include header/footer stories (same effect as IncludeHeadersFooters).</summary>
    public bool IncludeAllStories { get; init; }
    /// <summary>Which text view to read.</summary>
    public DocxTextView TextView { get; init; } = DocxTextView.Final;
    /// <summary>Maximum text characters per item; 0 drops text.</summary>
    public int MaxText { get; init; } = 4_000;
    /// <summary>Maximum ZIP entries accepted.</summary>
    public int MaxZipEntries { get; init; } = DocxPackageLimits.Default.MaxZipEntries;
    /// <summary>Maximum total uncompressed bytes accepted.</summary>
    public long MaxUncompressedBytes { get; init; } = DocxPackageLimits.Default.MaxUncompressedBytes;
    /// <summary>Maximum single-part bytes accepted.</summary>
    public long MaxSinglePartBytes { get; init; } = DocxPackageLimits.Default.MaxSinglePartBytes;
    /// <summary>Whether to leave the input stream open.</summary>
    public bool LeaveInputOpen { get; init; } = DocxPackageLimits.Default.LeaveInputOpen;
}

/// <summary>Options for outlining document structure.</summary>
public sealed class DocxOutlineOptions : IHasPackageLimits
{
    /// <summary>Whether to include header/footer stories.</summary>
    public bool IncludeHeadersFooters { get; init; }
    /// <summary>Which text view to read.</summary>
    public DocxTextView TextView { get; init; } = DocxTextView.Final;
    /// <summary>Maximum ZIP entries accepted.</summary>
    public int MaxZipEntries { get; init; } = DocxPackageLimits.Default.MaxZipEntries;
    /// <summary>Maximum total uncompressed bytes accepted.</summary>
    public long MaxUncompressedBytes { get; init; } = DocxPackageLimits.Default.MaxUncompressedBytes;
    /// <summary>Maximum single-part bytes accepted.</summary>
    public long MaxSinglePartBytes { get; init; } = DocxPackageLimits.Default.MaxSinglePartBytes;
    /// <summary>Whether to leave the input stream open.</summary>
    public bool LeaveInputOpen { get; init; } = DocxPackageLimits.Default.LeaveInputOpen;
}

/// <summary>Options for text search.</summary>
public sealed class DocxFindOptions : IHasPackageLimits
{
    /// <summary>Whether to include header/footer stories.</summary>
    public bool IncludeHeadersFooters { get; init; }
    /// <summary>Which text view to read.</summary>
    public DocxTextView TextView { get; init; } = DocxTextView.Final;
    /// <summary>Maximum text characters per item; 0 drops text.</summary>
    public int MaxText { get; init; } = 4_000;
    /// <summary>Maximum ZIP entries accepted.</summary>
    public int MaxZipEntries { get; init; } = DocxPackageLimits.Default.MaxZipEntries;
    /// <summary>Maximum total uncompressed bytes accepted.</summary>
    public long MaxUncompressedBytes { get; init; } = DocxPackageLimits.Default.MaxUncompressedBytes;
    /// <summary>Maximum single-part bytes accepted.</summary>
    public long MaxSinglePartBytes { get; init; } = DocxPackageLimits.Default.MaxSinglePartBytes;
    /// <summary>Whether to leave the input stream open.</summary>
    public bool LeaveInputOpen { get; init; } = DocxPackageLimits.Default.LeaveInputOpen;
}

/// <summary>Options for dumping one target.</summary>
public sealed class DocxDumpOptions : IHasPackageLimits
{
    /// <summary>Whether to include run-level markup detail.</summary>
    public bool IncludeRuns { get; init; }
    /// <summary>Which text view to read.</summary>
    public DocxTextView TextView { get; init; } = DocxTextView.Final;
    /// <summary>Maximum text characters per item; 0 drops text.</summary>
    public int MaxText { get; init; } = 4_000;
    /// <summary>Maximum ZIP entries accepted.</summary>
    public int MaxZipEntries { get; init; } = DocxPackageLimits.Default.MaxZipEntries;
    /// <summary>Maximum total uncompressed bytes accepted.</summary>
    public long MaxUncompressedBytes { get; init; } = DocxPackageLimits.Default.MaxUncompressedBytes;
    /// <summary>Maximum single-part bytes accepted.</summary>
    public long MaxSinglePartBytes { get; init; } = DocxPackageLimits.Default.MaxSinglePartBytes;
    /// <summary>Whether to leave the input stream open.</summary>
    public bool LeaveInputOpen { get; init; } = DocxPackageLimits.Default.LeaveInputOpen;
}

/// <summary>Options for target context.</summary>
public sealed class DocxContextOptions : IHasPackageLimits
{
    /// <summary>Whether to include header/footer stories.</summary>
    public bool IncludeHeadersFooters { get; init; }
    /// <summary>Which text view to read.</summary>
    public DocxTextView TextView { get; init; } = DocxTextView.Final;
    /// <summary>Neighboring targets shown on each side (clamped at 0).</summary>
    public int Radius { get; init; } = 1;
    /// <summary>Maximum text characters per item; 0 drops text.</summary>
    public int MaxText { get; init; }
    /// <summary>Maximum ZIP entries accepted.</summary>
    public int MaxZipEntries { get; init; } = DocxPackageLimits.Default.MaxZipEntries;
    /// <summary>Maximum total uncompressed bytes accepted.</summary>
    public long MaxUncompressedBytes { get; init; } = DocxPackageLimits.Default.MaxUncompressedBytes;
    /// <summary>Maximum single-part bytes accepted.</summary>
    public long MaxSinglePartBytes { get; init; } = DocxPackageLimits.Default.MaxSinglePartBytes;
    /// <summary>Whether to leave the input stream open.</summary>
    public bool LeaveInputOpen { get; init; } = DocxPackageLimits.Default.LeaveInputOpen;
}

/// <summary>Options for the style inventory.</summary>
public sealed class DocxStylesOptions : IHasPackageLimits
{
    /// <summary>Maximum ZIP entries accepted.</summary>
    public int MaxZipEntries { get; init; } = DocxPackageLimits.Default.MaxZipEntries;
    /// <summary>Maximum total uncompressed bytes accepted.</summary>
    public long MaxUncompressedBytes { get; init; } = DocxPackageLimits.Default.MaxUncompressedBytes;
    /// <summary>Maximum single-part bytes accepted.</summary>
    public long MaxSinglePartBytes { get; init; } = DocxPackageLimits.Default.MaxSinglePartBytes;
    /// <summary>Whether to leave the input stream open.</summary>
    public bool LeaveInputOpen { get; init; } = DocxPackageLimits.Default.LeaveInputOpen;
}

/// <summary>Options for the image inventory.</summary>
public sealed class DocxMediaOptions : IHasPackageLimits
{
    /// <summary>Maximum ZIP entries accepted.</summary>
    public int MaxZipEntries { get; init; } = DocxPackageLimits.Default.MaxZipEntries;
    /// <summary>Maximum total uncompressed bytes accepted.</summary>
    public long MaxUncompressedBytes { get; init; } = DocxPackageLimits.Default.MaxUncompressedBytes;
    /// <summary>Maximum single-part bytes accepted.</summary>
    public long MaxSinglePartBytes { get; init; } = DocxPackageLimits.Default.MaxSinglePartBytes;
    /// <summary>Whether to leave the input stream open.</summary>
    public bool LeaveInputOpen { get; init; } = DocxPackageLimits.Default.LeaveInputOpen;
}

/// <summary>Options for package validation.</summary>
public sealed class DocxValidateOptions : IHasPackageLimits
{
    /// <summary>Validation profile.</summary>
    public DocxValidationProfile Profile { get; init; } = DocxValidationProfile.Structural;
    /// <summary>Maximum diagnostics returned; excess is capped with a warning.</summary>
    public int MaxDiagnostics { get; init; } = 500;
    /// <summary>Maximum ZIP entries accepted.</summary>
    public int MaxZipEntries { get; init; } = DocxPackageLimits.Default.MaxZipEntries;
    /// <summary>Maximum total uncompressed bytes accepted.</summary>
    public long MaxUncompressedBytes { get; init; } = DocxPackageLimits.Default.MaxUncompressedBytes;
    /// <summary>Maximum single-part bytes accepted.</summary>
    public long MaxSinglePartBytes { get; init; } = DocxPackageLimits.Default.MaxSinglePartBytes;
    /// <summary>Whether to leave the input stream open.</summary>
    public bool LeaveInputOpen { get; init; } = DocxPackageLimits.Default.LeaveInputOpen;
}

/// <summary>Options for the change inventory.</summary>
public sealed class DocxChangesOptions : IHasPackageLimits
{
    /// <summary>Whether to include comment text.</summary>
    public bool IncludeCommentText { get; init; }
    /// <summary>Maximum comment text characters.</summary>
    public int MaxCommentText { get; init; } = 240;
    /// <summary>Prior operation reports used to annotate changes.</summary>
    public IReadOnlyList<DocxPatchOperationReport> OperationReports { get; init; } = [];
    /// <summary>Maximum ZIP entries accepted.</summary>
    public int MaxZipEntries { get; init; } = DocxPackageLimits.Default.MaxZipEntries;
    /// <summary>Maximum total uncompressed bytes accepted.</summary>
    public long MaxUncompressedBytes { get; init; } = DocxPackageLimits.Default.MaxUncompressedBytes;
    /// <summary>Maximum single-part bytes accepted.</summary>
    public long MaxSinglePartBytes { get; init; } = DocxPackageLimits.Default.MaxSinglePartBytes;
    /// <summary>Whether to leave the input stream open.</summary>
    public bool LeaveInputOpen { get; init; } = DocxPackageLimits.Default.LeaveInputOpen;
}

/// <summary>Patch parsing options (reserved for future use).</summary>
public sealed class DocxPatchParseOptions
{
}

/// <summary>Options for patch check and apply.</summary>
public sealed class DocxEditOptions : IHasPackageLimits
{
    /// <summary>Track-change output mode.</summary>
    public TrackChangesMode TrackChanges { get; init; } = TrackChangesMode.Off;
    /// <summary>Author recorded on generated revisions.</summary>
    public string Author { get; init; } = "docxedit";
    /// <summary>Timestamp recorded on generated revisions.</summary>
    public DateTimeOffset TimestampUtc { get; init; } = DateTimeOffset.UtcNow;

    /// <summary>Resolves image asset references for image operations.</summary>
    public IDocxAssetProvider? AssetProvider { get; init; }

    /// <summary>Whether to leave the input stream open.</summary>
    public bool LeaveInputOpen { get; init; } = DocxPackageLimits.Default.LeaveInputOpen;
    /// <summary>Whether to leave the output stream open.</summary>
    public bool LeaveOutputOpen { get; init; } = true;

    /// <summary>Maximum ZIP entries accepted.</summary>
    public int MaxZipEntries { get; init; } = DocxPackageLimits.Default.MaxZipEntries;
    /// <summary>Maximum total uncompressed bytes accepted.</summary>
    public long MaxUncompressedBytes { get; init; } = DocxPackageLimits.Default.MaxUncompressedBytes;
    /// <summary>Maximum single-part bytes accepted.</summary>
    public long MaxSinglePartBytes { get; init; } = DocxPackageLimits.Default.MaxSinglePartBytes;

    /// <summary>Whether macro-enabled documents load instead of failing.</summary>
    public bool AllowMacroEnabledDocuments { get; init; }
    /// <summary>Whether successful edits mark fields dirty.</summary>
    public bool MarkFieldsDirtyWhenEditing { get; init; } = true;
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
