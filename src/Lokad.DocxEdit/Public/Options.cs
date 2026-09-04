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

public sealed class DocxReadOptions : IHasPackageLimits
{
    public bool IncludeHeadersFooters { get; init; }
    public bool IncludeAllStories { get; init; }
    public DocxTextView TextView { get; init; } = DocxTextView.Final;
    public int MaxText { get; init; } = 4_000;
    public int MaxZipEntries { get; init; } = DocxPackageLimits.Default.MaxZipEntries;
    public long MaxUncompressedBytes { get; init; } = DocxPackageLimits.Default.MaxUncompressedBytes;
    public long MaxSinglePartBytes { get; init; } = DocxPackageLimits.Default.MaxSinglePartBytes;
    public bool LeaveInputOpen { get; init; } = DocxPackageLimits.Default.LeaveInputOpen;
}

public sealed class DocxOutlineOptions : IHasPackageLimits
{
    public bool IncludeHeadersFooters { get; init; }
    public DocxTextView TextView { get; init; } = DocxTextView.Final;
    public int MaxZipEntries { get; init; } = DocxPackageLimits.Default.MaxZipEntries;
    public long MaxUncompressedBytes { get; init; } = DocxPackageLimits.Default.MaxUncompressedBytes;
    public long MaxSinglePartBytes { get; init; } = DocxPackageLimits.Default.MaxSinglePartBytes;
    public bool LeaveInputOpen { get; init; } = DocxPackageLimits.Default.LeaveInputOpen;
}

public sealed class DocxFindOptions : IHasPackageLimits
{
    public bool IncludeHeadersFooters { get; init; }
    public DocxTextView TextView { get; init; } = DocxTextView.Final;
    public int MaxText { get; init; } = 4_000;
    public int MaxZipEntries { get; init; } = DocxPackageLimits.Default.MaxZipEntries;
    public long MaxUncompressedBytes { get; init; } = DocxPackageLimits.Default.MaxUncompressedBytes;
    public long MaxSinglePartBytes { get; init; } = DocxPackageLimits.Default.MaxSinglePartBytes;
    public bool LeaveInputOpen { get; init; } = DocxPackageLimits.Default.LeaveInputOpen;
}

public sealed class DocxDumpOptions : IHasPackageLimits
{
    public bool IncludeRuns { get; init; }
    public DocxTextView TextView { get; init; } = DocxTextView.Final;
    public int MaxText { get; init; } = 4_000;
    public int MaxZipEntries { get; init; } = DocxPackageLimits.Default.MaxZipEntries;
    public long MaxUncompressedBytes { get; init; } = DocxPackageLimits.Default.MaxUncompressedBytes;
    public long MaxSinglePartBytes { get; init; } = DocxPackageLimits.Default.MaxSinglePartBytes;
    public bool LeaveInputOpen { get; init; } = DocxPackageLimits.Default.LeaveInputOpen;
}

public sealed class DocxContextOptions : IHasPackageLimits
{
    public bool IncludeHeadersFooters { get; init; }
    public DocxTextView TextView { get; init; } = DocxTextView.Final;
    public int Radius { get; init; } = 1;
    public int MaxText { get; init; }
    public int MaxZipEntries { get; init; } = DocxPackageLimits.Default.MaxZipEntries;
    public long MaxUncompressedBytes { get; init; } = DocxPackageLimits.Default.MaxUncompressedBytes;
    public long MaxSinglePartBytes { get; init; } = DocxPackageLimits.Default.MaxSinglePartBytes;
    public bool LeaveInputOpen { get; init; } = DocxPackageLimits.Default.LeaveInputOpen;
}

public sealed class DocxStylesOptions : IHasPackageLimits
{
    public int MaxZipEntries { get; init; } = DocxPackageLimits.Default.MaxZipEntries;
    public long MaxUncompressedBytes { get; init; } = DocxPackageLimits.Default.MaxUncompressedBytes;
    public long MaxSinglePartBytes { get; init; } = DocxPackageLimits.Default.MaxSinglePartBytes;
    public bool LeaveInputOpen { get; init; } = DocxPackageLimits.Default.LeaveInputOpen;
}

public sealed class DocxMediaOptions : IHasPackageLimits
{
    public int MaxZipEntries { get; init; } = DocxPackageLimits.Default.MaxZipEntries;
    public long MaxUncompressedBytes { get; init; } = DocxPackageLimits.Default.MaxUncompressedBytes;
    public long MaxSinglePartBytes { get; init; } = DocxPackageLimits.Default.MaxSinglePartBytes;
    public bool LeaveInputOpen { get; init; } = DocxPackageLimits.Default.LeaveInputOpen;
}

public sealed class DocxValidateOptions : IHasPackageLimits
{
    public DocxValidationProfile Profile { get; init; } = DocxValidationProfile.Structural;
    public int MaxDiagnostics { get; init; } = 500;
    public int MaxZipEntries { get; init; } = DocxPackageLimits.Default.MaxZipEntries;
    public long MaxUncompressedBytes { get; init; } = DocxPackageLimits.Default.MaxUncompressedBytes;
    public long MaxSinglePartBytes { get; init; } = DocxPackageLimits.Default.MaxSinglePartBytes;
    public bool LeaveInputOpen { get; init; } = DocxPackageLimits.Default.LeaveInputOpen;
}

public sealed class DocxChangesOptions : IHasPackageLimits
{
    public bool IncludeCommentText { get; init; }
    public int MaxCommentText { get; init; } = 240;
    public IReadOnlyList<DocxPatchOperationReport> OperationReports { get; init; } = [];
    public int MaxZipEntries { get; init; } = DocxPackageLimits.Default.MaxZipEntries;
    public long MaxUncompressedBytes { get; init; } = DocxPackageLimits.Default.MaxUncompressedBytes;
    public long MaxSinglePartBytes { get; init; } = DocxPackageLimits.Default.MaxSinglePartBytes;
    public bool LeaveInputOpen { get; init; } = DocxPackageLimits.Default.LeaveInputOpen;
}

public sealed class DocxPatchParseOptions
{
}

public sealed class DocxEditOptions : IHasPackageLimits
{
    public TrackChangesMode TrackChanges { get; init; } = TrackChangesMode.Off;
    public string Author { get; init; } = "docxedit";
    public DateTimeOffset TimestampUtc { get; init; } = DateTimeOffset.UtcNow;

    public IDocxAssetProvider? AssetProvider { get; init; }

    public bool LeaveInputOpen { get; init; } = DocxPackageLimits.Default.LeaveInputOpen;
    public bool LeaveOutputOpen { get; init; } = true;

    public int MaxZipEntries { get; init; } = DocxPackageLimits.Default.MaxZipEntries;
    public long MaxUncompressedBytes { get; init; } = DocxPackageLimits.Default.MaxUncompressedBytes;
    public long MaxSinglePartBytes { get; init; } = DocxPackageLimits.Default.MaxSinglePartBytes;

    public bool AllowMacroEnabledDocuments { get; init; }
    public bool MarkFieldsDirtyWhenEditing { get; init; } = true;
}

public enum TrackChangesMode
{
    Off,
    Preserve,
    Suggest,
    Require
}

public enum DocxTextView
{
    Final,
    Original,
    Markup
}

public enum DocxValidationProfile
{
    Structural,
    Package
}
