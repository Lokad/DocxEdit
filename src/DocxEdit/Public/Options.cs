namespace DocxEdit;

public sealed class DocxReadOptions
{
    public bool IncludeHeadersFooters { get; init; }
    public bool IncludeAllStories { get; init; }
    public DocxTextView TextView { get; init; } = DocxTextView.Final;
    public int MaxText { get; init; } = 4_000;
    public int MaxZipEntries { get; init; } = 10_000;
    public long MaxUncompressedBytes { get; init; } = 512L * 1024L * 1024L;
    public long MaxSinglePartBytes { get; init; } = 128L * 1024L * 1024L;
    public bool LeaveInputOpen { get; init; } = true;
}

public sealed class DocxOutlineOptions
{
    public bool IncludeHeadersFooters { get; init; }
    public int MaxZipEntries { get; init; } = 10_000;
    public long MaxUncompressedBytes { get; init; } = 512L * 1024L * 1024L;
    public long MaxSinglePartBytes { get; init; } = 128L * 1024L * 1024L;
    public bool LeaveInputOpen { get; init; } = true;
}

public sealed class DocxFindOptions
{
    public bool IncludeHeadersFooters { get; init; }
    public DocxTextView TextView { get; init; } = DocxTextView.Final;
    public int MaxText { get; init; } = 4_000;
    public int MaxZipEntries { get; init; } = 10_000;
    public long MaxUncompressedBytes { get; init; } = 512L * 1024L * 1024L;
    public long MaxSinglePartBytes { get; init; } = 128L * 1024L * 1024L;
    public bool LeaveInputOpen { get; init; } = true;
}

public sealed class DocxDumpOptions
{
    public bool IncludeRuns { get; init; }
    public DocxTextView TextView { get; init; } = DocxTextView.Final;
    public int MaxText { get; init; } = 4_000;
    public int MaxZipEntries { get; init; } = 10_000;
    public long MaxUncompressedBytes { get; init; } = 512L * 1024L * 1024L;
    public long MaxSinglePartBytes { get; init; } = 128L * 1024L * 1024L;
    public bool LeaveInputOpen { get; init; } = true;
}

public sealed class DocxContextOptions
{
    public bool IncludeHeadersFooters { get; init; }
    public DocxTextView TextView { get; init; } = DocxTextView.Final;
    public int Radius { get; init; } = 1;
    public int MaxText { get; init; }
    public int MaxZipEntries { get; init; } = 10_000;
    public long MaxUncompressedBytes { get; init; } = 512L * 1024L * 1024L;
    public long MaxSinglePartBytes { get; init; } = 128L * 1024L * 1024L;
    public bool LeaveInputOpen { get; init; } = true;
}

public sealed class DocxStylesOptions
{
    public int MaxZipEntries { get; init; } = 10_000;
    public long MaxUncompressedBytes { get; init; } = 512L * 1024L * 1024L;
    public long MaxSinglePartBytes { get; init; } = 128L * 1024L * 1024L;
    public bool LeaveInputOpen { get; init; } = true;
}

public sealed class DocxMediaOptions
{
    public int MaxZipEntries { get; init; } = 10_000;
    public long MaxUncompressedBytes { get; init; } = 512L * 1024L * 1024L;
    public long MaxSinglePartBytes { get; init; } = 128L * 1024L * 1024L;
    public bool LeaveInputOpen { get; init; } = true;
}

public sealed class DocxValidateOptions
{
    public DocxValidationProfile Profile { get; init; } = DocxValidationProfile.Structural;
    public int MaxDiagnostics { get; init; } = 500;
    public int MaxZipEntries { get; init; } = 10_000;
    public long MaxUncompressedBytes { get; init; } = 512L * 1024L * 1024L;
    public long MaxSinglePartBytes { get; init; } = 128L * 1024L * 1024L;
    public bool LeaveInputOpen { get; init; } = true;
}

public sealed class DocxChangesOptions
{
    public bool IncludeCommentText { get; init; }
    public int MaxCommentText { get; init; } = 240;
    public int MaxZipEntries { get; init; } = 10_000;
    public long MaxUncompressedBytes { get; init; } = 512L * 1024L * 1024L;
    public long MaxSinglePartBytes { get; init; } = 128L * 1024L * 1024L;
    public bool LeaveInputOpen { get; init; } = true;
}

public sealed class DocxPatchParseOptions
{
}

public sealed class DocxEditOptions
{
    public TrackChangesMode TrackChanges { get; init; } = TrackChangesMode.Off;
    public string Author { get; init; } = "docxedit";
    public DateTimeOffset TimestampUtc { get; init; } = DateTimeOffset.UtcNow;

    public IDocxAssetProvider? AssetProvider { get; init; }

    public bool LeaveInputOpen { get; init; } = true;
    public bool LeaveOutputOpen { get; init; } = true;

    public int MaxZipEntries { get; init; } = 10_000;
    public long MaxUncompressedBytes { get; init; } = 512L * 1024L * 1024L;
    public long MaxSinglePartBytes { get; init; } = 128L * 1024L * 1024L;

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
