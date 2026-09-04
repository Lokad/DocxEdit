namespace Lokad.DocxEdit;

/// <summary>
/// Closed entity kinds addressable by an explicit target ID (SPEC section 8.2).
/// </summary>
internal enum DocxTargetKind
{
    Paragraph,
    Table,
    Row,
    Cell,
    MergeGroup,
    Image,
    Hyperlink,
    Field,
    Bookmark,
    ContentControl,
    Section
}

/// <summary>
/// Structured explicit target ID, parsed once instead of re-matching the raw string
/// at every resolution choke point.
/// </summary>
/// <remarks>
/// <para>
/// Parse only through <c>DocxPatchEngine.TryParseTargetId</c>, which delegates to the
/// established fixed-shape matchers, so acceptance is identical to matching the raw
/// string. <c>Primary</c> holds the entity ordinal (the table ordinal for
/// <c>Row</c>/<c>Cell</c>/<c>MergeGroup</c>); <c>Secondary</c> holds the row ordinal
/// for <c>Row</c>/<c>Cell</c> or the merge-group ordinal for <c>MergeGroup</c>;
/// <c>Tertiary</c> holds the cell ordinal for <c>Cell</c>. Unused slots are 0.
/// <c>Story</c> is the story prefix (<c>M</c>, <c>H</c>, <c>F</c>); <c>StoryPart</c>
/// is the 1-based header/footer part ordinal, 0 for the main story.
/// </para>
/// </remarks>
internal readonly record struct DocxTargetId(
    char Story,
    int StoryPart,
    DocxTargetKind Kind,
    int Primary,
    int Secondary,
    int Tertiary)
{
    /// <summary>
    /// Formats the canonical wire form (SPEC section 8.2).
    /// </summary>
    /// <remarks>
    /// <para>
    /// Counters zero-pad exactly like the scanners, so canonical IDs round-trip.
    /// Non-canonical parses admitted by the fixed-shape matchers (for example
    /// sign-carrying ordinals) do not necessarily reproduce their input.
    /// </para>
    /// </remarks>
    /// <summary>
    /// Containing-table ID for <c>Row</c>, <c>Cell</c>, and <c>MergeGroup</c>; meaningless otherwise.
    /// </summary>
    public DocxTargetId TableId => new DocxTargetId(Story, StoryPart, DocxTargetKind.Table, Primary, 0, 0);

    /// <summary>
    /// Containing-row ID for <c>Cell</c>; meaningless otherwise.
    /// </summary>
    public DocxTargetId RowId => new DocxTargetId(Story, StoryPart, DocxTargetKind.Row, Primary, Secondary, 0);

    public string ToWireValue()
    {
        string head = Story == 'M' ? "M" : $"{Story}{StoryPart:D3}";
        return Kind switch
        {
            DocxTargetKind.Paragraph => $"{head}.P{Primary:D4}",
            DocxTargetKind.Table => $"{head}.T{Primary:D4}",
            DocxTargetKind.Row => $"{head}.T{Primary:D4}.R{Secondary:D2}",
            DocxTargetKind.Cell => $"{head}.T{Primary:D4}.R{Secondary:D2}.C{Tertiary:D2}",
            DocxTargetKind.MergeGroup => $"{head}.T{Primary:D4}.MG{Secondary:D4}",
            DocxTargetKind.Image => $"{head}.I{Primary:D4}",
            DocxTargetKind.Hyperlink => $"{head}.L{Primary:D4}",
            DocxTargetKind.Field => $"{head}.F{Primary:D4}",
            DocxTargetKind.Bookmark => $"{head}.B{Primary:D4}",
            DocxTargetKind.ContentControl => $"{head}.CC{Primary:D4}",
            DocxTargetKind.Section => $"{head}.S{Primary:D4}",
            _ => throw new ArgumentOutOfRangeException(nameof(Kind), $"Unsupported target kind '{Kind}'.")
        };
    }
}
