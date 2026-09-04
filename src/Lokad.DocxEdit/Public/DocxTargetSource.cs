namespace Lokad.DocxEdit;

/// <summary>
/// How a tracked-change target was resolved.
/// </summary>
/// <remarks>
/// <para>
/// The wire form is the lowercase lexical value; convert with
/// <see cref="DocxTargetSourceExtensions"/> at I/O boundaries.
/// </para>
/// </remarks>
public enum DocxTargetSource
{
    /// <summary>No source; the change is targetless.</summary>
    None = 0,

    /// <summary>Found via an ancestor of the markup element.</summary>
    Ancestor = 1,

    /// <summary>Nearest adjacent modeled block for a range boundary outside one.</summary>
    AdjacentRange = 2,

    /// <summary>Linked through matching comment anchor metadata.</summary>
    CommentAnchor = 3
}

/// <summary>
/// Converts <see cref="DocxTargetSource"/> to and from its lowercase wire form.
/// </summary>
public static class DocxTargetSourceExtensions
{
    /// <summary>Returns the lowercase wire form.</summary>
    public static string ToWireValue(this DocxTargetSource source)
    {
        return source switch
        {
            DocxTargetSource.None => "none",
            DocxTargetSource.Ancestor => "ancestor",
            DocxTargetSource.AdjacentRange => "adjacent-range",
            DocxTargetSource.CommentAnchor => "comment-anchor",
            _ => throw new ArgumentOutOfRangeException(nameof(source), $"Unsupported target source '{source}'.")
        };
    }

    /// <summary>Tries to parse the lowercase wire form with an ordinal, case-sensitive comparison.</summary>
    public static bool TryParseWireValue(string? value, out DocxTargetSource source)
    {
        switch (value)
        {
            case "none":
                source = DocxTargetSource.None;
                return true;
            case "ancestor":
                source = DocxTargetSource.Ancestor;
                return true;
            case "adjacent-range":
                source = DocxTargetSource.AdjacentRange;
                return true;
            case "comment-anchor":
                source = DocxTargetSource.CommentAnchor;
                return true;
            default:
                source = DocxTargetSource.None;
                return false;
        }
    }
}
