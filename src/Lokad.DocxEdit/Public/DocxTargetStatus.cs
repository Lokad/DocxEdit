namespace Lokad.DocxEdit;

/// <summary>
/// How a tracked change is anchored to an edit target.
/// </summary>
/// <remarks>
/// <para>
/// The wire form is the lowercase lexical value; convert with
/// <see cref="DocxTargetStatusExtensions"/> at I/O boundaries.
/// </para>
/// </remarks>
public enum DocxTargetStatus
{
    /// <summary>No modeled target; see <c>TargetReason</c> for why.</summary>
    Targetless = 0,

    /// <summary>Anchored to a modeled paragraph, table, cell, or section target.</summary>
    Targeted = 1,

    /// <summary>Linked through matching comment anchor metadata.</summary>
    CommentAnchor = 2
}

/// <summary>
/// Converts <see cref="DocxTargetStatus"/> to and from its lowercase wire form.
/// </summary>
public static class DocxTargetStatusExtensions
{
    /// <summary>Returns the lowercase wire form.</summary>
    public static string ToWireValue(this DocxTargetStatus status)
    {
        return status switch
        {
            DocxTargetStatus.Targetless => "targetless",
            DocxTargetStatus.Targeted => "targeted",
            DocxTargetStatus.CommentAnchor => "comment-anchor",
            _ => throw new ArgumentOutOfRangeException(nameof(status), $"Unsupported target status '{status}'.")
        };
    }

    /// <summary>Tries to parse the lowercase wire form with an ordinal, case-sensitive comparison.</summary>
    public static bool TryParseWireValue(string? value, out DocxTargetStatus status)
    {
        switch (value)
        {
            case "targetless":
                status = DocxTargetStatus.Targetless;
                return true;
            case "targeted":
                status = DocxTargetStatus.Targeted;
                return true;
            case "comment-anchor":
                status = DocxTargetStatus.CommentAnchor;
                return true;
            default:
                status = DocxTargetStatus.Targetless;
                return false;
        }
    }
}
