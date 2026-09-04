namespace Lokad.DocxEdit;

/// <summary>
/// Why a tracked change has no modeled target. Null on targeted records.</summary>
/// <remarks>
/// <para>
/// The wire form is the lowercase lexical value; convert with
/// <see cref="DocxTargetReasonExtensions"/> at I/O boundaries.
/// </para>
/// </remarks>
public enum DocxTargetReason
{
    /// <summary>Range boundary handled via the nearest adjacent modeled block.</summary>
    RangeBoundary = 0,

    /// <summary>Range boundary with no adjacent modeled block found.</summary>
    RangeBoundaryNoAdjacentTarget = 1,

    /// <summary>Markup is a direct child of the document body.</summary>
    BodyLevelMarkup = 2,

    /// <summary>Markup is in a comment story without a modeled target.</summary>
    CommentStory = 3,

    /// <summary>Markup is inside an unmodeled body structure.</summary>
    UnmodeledBodyStructure = 4,

    /// <summary>Markup is in an unmodeled Word XML part.</summary>
    UnmodeledWordPart = 5
}

/// <summary>
/// Converts <see cref="DocxTargetReason"/> to and from its lowercase wire form.
/// </summary>
public static class DocxTargetReasonExtensions
{
    /// <summary>Returns the lowercase wire form.</summary>
    public static string ToWireValue(this DocxTargetReason reason)
    {
        return reason switch
        {
            DocxTargetReason.RangeBoundary => "range-boundary",
            DocxTargetReason.RangeBoundaryNoAdjacentTarget => "range-boundary-no-adjacent-target",
            DocxTargetReason.BodyLevelMarkup => "body-level-markup",
            DocxTargetReason.CommentStory => "comment-story",
            DocxTargetReason.UnmodeledBodyStructure => "unmodeled-body-structure",
            DocxTargetReason.UnmodeledWordPart => "unmodeled-word-part",
            _ => throw new ArgumentOutOfRangeException(nameof(reason), $"Unsupported target reason '{reason}'.")
        };
    }

    /// <summary>Tries to parse the lowercase wire form with an ordinal, case-sensitive comparison.</summary>
    public static bool TryParseWireValue(string? value, out DocxTargetReason reason)
    {
        switch (value)
        {
            case "range-boundary":
                reason = DocxTargetReason.RangeBoundary;
                return true;
            case "range-boundary-no-adjacent-target":
                reason = DocxTargetReason.RangeBoundaryNoAdjacentTarget;
                return true;
            case "body-level-markup":
                reason = DocxTargetReason.BodyLevelMarkup;
                return true;
            case "comment-story":
                reason = DocxTargetReason.CommentStory;
                return true;
            case "unmodeled-body-structure":
                reason = DocxTargetReason.UnmodeledBodyStructure;
                return true;
            case "unmodeled-word-part":
                reason = DocxTargetReason.UnmodeledWordPart;
                return true;
            default:
                reason = DocxTargetReason.UnmodeledWordPart;
                return false;
        }
    }
}
