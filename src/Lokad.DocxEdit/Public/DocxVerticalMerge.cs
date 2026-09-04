namespace Lokad.DocxEdit;

/// <summary>
/// Vertical-merge state of a table cell, read from <c>w:vMerge</c>.
/// </summary>
/// <remarks>
/// <para>
/// The wire form is the lowercase lexical value; convert with
/// <see cref="DocxVerticalMergeExtensions"/> at I/O boundaries. Absence
/// (no <c>w:vMerge</c> element) is represented by null, never by a member.
/// </para>
/// </remarks>
public enum DocxVerticalMerge
{
    /// <summary>Merge chain root (<c>w:val="restart"</c>).</summary>
    Restart,

    /// <summary>Merge chain continuation: explicit <c>continue</c>, missing value, or unrecognized value.</summary>
    Continue
}

/// <summary>
/// Converts <see cref="DocxVerticalMerge"/> to and from its lowercase wire form.
/// </summary>
public static class DocxVerticalMergeExtensions
{
    /// <summary>Returns the lowercase wire form.</summary>
    public static string ToWireValue(this DocxVerticalMerge merge)
    {
        return merge switch
        {
            DocxVerticalMerge.Restart => "restart",
            DocxVerticalMerge.Continue => "continue",
            _ => throw new ArgumentOutOfRangeException(nameof(merge), $"Unsupported vertical merge '{merge}'."),
        };
    }

    /// <summary>Tries to parse the lowercase wire form with an ordinal, case-sensitive comparison.</summary>
    public static bool TryParseWireValue(string? value, out DocxVerticalMerge merge)
    {
        switch (value)
        {
            case "restart":
                merge = DocxVerticalMerge.Restart;
                return true;
            case "continue":
                merge = DocxVerticalMerge.Continue;
                return true;
            default:
                merge = DocxVerticalMerge.Continue;
                return false;
        }
    }
}
