namespace Lokad.DocxEdit;

/// <summary>
/// Where a numbering label definition came from.
/// </summary>
/// <remarks>
/// <para>
/// The wire form is the lowercase lexical value; convert with
/// <see cref="DocxLabelSourceExtensions"/> at I/O boundaries.
/// </para>
/// </remarks>
public enum DocxLabelSource
{
    /// <summary>Numbering definition referenced directly.</summary>
    Direct = 0,

    /// <summary>Numbering definition from the paragraph style.</summary>
    Style = 1,

    /// <summary>Numbering definition inherited through a base style.</summary>
    StyleInherited = 2
}

/// <summary>
/// Converts <see cref="DocxLabelSource"/> to and from its lowercase wire form.
/// </summary>
public static class DocxLabelSourceExtensions
{
    /// <summary>Returns the lowercase wire form.</summary>
    public static string ToWireValue(this DocxLabelSource source)
    {
        return source switch
        {
            DocxLabelSource.Direct => "direct",
            DocxLabelSource.Style => "style",
            DocxLabelSource.StyleInherited => "style-inherited",
            _ => throw new ArgumentOutOfRangeException(nameof(source), $"Unsupported label source '{source}'."),
        };
    }

    /// <summary>Tries to parse the lowercase wire form with an ordinal, case-sensitive comparison.</summary>
    public static bool TryParseWireValue(string? value, out DocxLabelSource source)
    {
        switch (value)
        {
            case "direct":
                source = DocxLabelSource.Direct;
                return true;
            case "style":
                source = DocxLabelSource.Style;
                return true;
            case "style-inherited":
                source = DocxLabelSource.StyleInherited;
                return true;
            default:
                source = DocxLabelSource.Direct;
                return false;
        }
    }
}
