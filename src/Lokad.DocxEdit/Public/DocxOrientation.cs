namespace Lokad.DocxEdit;

/// <summary>
/// Page orientation of a section, read from the OOXML <c>w:pgSz w:orient</c> attribute.
/// </summary>
/// <remarks>
/// <para>
/// The wire form is always the lowercase OOXML lexical value (<c>portrait</c> or
/// <c>landscape</c>); convert with <see cref="DocxOrientationExtensions"/> at I/O
/// boundaries so text, JSON, and patch output stay string-shaped.
/// </para>
/// </remarks>
public enum DocxOrientation
{
    /// <summary>
    /// Upright page orientation; also the OOXML default when <c>w:orient</c> is absent.
    /// </summary>
    Portrait = 0,

    /// <summary>
    /// Rotated page orientation.
    /// </summary>
    Landscape = 1
}

/// <summary>
/// Converts <see cref="DocxOrientation"/> to and from its lowercase wire form.
/// </summary>
public static class DocxOrientationExtensions
{
    /// <summary>
    /// Returns the lowercase wire form (<c>portrait</c> or <c>landscape</c>).
    /// </summary>
    public static string ToWireValue(this DocxOrientation orientation)
    {
        return orientation switch
        {
            DocxOrientation.Portrait => "portrait",
            DocxOrientation.Landscape => "landscape",
            _ => throw new ArgumentOutOfRangeException(nameof(orientation), $"Unsupported orientation '{orientation}'.")
        };
    }

    /// <summary>
    /// Tries to parse the lowercase wire form with an ordinal, case-sensitive comparison.
    /// A null or unrecognized value returns <c>false</c> and yields
    /// <see cref="DocxOrientation.Portrait"/>.
    /// </summary>
    public static bool TryParseWireValue(string? value, out DocxOrientation orientation)
    {
        switch (value)
        {
            case "portrait":
                orientation = DocxOrientation.Portrait;
                return true;
            case "landscape":
                orientation = DocxOrientation.Landscape;
                return true;
            default:
                orientation = DocxOrientation.Portrait;
                return false;
        }
    }
}
