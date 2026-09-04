namespace Lokad.DocxEdit;

/// <summary>
/// How far a numbering label got resolved.
/// </summary>
/// <remarks>
/// <para>
/// The wire form is the lowercase lexical value; convert with
/// <see cref="DocxLabelStatusExtensions"/> at I/O boundaries.
/// </para>
/// </remarks>
public enum DocxLabelStatus
{
    /// <summary>Labeling did not run for this list.</summary>
    NotResolved = 0,

    /// <summary>Label text could not be produced.</summary>
    Unsupported = 1,

    /// <summary>Label text fully resolved.</summary>
    Resolved = 2,

    /// <summary>Label text resolved with warnings.</summary>
    Partial = 3
}

/// <summary>
/// Converts <see cref="DocxLabelStatus"/> to and from its lowercase wire form.
/// </summary>
public static class DocxLabelStatusExtensions
{
    /// <summary>Returns the lowercase wire form.</summary>
    public static string ToWireValue(this DocxLabelStatus status)
    {
        return status switch
        {
            DocxLabelStatus.NotResolved => "not-resolved",
            DocxLabelStatus.Unsupported => "unsupported",
            DocxLabelStatus.Resolved => "resolved",
            DocxLabelStatus.Partial => "partial",
            _ => throw new ArgumentOutOfRangeException(nameof(status), $"Unsupported label status '{status}'."),
        };
    }

    /// <summary>Tries to parse the lowercase wire form with an ordinal, case-sensitive comparison.</summary>
    public static bool TryParseWireValue(string? value, out DocxLabelStatus status)
    {
        switch (value)
        {
            case "not-resolved":
                status = DocxLabelStatus.NotResolved;
                return true;
            case "unsupported":
                status = DocxLabelStatus.Unsupported;
                return true;
            case "resolved":
                status = DocxLabelStatus.Resolved;
                return true;
            case "partial":
                status = DocxLabelStatus.Partial;
                return true;
            default:
                status = DocxLabelStatus.NotResolved;
                return false;
        }
    }
}
