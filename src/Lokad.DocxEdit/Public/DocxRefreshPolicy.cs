namespace Lokad.DocxEdit;

/// <summary>
/// How a field result could be refreshed, as classified from the field code.
/// </summary>
/// <remarks>
/// <para>
/// The wire form is the lowercase lexical value; convert with
/// <see cref="DocxRefreshPolicyExtensions"/> at I/O boundaries.
/// </para>
/// </remarks>
public enum DocxRefreshPolicy
{
    /// <summary>Field type is not modeled for deterministic refresh.</summary>
    Unsupported = 0,

    /// <summary>Refreshes from an unambiguous same-part bookmark (REF family).</summary>
    SamePartBookmark = 1,

    /// <summary>Refreshes from literal field-code arguments (QUOTE).</summary>
    Literal = 2,

    /// <summary>Requires Word layout or pagination state.</summary>
    WordLayout = 3,

    /// <summary>Requires document property state.</summary>
    DocumentProperty = 4,

    /// <summary>Requires mail merge data or mail merge state.</summary>
    MailMergeData = 5,

    /// <summary>Requires Word formula evaluation.</summary>
    Formula = 6,

    /// <summary>Requires Word conditional field evaluation.</summary>
    Conditional = 7,

    /// <summary>Requires Word date/time evaluation.</summary>
    DateTimeEvaluation = 8,

    /// <summary>Requires hyperlink or external target state.</summary>
    External = 9
}

/// <summary>
/// Converts <see cref="DocxRefreshPolicy"/> to and from its lowercase wire form.
/// </summary>
public static class DocxRefreshPolicyExtensions
{
    /// <summary>Returns the lowercase wire form.</summary>
    public static string ToWireValue(this DocxRefreshPolicy policy)
    {
        return policy switch
        {
            DocxRefreshPolicy.Unsupported => "unsupported",
            DocxRefreshPolicy.SamePartBookmark => "same-part-bookmark",
            DocxRefreshPolicy.Literal => "literal",
            DocxRefreshPolicy.WordLayout => "word-layout",
            DocxRefreshPolicy.DocumentProperty => "document-property",
            DocxRefreshPolicy.MailMergeData => "mail-merge-data",
            DocxRefreshPolicy.Formula => "formula",
            DocxRefreshPolicy.Conditional => "conditional",
            DocxRefreshPolicy.DateTimeEvaluation => "date-time",
            DocxRefreshPolicy.External => "external",
            _ => throw new ArgumentOutOfRangeException(nameof(policy), $"Unsupported refresh policy '{policy}'.")
        };
    }

    /// <summary>Tries to parse the lowercase wire form with an ordinal, case-sensitive comparison.</summary>
    public static bool TryParseWireValue(string? value, out DocxRefreshPolicy policy)
    {
        switch (value)
        {
            case "unsupported":
                policy = DocxRefreshPolicy.Unsupported;
                return true;
            case "same-part-bookmark":
                policy = DocxRefreshPolicy.SamePartBookmark;
                return true;
            case "literal":
                policy = DocxRefreshPolicy.Literal;
                return true;
            case "word-layout":
                policy = DocxRefreshPolicy.WordLayout;
                return true;
            case "document-property":
                policy = DocxRefreshPolicy.DocumentProperty;
                return true;
            case "mail-merge-data":
                policy = DocxRefreshPolicy.MailMergeData;
                return true;
            case "conditional":
                policy = DocxRefreshPolicy.Conditional;
                return true;
            case "formula":
                policy = DocxRefreshPolicy.Formula;
                return true;
            case "date-time":
                policy = DocxRefreshPolicy.DateTimeEvaluation;
                return true;
            case "external":
                policy = DocxRefreshPolicy.External;
                return true;
            default:
                policy = DocxRefreshPolicy.Unsupported;
                return false;
        }
    }
}
