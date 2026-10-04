namespace Lokad.DocxEdit;

/// <summary>Paper preset for a newly created document, independent of orientation.</summary>
public enum DocxPaperSize
{
    /// <summary>A4: 210 by 297 millimeters (11906 by 16838 twips).</summary>
    A4 = 0,
    /// <summary>US Letter: 8.5 by 11 inches (12240 by 15840 twips).</summary>
    Letter = 1
}

/// <summary>Converts paper presets to and from their lowercase wire values.</summary>
public static class DocxPaperSizeExtensions
{
    /// <summary>Returns a4 or letter; throws for an unsupported enum value.</summary>
    public static string ToWireValue(this DocxPaperSize paperSize) => paperSize switch
    {
        DocxPaperSize.A4 => "a4",
        DocxPaperSize.Letter => "letter",
        _ => throw new ArgumentOutOfRangeException(nameof(paperSize))
    };

    /// <summary>Parses a4 or letter with an ordinal, case-sensitive comparison.</summary>
    public static bool TryParseWireValue(string? value, out DocxPaperSize paperSize)
    {
        paperSize = value == "letter" ? DocxPaperSize.Letter : DocxPaperSize.A4;
        return value is "a4" or "letter";
    }
}
