using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Lokad.DocxEdit;

/// <summary>
/// Stable change-record ID: the parsed form of one emitted change ID string
/// (<c>M.CH0001</c>, <c>H001.CH0002</c>, <c>C001.CH0001</c>).
/// </summary>
/// <remarks>
/// <para>
/// Unlike <see cref="DocxTargetId"/>, change IDs are never addressable: no
/// selector or patch field parses them. They identify <c>DocxChangeInfo</c>
/// records and link range starts to ends via <c>PairedChangeId</c>. The story
/// alphabet is wider than the target grammar: comments (<c>C</c>) and fallback
/// (<c>P</c>) parts emit changes but host no targets.
/// </para>
/// </remarks>
[JsonConverter(typeof(DocxChangeIdJsonConverter))]
public readonly record struct DocxChangeId(char Story, int StoryPart, int Ordinal)
{
    /// <summary>
    /// Parses the wire form: a main-story (<c>M.CH0001</c>) or part-story
    /// (<c>H001.CH0002</c>) head plus a zero-padded ordinal.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The ordinal is an ASCII-digit run with canonical minimum width 4 (story
    /// parts: minimum width 3) and no upper width limit, so producer output past
    /// a width boundary still parses. Ordinals must be at least 1: zero, signs,
    /// whitespace, and non-digits are rejected, as are short runs, unknown
    /// story letters, and trailing content.
    /// </para>
    /// </remarks>
    public static bool TryParse(string? changeId, out DocxChangeId parsed)
    {
        parsed = default;
        if (string.IsNullOrEmpty(changeId))
        {
            return false;
        }

        ReadOnlySpan<char> span = changeId.AsSpan();
        char story;
        int storyPart;
        if (span.Length >= 1 && span[0] == 'M')
        {
            story = 'M';
            storyPart = 0;
            span = span[1..];
        }
        else if (span.Length >= 1 && span[0] is 'C' or 'H' or 'F' or 'P')
        {
            story = span[0];
            ReadOnlySpan<char> tail = span[1..];
            if (!TakeOrdinal(ref tail, 3, out storyPart))
            {
                return false;
            }

            span = tail;
        }
        else
        {
            return false;
        }

        if (span.Length < 3 || span[0] != '.' || span[1] != 'C' || span[2] != 'H')
        {
            return false;
        }

        ReadOnlySpan<char> ordinal = span[3..];
        if (!TakeOrdinal(ref ordinal, 4, out int value))
        {
            return false;
        }

        if (!ordinal.IsEmpty)
        {
            return false;
        }

        parsed = new DocxChangeId(story, storyPart, value);
        return true;

        static bool TakeOrdinal(ref ReadOnlySpan<char> span, int minDigits, out int ordinal)
        {
            ordinal = 0;
            int width = 0;
            while (width < span.Length && char.IsAsciiDigit(span[width]))
            {
                width++;
            }

            if (width < minDigits)
            {
                return false;
            }

            if (!int.TryParse(span[..width], NumberStyles.Integer, CultureInfo.InvariantCulture, out ordinal) || ordinal < 1)
            {
                ordinal = 0;
                return false;
            }

            span = span[width..];
            return true;
        }
    }

    /// <summary>Returns the canonical wire form (same as <see cref="ToWireValue"/>).</summary>
    public override string ToString()
    {
        return ToWireValue();
    }

    /// <summary>
    /// Formats the canonical wire form: zero-padded exactly like the scanner,
    /// so parsed IDs round-trip.
    /// </summary>
    public string ToWireValue()
    {
        if (Story != 'M' && Story != 'H' && Story != 'F' && Story != 'C' && Story != 'P')
        {
            throw new ArgumentOutOfRangeException(nameof(Story), "Unsupported story. Expected 'M', 'H', 'F', 'C', or 'P'.");
        }

        if (Story == 'M' ? StoryPart != 0 : StoryPart < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(StoryPart), "The main story uses part 0; part stories start at 1.");
        }

        if (Ordinal < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(Ordinal), "Change ordinals start at 1.");
        }

        return Story == 'M' ? $"M.CH{Ordinal:D4}" : $"{Story}{StoryPart:D3}.CH{Ordinal:D4}";
    }
}

/// <summary>
/// Converts <see cref="DocxChangeId"/> to and from its wire form, so JSON
/// output stays string-shaped under both default and CLI options.
/// </summary>
public sealed class DocxChangeIdJsonConverter : JsonConverter<DocxChangeId>
{
    /// <summary>Reads one wire-form change ID string.</summary>
    public override DocxChangeId Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        string? value = reader.GetString();
        if (DocxChangeId.TryParse(value, out DocxChangeId changeId))
        {
            return changeId;
        }

        throw new JsonException($"Unsupported change ID '{value}'. Expected M.CH0001 or H001.CH0002 shape.");
    }

    /// <summary>Writes one wire-form change ID string.</summary>
    public override void Write(Utf8JsonWriter writer, DocxChangeId value, JsonSerializerOptions options)
    {
        writer.WriteStringValue(value.ToWireValue());
    }
}
