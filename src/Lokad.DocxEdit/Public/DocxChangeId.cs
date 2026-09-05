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
    /// Parses the wire form with the exact fixed-shape grammar: a main-story
    /// (<c>M.CH0001</c>) or part-story (<c>H001.CH0002</c>) head plus a
    /// zero-padded ordinal.
    /// </summary>
    public static bool TryParse(string? changeId, out DocxChangeId parsed)
{
        parsed = default;
        if (string.IsNullOrEmpty(changeId))
        {
            return false;
        }

        if (changeId.Length == 8 &&
            changeId[0] == 'M' &&
            changeId[1..4] == ".CH" &&
            int.TryParse(changeId[4..8], out int ordinal))
        {
            parsed = new DocxChangeId('M', 0, ordinal);
            return true;
        }

        if (changeId.Length == 11 &&
            changeId[0] is 'C' or 'H' or 'F' or 'P' &&
            changeId[4..7] == ".CH" &&
            int.TryParse(changeId[1..4], out int storyPart) &&
            int.TryParse(changeId[7..11], out int storyOrdinal))
        {
            parsed = new DocxChangeId(changeId[0], storyPart, storyOrdinal);
            return true;
        }

        return false;
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
