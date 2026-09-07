using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Lokad.DocxEdit;

/// <summary>
/// Closed entity kinds addressable by an explicit target ID (SPEC section 8.2).
/// </summary>
public enum DocxTargetKind
{
    /// <summary>Top-level paragraph (<c>P</c>).</summary>
    Paragraph,

    /// <summary>Top-level table (<c>T</c>).</summary>
    Table,

    /// <summary>Table row (<c>R</c>).</summary>
    Row,

    /// <summary>Table cell (<c>C</c>).</summary>
    Cell,

    /// <summary>Vertical-merge group (<c>MG</c>).</summary>
    MergeGroup,

    /// <summary>Image (<c>I</c>).</summary>
    Image,

    /// <summary>Hyperlink (<c>L</c>).</summary>
    Hyperlink,

    /// <summary>Field (<c>F</c>).</summary>
    Field,

    /// <summary>Bookmark (<c>B</c>).</summary>
    Bookmark,

    /// <summary>Content control (<c>CC</c>).</summary>
    ContentControl,

    /// <summary>Section (<c>S</c>).</summary>
    Section
}

/// <summary>
/// Structured explicit target ID: the parsed form of one target ID string
/// (<c>M.P0001</c>, <c>H001.T0001</c>, <c>M.T0001.R02.C03</c>, ...).
/// </summary>
/// <remarks>
/// <para>
/// <c>Primary</c> holds the entity ordinal (the table ordinal for
/// <c>Row</c>/<c>Cell</c>/<c>MergeGroup</c>); <c>Secondary</c> holds the row ordinal
/// for <c>Row</c>/<c>Cell</c> or the merge-group ordinal for <c>MergeGroup</c>;
/// <c>Tertiary</c> holds the cell ordinal for <c>Cell</c>. Unused slots are 0.
/// <c>Story</c> is the story prefix (<c>M</c>, <c>H</c>, <c>F</c>); <c>StoryPart</c>
/// is the 1-based header/footer part ordinal, 0 for the main story.
/// </para>
/// <para>
/// IDs enumerate physical document order at inspection time and are stable for
/// the same document bytes and scanner version, but they are not permanently
/// stable: within a patch, operations resolve sequentially, so an earlier
/// structural edit shifts later positional IDs (SPEC section 8.2).
/// </para>
/// </remarks>
[JsonConverter(typeof(DocxTargetIdJsonConverter))]
public readonly record struct DocxTargetId(
    char Story,
    int StoryPart,
    DocxTargetKind Kind,
    int Primary,
    int Secondary,
    int Tertiary)
{
    /// <summary>
    /// Containing-table ID for <c>Row</c>, <c>Cell</c>, and <c>MergeGroup</c>; meaningless otherwise.
    /// </summary>
    public DocxTargetId TableId => new DocxTargetId(Story, StoryPart, DocxTargetKind.Table, Primary, 0, 0);

    /// <summary>
    /// Containing-row ID for <c>Cell</c>; meaningless otherwise.
    /// </summary>
    public DocxTargetId RowId => new DocxTargetId(Story, StoryPart, DocxTargetKind.Row, Primary, Secondary, 0);

    /// <summary>
    /// Parses the wire form (SPEC section 8.2).
    /// </summary>
    /// <remarks>
    /// <para>
    /// Markers are fixed; ordinals are ASCII-digit runs with canonical minimum
    /// widths (4 for entities, 2 for rows and cells, 4 for merge groups, 3 for
    /// story parts) and no upper width limit, so producer output past a width
    /// boundary still parses. Every ordinal must be at least 1: zero, signs,
    /// whitespace, and non-digits are rejected, as are short runs, unknown
    /// markers, and trailing content.
    /// </para>
    /// </remarks>
    public static bool TryParse(string? target, out DocxTargetId targetId)
    {
        targetId = default;
        if (string.IsNullOrEmpty(target))
        {
            return false;
        }

        ReadOnlySpan<char> span = target.AsSpan();
        if (!TryParseStoryHead(ref span, out char story, out int storyPart))
        {
            return false;
        }

        if (!TryParseEntity(ref span, out DocxTargetKind kind, out int primary, out int secondary, out int tertiary))
        {
            return false;
        }

        if (!span.IsEmpty)
        {
            return false;
        }

        targetId = new DocxTargetId(story, storyPart, kind, primary, secondary, tertiary);
        return true;

        static bool TryParseStoryHead(ref ReadOnlySpan<char> span, out char story, out int storyPart)
        {
            story = (char)0;
            storyPart = 0;
            if (span.IsEmpty)
            {
                return false;
            }

            if (span[0] == 'M')
            {
                story = 'M';
                span = span[1..];
                return true;
            }

            if (span[0] is 'H' or 'F')
            {
                story = span[0];
                ReadOnlySpan<char> tail = span[1..];
                if (!TakeOrdinal(ref tail, 3, out storyPart))
                {
                    return false;
                }

                span = tail;
                return true;
            }

            return false;
        }

        static bool TryParseEntity(ref ReadOnlySpan<char> span, out DocxTargetKind kind, out int primary, out int secondary, out int tertiary)
        {
            kind = default;
            primary = 0;
            secondary = 0;
            tertiary = 0;
            if (span.Length < 2 || span[0] != '.')
            {
                return false;
            }

            switch (span[1])
            {
                case 'P':
                    kind = DocxTargetKind.Paragraph;
                    break;
                case 'I':
                    kind = DocxTargetKind.Image;
                    break;
                case 'L':
                    kind = DocxTargetKind.Hyperlink;
                    break;
                case 'F':
                    kind = DocxTargetKind.Field;
                    break;
                case 'B':
                    kind = DocxTargetKind.Bookmark;
                    break;
                case 'S':
                    kind = DocxTargetKind.Section;
                    break;
                case 'T':
                    return TryParseTableTail(span[2..], out kind, out primary, out secondary, out tertiary, out span);
                case 'C':
                    if (span.Length >= 3 && span[2] == 'C')
                    {
                        kind = DocxTargetKind.ContentControl;
                        ReadOnlySpan<char> tail = span[3..];
                        if (!TakeOrdinal(ref tail, 4, out primary))
                        {
                            return false;
                        }

                        span = tail;
                        return true;
                    }

                    return false;
                default:
                    return false;
            }

            ReadOnlySpan<char> ordinal = span[2..];
            if (!TakeOrdinal(ref ordinal, 4, out primary))
            {
                return false;
            }

            span = ordinal;
            return true;
        }

        static bool TryParseTableTail(ReadOnlySpan<char> span, out DocxTargetKind kind, out int primary, out int secondary, out int tertiary, out ReadOnlySpan<char> rest)
        {
            kind = default;
            primary = 0;
            secondary = 0;
            tertiary = 0;
            rest = span;
            if (!TakeOrdinal(ref rest, 4, out primary))
            {
                return false;
            }

            if (rest.IsEmpty)
            {
                kind = DocxTargetKind.Table;
                return true;
            }

            if (rest.Length >= 2 && rest[0] == '.' && rest[1] == 'R')
            {
                ReadOnlySpan<char> tail = rest[2..];
                if (!TakeOrdinal(ref tail, 2, out secondary))
                {
                    return false;
                }

                if (tail.IsEmpty)
                {
                    kind = DocxTargetKind.Row;
                    rest = tail;
                    return true;
                }

                if (tail.Length >= 2 && tail[0] == '.' && tail[1] == 'C')
                {
                    ReadOnlySpan<char> cell = tail[2..];
                    if (!TakeOrdinal(ref cell, 2, out tertiary))
                    {
                        return false;
                    }

                    kind = DocxTargetKind.Cell;
                    rest = cell;
                    return true;
                }

                return false;
            }

            if (rest.Length >= 3 && rest[0] == '.' && rest[1] == 'M' && rest[2] == 'G')
            {
                ReadOnlySpan<char> tail = rest[3..];
                if (!TakeOrdinal(ref tail, 4, out secondary))
                {
                    return false;
                }

                kind = DocxTargetKind.MergeGroup;
                rest = tail;
                return true;
            }

            return false;
        }

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

    /// <summary>
    /// Formats the canonical wire form (SPEC section 8.2).
    /// </summary>
    /// <remarks>
    /// <para>
    /// Counters zero-pad to canonical minimum widths exactly like the scanners,
    /// so canonical IDs round-trip. Longer counters extend the width instead of
    /// wrapping. Stories, parts, kinds, and used ordinals are range-checked;
    /// default or out-of-range values throw instead of emitting an unparseable ID.
    /// </para>
    /// </remarks>
    public string ToWireValue()
    {
        if (Story != 'M' && Story != 'H' && Story != 'F')
        {
            throw new ArgumentOutOfRangeException(nameof(Story), $"Unsupported story '{Story}'. Expected 'M', 'H', or 'F'.");
        }

        if (Story == 'M' ? StoryPart != 0 : StoryPart < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(StoryPart), "The main story uses part 0; header and footer parts start at 1.");
        }

        if (Primary < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(Primary), "Target ordinals start at 1.");
        }

        if (Kind is DocxTargetKind.Row or DocxTargetKind.Cell or DocxTargetKind.MergeGroup && Secondary < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(Secondary), "Row, cell, and merge-group targets need a row ordinal starting at 1.");
        }

        if (Kind is DocxTargetKind.Cell && Tertiary < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(Tertiary), "Cell targets need a cell ordinal starting at 1.");
        }

        string head = Story == 'M' ? "M" : $"{Story}{StoryPart:D3}";
        return Kind switch
        {
            DocxTargetKind.Paragraph => $"{head}.P{Primary:D4}",
            DocxTargetKind.Table => $"{head}.T{Primary:D4}",
            DocxTargetKind.Row => $"{head}.T{Primary:D4}.R{Secondary:D2}",
            DocxTargetKind.Cell => $"{head}.T{Primary:D4}.R{Secondary:D2}.C{Tertiary:D2}",
            DocxTargetKind.MergeGroup => $"{head}.T{Primary:D4}.MG{Secondary:D4}",
            DocxTargetKind.Image => $"{head}.I{Primary:D4}",
            DocxTargetKind.Hyperlink => $"{head}.L{Primary:D4}",
            DocxTargetKind.Field => $"{head}.F{Primary:D4}",
            DocxTargetKind.Bookmark => $"{head}.B{Primary:D4}",
            DocxTargetKind.ContentControl => $"{head}.CC{Primary:D4}",
            DocxTargetKind.Section => $"{head}.S{Primary:D4}",
            _ => throw new ArgumentOutOfRangeException(nameof(Kind), $"Unsupported target kind '{Kind}'.")
        };
    }

    /// <summary>Returns the canonical wire form (same as <see cref="ToWireValue"/>).</summary>
    /// <remarks>An explicit override: the synthesized record display would recurse through
    /// the <c>TableId</c>/<c>RowId</c> projections, so display is defined as the wire form.</remarks>
    public override string ToString()
    {
        return ToWireValue();
    }

    /// <summary>Allocates the next merge-group ID for a table, incrementing the counter.</summary>
    internal static DocxTargetId AllocateMergeGroupId(DocxTargetId tableId, ref int mergeGroupIndex)
    {
        return new DocxTargetId(tableId.Story, tableId.StoryPart, DocxTargetKind.MergeGroup, tableId.Primary, mergeGroupIndex++, 0);
    }

    /// <summary>Splits a story ID prefix (<c>M</c>, <c>H001</c>) for producer emission.</summary>
    internal static (char Story, int StoryPart) ParseStoryPrefix(string idPrefix)
    {
        char story = idPrefix[0];
        int storyPart = idPrefix.Length == 1 ? 0 : int.Parse(idPrefix[1..], CultureInfo.InvariantCulture);
        return (story, storyPart);
    }
}

/// <summary>
/// Converts <see cref="DocxTargetId"/> to and from its wire form, so JSON
/// output stays string-shaped under both default and CLI options.
/// </summary>
public sealed class DocxTargetIdJsonConverter : JsonConverter<DocxTargetId>
{
    /// <summary>Reads one wire-form target ID string.</summary>
    public override DocxTargetId Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        string? value = reader.GetString();
        if (DocxTargetId.TryParse(value, out DocxTargetId targetId))
        {
            return targetId;
        }

        throw new JsonException($"Unsupported target ID '{value}'. Expected M.P0001 or H001.P0002 shape.");
    }

    /// <summary>Writes one wire-form target ID string.</summary>
    public override void Write(Utf8JsonWriter writer, DocxTargetId value, JsonSerializerOptions options)
    {
        writer.WriteStringValue(value.ToWireValue());
    }
}
