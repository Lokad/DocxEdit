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
/// Structured explicit target ID: the parsed form of one stable ID string
/// (<c>M.P0001</c>, <c>H001.T0001</c>, <c>M.T0001.R02.C03</c>, …).
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
    /// Parses the wire form (SPEC section 8.2) with the exact fixed-shape grammar:
    /// fixed lengths, ordinal prefixes and markers, and integer ordinals.
    /// </summary>
    public static bool TryParse(string? target, out DocxTargetId targetId)
    {
        targetId = default;
        if (string.IsNullOrEmpty(target))
        {
            return false;
        }

        if (TryParseMain(target, out targetId))
        {
            return true;
        }

        return TryParseStory(target, 'H', out targetId) || TryParseStory(target, 'F', out targetId);
    }

    private static bool TryParseMain(string target, out DocxTargetId targetId)
    {
        targetId = default;
        if (target.Length == 7)
        {
            if (target.StartsWith("M.P", StringComparison.Ordinal) && int.TryParse(target[3..], out int paragraphOrdinal))
            {
                targetId = new DocxTargetId('M', 0, DocxTargetKind.Paragraph, paragraphOrdinal, 0, 0);
                return true;
            }

            if (target.StartsWith("M.T", StringComparison.Ordinal) && int.TryParse(target[3..], out int tableOrdinal))
            {
                targetId = new DocxTargetId('M', 0, DocxTargetKind.Table, tableOrdinal, 0, 0);
                return true;
            }

            if (target.StartsWith("M.I", StringComparison.Ordinal) && int.TryParse(target[3..], out int imageOrdinal))
            {
                targetId = new DocxTargetId('M', 0, DocxTargetKind.Image, imageOrdinal, 0, 0);
                return true;
            }

            if (target.StartsWith("M.L", StringComparison.Ordinal) && int.TryParse(target[3..], out int hyperlinkOrdinal))
            {
                targetId = new DocxTargetId('M', 0, DocxTargetKind.Hyperlink, hyperlinkOrdinal, 0, 0);
                return true;
            }

            if (target.StartsWith("M.F", StringComparison.Ordinal) && int.TryParse(target[3..], out int fieldOrdinal))
            {
                targetId = new DocxTargetId('M', 0, DocxTargetKind.Field, fieldOrdinal, 0, 0);
                return true;
            }

            if (target.StartsWith("M.B", StringComparison.Ordinal) && int.TryParse(target[3..], out int bookmarkOrdinal))
            {
                targetId = new DocxTargetId('M', 0, DocxTargetKind.Bookmark, bookmarkOrdinal, 0, 0);
                return true;
            }

            if (target.StartsWith("M.S", StringComparison.Ordinal) && int.TryParse(target[3..], out int sectionOrdinal))
            {
                targetId = new DocxTargetId('M', 0, DocxTargetKind.Section, sectionOrdinal, 0, 0);
                return true;
            }
        }

        if (target.Length == 8 &&
            target.StartsWith("M.CC", StringComparison.Ordinal) &&
            int.TryParse(target[4..], out int contentControlOrdinal))
        {
            targetId = new DocxTargetId('M', 0, DocxTargetKind.ContentControl, contentControlOrdinal, 0, 0);
            return true;
        }

        if (target.Length == 11 &&
            target.StartsWith("M.T", StringComparison.Ordinal) &&
            target[7..9] == ".R" &&
            int.TryParse(target[3..7], out int rowTableOrdinal) &&
            int.TryParse(target[9..11], out int rowOrdinal))
        {
            targetId = new DocxTargetId('M', 0, DocxTargetKind.Row, rowTableOrdinal, rowOrdinal, 0);
            return true;
        }

        if (target.Length == 14 &&
            target.StartsWith("M.T", StringComparison.Ordinal) &&
            target[7..10] == ".MG" &&
            int.TryParse(target[3..7], out int mergeGroupTableOrdinal) &&
            int.TryParse(target[10..14], out int mergeGroupOrdinal))
        {
            targetId = new DocxTargetId('M', 0, DocxTargetKind.MergeGroup, mergeGroupTableOrdinal, mergeGroupOrdinal, 0);
            return true;
        }

        if (target.Length == 15 &&
            target.StartsWith("M.T", StringComparison.Ordinal) &&
            target[7..9] == ".R" &&
            target[11..13] == ".C" &&
            int.TryParse(target[3..7], out int cellTableOrdinal) &&
            int.TryParse(target[9..11], out int cellRowOrdinal) &&
            int.TryParse(target[13..15], out int cellOrdinal))
        {
            targetId = new DocxTargetId('M', 0, DocxTargetKind.Cell, cellTableOrdinal, cellRowOrdinal, cellOrdinal);
            return true;
        }

        return false;
    }

    private static bool TryParseStory(string target, char storyPrefix, out DocxTargetId targetId)
    {
        targetId = default;
        if (target.Length == 10 && target[0] == storyPrefix)
        {
            if (target[4..6] == ".P" &&
                int.TryParse(target[1..4], out int storyOrdinal) &&
                int.TryParse(target[6..], out int paragraphOrdinal))
            {
                targetId = new DocxTargetId(storyPrefix, storyOrdinal, DocxTargetKind.Paragraph, paragraphOrdinal, 0, 0);
                return true;
            }

            if (target[4..6] == ".T" &&
                int.TryParse(target[1..4], out storyOrdinal) &&
                int.TryParse(target[6..], out int tableOrdinal))
            {
                targetId = new DocxTargetId(storyPrefix, storyOrdinal, DocxTargetKind.Table, tableOrdinal, 0, 0);
                return true;
            }

            if (target[4..6] == ".I" &&
                int.TryParse(target[1..4], out storyOrdinal) &&
                int.TryParse(target[6..], out int imageOrdinal))
            {
                targetId = new DocxTargetId(storyPrefix, storyOrdinal, DocxTargetKind.Image, imageOrdinal, 0, 0);
                return true;
            }

            if (target[4..6] == ".L" &&
                int.TryParse(target[1..4], out storyOrdinal) &&
                int.TryParse(target[6..], out int hyperlinkOrdinal))
            {
                targetId = new DocxTargetId(storyPrefix, storyOrdinal, DocxTargetKind.Hyperlink, hyperlinkOrdinal, 0, 0);
                return true;
            }

            if (target[4..6] == ".F" &&
                int.TryParse(target[1..4], out storyOrdinal) &&
                int.TryParse(target[6..], out int fieldOrdinal))
            {
                targetId = new DocxTargetId(storyPrefix, storyOrdinal, DocxTargetKind.Field, fieldOrdinal, 0, 0);
                return true;
            }

            if (target[4..6] == ".B" &&
                int.TryParse(target[1..4], out storyOrdinal) &&
                int.TryParse(target[6..], out int bookmarkOrdinal))
            {
                targetId = new DocxTargetId(storyPrefix, storyOrdinal, DocxTargetKind.Bookmark, bookmarkOrdinal, 0, 0);
                return true;
            }
        }

        if (target.Length == 11 &&
            target[0] == storyPrefix &&
            target[4..7] == ".CC" &&
            int.TryParse(target[1..4], out int storyPartOrdinal) &&
            int.TryParse(target[7..], out int contentControlOrdinal))
        {
            targetId = new DocxTargetId(storyPrefix, storyPartOrdinal, DocxTargetKind.ContentControl, contentControlOrdinal, 0, 0);
            return true;
        }

        if (target.Length == 14 &&
            target[0] == storyPrefix &&
            target[4..6] == ".T" &&
            target[10..12] == ".R" &&
            int.TryParse(target[1..4], out storyPartOrdinal) &&
            int.TryParse(target[6..10], out int rowTableOrdinal) &&
            int.TryParse(target[12..], out int rowOrdinal))
        {
            targetId = new DocxTargetId(storyPrefix, storyPartOrdinal, DocxTargetKind.Row, rowTableOrdinal, rowOrdinal, 0);
            return true;
        }

        if (target.Length == 18 &&
            target[0] == storyPrefix &&
            target[4..6] == ".T" &&
            target[10..12] == ".R" &&
            target[14..16] == ".C" &&
            int.TryParse(target[1..4], out storyPartOrdinal) &&
            int.TryParse(target[6..10], out int cellTableOrdinal) &&
            int.TryParse(target[12..14], out int cellRowOrdinal) &&
            int.TryParse(target[16..], out int cellOrdinal))
        {
            targetId = new DocxTargetId(storyPrefix, storyPartOrdinal, DocxTargetKind.Cell, cellTableOrdinal, cellRowOrdinal, cellOrdinal);
            return true;
        }

        if (target.Length == 17 &&
            target[0] == storyPrefix &&
            target[4..6] == ".T" &&
            target[10..13] == ".MG" &&
            int.TryParse(target[1..4], out storyPartOrdinal) &&
            int.TryParse(target[6..10], out int mergeGroupTableOrdinal) &&
            int.TryParse(target[13..17], out int mergeGroupOrdinal))
        {
            targetId = new DocxTargetId(storyPrefix, storyPartOrdinal, DocxTargetKind.MergeGroup, mergeGroupTableOrdinal, mergeGroupOrdinal, 0);
            return true;
        }

        return false;
    }

    /// <summary>
    /// Formats the canonical wire form (SPEC section 8.2).
    /// </summary>
    /// <remarks>
    /// <para>
    /// Counters zero-pad exactly like the scanners, so canonical IDs round-trip.
    /// Non-canonical parses admitted by the fixed-shape matchers (for example
    /// sign-carrying ordinals) do not necessarily reproduce their input.
    /// </para>
    /// </remarks>
    public string ToWireValue()
    {
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
