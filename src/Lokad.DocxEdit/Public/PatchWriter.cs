using System.Text;

namespace Lokad.DocxEdit;

/// <summary>One patch operation draft for canonical writing: an operation name plus ordered field values; repeated fields keep every occurrence in order.</summary>
/// <param name="OperationName">Operation name as written after op.</param>
/// <param name="Fields">Field values in file order.</param>
public sealed record DocxPatchOperationDraft(
    string OperationName,
    IReadOnlyList<KeyValuePair<string, string>> Fields);

/// <summary>Canonical docxpatch writer for machine-generated patches.</summary>
/// <remarks>
/// <para>Rendering inverts the parser literal rules without loss. Values that survive trimming and quote detection stay bare; other single-line values use double-quoted escapes; multi-line values use heredoc blocks unless a line would end the block early, in which case they fall back to quoted escapes. Whole patches use LF newlines with the docxpatch 1 preamble.</para>
/// </remarks>
public static class DocxPatchWriter
{
    private static readonly string[] LineSplit = ["\n"];

    /// <summary>Renders one canonical field line or heredoc block for a field name and value.</summary>
    /// <param name="field">Field name as written.</param>
    /// <param name="value">Field value; carriage returns normalize to line feeds.</param>
    public static string WriteField(string field, string value)
    {
        ArgumentNullException.ThrowIfNull(field);
        ArgumentNullException.ThrowIfNull(value);
        string normalized = value.Replace("\r\n", "\n", StringComparison.Ordinal).Replace("\r", "\n", StringComparison.Ordinal);
        if (normalized.Length == 0)
        {
            return field + " \"\"";
        }

        if (normalized.Split(LineSplit, StringSplitOptions.None).Length > 1)
        {
            if (IsSafeHeredocBody(normalized))
            {
                return field + " <<<\n" + normalized + "\n>>>";
            }

            return field + " \"" + EscapeQuoted(normalized, multiline: true) + "\"";
        }

        if (NeedsQuoting(normalized))
        {
            return field + " \"" + EscapeQuoted(normalized, multiline: false) + "\"";
        }

        return field + " " + normalized;
    }

    /// <summary>Renders one canonical operation block from a draft.</summary>
    public static string WriteOperation(DocxPatchOperationDraft operation)
    {
        ArgumentNullException.ThrowIfNull(operation);
        StringBuilder builder = new();
        builder.Append("op ").Append(operation.OperationName).Append("\n");
        foreach (KeyValuePair<string, string> field in operation.Fields)
        {
            builder.Append(WriteField(field.Key, field.Value)).Append("\n");
        }

        builder.Append("end");
        return builder.ToString();
    }

    /// <summary>Renders a canonical patch document from operation drafts in order.</summary>
    public static string WritePatch(IEnumerable<DocxPatchOperationDraft> operations)
    {
        ArgumentNullException.ThrowIfNull(operations);
        return "docxpatch 1\n\n" + string.Join("\n\n", operations.Select(WriteOperation)) + "\n";
    }

    /// <summary>Whether a value can render as a heredoc body without a delimiter-like line ending the block early.</summary>
    /// <param name="value">Field value; carriage returns normalize to line feeds.</param>
    internal static bool IsSafeHeredocBody(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        string normalized = value.Replace("\r\n", "\n", StringComparison.Ordinal).Replace("\r", "\n", StringComparison.Ordinal);
        return !normalized.Split(LineSplit, StringSplitOptions.None).Any(static line => line.Trim() == ">>>");
    }

    private static bool NeedsQuoting(string value)
    {
        if (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1]))
        {
            return true;
        }

        if (value.StartsWith("\"", StringComparison.Ordinal) || value.EndsWith("\"", StringComparison.Ordinal))
        {
            return true;
        }

        if (value == "<<<")
        {
            return true;
        }

        return value.Contains("\"", StringComparison.Ordinal);
    }

    private static string EscapeQuoted(string value, bool multiline)
    {
        string escaped = value
            .Replace("\\", "\\\\", StringComparison.Ordinal)
            .Replace("\"", "\\\"", StringComparison.Ordinal)
            .Replace("\t", "\\t", StringComparison.Ordinal);
        return multiline ? escaped.Replace("\n", "\\n", StringComparison.Ordinal) : escaped;
    }
}
