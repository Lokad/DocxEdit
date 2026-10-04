using System.Globalization;

namespace Lokad.DocxEdit.Ooxml;

/// <summary>
/// Shared readers for XML lexical values: quoted-text escaping, optional
/// integers, and on/off booleans. Single home for the helpers previously
/// duplicated across the model scanners and the two text renderers.
/// </summary>
internal static class XmlValues
{
    /// <summary>
    /// Escapes text for the quoted, line-oriented agent output (`text="..."`).
    /// Every control character that could break line parsing is escaped,
    /// including carriage returns: raw `\r` output used to differ between the
    /// two renderers and could garble line-oriented consumers.
    /// </summary>
    public static string EscapeText(string text)
    {
        return text
            .Replace("\\", "\\\\", StringComparison.Ordinal)
            .Replace("\"", "\\\"", StringComparison.Ordinal)
            .Replace("\r", "\\r", StringComparison.Ordinal)
            .Replace("\n", "\\n", StringComparison.Ordinal)
            .Replace("\t", "\\t", StringComparison.Ordinal);
    }

    /// <summary>
    /// Reads an optional integer attribute value. Returns null for missing,
    /// empty, or non-integer values; otherwise the parsed value (leading and
    /// trailing whitespace is accepted). XML number signs are culture-invariant.
    /// </summary>
    public static int? TryReadInt(string? value)
    {
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) ? parsed : null;
    }


    /// <summary>
    /// Pairs each item with its non-blank string key, dropping items whose key
    /// is missing or whitespace. Use before grouping by a nullable key so the
    /// group key stays non-nullable without null-forgiving operators; combine
    /// with the two-argument `GroupBy` overload to keep grouping whole items.
    /// </summary>
    public static IEnumerable<(string Key, T Item)> WithNonBlankKey<T>(
        this IEnumerable<T> items,
        Func<T, string?> keySelector)
    {
        foreach (T item in items)
        {
            string? key = keySelector(item);
            if (string.IsNullOrWhiteSpace(key))
            {
                continue;
            }

            yield return (key, item);
        }
    }
}

/// <summary>
/// Reads an ST_OnOff-style boolean attribute value (`1`/`true`/`on`,
/// case-insensitively; `0`/`false`/`off` and anything else read as false).
/// The case-insensitive `on` acceptance is a lenient superset of strict schema
/// values; no in-repo fixture depends on rejecting it.
/// </summary>
internal static class WmlBoolean
{
    /// <summary>
    /// Reads an on/off value with explicit absence semantics.
    /// </summary>
    /// <param name="value">Raw attribute value, or null when absent.</param>
    /// <param name="valueWhenMissing">
    /// Result when the attribute is absent. Absence-means-true is a real OOXML
    /// rule for some attributes and surprising everywhere else, so every call
    /// site states its choice: style `w:default` passes <c>false</c>, while
    /// nullable field/history flags keep absence as null through a separate wrapper.
    /// </param>
    public static bool IsTrue(string? value, bool valueWhenMissing)
    {
        return value is null
            ? valueWhenMissing
            : value.Equals("1", StringComparison.Ordinal) ||
                value.Equals("true", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("on", StringComparison.OrdinalIgnoreCase);
    }
}
