namespace Lokad.DocxEdit.Model;

/// <summary>
/// Applies deterministic list labels with per-numbering-ID counters. Single-scan use only:
/// the counters accumulate across calls, so the scanner creates one labeler per scan.
/// Not thread-safe.
/// </summary>
internal sealed class DocxNumberingLabeler
{
    private readonly DocxNumberingCatalog _numbering;
    private readonly Dictionary<string, Dictionary<int, int>> _countersByNumberingId = new(StringComparer.Ordinal);

    public DocxNumberingLabeler(DocxNumberingCatalog numbering)
    {
        _numbering = numbering;
    }

    public DocxListInfo? ApplyLabel(DocxListInfo? list)
    {
        if (list is null)
        {
            return null;
        }

        Dictionary<int, int> counters = GetCounters(list.NumberingId);
        int startValue = list.StartValue ?? 1;
        counters[list.Level] = counters.TryGetValue(list.Level, out int current)
            ? current + 1
            : startValue;
        ResetDeeperLevels(counters, list.Level);

        List<string> warnings = [];
        List<DocxListLabelComponent> components = [];
        string? label = FormatLabel(list, counters, warnings, components);
        return list with
        {
            LabelText = label,
            LabelComponents = components.ToArray(),
            LabelStatus = label is null
                ? DocxLabelStatus.Unsupported
                : warnings.Count == 0 ? DocxLabelStatus.Resolved : DocxLabelStatus.Partial,
            LabelWarnings = warnings.ToArray()
        };
    }

    private Dictionary<int, int> GetCounters(string numberingId)
    {
        if (!_countersByNumberingId.TryGetValue(numberingId, out Dictionary<int, int>? counters))
        {
            counters = [];
            _countersByNumberingId[numberingId] = counters;
        }

        return counters;
    }

    private static void ResetDeeperLevels(Dictionary<int, int> counters, int level)
    {
        foreach (int key in counters.Keys.Where(key => key > level).ToArray())
        {
            counters.Remove(key);
        }
    }

    private string? FormatLabel(
        DocxListInfo list,
        IReadOnlyDictionary<int, int> counters,
        List<string> warnings,
        List<DocxListLabelComponent> components)
    {
        if (string.IsNullOrWhiteSpace(list.Format) || string.IsNullOrWhiteSpace(list.LevelText))
        {
            warnings.Add("missing-numbering-definition");
            return null;
        }

        if (string.Equals(list.Format, "none", StringComparison.Ordinal))
        {
            return string.Empty;
        }

        if (string.Equals(list.Format, "bullet", StringComparison.Ordinal))
        {
            return list.LevelText;
        }

        string label = ExpandTokens(list, counters, warnings, components);
        if (!label.Contains('%', StringComparison.Ordinal))
        {
            return label;
        }

        warnings.Add("unexpanded-level-token");
        return label;
    }

    private string ExpandTokens(
        DocxListInfo list,
        IReadOnlyDictionary<int, int> counters,
        List<string> warnings,
        List<DocxListLabelComponent> components)
    {
        string levelText = list.LevelText ?? string.Empty;
        var builder = new System.Text.StringBuilder(levelText.Length);
        for (int index = 0; index < levelText.Length; index++)
        {
            char current = levelText[index];
            if (current != '%' || index + 1 >= levelText.Length || !char.IsDigit(levelText[index + 1]))
            {
                builder.Append(current);
                continue;
            }

            int tokenStart = ++index;
            while (index + 1 < levelText.Length && char.IsDigit(levelText[index + 1]))
            {
                index++;
            }

            string tokenText = levelText[tokenStart..(index + 1)];
            if (!int.TryParse(tokenText, out int oneBasedLevel) || oneBasedLevel <= 0)
            {
                builder.Append('%').Append(tokenText);
                warnings.Add("invalid-level-token");
                continue;
            }

            int tokenLevel = oneBasedLevel - 1;
            if (!counters.TryGetValue(tokenLevel, out int value))
            {
                builder.Append('%').Append(tokenText);
                warnings.Add($"missing-counter-level-{tokenLevel}");
                continue;
            }

            DocxListInfo tokenLevelInfo = tokenLevel == list.Level
                ? list
                : _numbering.Resolve(list.NumberingId, tokenLevel, list.Source);
            string format = list.IsLegal ? "decimal" : tokenLevelInfo.Format ?? string.Empty;
            string? formatted = FormatCounter(value, format);
            if (formatted is null)
            {
                builder.Append('%').Append(tokenText);
                warnings.Add($"unsupported-format-level-{tokenLevel}");
                continue;
            }

            components.Add(new DocxListLabelComponent(tokenLevel, value, formatted, format));
            builder.Append(formatted);
        }

        return builder.ToString();
    }

    private static string? FormatCounter(int value, string? format)
    {
        return format switch
        {
            "decimal" => value.ToString(System.Globalization.CultureInfo.InvariantCulture),
            "decimalZero" => value is >= 0 and < 10
                ? "0" + value.ToString(System.Globalization.CultureInfo.InvariantCulture)
                : value.ToString(System.Globalization.CultureInfo.InvariantCulture),
            "lowerLetter" => ToLetters(value, upperCase: false),
            "upperLetter" => ToLetters(value, upperCase: true),
            "lowerRoman" => ToRoman(value)?.ToLowerInvariant(),
            "upperRoman" => ToRoman(value),
            _ => null
        };
    }

    private static string? ToLetters(int value, bool upperCase)
    {
        if (value <= 0)
        {
            return null;
        }

        var chars = new Stack<char>();
        int current = value;
        while (current > 0)
        {
            current--;
            chars.Push((char)((upperCase ? 'A' : 'a') + current % 26));
            current /= 26;
        }

        return new string(chars.ToArray());
    }

    private static string? ToRoman(int value)
    {
        if (value is <= 0 or > 3999)
        {
            return null;
        }

        Span<(int Value, string Text)> numerals =
        [
            (1000, "M"),
            (900, "CM"),
            (500, "D"),
            (400, "CD"),
            (100, "C"),
            (90, "XC"),
            (50, "L"),
            (40, "XL"),
            (10, "X"),
            (9, "IX"),
            (5, "V"),
            (4, "IV"),
            (1, "I")
        ];
        var builder = new System.Text.StringBuilder();
        int remaining = value;
        foreach ((int numeralValue, string numeralText) in numerals)
        {
            while (remaining >= numeralValue)
            {
                builder.Append(numeralText);
                remaining -= numeralValue;
            }
        }

        return builder.ToString();
    }
}
