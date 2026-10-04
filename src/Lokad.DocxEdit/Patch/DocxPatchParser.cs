using System.Text;

namespace Lokad.DocxEdit;

internal static class DocxPatchParser
{
    private static readonly IReadOnlyDictionary<string, OperationDefinition> OperationDefinitions = DocxPatchEngine.AllOperations
        .ToDictionary(registration => registration.Name, registration => new OperationDefinition(
            registration.Fields.Select(static field => field.Name).ToArray(),
            registration.Fields.Where(static field => field.Kind == FieldValueKind.Boolean).Select(static field => field.Name).ToArray(),
            registration.Fields.Where(static field => field.Kind == FieldValueKind.Integer).Select(static field => field.Name).ToArray(),
            registration.Fields.Where(static field => field.Repeatable).Select(static field => field.Name).ToArray()), StringComparer.Ordinal);

    public static DocxPatch Parse(string text)
    {
        ArgumentNullException.ThrowIfNull(text);

        if (text.StartsWith("\uFEFF", StringComparison.Ordinal))
        {
            text = text.Substring(1);

        }
        string normalized = text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n');
        string[] lines = normalized.Split('\n');
        int firstContentLine = Array.FindIndex(lines, static line => IsPreambleCandidate(line));
        if (firstContentLine < 0)
        {
            return Error("E2001", "Patch is empty.", 1, 1);
        }

        string preamble = lines[firstContentLine].Trim();
        if (!preamble.StartsWith("docxpatch ", StringComparison.Ordinal))
        {
            return Error("E2002", "Patch is missing required 'docxpatch 1' preamble.", firstContentLine + 1, 1);
        }

        string versionText = preamble["docxpatch ".Length..].Trim();
        if (!int.TryParse(versionText, out int majorVersion) || majorVersion != 1)
        {
            return Error("E2003", $"Unsupported docxpatch major version '{versionText}'.", firstContentLine + 1, "docxpatch ".Length + 1);
        }

        var operations = new List<DocxPatchOperation>();
        int index = 1;
        for (int i = firstContentLine + 1; i < lines.Length; i++)
        {
            string trimmed = lines[i].Trim();
            if (trimmed.Length == 0 || trimmed.StartsWith('#'))
            {
                continue;
            }

            if (!trimmed.StartsWith("op ", StringComparison.Ordinal))
            {
                return Error("E2005", $"Unexpected patch line '{trimmed}'.", i + 1, 1);
            }

            string operationName = trimmed[3..].Trim();
            if (!OperationDefinitions.TryGetValue(operationName, out OperationDefinition? operationDefinition))
            {
                string? operationSuggestion = SuggestNearestName(operationName, OperationDefinitions.Keys.OrderBy(static name => name, StringComparer.Ordinal));
                return Error("E2010", operationSuggestion is null ? $"Unknown operation '{operationName}'." : $"Unknown operation '{operationName}'. Did you mean '{operationSuggestion}'?", i + 1, 4);
            }

            var fields = new Dictionary<string, string>(StringComparer.Ordinal);
            var fieldValues = new List<DocxPatchField>();
            i++;
            for (; i < lines.Length; i++)
            {
                string operationLine = lines[i];
                string operationTrimmed = operationLine.Trim();
                if (operationTrimmed.Length == 0 || operationTrimmed.StartsWith('#'))
                {
                    continue;
                }

                if (operationTrimmed == "end")
                {
                    break;
                }

                int leadingWhitespace = operationLine.Length - operationLine.TrimStart().Length;
                string fieldLine = operationLine[leadingWhitespace..];
                int separator = fieldLine.IndexOfAny([' ', '\t']);
                if (separator <= 0)
                {
                    return Error("E2007", $"Invalid field line '{operationTrimmed}'.", i + 1, 1, operationName);
                }

                string key = fieldLine[..separator].Trim();
                string rawValue = fieldLine[(separator + 1)..].Trim();
                int keyColumn = leadingWhitespace + 1;
                int fieldLineNumber = i + 1;
                if (key == "expect-hash" && operationName is not ("replace-equation" or "delete-equation"))
                {
                    return Error("E2004", "The expect-hash feature is not supported.", fieldLineNumber, keyColumn, operationName);
                }

                if (!operationDefinition.AllowedFields.Contains(key))
                {
                    string? fieldSuggestion = SuggestNearestName(key, operationDefinition.AllowedFields.OrderBy(static name => name, StringComparer.Ordinal));
                    return Error("E2011", fieldSuggestion is null ? $"Unknown field '{key}' for operation '{operationName}'." : $"Unknown field '{key}' for operation '{operationName}'. Did you mean '{fieldSuggestion}'?", fieldLineNumber, keyColumn, operationName);
                }

                if (fields.ContainsKey(key) && !operationDefinition.RepeatableFields.Contains(key))
                {
                    return Error("E2015", $"Field '{key}' is specified more than once in operation '{operationName}'. Only fields documented as repeatable may repeat.", fieldLineNumber, keyColumn, operationName);
                }

                string value = rawValue;
                if (value == "<<<")
                {
                    int startLine = i + 2;
                    var heredoc = new List<string>();
                    i++;
                    for (; i < lines.Length && lines[i].Trim() != ">>>"; i++)
                    {
                        heredoc.Add(lines[i]);
                    }

                    if (i >= lines.Length)
                    {
                        return Error("E2008", $"Unterminated heredoc for field '{key}'.", startLine, 1, operationName);
                    }

                    value = string.Join('\n', heredoc);
                }
                else if ((rawValue.StartsWith((char)34) && !rawValue.EndsWith((char)34)) ||
                    (rawValue.EndsWith((char)34) && !rawValue.StartsWith((char)34) && rawValue.IndexOf((char)34) == rawValue.Length - 1))
                {
                    return Error("E2007", $"Field \"{key}\" has an unmatched double quote; either wrap the whole value in double quotes (escaping inner quotes) or remove the stray quote.", fieldLineNumber, keyColumn, operationName);
                }
                else
                {
                    value = DecodeFieldValue(value);
                }

                if (operationDefinition.BooleanFields.Contains(key) && !IsBooleanLiteral(value))
                {
                    return Error("E2012", $"Field '{key}' must be true or false.", fieldLineNumber, keyColumn, operationName);
                }

                // replace-text selects every match with the literal occurrence value all.
                // Integer patch fields are counts and ordinals: execution rejects
                // non-positive values with E4205, so the parser reports the same
                // failure at the field position. The literal occurrence value all
                // stays accepted for replace-text only.
                if (operationDefinition.IntegerFields.Contains(key) && int.TryParse(value, out int integerValue) && integerValue <= 0)
                {
                    return Error("E4205", $"Field \u0027{key}\u0027 must be greater than 0.", fieldLineNumber, keyColumn, operationName);
                }
                if (operationDefinition.IntegerFields.Contains(key) && !int.TryParse(value, out _) && !IsAllOccurrence(operationName, key, value))
                {
                    return Error("E2013", $"Field '{key}' must be an integer.", fieldLineNumber, keyColumn, operationName);
                }

                fields[key] = value;
                fieldValues.Add(new DocxPatchField(key, value, fieldLineNumber, keyColumn));
            }

            if (i >= lines.Length || lines[i].Trim() != "end")
            {
                return Error("E2009", $"Operation '{operationName}' is missing 'end'.", i + 1, 1, operationName);
            }

            operations.Add(new DocxPatchOperation(index++, operationName, fields)
            {
                FieldValues = fieldValues
            });
        }

        return new DocxPatch(true, majorVersion, operations, []);

        static bool IsBooleanLiteral(string value)
        {
            return string.Equals(value, "true", StringComparison.Ordinal) ||
                string.Equals(value, "false", StringComparison.Ordinal);
        }

        static string DecodeFieldValue(string fieldValue)
        {
            if (fieldValue.Length >= 2 && fieldValue[0] == '"' && fieldValue[^1] == '"')
            {
                return UnescapePatchValue(fieldValue[1..^1]);
            }

            return fieldValue;
        }
    }

    private static bool IsAllOccurrence(string operationName, string key, string value)
    {
        return string.Equals(operationName, "replace-text", StringComparison.Ordinal) &&
            string.Equals(key, "occurrence", StringComparison.Ordinal) &&
            string.Equals(value, "all", StringComparison.Ordinal);
    }

    private static bool IsPreambleCandidate(string line)
    {
        string trimmed = line.Trim();
        return trimmed.Length != 0 && !trimmed.StartsWith('#');
    }

    internal static string UnescapePatchValue(string value)
    {
        if (value.IndexOf('\\') < 0)
        {
            return value;
        }

        StringBuilder decoded = new(value.Length);
        for (int index = 0; index < value.Length; index++)
        {
            if (value[index] == '\\' && index + 1 < value.Length)
            {
                index++;
                decoded.Append(value[index] switch
                {
                    '"' => "\"",
                    '\\' => "\\",
                    'n' => "\n",
                    't' => "\t",
                    _ => "\\" + value[index],
                });
            }
            else
            {
                decoded.Append(value[index]);
            }
        }

        return decoded.ToString();
    }

    private static DocxPatch Error(string code, string message, int line, int column, string? helpTopic = null)
    {
        return new DocxPatch(false, 0, [], [new DocxDiagnostic(DocxSeverity.Error, code, message) with { Line = line, Column = column, HelpTopic = helpTopic }]);
    }

    internal static string? SuggestNearestName(string value, IEnumerable<string> candidates)
    {
        string? best = null;
        int bestDistance = 4;
        foreach (string candidate in candidates)
        {
            int distance = EditDistance(value.ToLowerInvariant(), candidate.ToLowerInvariant());
            if (distance < bestDistance)
            {
                bestDistance = distance;
                best = candidate;
            }
        }

        return best;
    }

    private static int EditDistance(string first, string second)
    {
        int[] previous = new int[second.Length + 1];
        for (int j = 0; j <= second.Length; j++)
        {
            previous[j] = j;
        }

        for (int i = 1; i <= first.Length; i++)
        {
            int[] current = new int[second.Length + 1];
            current[0] = i;
            for (int j = 1; j <= second.Length; j++)
            {
                int substitution = previous[j - 1] + (first[i - 1] == second[j - 1] ? 0 : 1);
                current[j] = Math.Min(Math.Min(previous[j] + 1, current[j - 1] + 1), substitution);
            }

            previous = current;
        }

        return previous[second.Length];
    }

    private sealed record OperationDefinition(
        IReadOnlySet<string> AllowedFields,
        IReadOnlySet<string> BooleanFields,
        IReadOnlySet<string> IntegerFields,
        IReadOnlySet<string> RepeatableFields)
    {
        public OperationDefinition(string[] allowedFields, string[] booleanFields, string[] integerFields, string[] repeatableFields)
            : this(
                new HashSet<string>(allowedFields, StringComparer.Ordinal),
                new HashSet<string>(booleanFields, StringComparer.Ordinal),
                new HashSet<string>(integerFields, StringComparer.Ordinal),
                new HashSet<string>(repeatableFields, StringComparer.Ordinal))
        {
        }
    }
}
