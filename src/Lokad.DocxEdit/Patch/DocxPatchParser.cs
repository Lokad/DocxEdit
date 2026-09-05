using System.Text;

namespace Lokad.DocxEdit;

internal static class DocxPatchParser
{
    private static readonly IReadOnlyDictionary<string, OperationDefinition> OperationDefinitions = DocxPatchEngine.AllOperations
        .ToDictionary(registration => registration.Name, registration => new OperationDefinition(registration.AllowedFields, registration.BooleanFields, registration.IntegerFields), StringComparer.Ordinal);

    public static DocxPatch Parse(string text)
    {
        ArgumentNullException.ThrowIfNull(text);

        string normalized = text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n');
        string[] lines = normalized.Split('\n');
        int firstContentLine = Array.FindIndex(lines, line => line.Trim().Length != 0);
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

        if (normalized.Contains("\nexpect-hash", StringComparison.Ordinal) ||
            normalized.Contains("\nexpect-hash ", StringComparison.Ordinal))
        {
            int line = FindLine(lines, "expect-hash");
            return Error("E2004", "The expect-hash feature is not supported.", line, 1);
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
            if (operationName.Length == 0)
            {
                return Error("E2006", "Operation name is required.", i + 1, 4);
            }

            if (!OperationDefinitions.TryGetValue(operationName, out OperationDefinition? operationDefinition))
            {
                return Error("E2010", $"Unknown operation '{operationName}'.", i + 1, 4);
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
                int separator = fieldLine.IndexOf(' ');
                if (separator <= 0)
                {
                    return Error("E2007", $"Invalid field line '{operationTrimmed}'.", i + 1, 1);
                }

                string key = fieldLine[..separator].Trim();
                string rawValue = fieldLine[(separator + 1)..].Trim();
                int keyColumn = leadingWhitespace + 1;
                if (!operationDefinition.AllowedFields.Contains(key))
                {
                    return Error("E2011", $"Unknown field '{key}' for operation '{operationName}'.", i + 1, keyColumn);
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
                        return Error("E2008", $"Unterminated heredoc for field '{key}'.", startLine, 1);
                    }

                    value = string.Join('\n', heredoc);
                }
                else
                {
                    value = DecodeFieldValue(value);
                }

                if (operationDefinition.BooleanFields.Contains(key) && !IsBooleanLiteral(value))
                {
                    return Error("E2012", $"Field '{key}' must be true or false.", i + 1, keyColumn);
                }

                if (operationDefinition.IntegerFields.Contains(key) && !int.TryParse(value, out _))
                {
                    return Error("E2013", $"Field '{key}' must be an integer.", i + 1, keyColumn);
                }

                fields[key] = value;
                fieldValues.Add(new DocxPatchField(key, value, i + 1, keyColumn));
            }

            if (i >= lines.Length || lines[i].Trim() != "end")
            {
                return Error("E2009", $"Operation '{operationName}' is missing 'end'.", i + 1, 1);
            }

            operations.Add(new DocxPatchOperation(index++, operationName, fields)
            {
                FieldValues = fieldValues
            });
        }

        return new DocxPatch(true, majorVersion, operations, []);

        static int FindLine(string[] parserLines, string prefix)
        {
            for (int i = 0; i < parserLines.Length; i++)
            {
                if (parserLines[i].TrimStart().StartsWith(prefix, StringComparison.Ordinal))
                {
                    return i + 1;
                }
            }

            return 1;
        }

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

    private static DocxPatch Error(string code, string message, int line, int column)
    {
        return new DocxPatch(false, 0, [], [new DocxDiagnostic(DocxSeverity.Error, code, message) with { Line = line, Column = column }]);
    }

    private sealed record OperationDefinition(
        IReadOnlySet<string> AllowedFields,
        IReadOnlySet<string> BooleanFields,
        IReadOnlySet<string> IntegerFields)
    {
        public OperationDefinition(string[] allowedFields, string[] booleanFields, string[] integerFields)
            : this(
                new HashSet<string>(allowedFields, StringComparer.Ordinal),
                new HashSet<string>(booleanFields, StringComparer.Ordinal),
                new HashSet<string>(integerFields, StringComparer.Ordinal))
        {
        }
    }
}
