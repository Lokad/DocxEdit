namespace DocxEdit;

internal static class DocxPatchParser
{
    private static readonly IReadOnlyDictionary<string, OperationDefinition> OperationDefinitions = new Dictionary<string, OperationDefinition>(StringComparer.Ordinal)
    {
        ["replace-text"] = new(["target", "expect-text", "find", "with", "preserve-runs", "occurrence"], ["preserve-runs"], ["occurrence"]),
        ["replace-paragraph"] = new(["target", "expect-text", "style", "text"], [], []),
        ["insert-before"] = new(["target", "style", "copy-paragraph-properties", "text"], ["copy-paragraph-properties"], []),
        ["insert-after"] = new(["target", "style", "copy-paragraph-properties", "text"], ["copy-paragraph-properties"], []),
        ["delete-block"] = new(["target", "expect-text"], [], []),
        ["set-style"] = new(["target", "style"], [], []),
        ["set-content-control-text"] = new(["target", "text"], [], []),
        ["set-content-control-checkbox"] = new(["target", "checked"], [], []),
        ["set-content-control-choice"] = new(["target", "value", "display-text"], [], []),
        ["set-content-control-date"] = new(["target", "value", "display-text"], [], []),
        ["replace-bookmark-text"] = new(["target", "text"], [], []),
        ["rename-bookmark"] = new(["target", "name"], [], []),
        ["delete-bookmark"] = new(["target"], [], []),
        ["set-comment-text"] = new(["target", "text"], [], []),
        ["resolve-comment"] = new(["target"], [], []),
        ["reopen-comment"] = new(["target"], [], []),
        ["delete-comment"] = new(["target"], [], []),
        ["set-field-dirty"] = new(["target", "dirty"], ["dirty"], []),
        ["set-field-lock"] = new(["target", "locked"], ["locked"], []),
        ["set-hyperlink-target"] = new(["target", "uri", "anchor", "tooltip", "target-frame", "history"], ["history"], []),
        ["set-hyperlink-text"] = new(["target", "text"], [], []),
        ["insert-hyperlink-after"] = new(["target", "text", "uri", "anchor", "tooltip", "target-frame", "history"], ["history"], []),
        ["remove-hyperlink"] = new(["target"], [], []),
        ["set-cell"] = new(["target", "expect-text", "expect-row-count", "expect-column-count", "text", "force"], ["force"], ["expect-row-count", "expect-column-count"]),
        ["append-row"] = new(["target", "expect-row-count", "expect-column-count", "cell"], [], ["expect-row-count", "expect-column-count"]),
        ["insert-row-before"] = new(["target", "expect-row-count", "expect-column-count", "expect-cell-count", "cell", "force"], ["force"], ["expect-row-count", "expect-column-count", "expect-cell-count"]),
        ["insert-row-after"] = new(["target", "expect-row-count", "expect-column-count", "expect-cell-count", "cell", "force"], ["force"], ["expect-row-count", "expect-column-count", "expect-cell-count"]),
        ["delete-row"] = new(["target", "expect-row-count", "expect-column-count", "expect-cell-count", "expect-contains", "force"], ["force"], ["expect-row-count", "expect-column-count", "expect-cell-count"]),
        ["replace-image"] = new(["target", "asset", "expect-content-type", "alt"], [], []),
        ["insert-image-after"] = new(["target", "asset", "expect-content-type", "width", "height", "alt"], [], []),
        ["set-image-alt"] = new(["target", "expect-content-type", "alt"], [], []),
        ["set-image-metadata"] = new(["target", "expect-content-type", "alt", "title", "name"], [], []),
        ["delete-image"] = new(["target", "expect-content-type"], [], []),
        ["set-section-columns"] = new(["target", "expect-columns", "expect-orientation", "count", "space"], [], ["expect-columns", "count"]),
        ["set-section-orientation"] = new(["target", "expect-columns", "expect-orientation", "orientation"], [], ["expect-columns"])
    };

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
                string value = fieldLine[(separator + 1)..].Trim();
                int keyColumn = leadingWhitespace + 1;
                if (!operationDefinition.AllowedFields.Contains(key))
                {
                    return Error("E2011", $"Unknown field '{key}' for operation '{operationName}'.", i + 1, keyColumn);
                }

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
    }

    private static DocxPatch Error(string code, string message, int line, int column)
    {
        return new DocxPatch(false, 0, [], [new DocxDiagnostic(DocxSeverity.Error, code, message, Line: line, Column: column)]);
    }

    private static int FindLine(string[] lines, string prefix)
    {
        for (int i = 0; i < lines.Length; i++)
        {
            if (lines[i].TrimStart().StartsWith(prefix, StringComparison.Ordinal))
            {
                return i + 1;
            }
        }

        return 1;
    }

    private static bool IsBooleanLiteral(string value)
    {
        return string.Equals(value, "true", StringComparison.Ordinal) ||
            string.Equals(value, "false", StringComparison.Ordinal);
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
