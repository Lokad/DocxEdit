namespace DocxEdit;

internal static class DocxPatchParser
{
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

                int separator = operationLine.IndexOf(' ');
                if (separator <= 0)
                {
                    return Error("E2007", $"Invalid field line '{operationTrimmed}'.", i + 1, 1);
                }

                string key = operationLine[..separator].Trim();
                string value = operationLine[(separator + 1)..].Trim();
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

                fields[key] = value;
                fieldValues.Add(new DocxPatchField(key, value));
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
}
