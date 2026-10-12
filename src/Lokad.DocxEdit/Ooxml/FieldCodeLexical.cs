using System.Text;

namespace Lokad.DocxEdit.Ooxml;

// Shared lexical rules; field metadata and edit eligibility remain separate.
internal static class FieldCodeLexical
{
    public static string Normalize(string code)
    {
        return string.Join(
            " ",
            code.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));
    }

    public static string[] Tokenize(string code)
    {
        var tokens = new List<string>();
        var current = new StringBuilder();
        bool inQuote = false;
        foreach (char character in code)
        {
            if (character == '"')
            {
                inQuote = !inQuote;
                continue;
            }

            if (char.IsWhiteSpace(character) && !inQuote)
            {
                AddCurrentFieldToken(tokens, current);
                continue;
            }

            current.Append(character);
        }

        AddCurrentFieldToken(tokens, current);
        return tokens.ToArray();
    }

    private static void AddCurrentFieldToken(List<string> tokens, StringBuilder current)
    {
        if (current.Length == 0)
        {
            return;
        }

        tokens.Add(current.ToString());
        current.Clear();
    }
}
