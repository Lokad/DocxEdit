using System.IO;
using System.Text.RegularExpressions;
using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D10/D14: every diagnostic code the library can emit is cataloged in
// docs/diagnostics.md, and the catalog names no phantom codes.
public static class DiagnosticCatalogTests
{
    private static HashSet<string> ExtractCodes(string text)
    {
        var codes = new HashSet<string>(StringComparer.Ordinal);
        foreach (Match match in Regex.Matches(text, "\"[EIW][0-9]{4}\""))
        {
            codes.Add(match.Value.Substring(1, match.Value.Length - 2));
        }

        return codes;
    }

    [Fact]
    public static void EmittedAndCatalogedDiagnosticCodesAgree()
    {
        string root = FindRepoRoot();
        var emitted = new HashSet<string>(StringComparer.Ordinal);
        foreach (string path in Directory.EnumerateFiles(Path.Combine(root, "src", "Lokad.DocxEdit"), "*.cs", SearchOption.AllDirectories))
        {
            foreach (string code in ExtractCodes(File.ReadAllText(path)))
            {
                emitted.Add(code);
            }
        }

        string doc = File.ReadAllText(Path.Combine(root, "docs", "diagnostics.md"));
        var documented = new HashSet<string>(StringComparer.Ordinal);
        foreach (Match match in Regex.Matches(doc, "[EIW][0-9]{4}"))
        {
            documented.Add(match.Value);
        }

        Assert.NotEmpty(emitted);
        foreach (string code in emitted)
        {
            Assert.Contains(code, documented);
        }

        foreach (string code in documented)
        {
            Assert.Contains(code, emitted);
        }
    }
}