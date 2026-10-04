using System.Text.Json;

namespace Lokad.DocxEdit.Tests;

[Collection("ConsoleCli")]
public static class EquationHelpTests
{
    [Fact]
    public static void CliDiscoversSyntaxWithoutRepositoryDocs()
    {
        Assert.Contains("|equations", CliTests.RunCli("--help").Output);
        Assert.Contains("docxedit help equations", CliTests.RunCli("help", "patch").Output);
        var topic = CliTests.RunCli("help", "equations");
        Assert.Equal(0, topic.ExitCode);
        Assert.Equal(DocxHelp.RenderTopic("equations"), topic.Output);
        Assert.Contains("REQUIRE a braced body", topic.Output);
        Assert.Contains("32768 source characters", topic.Output);
        Assert.Contains("64 parser nesting levels", topic.Output);
        Assert.Contains("1024 matrix cells", topic.Output);
        Assert.Contains("no tracked equation revisions", topic.Output);

        // Run the actual patch copied from the topic, not a separate example.
        int start = topic.Output.IndexOf("docxpatch 1", StringComparison.Ordinal);
        int end = topic.Output.IndexOf("Check and apply:", start, StringComparison.Ordinal);
        CheckExpressionPatch(topic.Output[start..end]);
    }

    [Theory]
    [InlineData("insert-equation")]
    [InlineData("replace-equation")]
    public static void FocusedHelpAndJsonCatalogCarryTheSamePrimer(string operation)
    {
        string help = CliTests.RunCli("help", operation).Output;
        Assert.Contains("Equation input:", help);
        Assert.Contains("REQUIRE a braced body", help);
        Assert.Contains("Literal backslashes", help);
        Assert.Contains("docxedit help equations", help);

        var catalog = CliTests.RunCli("catalog", "--json");
        Assert.Equal(0, catalog.ExitCode);
        using var json = JsonDocument.Parse(catalog.Output);
        JsonElement entry = json.RootElement.GetProperty("PatchOperations").EnumerateArray()
            .Single(e => e.GetProperty("Name").GetString() == operation);
        var notes = entry.GetProperty("Notes").EnumerateArray().Single(e => e.GetProperty("Heading").GetString() == "Equation input");
        foreach (JsonElement line in notes.GetProperty("Lines").EnumerateArray())
            Assert.Contains(line.GetString()!, help);
    }

    [Theory]
    [InlineData(@"x_i^2")]
    [InlineData(@"\frac{a}{b}")]
    [InlineData(@"\sqrt{x}")]
    [InlineData(@"\sqrt[3]{x}")]
    [InlineData(@"\sum_{i=1}^{n}{i^2}")]
    [InlineData(@"\prod_{i=1}^{n}{i}")]
    [InlineData(@"\int_0^1{x^2}")]
    [InlineData(@"\begin{pmatrix}a&b\\c&d\end{pmatrix}")]
    [InlineData(@"\left(\frac{x}{y}\right)")]
    [InlineData(@"\text{if }x\geq0")]
    [InlineData(@"\mathrm{d}")]
    [InlineData(@"\mathbf{x}")]
    [InlineData(@"\mathit{x}")]
    public static void AdvertisedExpressionsPassDocumentCheck(string expression)
    {
        Assert.Contains(expression, DocxHelp.RenderTopic("equations"));
        Assert.Contains(expression, DocxHelp.RenderTopic("insert-equation"));
        CheckExpressionPatch("docxpatch 1\nop insert-equation\ntarget M.P0001\nlatex " + expression + "\nend");
    }

    [Fact]
    public static void EveryAdvertisedNamedSymbolIsAccepted()
    {
        string help = DocxHelp.RenderTopic("equations");
        string symbols = help.Split("Named symbols (case-sensitive; literal Unicode also works):", StringSplitOptions.None)[1]
            .Split("Minimal patch", StringSplitOptions.None)[0];
        string[] commands = symbols.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        Assert.True(commands.Length > 50);
        CheckExpressionPatch("docxpatch 1\nop insert-equation\ntarget M.P0001\nlatex " + string.Join(" ", commands) + "\nend");
    }

    private static void CheckExpressionPatch(string patch)
    {
        var editor = new DocxEditor();
        using var input = new MemoryStream();
        Assert.True(editor.Create(input).Success);
        input.Position = 0;
        DocxCheckResult result = editor.Check(input, new StringReader(patch));
        Assert.True(result.Success, string.Join("; ", result.Diagnostics.Select(d => d.Message)));
    }
}
