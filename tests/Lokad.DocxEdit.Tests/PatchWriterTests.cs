using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D12: the canonical patch writer inverts the parser literal rules without loss.
public static class PatchWriterTests
{
    [Theory]
    [InlineData("Alpha")]
    [InlineData(" leading")]
    [InlineData("trailing ")]
    [InlineData("  both  ")]
    [InlineData("say \"hi\"")]
    [InlineData("back\\slash")]
    [InlineData("tab\there")]
    [InlineData("")]
    [InlineData("line one\nline two")]
    [InlineData("a\n\nb")]
    [InlineData("before\n>>>\nafter")]
    [InlineData(">>>")]
    [InlineData("<<<")]
    [InlineData("trailing\n")]
    [InlineData("\nleading")]
    [InlineData("héllo ✓")]
    [InlineData("\"")]
    [InlineData("ends with quote\"")]
    [InlineData("\"starts with quote")]
    [InlineData("with\ttab and \"quote\"")]
    public static void FieldValuesRoundTrip(string value)
    {
        var draft = new DocxPatchOperationDraft("replace-text", new[]
        {
            new KeyValuePair<string, string>("target", "M.P0001"),
            new KeyValuePair<string, string>("find", "Alpha"),
            new KeyValuePair<string, string>("with", value),
        });

        string patch = DocxPatchWriter.WritePatch(new[] { draft });
        DocxPatch parsed = new DocxEditor().ParsePatch(new StringReader(patch));

        Assert.True(parsed.Success, string.Join("|", parsed.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        Assert.Equal(value, Assert.Single(parsed.Operations).Fields["with"]);
    }

    [Fact]
    public static void CarriageReturnsNormalizeToLineFeeds()
    {
        var draft = new DocxPatchOperationDraft("replace-text", new[]
        {
            new KeyValuePair<string, string>("target", "M.P0001"),
            new KeyValuePair<string, string>("find", "Alpha"),
            new KeyValuePair<string, string>("with", "line\r\nbreak\rlast"),
        });

        string patch = DocxPatchWriter.WritePatch(new[] { draft });
        DocxPatch parsed = new DocxEditor().ParsePatch(new StringReader(patch));

        Assert.True(parsed.Success, string.Join("|", parsed.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        Assert.Equal("line\nbreak\nlast", Assert.Single(parsed.Operations).Fields["with"]);
    }

    [Fact]
    public static void BareValuesStayBare()
    {
        Assert.Equal("with Alpha", DocxPatchWriter.WriteField("with", "Alpha"));
        Assert.Equal("cell A2", DocxPatchWriter.WriteField("cell", "A2"));
    }

    [Fact]
    public static void EmptyValueRendersQuotedEmpty()
    {
        Assert.Equal("with \"\"", DocxPatchWriter.WriteField("with", ""));
    }

    [Fact]
    public static void PaddedValueRendersQuoted()
    {
        Assert.Equal("with \" padded \"", DocxPatchWriter.WriteField("with", " padded "));
    }

    [Fact]
    public static void MultilineValueRendersHeredoc()
    {
        Assert.Equal("with <<<\na\nb\n>>>", DocxPatchWriter.WriteField("with", "a\nb"));
    }

    [Fact]
    public static void DelimiterLineFallsBackToQuotedEscapes()
    {
        Assert.Equal("with \"before\\n>>>\\nafter\"", DocxPatchWriter.WriteField("with", "before\n>>>\nafter"));
    }

    [Fact]
    public static void WritePatchRendersExactDocument()
    {
        var first = new DocxPatchOperationDraft("replace-text", new[]
        {
            new KeyValuePair<string, string>("target", "M.P0001"),
            new KeyValuePair<string, string>("find", "Alpha"),
            new KeyValuePair<string, string>("with", "Omega"),
        });
        var second = new DocxPatchOperationDraft("delete-block", new[]
        {
            new KeyValuePair<string, string>("target", "M.P0002"),
        });

        Assert.Equal(
            "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Omega\nend\n\nop delete-block\ntarget M.P0002\nend\n",
            DocxPatchWriter.WritePatch(new[] { first, second }));
    }

    [Fact]
    public static void RepeatedFieldsKeepOrder()
    {
        var operation = new DocxPatchOperationDraft("append-row", new[]
        {
            new KeyValuePair<string, string>("target", "M.T0001"),
            new KeyValuePair<string, string>("cell", "A2"),
            new KeyValuePair<string, string>("cell", "B2"),
        });

        Assert.Equal(
            "op append-row\ntarget M.T0001\ncell A2\ncell B2\nend",
            DocxPatchWriter.WriteOperation(operation));
    }

    [Fact]
    public static void WrittenPatchChecksClean()
    {
        var draft = new DocxPatchOperationDraft("replace-text", new[]
        {
            new KeyValuePair<string, string>("target", "M.P0001"),
            new KeyValuePair<string, string>("find", "Alpha"),
            new KeyValuePair<string, string>("with", "Omega"),
        });

        using MemoryStream input = CreateDocx("Alpha");
        using var patch = new StringReader(DocxPatchWriter.WritePatch(new[] { draft }));
        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void NullInputsThrow()
    {
        Assert.Throws<ArgumentNullException>(() => DocxPatchWriter.WriteField(null!, "x"));
        Assert.Throws<ArgumentNullException>(() => DocxPatchWriter.WriteField("with", null!));
        Assert.Throws<ArgumentNullException>(() => DocxPatchWriter.WriteOperation(null!));
        Assert.Throws<ArgumentNullException>(() => DocxPatchWriter.WritePatch(null!));
    }

    [Fact]
    public static void SemanticSelectorValueRoundTrips()
    {
        Assert.Equal("target bookmark:ReviewAnchor", DocxPatchWriter.WriteField("target", "bookmark:ReviewAnchor"));

        var draft = new DocxPatchOperationDraft("replace-bookmark-text", new[]
        {
            new KeyValuePair<string, string>("target", "bookmark:ReviewAnchor"),
            new KeyValuePair<string, string>("text", "New text"),
        });

        string patch = DocxPatchWriter.WritePatch(new[] { draft });
        DocxPatch parsed = new DocxEditor().ParsePatch(new StringReader(patch));

        Assert.True(parsed.Success, string.Join("|", parsed.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        Assert.Equal("bookmark:ReviewAnchor", Assert.Single(parsed.Operations).Fields["target"]);
    }

    [Fact]
    public static void ReservedLiteralRendersQuoted()
    {
        Assert.Equal("with \"<<<\"", DocxPatchWriter.WriteField("with", "<<<"));
    }
}