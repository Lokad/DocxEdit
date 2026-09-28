using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D15: check and apply record the identical effective execution policy,
// including on failed results, so a report pins the policy it ran under.
public static class EffectivePolicyTests
{
    private static DocxEditOptions SuggestOptions()
    {
        return new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest };
    }

    [Fact]
    public static void CheckRecordsEffectivePolicy()
    {
        const string patchText = """
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            with Omega
            end
            """;
        using MemoryStream suggestInput = CreateDocx("Alpha Beta");
        using var suggestPatch = new StringReader(patchText);
        DocxCheckResult suggest = new DocxEditor().Check(suggestInput, suggestPatch, SuggestOptions());
        Assert.True(suggest.Success);
        Assert.Equal(TrackChangesMode.Suggest, suggest.TrackChanges);

        using MemoryStream defaultInput = CreateDocx("Alpha Beta");
        using var defaultPatch = new StringReader(patchText);
        DocxCheckResult @default = new DocxEditor().Check(defaultInput, defaultPatch);
        Assert.True(@default.Success);
        Assert.Equal(TrackChangesMode.Off, @default.TrackChanges);
    }

    [Fact]
    public static void ApplyRecordsEffectivePolicy()
    {
        const string patchText = """
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            with Omega
            end
            """;
        var options = new DocxEditOptions { TrackChanges = TrackChangesMode.Require };
        using MemoryStream input = CreateDocx("Alpha Beta");
        using var patch = new StringReader(patchText);
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output, options);

        Assert.True(apply.Success);
        Assert.Equal(TrackChangesMode.Require, apply.TrackChanges);
        Assert.Equal("docxedit", apply.Author);
    }

    [Fact]
    public static void PolicyRecordedOnParseFailure()
    {
        using MemoryStream input = CreateDocx("Alpha Beta");
        using var patch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nend\n");
        DocxCheckResult check = new DocxEditor().Check(input, patch, SuggestOptions());

        Assert.False(check.Success);
        Assert.Equal(TrackChangesMode.Suggest, check.TrackChanges);
    }

    [Fact]
    public static void PolicyRecordedOnOperationFailure()
    {
        const string patchText = """
            docxpatch 1

            op replace-text
            target M.P0001
            expect-text Wrong guarded text
            find Alpha
            with Omega
            end
            """;
        using MemoryStream input = CreateDocx("Alpha Beta");
        using var patch = new StringReader(patchText);
        DocxCheckResult check = new DocxEditor().Check(input, patch, SuggestOptions());

        Assert.False(check.Success);
        Assert.Equal(TrackChangesMode.Suggest, check.TrackChanges);
        Assert.Equal("docxedit", check.Author);
    }

    [Fact]
    public static void CheckAndApplyAgreeOnRecordedPolicy()
    {
        const string patchText = """
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            with Omega
            end
            """;
        var options = new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest, Author = "Agent" };
        using MemoryStream checkInput = CreateDocx("Alpha Beta");
        DocxCheckResult check = new DocxEditor().Check(checkInput, new StringReader(patchText), options);
        using MemoryStream applyInput = CreateDocx("Alpha Beta");
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, new StringReader(patchText), output, options);

        Assert.True(check.Success);
        Assert.True(apply.Success);
        Assert.Equal(check.TrackChanges, apply.TrackChanges);
        Assert.Equal(check.Author, apply.Author);
        Assert.Equal(check.TimestampUtc, apply.TimestampUtc);
    }

    [Fact]
    public static void TrackChangesWireValuesRoundTrip()
    {
        Assert.Equal("off", TrackChangesMode.Off.ToWireValue());
        Assert.Equal("preserve", TrackChangesMode.Preserve.ToWireValue());
        Assert.Equal("suggest", TrackChangesMode.Suggest.ToWireValue());
        Assert.Equal("require", TrackChangesMode.Require.ToWireValue());
        Assert.True(TrackChangesModeExtensions.TryParseWireValue("suggest", out TrackChangesMode suggest));
        Assert.Equal(TrackChangesMode.Suggest, suggest);
        Assert.False(TrackChangesModeExtensions.TryParseWireValue("sometimes", out _));
        Assert.False(TrackChangesModeExtensions.TryParseWireValue(null, out _));
    }
}