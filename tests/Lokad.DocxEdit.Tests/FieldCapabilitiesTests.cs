using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D17: field capabilities reuse the shape, guard-support, and track-support
// predicates from the field engine.
public static class FieldCapabilitiesTests
{
    private const string MixedBody = """
                    <w:p>
                      <w:fldSimple w:instr=" DATE " w:dirty="false" w:fldLock="0">
                        <w:r><w:t>June 12</w:t></w:r>
                      </w:fldSimple>
                    </w:p>
                    <w:p>
                      <w:r><w:fldChar w:fldCharType="begin" w:dirty="false" w:fldLock="0"/></w:r>
                      <w:r><w:instrText> PAGE </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:t>1</w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                    </w:p>
        """;

    private static DocxOperationCapability FindOperation(DocxTargetCapabilities capabilities, string operation)
    {
        return capabilities.Operations.First(candidate => string.Equals(candidate.Operation, operation, StringComparison.Ordinal));
    }

    private static DocxTargetCapabilities GetCapabilities(MemoryStream input, string targetId, TrackChangesMode mode)
    {
        input.Position = 0;
        DocxCapabilitiesResult result = new DocxEditor().GetCapabilities(input, targetId, new DocxCapabilitiesOptions { TrackChanges = mode });
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        Assert.Equal(targetId, result.TargetId);
        Assert.NotNull(result.Capabilities);
        return result.Capabilities!;
    }

    private static DocxCheckResult RunCheck(MemoryStream input, string patchText, TrackChangesMode mode)
    {
        input.Position = 0;
        using var patch = new StringReader(patchText);
        return new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = mode });
    }

    [Fact]
    public static void SimpleRefFieldVerdicts()
    {
        using MemoryStream input = CreateDocxWithRefField();
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.F0001", TrackChangesMode.Off);

        Assert.Equal("M.F0001", capabilities.TargetId);
        Assert.Equal("field", capabilities.Kind);
        Assert.Equal("main", capabilities.Story);
        string[] expectedOrder = ["set-field-dirty", "set-field-lock", "set-field-code", "set-field-result", "refresh-field-result"];
        Assert.Equal(expectedOrder, capabilities.Operations.Select(static operation => operation.Operation).ToArray());
        Assert.Equal("supported", FindOperation(capabilities, "set-field-dirty").Support);
        Assert.Equal("supported", FindOperation(capabilities, "set-field-lock").Support);
        Assert.Equal("supported", FindOperation(capabilities, "set-field-code").Support);
        Assert.Equal("supported", FindOperation(capabilities, "set-field-result").Support);
        Assert.Equal("supported", FindOperation(capabilities, "refresh-field-result").Support);

        using MemoryStream resultInput = CreateDocxWithRefField();
        DocxCheckResult result = RunCheck(resultInput, "docxpatch 1\n\nop set-field-result\ntarget M.F0001\ntext Acme Corporation\nend\n", TrackChangesMode.Off);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));

        using MemoryStream refreshInput = CreateDocxWithRefField();
        DocxCheckResult refresh = RunCheck(refreshInput, "docxpatch 1\n\nop refresh-field-result\ntarget M.F0001\nend\n", TrackChangesMode.Off);
        Assert.True(refresh.Success, string.Join("|", refresh.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void RequireModeRefusesMetadataOps()
    {
        using MemoryStream input = CreateDocxWithRefField();
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.F0001", TrackChangesMode.Require);

        Assert.Equal("unsupported", FindOperation(capabilities, "set-field-dirty").Support);
        Assert.Equal("unsupported", FindOperation(capabilities, "set-field-lock").Support);
        Assert.Equal("unsupported", FindOperation(capabilities, "set-field-code").Support);
        Assert.Equal("conditional", FindOperation(capabilities, "set-field-result").Support);
        Assert.Equal("unsupported", FindOperation(capabilities, "refresh-field-result").Support);

        using MemoryStream dirtyCheck = CreateDocxWithRefField();
        DocxCheckResult dirty = RunCheck(dirtyCheck, "docxpatch 1\n\nop set-field-dirty\ntarget M.F0001\ndirty true\nend\n", TrackChangesMode.Require);
        Assert.False(dirty.Success);
        Assert.Contains(dirty.Diagnostics, static diagnostic => diagnostic.Code == "E6001");
    }

    [Fact]
    public static void UnknownFieldFails()
    {
        using MemoryStream input = CreateDocxWithRefField();
        input.Position = 0;
        DocxCapabilitiesResult missing = new DocxEditor().GetCapabilities(input, "M.F0009");
        Assert.False(missing.Success);
        Assert.Null(missing.Capabilities);
        Assert.Contains(missing.Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }

    [Fact]
    public static void ComplexFieldLimitsCodeAndRefresh()
    {
        using MemoryStream input = CreateDocxWithBody(MixedBody);
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.F0002", TrackChangesMode.Off);

        Assert.Equal("unsupported", FindOperation(capabilities, "set-field-code").Support);
        Assert.Equal("unsupported", FindOperation(capabilities, "refresh-field-result").Support);
        Assert.Equal("supported", FindOperation(capabilities, "set-field-result").Support);

        using MemoryStream codeCheck = CreateDocxWithBody(MixedBody);
        DocxCheckResult code = RunCheck(codeCheck, "docxpatch 1\n\nop set-field-code\ntarget M.F0002\ncode PAGE\nend\n", TrackChangesMode.Off);
        Assert.False(code.Success);
        Assert.Contains(code.Diagnostics, static diagnostic => diagnostic.Code == "E4313");

        using MemoryStream refreshCheck = CreateDocxWithBody(MixedBody);
        DocxCheckResult refresh = RunCheck(refreshCheck, "docxpatch 1\n\nop refresh-field-result\ntarget M.F0002\nend\n", TrackChangesMode.Off);
        Assert.False(refresh.Success);
        Assert.Contains(refresh.Diagnostics, static diagnostic => diagnostic.Code == "E4313");

        using MemoryStream resultCheck = CreateDocxWithBody(MixedBody);
        DocxCheckResult result = RunCheck(resultCheck, "docxpatch 1\n\nop set-field-result\ntarget M.F0002\ntext 2\nend\n", TrackChangesMode.Off);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }
}