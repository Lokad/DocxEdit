using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D17: image capabilities reuse the drawing-shape predicates from the image
// engine.
public static class ImageCapabilitiesTests
{
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
    public static void SimpleInlineImageVerdicts()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.I0001", TrackChangesMode.Off);

        Assert.Equal("M.I0001", capabilities.TargetId);
        Assert.Equal("image", capabilities.Kind);
        Assert.Equal("main", capabilities.Story);
        string[] expectedOrder = ["replace-image", "set-image-alt", "set-image-metadata", "set-image-size", "set-image-wrap", "set-image-position", "set-image-crop", "delete-image"];
        Assert.Equal(expectedOrder, capabilities.Operations.Select(static operation => operation.Operation).ToArray());
        Assert.Equal("conditional", FindOperation(capabilities, "replace-image").Support);
        Assert.Equal("supported", FindOperation(capabilities, "set-image-alt").Support);
        Assert.Equal("supported", FindOperation(capabilities, "set-image-metadata").Support);
        Assert.Equal("supported", FindOperation(capabilities, "set-image-size").Support);
        Assert.Equal("unsupported", FindOperation(capabilities, "set-image-wrap").Support);
        Assert.Equal("unsupported", FindOperation(capabilities, "set-image-position").Support);
        Assert.Equal("supported", FindOperation(capabilities, "set-image-crop").Support);
        Assert.Equal("supported", FindOperation(capabilities, "delete-image").Support);

        using MemoryStream altInput = CreateDocxWithImage("png", "image/png", "old-png");
        DocxCheckResult alt = RunCheck(altInput, "docxpatch 1\n\nop set-image-alt\ntarget M.I0001\nalt New chart\nend\n", TrackChangesMode.Off);
        Assert.True(alt.Success, string.Join("|", alt.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));

        using MemoryStream wrapInput = CreateDocxWithImage("png", "image/png", "old-png");
        DocxCheckResult wrap = RunCheck(wrapInput, "docxpatch 1\n\nop set-image-wrap\ntarget M.I0001\nmode square\nend\n", TrackChangesMode.Off);
        Assert.False(wrap.Success);
        Assert.Contains(wrap.Diagnostics, static diagnostic => diagnostic.Code == "E5205");

        using MemoryStream deleteInput = CreateDocxWithImage("png", "image/png", "old-png");
        DocxCheckResult deleted = RunCheck(deleteInput, "docxpatch 1\n\nop delete-image\ntarget M.I0001\nend\n", TrackChangesMode.Off);
        Assert.True(deleted.Success, string.Join("|", deleted.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void RequireModeRefusesAllImageOps()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.I0001", TrackChangesMode.Require);

        foreach (string operation in new[] { "replace-image", "set-image-alt", "set-image-metadata", "set-image-size", "set-image-wrap", "set-image-position", "set-image-crop", "delete-image" })
        {
            Assert.Equal("unsupported", FindOperation(capabilities, operation).Support);
        }

        using MemoryStream deleteCheck = CreateDocxWithImage("png", "image/png", "old-png");
        DocxCheckResult deleted = RunCheck(deleteCheck, "docxpatch 1\n\nop delete-image\ntarget M.I0001\nend\n", TrackChangesMode.Require);
        Assert.False(deleted.Success);
        Assert.Contains(deleted.Diagnostics, static diagnostic => diagnostic.Code == "E6001");
    }

    [Fact]
    public static void UnknownImageFails()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        input.Position = 0;
        DocxCapabilitiesResult missing = new DocxEditor().GetCapabilities(input, "M.I0009");
        Assert.False(missing.Success);
        Assert.Null(missing.Capabilities);
        Assert.Contains(missing.Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }
}