using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D17: content-control capabilities reuse the execution gates for kind,
// locks, containers, protected content, and tracked shapes.
public static class ContentControlCapabilitiesTests
{
    private const string PlainBody = """
                <w:p>
                  <w:sdt>
                    <w:sdtPr><w:text/><w:tag w:val="client"/></w:sdtPr>
                    <w:sdtContent><w:r><w:t>Old Client</w:t></w:r></w:sdtContent>
                  </w:sdt>
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
    public static void PlainTextControlSupportsTextAndNamesKindAlternatives()
    {
        using MemoryStream input = CreateDocxWithBody(PlainBody);
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.CC0001", TrackChangesMode.Off);

        Assert.Equal("M.CC0001", capabilities.TargetId);
        Assert.Equal("content-control", capabilities.Kind);
        Assert.Equal("main", capabilities.Story);
        string[] expectedOrder = ["set-content-control-text", "set-content-control-checkbox", "set-content-control-choice", "set-content-control-date"];
        Assert.Equal(expectedOrder, capabilities.Operations.Select(static operation => operation.Operation).ToArray());
        Assert.Equal("supported", FindOperation(capabilities, "set-content-control-text").Support);
        Assert.Equal("unsupported", FindOperation(capabilities, "set-content-control-checkbox").Support);
        Assert.Equal("set-content-control-text", FindOperation(capabilities, "set-content-control-checkbox").Alternative);
        Assert.Equal("unsupported", FindOperation(capabilities, "set-content-control-choice").Support);
        Assert.Equal("unsupported", FindOperation(capabilities, "set-content-control-date").Support);
        foreach (DocxOperationCapability operation in capabilities.Operations)
        {
            Assert.Equal(operation.Operation, operation.HelpTopic);
            Assert.False(string.IsNullOrWhiteSpace(operation.Reason));
        }

        using MemoryStream checkInput = CreateDocxWithBody(PlainBody);
        DocxCheckResult result = RunCheck(checkInput, "docxpatch 1\n\nop set-content-control-text\ntarget M.CC0001\ntext New Client\nend\n", TrackChangesMode.Off);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void LockedControlRefusesEveryEdit()
    {
        const string body = """
                    <w:p>
                      <w:sdt>
                        <w:sdtPr><w:text/><w:lock w:val="sdtContentLocked"/></w:sdtPr>
                        <w:sdtContent><w:r><w:t>Old Client</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """;
        using MemoryStream input = CreateDocxWithBody(body);
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.CC0001", TrackChangesMode.Off);

        // Execution checks kind before lock, so only the text operation
        // reaches the lock refusal on this plain-text control.
        Assert.Contains("sdtContentLocked", FindOperation(capabilities, "set-content-control-text").Reason, StringComparison.Ordinal);
        foreach (DocxOperationCapability operation in capabilities.Operations)
        {
            Assert.Equal("unsupported", operation.Support);
        }

        using MemoryStream checkInput = CreateDocxWithBody(body);
        DocxCheckResult result = RunCheck(checkInput, "docxpatch 1\n\nop set-content-control-text\ntarget M.CC0001\ntext New Client\nend\n", TrackChangesMode.Off);
        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E4310");
    }    [Fact]
    public static void CheckboxControlSupportsCheckboxExceptUnderRequire()
    {
        const string body = """
                    <w:p>
                      <w:sdt>
                        <w:sdtPr><w:checkBox><w:checked w:val="0"/></w:checkBox><w:tag w:val="accepted"/></w:sdtPr>
                        <w:sdtContent><w:r><w:t>Unchecked</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """;
        using MemoryStream offInput = CreateDocxWithBody(body);
        DocxTargetCapabilities off = GetCapabilities(offInput, "M.CC0001", TrackChangesMode.Off);
        Assert.Equal("supported", FindOperation(off, "set-content-control-checkbox").Support);
        Assert.Equal("unsupported", FindOperation(off, "set-content-control-text").Support);
        Assert.Equal("set-content-control-checkbox", FindOperation(off, "set-content-control-text").Alternative);

        using MemoryStream requireInput = CreateDocxWithBody(body);
        DocxTargetCapabilities require = GetCapabilities(requireInput, "M.CC0001", TrackChangesMode.Require);
        DocxOperationCapability checkbox = FindOperation(require, "set-content-control-checkbox");
        Assert.Equal("unsupported", checkbox.Support);
        Assert.Contains("E6001", checkbox.Reason, StringComparison.Ordinal);

        using MemoryStream checkInput = CreateDocxWithBody(body);
        DocxCheckResult refused = RunCheck(checkInput, "docxpatch 1\n\nop set-content-control-checkbox\ntarget M.CC0001\nchecked true\nend\n", TrackChangesMode.Require);
        Assert.False(refused.Success);
        Assert.Contains(refused.Diagnostics, static diagnostic => diagnostic.Code == "E6001");
    }

    [Fact]
    public static void DropdownControlSupportsChoiceSelection()
    {
        const string body = """
                    <w:p>
                      <w:sdt>
                        <w:sdtPr><w:dropDownList><w:listItem w:displayText="North" w:value="north"/><w:listItem w:displayText="South" w:value="south"/></w:dropDownList></w:sdtPr>
                        <w:sdtContent><w:r><w:t>South</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """;
        using MemoryStream input = CreateDocxWithBody(body);
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.CC0001", TrackChangesMode.Off);
        Assert.Equal("supported", FindOperation(capabilities, "set-content-control-choice").Support);
        Assert.Equal("unsupported", FindOperation(capabilities, "set-content-control-date").Support);

        using MemoryStream checkInput = CreateDocxWithBody(body);
        DocxCheckResult result = RunCheck(checkInput, "docxpatch 1\n\nop set-content-control-choice\ntarget M.CC0001\nvalue north\nend\n", TrackChangesMode.Off);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void RichTextWithFieldRefusesTextReplacement()
    {
        const string body = """
                    <w:p>
                      <w:sdt>
                        <w:sdtPr><w:richText/><w:tag w:val="summary"/></w:sdtPr>
                        <w:sdtContent>
                          <w:p><w:fldSimple w:instr=" REF Mark "><w:r><w:t>Old summary</w:t></w:r></w:fldSimple></w:p>
                        </w:sdtContent>
                      </w:sdt>
                    </w:p>
            """;
        using MemoryStream input = CreateDocxWithBody(body);
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.CC0001", TrackChangesMode.Off);
        DocxOperationCapability text = FindOperation(capabilities, "set-content-control-text");
        Assert.Equal("unsupported", text.Support);
        Assert.Contains("field", text.Reason, StringComparison.Ordinal);

        using MemoryStream checkInput = CreateDocxWithBody(body);
        DocxCheckResult result = RunCheck(checkInput, "docxpatch 1\n\nop set-content-control-text\ntarget M.CC0001\nexpect-text Old summary\ntext New summary\nend\n", TrackChangesMode.Off);
        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E4310");
    }

    [Fact]
    public static void PictureControlPointsToImageGuidance()
    {
        const string body = """
                    <w:p>
                      <w:sdt>
                        <w:sdtPr><w:picture/></w:sdtPr>
                        <w:sdtContent><w:r><w:t>Image placeholder</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """;
        using MemoryStream input = CreateDocxWithBody(body);
        DocxTargetCapabilities capabilities = GetCapabilities(input, "M.CC0001", TrackChangesMode.Off);
        DocxOperationCapability text = FindOperation(capabilities, "set-content-control-text");
        Assert.Equal("unsupported", text.Support);
        Assert.Contains("picture", text.Reason, StringComparison.Ordinal);
        Assert.Null(text.Alternative);
    }

    [Fact]
    public static void UnknownControlOrdinalFailsWithE1201()
    {
        using MemoryStream input = CreateDocxWithBody(PlainBody);
        input.Position = 0;
        DocxCapabilitiesResult result = new DocxEditor().GetCapabilities(input, "M.CC0009");
        Assert.False(result.Success);
        Assert.Null(result.Capabilities);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }
}
