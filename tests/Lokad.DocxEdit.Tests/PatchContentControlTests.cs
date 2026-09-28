using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml.Linq;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

public static class PatchContentControlTests
{

    [Fact]
    public static void CheckContentControlSelectorResolvesAndRejectsProtectedTextEdit()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:alias w:val="Client Name"/>
                          <w:tag w:val="client-name"/>
                        </w:sdtPr>
                        <w:sdtContent>
                          <w:r><w:t>Client</w:t></w:r>
                        </w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target content-control:"client-name"
            find Client
            with Customer
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E4305" && diagnostic.TargetId == "content-control:\"client-name\"");
    }

    [Fact]
    public static void ApplySetContentControlTextPreservesSdtProperties()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:text/>
                          <w:id w:val="99"/>
                          <w:tag w:val="client-name"/>
                          <w:alias w:val="Client Name"/>
                          <w:lock w:val="unlocked"/>
                          <w:placeholder><w:docPart w:val="DefaultPlaceholder"/></w:placeholder>
                          <w:dataBinding w:xpath="/root/client" w:storeItemID="{11111111-1111-1111-1111-111111111111}" w:prefixMappings="xmlns:ns='urn:test'"/>
                        </w:sdtPr>
                        <w:sdtContent>
                          <w:r><w:t>Old Client</w:t></w:r>
                        </w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-content-control-text
            target M.CC0001
            text New Client
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal("New Client", Assert.Single(read.Paragraphs).Text);
        DocxContentControlInfo control = Assert.Single(read.ContentControls);
        Assert.Equal("plain-text", control.Kind);
        Assert.Equal("99", control.OoxmlId);
        Assert.Equal("client-name", control.Tag);
        Assert.Equal("Client Name", control.Alias);
        Assert.Equal("unlocked", control.Lock);
        Assert.Equal("DefaultPlaceholder", control.PlaceholderDocPart);
        Assert.Equal("/root/client", control.DataBindingXPath);
        Assert.Equal("{11111111-1111-1111-1111-111111111111}", control.DataBindingStoreItemId);
        Assert.Equal("xmlns:ns='urn:test'", control.DataBindingPrefixMappings);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:sdt>", xml, StringComparison.Ordinal);
        Assert.Contains("w:id w:val=\"99\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:tag w:val=\"client-name\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:lock w:val=\"unlocked\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:docPart w:val=\"DefaultPlaceholder\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:dataBinding w:xpath=\"/root/client\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetRichTextContentControlRequiresGuardAndPreservesWrapper()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:id w:val="77"/>
                          <w:tag w:val="summary"/>
                          <w:richText/>
                        </w:sdtPr>
                        <w:sdtContent>
                          <w:p><w:r><w:t>Old summary</w:t></w:r></w:p>
                        </w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-content-control-text
            target M.CC0001
            expect-text Old summary
            text New summary
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        DocxContentControlInfo control = Assert.Single(read.ContentControls);
        Assert.Equal("rich-text", control.Kind);
        Assert.Equal("rich-text", control.SafeEditStatus);
        Assert.Equal("New summary", Assert.Single(read.Paragraphs).Text);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:richText", xml, StringComparison.Ordinal);
        Assert.Contains("w:tag w:val=\"summary\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckSetRichTextContentControlRejectsMissingGuard()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr><w:richText/></w:sdtPr>
                        <w:sdtContent>
                          <w:p><w:r><w:t>Old summary</w:t></w:r></w:p>
                        </w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-content-control-text
            target M.CC0001
            text New summary
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal("E4205", diagnostic.Code);
        Assert.Equal("M.CC0001", diagnostic.TargetId);
        Assert.Contains("requires expect-text", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckSetContentControlTextRejectsPictureControlWithImageGuidance()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr><w:picture/></w:sdtPr>
                        <w:sdtContent><w:r><w:t>Image placeholder</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-content-control-text
            target M.CC0001
            text Replacement
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal("E4310", diagnostic.Code);
        Assert.Equal("M.CC0001", diagnostic.TargetId);
        Assert.Contains("kind 'picture'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("read/media", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("image operations", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckSetContentControlTextRejectsGroupControlWithChildGuidance()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr><w:group/></w:sdtPr>
                        <w:sdtContent>
                          <w:sdt>
                            <w:sdtPr><w:text/></w:sdtPr>
                            <w:sdtContent><w:r><w:t>Editable child</w:t></w:r></w:sdtContent>
                          </w:sdt>
                        </w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var groupPatch = new StringReader("""
            docxpatch 1

            op set-content-control-text
            target M.CC0001
            text Replacement
            end
            """);

        DocxCheckResult groupResult = new DocxEditor().Check(input, groupPatch);

        Assert.False(groupResult.Success);
        DocxDiagnostic diagnostic = Assert.Single(groupResult.Diagnostics);
        Assert.Equal("E4310", diagnostic.Code);
        Assert.Equal("M.CC0001", diagnostic.TargetId);
        Assert.Contains("kind 'group'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("target an editable child content control", diagnostic.Message, StringComparison.Ordinal);

        input.Position = 0;
        using var childPatch = new StringReader("""
            docxpatch 1

            op set-content-control-text
            target M.CC0002
            text Replacement
            end
            """);

        DocxCheckResult childResult = new DocxEditor().Check(input, childPatch);

        Assert.True(childResult.Success);
    }

    [Fact]
    public static void CheckSetRichTextContentControlRejectsNestedContentControlBoundary()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr><w:richText/></w:sdtPr>
                        <w:sdtContent>
                          <w:p>
                            <w:r><w:t>Outer </w:t></w:r>
                            <w:sdt>
                              <w:sdtPr><w:text/></w:sdtPr>
                              <w:sdtContent><w:r><w:t>Inner</w:t></w:r></w:sdtContent>
                            </w:sdt>
                          </w:p>
                        </w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-content-control-text
            target M.CC0001
            expect-text Outer Inner
            text Replacement
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal("E4310", diagnostic.Code);
        Assert.Equal("M.CC0001", diagnostic.TargetId);
        Assert.Contains("protected OOXML boundary 'content-control'", diagnostic.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("set-content-control-text", "<w:text/>", "text New Client")]
    [InlineData("set-content-control-text", "<w:richText/>", "expect-text Old Client\ntext New Client")]
    [InlineData("set-content-control-checkbox", "<w:checkBox><w:checked w:val=\"0\"/></w:checkBox>", "checked true")]
    [InlineData("set-content-control-choice", "<w:dropDownList><w:listItem w:displayText=\"North\" w:value=\"north\"/><w:listItem w:displayText=\"South\" w:value=\"south\"/></w:dropDownList>", "value south")]
    [InlineData("set-content-control-date", "<w:date><w:fullDate w:val=\"2026-06-12T00:00:00Z\"/></w:date>", "value 2026-07-01T00:00:00Z")]
    public static void CheckContentControlEditsRejectLockedControls(
        string operationName,
        string kindXml,
        string operationFields)
    {
        using MemoryStream input = CreateDocxWithBody($"""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          {kindXml}
                          <w:lock w:val="sdtContentLocked"/>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>Old Client</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            target M.CC0001
            {operationFields}
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal("E4310", diagnostic.Code);
        Assert.Equal("M.CC0001", diagnostic.TargetId);
        Assert.Contains("locked by w:lock='sdtContentLocked'", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetContentControlCheckboxUpdatesStateAndDisplay()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:checkBox>
                            <w:checked w:val="0"/>
                            <w:checkedState w:val="2612"/>
                            <w:uncheckedState w:val="2610"/>
                          </w:checkBox>
                          <w:tag w:val="accepted"/>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>Unchecked</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-content-control-checkbox
            target M.CC0001
            checked true
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        DocxContentControlInfo control = Assert.Single(read.ContentControls);
        Assert.True(control.Checked);
        Assert.Equal(char.ConvertFromUtf32(0x2612), Assert.Single(read.Paragraphs).Text);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:checked w:val=\"1\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetContentControlChoiceUpdatesDropdownDisplay()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:dropDownList>
                            <w:listItem w:displayText="North" w:value="north"/>
                            <w:listItem w:displayText="South" w:value="south"/>
                          </w:dropDownList>
                          <w:tag w:val="region"/>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>North</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-content-control-choice
            target M.CC0001
            value south
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        DocxContentControlInfo control = Assert.Single(read.ContentControls);
        Assert.Equal("dropdown-list", control.Kind);
        Assert.Equal("South", Assert.Single(read.Paragraphs).Text);
        Assert.Equal(2, control.ListItems.Count);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:sdt>", xml, StringComparison.Ordinal);
        Assert.Contains("w:tag w:val=\"region\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetContentControlDateUpdatesValueAndDisplay()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:date>
                            <w:dateFormat w:val="yyyy-MM-dd"/>
                            <w:fullDate w:val="2026-06-12T00:00:00Z"/>
                          </w:date>
                          <w:tag w:val="deadline"/>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>2026-06-12</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-content-control-date
            target M.CC0001
            value 2026-07-01T00:00:00Z
            display-text 2026-07-01
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        DocxContentControlInfo control = Assert.Single(read.ContentControls);
        Assert.Equal("date", control.Kind);
        Assert.Equal("2026-07-01T00:00:00Z", control.DateValue);
        Assert.Equal("2026-07-01", Assert.Single(read.Paragraphs).Text);
    }

    [Theory]
    [InlineData("add-repeating-section-item", "index 1\ntext Added")]
    [InlineData("delete-repeating-section-item", "index 1")]
    public static void CheckRepeatingSectionItemOperationsFailWithExplicitUnsupportedDiagnostic(
        string operationName,
        string operationFields)
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:repeatingSection w:sectionTitle="Line items"/>
                        </w:sdtPr>
                        <w:sdtContent>
                          <w:sdt>
                            <w:sdtPr><w:repeatingSectionItem/></w:sdtPr>
                            <w:sdtContent><w:r><w:t>Existing</w:t></w:r></w:sdtContent>
                          </w:sdt>
                        </w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            target M.CC0001
            {operationFields}
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E4315" &&
            diagnostic.TargetId == "M.CC0001" &&
            diagnostic.Message.Contains("repeating-section item edits", StringComparison.Ordinal));
    }

    [Fact]
    public static void ApplySetContentControlTextWithTagSelector()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:text/>
                          <w:tag w:val="amount"/>
                        </w:sdtPr>
                        <w:sdtContent>
                          <w:r><w:t>Old</w:t></w:r>
                        </w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-content-control-text
            target content-control:"amount"
            text New
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Equal("New", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void ApplySetContentControlTextWithAliasSelector()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:text/>
                          <w:tag w:val="amount"/>
                          <w:alias w:val="Amount due"/>
                        </w:sdtPr>
                        <w:sdtContent>
                          <w:r><w:t>Old</w:t></w:r>
                        </w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-content-control-text
            target content-control:"Amount due"
            text New
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Equal("New", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void CheckSetContentControlTextWithUnknownTagFails()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:text/>
                          <w:tag w:val="amount"/>
                        </w:sdtPr>
                        <w:sdtContent>
                          <w:r><w:t>Old</w:t></w:r>
                        </w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-content-control-text
            target content-control:"missing"
            text New
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }

    [Fact]
    public static void CheckSetContentControlTextWithDuplicateTagFails()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:text/>
                          <w:tag w:val="amount"/>
                        </w:sdtPr>
                        <w:sdtContent>
                          <w:r><w:t>First</w:t></w:r>
                        </w:sdtContent>
                      </w:sdt>
                    </w:p>
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:text/>
                          <w:tag w:val="amount"/>
                        </w:sdtPr>
                        <w:sdtContent>
                          <w:r><w:t>Second</w:t></w:r>
                        </w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-content-control-text
            target content-control:"amount"
            text New
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, static d => d.Code == "E1202");
        Assert.Contains("M.CC0001", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("M.CC0002", diagnostic.Message, StringComparison.Ordinal);
    }

}
