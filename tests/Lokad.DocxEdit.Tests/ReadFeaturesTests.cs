using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml.Linq;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

public static class ReadFeaturesTests
{

    [Fact]
    public static void ReadWarnsOnUnsupportedNumberingShapes()
    {
        using MemoryStream stream = CreateDocxWithStylesAndNumbering(
            """
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr>
                      <w:r><w:t>Custom format</w:t></w:r>
                    </w:p>
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="10"/></w:numPr></w:pPr>
                      <w:r><w:t>Picture bullet</w:t></w:r>
                    </w:p>
            """,
            """
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
            """,
            """
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:numPicBullet w:numPicBulletId="1"/>
                  <w:abstractNum w:abstractNumId="7">
                    <w:lvl w:ilvl="0">
                      <w:numFmt w:val="chicago"/>
                      <w:lvlText w:val="%1."/>
                    </w:lvl>
                  </w:abstractNum>
                  <w:abstractNum w:abstractNumId="8">
                    <w:lvl w:ilvl="0">
                      <w:numFmt w:val="bullet"/>
                      <w:lvlText w:val="%1"/>
                      <w:lvlPicBulletId w:val="1"/>
                    </w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="9"><w:abstractNumId w:val="7"/></w:num>
                  <w:num w:numId="10"><w:abstractNumId w:val="8"/></w:num>
                </w:numbering>
            """);

        DocxReadResult result = new DocxEditor().Read(stream);

        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "W1024" &&
            diagnostic.Feature == "numbering" &&
            diagnostic.Fallback == "unsupported-picture-bullet");
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "W1025" &&
            diagnostic.Feature == "numbering" &&
            diagnostic.Fallback == "unsupported-numbering-format");
        Assert.Contains("unsupported-format-level-0", result.Paragraphs[0].List?.LabelWarnings ?? []);
    }

    [Fact]
    public static void ReadContentControlsExposeSafeEditReasonsForUnsupportedKinds()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:picture/>
                          <w:alias w:val="Logo"/>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>Image placeholder</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:group/>
                          <w:tag w:val="review-group"/>
                        </w:sdtPr>
                        <w:sdtContent>
                          <w:sdt>
                            <w:sdtPr><w:text/></w:sdtPr>
                            <w:sdtContent><w:r><w:t>Editable child</w:t></w:r></w:sdtContent>
                          </w:sdt>
                        </w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);

        DocxReadResult result = new DocxEditor().Read(stream);

        DocxContentControlInfo picture = result.ContentControls.Single(control => control.Kind == "picture");
        Assert.Equal("unsupported-picture", picture.SafeEditStatus);
        Assert.Contains("picture controls", picture.SafeEditReason, StringComparison.Ordinal);
        DocxContentControlInfo group = result.ContentControls.Single(control => control.Kind == "group");
        Assert.Equal("unsupported-group", group.SafeEditStatus);
        Assert.Contains("target an editable child content control", group.SafeEditReason, StringComparison.Ordinal);
        Assert.Equal(new[] { "M.CC0003" }, group.ChildContentControlIds);
        Assert.Contains("safe-edit-reason=\"picture controls preserve a picture container", result.Text, StringComparison.Ordinal);
        Assert.Contains("safe-edit-reason=\"group controls protect a container", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public static void ReadFindAndDumpReportUnsupportedFeatureDiagnosticsConsistently()
    {
        using MemoryStream readStream = CreateDocxWithBody("""
                    <w:p xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                      <w:hyperlink r:id="rLink"><w:r><w:t>Link</w:t></w:r></w:hyperlink>
                      <w:fldSimple w:instr="DATE"><w:r><w:t>Revenue</w:t></w:r></w:fldSimple>
                      <w:commentRangeStart w:id="1"/>
                      <w:ins w:id="2" w:author="A"><w:r><w:t>Inserted</w:t></w:r></w:ins>
                    </w:p>
            """);
        using MemoryStream findStream = CreateDocxWithBody("""
                    <w:p xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                      <w:hyperlink r:id="rLink"><w:r><w:t>Link</w:t></w:r></w:hyperlink>
                      <w:fldSimple w:instr="DATE"><w:r><w:t>Revenue</w:t></w:r></w:fldSimple>
                      <w:commentRangeStart w:id="1"/>
                      <w:ins w:id="2" w:author="A"><w:r><w:t>Inserted</w:t></w:r></w:ins>
                    </w:p>
            """);
        using MemoryStream dumpStream = CreateDocxWithBody("""
                    <w:p xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                      <w:hyperlink r:id="rLink"><w:r><w:t>Link</w:t></w:r></w:hyperlink>
                      <w:fldSimple w:instr="DATE"><w:r><w:t>Revenue</w:t></w:r></w:fldSimple>
                      <w:commentRangeStart w:id="1"/>
                      <w:ins w:id="2" w:author="A"><w:r><w:t>Inserted</w:t></w:r></w:ins>
                    </w:p>
            """);
        var editor = new DocxEditor();

        DocxReadResult read = editor.Read(readStream);
        DocxFindResult find = editor.Find(findStream, "Revenue");
        DocxDumpResult dump = editor.Dump(dumpStream, "M.P0001");

        AssertUnsupportedFeatureDiagnostics(read.Diagnostics);
        AssertUnsupportedFeatureDiagnostics(find.Diagnostics);
        AssertUnsupportedFeatureDiagnostics(dump.Diagnostics);
    }

    [Fact]
    public static void ReadWarnsAboutUnsupportedPreservedObjects()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                        xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                      <w:r>
                        <w:drawing>
                          <c:chart r:id="rChart"/>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:altChunk xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rAltChunk"/>
            """);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.True(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1009" && diagnostic.Feature == "chart");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1013" && diagnostic.Feature == "alt-chunk");
    }

    [Fact]
    public static void ReadWarnsAboutUnsupportedDrawingShapes()
    {
        using MemoryStream stream = CreateDocxWithBody(
            """
                    <w:p
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                        xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
                        xmlns:v="urn:schemas-microsoft-com:vml"
                        xmlns:o="urn:schemas-microsoft-com:office:office">
                      <w:r>
                        <w:drawing>
                          <wp:inline>
                            <a:graphic>
                              <a:graphicData>
                                <a:blip r:link="rLinkedImage"/>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                      <w:r>
                        <w:drawing>
                          <wp:inline>
                            <a:graphic>
                              <a:graphicData>
                                <wpg:wgp/>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                      <w:r><w:pict><v:shape id="v1"/></w:pict></w:r>
                      <w:r><w:object><o:OLEObject Type="Embed"/></w:object></w:r>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLinkedImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.test/linked.png" TargetMode="External"/>
                </Relationships>
                """, null);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        Assert.True(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1008" && diagnostic.Feature == "external-image");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1019" && diagnostic.Feature == "linked-image");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1020" && diagnostic.Feature == "vml");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1021" && diagnostic.Feature == "grouped-drawing");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W1022" && diagnostic.Feature == "ole-object");
    }

    [Fact]
    public static void ReadUnsupportedFeatureDiagnosticsCarryStableMetadata()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                        xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                      <w:bookmarkStart w:id="1" w:name="Bookmark"/>
                      <w:hyperlink r:id="rHyperlink"><w:r><w:t>Link</w:t></w:r></w:hyperlink>
                      <w:fldSimple w:instr="DATE"><w:r><w:t>Date</w:t></w:r></w:fldSimple>
                      <w:commentRangeStart w:id="1"/>
                      <w:sdt><w:sdtContent><w:r><w:t>Control</w:t></w:r></w:sdtContent></w:sdt>
                      <w:ins w:id="2" w:author="A"><w:r><w:t>Revision</w:t></w:r></w:ins>
                      <w:r>
                        <w:drawing>
                          <wp:anchor>
                            <a:graphic>
                              <a:graphicData>
                                <c:chart r:id="rChart"/>
                              </a:graphicData>
                            </a:graphic>
                          </wp:anchor>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:altChunk xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rAltChunk"/>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rExternalImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.test/image.png" TargetMode="External"/>
                </Relationships>
                """, null);
        var editor = new DocxEditor();

        DocxReadResult result = editor.Read(stream);

        DocxDiagnostic[] warnings = result.Diagnostics
            .Where(diagnostic => diagnostic.Code.StartsWith("W10", StringComparison.Ordinal))
            .ToArray();
        Assert.NotEmpty(warnings);
        Assert.All(warnings, diagnostic =>
        {
            Assert.Equal(DocxSeverity.Warning, diagnostic.Severity);
            Assert.Matches("^W10[0-9]{2}$", diagnostic.Code);
            Assert.Equal("/word/document.xml", diagnostic.PartName);
            Assert.Equal("main", diagnostic.Story);
            Assert.False(string.IsNullOrWhiteSpace(diagnostic.Feature));
            Assert.False(string.IsNullOrWhiteSpace(diagnostic.Fallback));
        });
        Assert.Contains(warnings, diagnostic => diagnostic.Code == "W1001" && diagnostic.Fallback == "selected-text-view");
        Assert.Contains(warnings, diagnostic => diagnostic.Code == "W1008" && diagnostic.Fallback == "omit-from-editable-images");
    }

    private static void AssertUnsupportedFeatureDiagnostics(IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "W1001" && diagnostic.Feature == "tracked-changes");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "W1002" && diagnostic.Feature == "hyperlink");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "W1003" && diagnostic.Feature == "field");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "W1004" && diagnostic.Feature == "comment");
    }
}
