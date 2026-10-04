using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml.Linq;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

public static class ReadValidateTests
{

    [Fact]
    public static void ValidateAcceptsBasicDocument()
    {
        using MemoryStream stream = CreateDocx();
        var editor = new DocxEditor();

        DocxValidateResult result = editor.Validate(stream);

        Assert.True(result.Success);
        Assert.Equal(DocxValidationProfile.Structural, result.Profile);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error);
        Assert.Contains("/word/document.xml", result.PartNames);
        Assert.Equal("/word/document.xml", result.MainDocumentPartName);
    }

    [Fact]
    public static void ValidatePackageProfileSkipsWordprocessingInvariants()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="1" w:name="Unclosed"/>
                      <w:r><w:t>Text</w:t></w:r>
                    </w:p>
            """);
        var editor = new DocxEditor();

        DocxValidateResult result = editor.Validate(stream, new DocxValidateOptions { Profile = DocxValidationProfile.Package });

        Assert.True(result.Success);
        Assert.Equal(DocxValidationProfile.Package, result.Profile);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "E9103");
    }

    [Theory]
    [InlineData("basic-document", true, null)]
    [InlineData("unclosed-bookmark", false, "E9103")]
    [InlineData("invalid-relationship-root", false, "E9102")]
    [InlineData("missing-style-definition", true, "W9116")]
    public static void ValidateFixtureCorpusCoversKnownGoodAndMalformedDocuments(
        string fixtureName,
        bool expectedSuccess,
        string? expectedDiagnosticCode)
    {
        using MemoryStream stream = CreateValidationFixture(fixtureName);

        DocxValidateResult result = new DocxEditor().Validate(stream);

        Assert.Equal(expectedSuccess, result.Success);
        if (expectedDiagnosticCode is null)
        {
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error);
        }
        else
        {
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == expectedDiagnosticCode);
        }
    }

    [Fact]
    public static void ValidateWarnsOnDuplicateSemanticSelectors()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="1" w:name="Shared"/>
                      <w:r><w:t>First</w:t></w:r>
                      <w:bookmarkEnd w:id="1"/>
                    </w:p>
                    <w:p>
                      <w:bookmarkStart w:id="2" w:name="Shared"/>
                      <w:r><w:t>Second</w:t></w:r>
                      <w:bookmarkEnd w:id="2"/>
                    </w:p>
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:tag w:val="shared_tag"/>
                          <w:alias w:val="Shared Alias"/>
                          <w:text/>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>One</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:tag w:val="shared_tag"/>
                          <w:alias w:val="Shared Alias"/>
                          <w:text/>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>Two</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        var editor = new DocxEditor();

        DocxValidateResult result = editor.Validate(stream);

        Assert.True(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "W9109" &&
            diagnostic.Feature == "bookmark" &&
            diagnostic.Fallback == "ambiguous-selector" &&
            diagnostic.Message.Contains("M.B0001, M.B0002", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "W9109" &&
            diagnostic.Feature == "content-control" &&
            diagnostic.Message.Contains("content-control tag", StringComparison.Ordinal) &&
            diagnostic.Message.Contains("M.CC0001, M.CC0002", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "W9109" &&
            diagnostic.Feature == "content-control" &&
            diagnostic.Message.Contains("content-control alias", StringComparison.Ordinal) &&
            diagnostic.Message.Contains("M.CC0001, M.CC0002", StringComparison.Ordinal));
    }

    [Fact]
    public static void ValidateWarnsOnMissingParagraphStyleDefinition()
    {
        using MemoryStream stream = CreateDocxWithStylesAndNumbering(
            """
                    <w:p>
                      <w:pPr><w:pStyle w:val="MissingStyle"/></w:pPr>
                      <w:r><w:t>Styled text</w:t></w:r>
                    </w:p>
            """,
            """
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="paragraph" w:styleId="KnownStyle">
                    <w:name w:val="Known Style"/>
                  </w:style>
                </w:styles>
                """,
            """
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
                """);

        DocxValidateResult result = new DocxEditor().Validate(stream);

        Assert.True(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "W9116" &&
            diagnostic.Feature == "style" &&
            diagnostic.Fallback == "missing-style-definition" &&
            diagnostic.Message.Contains("MissingStyle", StringComparison.Ordinal));
    }

    [Fact]
    public static void ValidateWarnsOnMissingNumberingDefinitions()
    {
        using MemoryStream stream = CreateDocxWithStylesAndNumbering(
            """
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="42"/></w:numPr></w:pPr>
                      <w:r><w:t>Numbered text</w:t></w:r>
                    </w:p>
            """,
            """
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
                """,
            """
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:num w:numId="9"><w:abstractNumId w:val="7"/></w:num>
                </w:numbering>
                """);

        DocxValidateResult result = new DocxEditor().Validate(stream);

        Assert.True(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "W9117" &&
            diagnostic.Feature == "numbering" &&
            diagnostic.Fallback == "missing-numbering-definition" &&
            diagnostic.Message.Contains("'42'", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "W9117" &&
            diagnostic.Feature == "numbering" &&
            diagnostic.Fallback == "missing-abstract-numbering-definition" &&
            diagnostic.Message.Contains("'7'", StringComparison.Ordinal));
    }

    [Fact]
    public static void ValidateReportsInvalidSettingsMetadata()
    {
        using MemoryStream stream = CreateDocxWithBody(
            """
                    <w:p><w:r><w:t>Body</w:t></w:r></w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
                </Relationships>
                """,
            extra: archive => AddEntry(archive, "word/settings.xml", """
                <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:updateFields w:val="maybe"/>
                </w:settings>
                """));

        DocxValidateResult result = new DocxEditor().Validate(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9118" &&
            diagnostic.PartName == "/word/settings.xml" &&
            diagnostic.Message.Contains("updateFields", StringComparison.Ordinal));
    }

    [Fact]
    public static void ValidateReportsInvalidHeaderFooterReferences()
    {
        using MemoryStream stream = CreateDocxWithBody(
            """
                    <w:p><w:r><w:t>Body</w:t></w:r></w:p>
                    <w:sectPr xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                      <w:headerReference w:type="default" r:id="rMissingHeader"/>
                      <w:footerReference w:type="default" r:id="rWrongFooter"/>
                    </w:sectPr>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rWrongFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/footer" TargetMode="External"/>
                </Relationships>
                """, null);

        DocxValidateResult result = new DocxEditor().Validate(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9119" &&
            diagnostic.Message.Contains("rMissingHeader", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9119" &&
            diagnostic.Message.Contains("rWrongFooter", StringComparison.Ordinal) &&
            diagnostic.Message.Contains("expected footer relationship", StringComparison.Ordinal));
    }

    [Fact]
    public static void ValidateReportsInvalidSectionProperties()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p><w:r><w:t>Body</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:cols w:num="0"/>
                      <w:pgSz w:orient="sideways"/>
                    </w:sectPr>
            """);

        DocxValidateResult result = new DocxEditor().Validate(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9120" &&
            diagnostic.Message.Contains("columns", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9120" &&
            diagnostic.Message.Contains("orient", StringComparison.Ordinal));
    }

    [Fact]
    public static void ValidateReportsInvalidFootnoteAndEndnoteRoots()
    {
        using MemoryStream stream = CreateDocxWithBody(
            """
                    <w:p><w:r><w:t>Body</w:t></w:r></w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
                  <Relationship Id="rEndnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
                </Relationships>
                """,
            extra: archive =>
            {
                AddEntry(archive, "word/footnotes.xml", """
                    <w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
                    """);
                AddEntry(archive, "word/endnotes.xml", """
                    <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
                    """);
            });

        DocxValidateResult result = new DocxEditor().Validate(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9102" &&
            diagnostic.PartName == "/word/footnotes.xml" &&
            diagnostic.Message.Contains("footnotes", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9102" &&
            diagnostic.PartName == "/word/endnotes.xml" &&
            diagnostic.Message.Contains("endnotes", StringComparison.Ordinal));
    }

    [Fact]
    public static void ValidateReportsInvalidRelationshipPartRoot()
    {
        using MemoryStream stream = CreateDocxWithBody(
            """
                    <w:p><w:r><w:t>Body</w:t></w:r></w:p>
            """,
            null,
            extra: archive => AddEntry(archive, "word/_rels/header1.xml.rels", """
                <BrokenRelationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
                """));

        DocxValidateResult result = new DocxEditor().Validate(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9102" &&
            diagnostic.PartName == "/word/_rels/header1.xml.rels" &&
            diagnostic.Message.Contains("Relationships", StringComparison.Ordinal));
    }

    [Fact]
    public static void ValidateReportsWordprocessingInvariants()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                        xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                      <w:bookmarkStart w:id="7" w:name="OpenBookmark"/>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r>
                        <w:drawing>
                          <wp:inline>
                            <wp:docPr id="5" name="Picture 5"/>
                            <a:graphic>
                              <a:graphicData>
                                <pic:pic>
                                  <pic:blipFill>
                                    <a:blip r:embed="rMissing"/>
                                  </pic:blipFill>
                                </pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                      <w:r>
                        <w:drawing>
                          <wp:inline>
                            <wp:docPr id="5" name="Duplicate picture 5"/>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:tbl>
                      <w:tr/>
                    </w:tbl>
            """);
        var editor = new DocxEditor();

        DocxValidateResult result = editor.Validate(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9103" &&
            diagnostic.PartName == "/word/document.xml" &&
            diagnostic.Message.Contains("bookmark", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E9104");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E9105");
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9107" &&
            diagnostic.Message.Contains("docPr id '5'", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9106" &&
            diagnostic.Message.Contains("row", StringComparison.Ordinal));
    }

    [Fact]
    public static void ValidateCapsDiagnosticsWithTruncationDiagnostic()
    {
        string bodyXml = string.Concat(Enumerable.Range(1, 6).Select(id => $"""
                    <w:p>
                      <w:bookmarkStart w:id="{id}" w:name="Bookmark{id}"/>
                      <w:r><w:t>Text {id}</w:t></w:r>
                    </w:p>
            """));
        using MemoryStream stream = CreateDocxWithBody(bodyXml);

        DocxValidateResult result = new DocxEditor().Validate(stream, new DocxValidateOptions { MaxDiagnostics = 3 });

        Assert.False(result.Success);
        Assert.Equal(3, result.Diagnostics.Count);
        Assert.Equal(2, result.Diagnostics.Count(diagnostic => diagnostic.Code == "E9103"));
        DocxDiagnostic capped = result.Diagnostics.Last();
        Assert.Equal("E9199", capped.Code);
        Assert.Equal(DocxSeverity.Error, capped.Severity);
        Assert.Contains("omitted 4 of 6", capped.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void ValidateReportsInvalidDrawingGeometry()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p
                        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                      <w:r>
                        <w:drawing>
                          <wp:inline>
                            <wp:extent cx="0" cy="100"/>
                            <a:graphic>
                              <a:graphicData>
                                <a:srcRect l="60000" r="40000"/>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                    </w:p>
            """);
        var editor = new DocxEditor();

        DocxValidateResult result = editor.Validate(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E9109");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E9110");
    }

    [Fact]
    public static void ValidateReportsInvalidDrawingImageTargets()
    {
        using MemoryStream stream = CreateDocxWithImageRelationshipIssues();
        var editor = new DocxEditor();

        DocxValidateResult result = editor.Validate(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9113" &&
            diagnostic.Message.Contains("expected image relationship", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9113" &&
            diagnostic.Message.Contains("non-image content type 'text/plain'", StringComparison.Ordinal));
    }

    [Fact]
    public static void ValidateReportsInvalidTableVisualGrid()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/></w:tcPr>
                          <w:p><w:r><w:t>Root</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:vMerge/></w:tcPr>
                          <w:p><w:r><w:t>Mismatch</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc><w:p><w:r><w:t>Overflow</w:t></w:r></w:p></w:tc>
                        <w:tc>
                          <w:tcPr><w:gridSpan w:val="0"/></w:tcPr>
                          <w:p><w:r><w:t>Invalid span</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:vMerge/></w:tcPr>
                          <w:p><w:r><w:t>Orphan</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        var editor = new DocxEditor();

        DocxValidateResult result = editor.Validate(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9114" &&
            diagnostic.Message.Contains("does not match active restart span", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9114" &&
            diagnostic.Message.Contains("exceeding declared tblGrid", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9114" &&
            diagnostic.Message.Contains("invalid gridSpan", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9114" &&
            diagnostic.Message.Contains("has no active restart", StringComparison.Ordinal));
    }

    [Fact]
    public static void ValidateReportsInvalidFieldBoundariesAndFlags()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:instrText> ORPHAN </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:fldChar w:fldCharType="invalid"/></w:r>
                      <w:fldSimple w:instr=" DATE " w:dirty="maybe">
                        <w:r><w:t>Date</w:t></w:r>
                      </w:fldSimple>
                      <w:r><w:fldChar w:fldCharType="begin" w:fldLock="maybe"/></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:t>Cached result too early</w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                    </w:p>
            """);
        var editor = new DocxEditor();

        DocxValidateResult result = editor.Validate(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9112" &&
            diagnostic.Message.Contains("instruction text", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9112" &&
            diagnostic.Message.Contains("w:dirty", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9112" &&
            diagnostic.Message.Contains("w:fldLock", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9112" &&
            diagnostic.Message.Contains("result text appears before", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9104" &&
            diagnostic.Message.Contains("separate appears without", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9104" &&
            diagnostic.Message.Contains("invalid fldCharType", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9104" &&
            diagnostic.Message.Contains("duplicate separate", StringComparison.Ordinal));
    }

    [Fact]
    public static void ValidateReportsInvalidContentControlMetadata()
    {
        using MemoryStream stream = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:id w:val="abc"/>
                          <w:lock w:val="maybe"/>
                          <w:checkBox><w:checked w:val="maybe"/></w:checkBox>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>One</w:t></w:r></w:sdtContent>
                      </w:sdt>
                      <w:sdt>
                        <w:sdtPr><w:id w:val="42"/></w:sdtPr>
                        <w:sdtContent><w:r><w:t>Two</w:t></w:r></w:sdtContent>
                      </w:sdt>
                      <w:sdt>
                        <w:sdtPr><w:id w:val="42"/></w:sdtPr>
                        <w:sdtContent><w:r><w:t>Three</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        var editor = new DocxEditor();

        DocxValidateResult result = editor.Validate(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9115" &&
            diagnostic.Message.Contains("invalid integer", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9115" &&
            diagnostic.Message.Contains("w:lock", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9115" &&
            diagnostic.Message.Contains("w:checked", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9115" &&
            diagnostic.Message.Contains("Duplicate content control", StringComparison.Ordinal));
    }

    [Fact]
    public static void ValidateReportsInvalidCommentsExtendedMetadata()
    {
        using MemoryStream stream = CreateDocxWithBodyAndComments(
            """
                    <w:p><w:r><w:t>Body</w:t></w:r></w:p>
            """,
            """
                <w:comments
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w:comment w:id="1" w:author="Reviewer">
                    <w:p w14:paraId="00AAA111"><w:r><w:t>Comment</w:t></w:r></w:p>
                  </w:comment>
                  <w:comment w:id="2" w:author="Reviewer">
                    <w:p w14:paraId="00CCC333"><w:r><w:t>Reply</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
            """,
            """
                <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w15:commentEx w15:paraId="00AAA111" w15:done="0"/>
                  <w15:commentEx w15:paraId="00AAA111" w15:done="1"/>
                  <w15:commentEx w15:paraId="00BBB222" w15:done="0"/>
                  <w15:commentEx w15:paraId="00CCC333" w15:paraIdParent="00MISSING" w15:done="0"/>
                  <w15:commentEx w15:done="0"/>
                </w15:commentsEx>
            """, null);
        var editor = new DocxEditor();

        DocxValidateResult result = editor.Validate(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9108" &&
            diagnostic.PartName == "/word/commentsExtended.xml" &&
            diagnostic.Message.Contains("Duplicate commentsExtended paraId '00AAA111'", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9108" &&
            diagnostic.Message.Contains("00BBB222", StringComparison.Ordinal) &&
            diagnostic.Message.Contains("no matching comment paragraph", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9108" &&
            diagnostic.Message.Contains("parent paraId '00MISSING'", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9108" &&
            diagnostic.Message.Contains("missing w15:paraId", StringComparison.Ordinal));
    }

    [Fact]
    public static void ValidateReportsInvalidCommentsIdsMetadata()
    {
        using MemoryStream stream = CreateDocxWithBodyAndComments(
            """
                    <w:p><w:r><w:t>Body</w:t></w:r></w:p>
            """,
            """
                <w:comments
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w:comment w:id="1" w:author="Reviewer">
                    <w:p w14:paraId="00AAA111"><w:r><w:t>Comment</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
            """,
            commentsExtendedXml: null,
            commentsIdsXml: """
                <w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">
                  <w16cid:commentId w16cid:paraId="00AAA111" w16cid:durableId="D1"/>
                  <w16cid:commentId w16cid:paraId="00AAA111" w16cid:durableId="D2"/>
                  <w16cid:commentId w16cid:paraId="00BBB222" w16cid:durableId="D1"/>
                  <w16cid:commentId w16cid:paraId="00CCC333"/>
                  <w16cid:commentId w16cid:durableId="D3"/>
                </w16cid:commentsIds>
            """);
        var editor = new DocxEditor();

        DocxValidateResult result = editor.Validate(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9122" &&
            diagnostic.PartName == "/word/commentsIds.xml" &&
            diagnostic.Message.Contains("Duplicate commentsIds paraId '00AAA111'", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9122" &&
            diagnostic.Message.Contains("Duplicate commentsIds durableId 'D1'", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9122" &&
            diagnostic.Message.Contains("00BBB222", StringComparison.Ordinal) &&
            diagnostic.Message.Contains("no matching comment paragraph", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9122" &&
            diagnostic.Message.Contains("missing w16cid:paraId", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9122" &&
            diagnostic.Message.Contains("missing w16cid:durableId", StringComparison.Ordinal));
    }

    [Fact]
    public static void ValidateReportsInvalidCommentBodyAndAnchorConsistency()
    {
        using MemoryStream stream = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
                    <w:p>
                      <w:commentRangeStart w:id="4"/>
                      <w:r><w:t>Missing body</w:t></w:r>
                      <w:commentRangeEnd w:id="4"/>
                      <w:r><w:commentReference w:id="4"/></w:r>
                    </w:p>
                    <w:p>
                      <w:commentRangeStart w:id="6"/>
                      <w:r><w:t>No reference</w:t></w:r>
                      <w:commentRangeEnd w:id="6"/>
                    </w:p>
                    <w:p>
                      <w:r><w:commentReference/></w:r>
                    </w:p>
            """,
            """
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="3" w:author="Reviewer">
                    <w:p><w:r><w:t>First</w:t></w:r></w:p>
                  </w:comment>
                  <w:comment w:id="3" w:author="Reviewer">
                    <w:p><w:r><w:t>Duplicate</w:t></w:r></w:p>
                  </w:comment>
                  <w:comment w:id="6" w:author="Reviewer">
                    <w:p><w:r><w:t>No reference body</w:t></w:r></w:p>
                  </w:comment>
                  <w:comment w:author="Reviewer">
                    <w:p><w:r><w:t>Missing ID</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
            """);
        var editor = new DocxEditor();

        DocxValidateResult result = editor.Validate(stream);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9111" &&
            diagnostic.Message.Contains("Duplicate comment body id '3'", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9111" &&
            diagnostic.Message.Contains("Comment markup id '4' has no matching comment body", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9111" &&
            diagnostic.Message.Contains("Comment range id '6' has no matching commentReference", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9111" &&
            diagnostic.Message.Contains("Comment body is missing w:id", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E9111" &&
            diagnostic.Message.Contains("commentReference is missing w:id", StringComparison.Ordinal));
    }

    private static MemoryStream CreateValidationFixture(string fixtureName)
    {
        return fixtureName switch
        {
            "basic-document" => CreateDocx(),
            "unclosed-bookmark" => CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="1" w:name="Unclosed"/>
                      <w:r><w:t>Text</w:t></w:r>
                    </w:p>
                """),
            "invalid-relationship-root" => CreateDocxWithBody(
                """
                    <w:p><w:r><w:t>Body</w:t></w:r></w:p>
                """,
                null,
                extra: archive => AddEntry(archive, "word/_rels/header1.xml.rels", """
                    <BrokenRelationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
                    """)),
            "missing-style-definition" => CreateDocxWithStylesAndNumbering(
                """
                    <w:p>
                      <w:pPr><w:pStyle w:val="MissingStyle"/></w:pPr>
                      <w:r><w:t>Styled text</w:t></w:r>
                    </w:p>
                """,
                """
                    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                      <w:style w:type="paragraph" w:styleId="KnownStyle"/>
                    </w:styles>
                    """,
                """
                    <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
                    """),
            _ => throw new ArgumentOutOfRangeException(nameof(fixtureName), fixtureName, "Unknown validation fixture.")
        };
    }

    private static MemoryStream CreateDocxWithImageRelationshipIssues()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="txt" ContentType="text/plain"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """);
            AddEntry(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rWrongType" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/image.png" TargetMode="External"/>
                  <Relationship Id="rTextPart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/not-image.txt"/>
                </Relationships>
                """);
            AddEntry(archive, "word/document.xml", """
                <w:document
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:drawing><a:blip r:embed="rWrongType"/></w:drawing></w:r>
                      <w:r><w:drawing><a:blip r:embed="rTextPart"/></w:drawing></w:r>
                    </w:p>
                  </w:body>
                </w:document>
                """);
            AddEntry(archive, "word/media/not-image.txt", "not image data");
        }

        stream.Position = 0;
        return stream;
    }
}
