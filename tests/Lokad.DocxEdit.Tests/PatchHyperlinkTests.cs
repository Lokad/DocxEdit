using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml.Linq;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

public static class PatchHyperlinkTests
{

    [Fact]
    public static void ApplyCanSetHyperlinkTargetAndText()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(
            """
                    <w:p>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rLink">
                        <w:r><w:t>Old link</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/old" TargetMode="External"/>
                </Relationships>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-hyperlink-target
            target M.L0001
            uri https://example.test/new
            tooltip Updated link
            target-frame _blank
            history false
            end

            op set-hyperlink-text
            target M.L0001
            text New link
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal("New link", Assert.Single(read.Paragraphs).Text);
        DocxHyperlinkInfo hyperlink = Assert.Single(read.Hyperlinks);
        Assert.Equal("https://example.test/new", hyperlink.Uri);
        Assert.Equal("Updated link", hyperlink.Tooltip);
        Assert.Equal("_blank", hyperlink.TargetFrame);
        Assert.False(hyperlink.History);
        Assert.Equal(8, hyperlink.DisplayTextLength);
        output.Position = 0;
        string relationships = ReadEntry(output, "word/_rels/document.xml.rels");
        Assert.Contains("Target=\"https://example.test/new\"", relationships, StringComparison.Ordinal);
        Assert.Contains("TargetMode=\"External\"", relationships, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetHyperlinkAnchorPreservesSharedRelationship()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(
            """
                    <w:p>
                      <w:bookmarkStart w:id="1" w:name="Bookmark"/>
                      <w:r><w:t>Anchor</w:t></w:r>
                      <w:bookmarkEnd w:id="1"/>
                    </w:p>
                    <w:p>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rShared">
                        <w:r><w:t>First</w:t></w:r>
                      </w:hyperlink>
                      <w:r><w:t xml:space="preserve"> </w:t></w:r>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rShared">
                        <w:r><w:t>Second</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/shared" TargetMode="External"/>
                </Relationships>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-hyperlink-target
            target M.L0001
            anchor Bookmark
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxHyperlinkInfo[] hyperlinks = new DocxEditor().Read(output).Hyperlinks.ToArray();
        Assert.Equal("Bookmark", hyperlinks[0].Anchor);
        Assert.Equal("https://example.test/shared", hyperlinks[1].Uri);
        output.Position = 0;
        string relationships = ReadEntry(output, "word/_rels/document.xml.rels");
        Assert.Contains("Id=\"rShared\"", relationships, StringComparison.Ordinal);
        Assert.Contains("Target=\"https://example.test/shared\"", relationships, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetHyperlinkAnchorRemovesUnusedRelationship()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(
            """
                    <w:p>
                      <w:bookmarkStart w:id="1" w:name="Bookmark"/>
                      <w:r><w:t>Anchor</w:t></w:r>
                      <w:bookmarkEnd w:id="1"/>
                    </w:p>
                    <w:p>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rLink">
                        <w:r><w:t>Link</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/old" TargetMode="External"/>
                </Relationships>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-hyperlink-target
            target M.L0001
            anchor Bookmark
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxHyperlinkInfo hyperlink = Assert.Single(new DocxEditor().Read(output).Hyperlinks);
        Assert.Equal("Bookmark", hyperlink.Anchor);
        Assert.Null(hyperlink.RelationshipId);
        output.Position = 0;
        string relationships = ReadEntry(output, "word/_rels/document.xml.rels");
        Assert.DoesNotContain("rLink", relationships, StringComparison.Ordinal);
        Assert.DoesNotContain("https://example.test/old", relationships, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyRemoveHyperlinkPreservesRelationshipStillInUse()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(
            """
                    <w:p>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rShared">
                        <w:r><w:t>First</w:t></w:r>
                      </w:hyperlink>
                      <w:r><w:t xml:space="preserve"> </w:t></w:r>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rShared">
                        <w:r><w:t>Second</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/shared" TargetMode="External"/>
                </Relationships>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op remove-hyperlink
            target M.L0001
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal("First Second", Assert.Single(read.Paragraphs).Text);
        Assert.Single(read.Hyperlinks);
        Assert.Equal("https://example.test/shared", read.Hyperlinks[0].Uri);
        output.Position = 0;
        string relationships = ReadEntry(output, "word/_rels/document.xml.rels");
        Assert.Contains("Id=\"rShared\"", relationships, StringComparison.Ordinal);
        Assert.Contains("Target=\"https://example.test/shared\"", relationships, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyRemoveHyperlinkGuardMismatchWritesNothing()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(
            """
                    <w:p>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rLink">
                        <w:r><w:t>Link</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/old" TargetMode="External"/>
                </Relationships>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op remove-hyperlink
            target M.L0001
            expect-text Stale link
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static d => d.Code == "E3201");
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public static void CheckRemoveHyperlinkGuardMismatchFails()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(
            """
                    <w:p>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rLink">
                        <w:r><w:t>Link</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/old" TargetMode="External"/>
                </Relationships>
                """);
        using var patch = new StringReader("""
            docxpatch 1

            op remove-hyperlink
            target M.L0001
            expect-text Stale link
            end
            """);

        DocxCheckResult check = new DocxEditor().Check(input, patch);
        Assert.False(check.Success);
        DocxDiagnostic failure = Assert.Single(check.Diagnostics, static d => d.Code == "E3201");
        Assert.Equal("remove-hyperlink", failure.HelpTopic);
        Assert.Equal(5, failure.Line);
    }

    [Fact]
    public static void ApplyRemoveHyperlinkGuardMatchSucceeds()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(
            """
                    <w:p>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rLink">
                        <w:r><w:t>Link</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/old" TargetMode="External"/>
                </Relationships>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op remove-hyperlink
            target M.L0001
            expect-text Link
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("Link", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        Assert.Empty(new DocxEditor().Read(output).Hyperlinks);
        output.Position = 0;
        string relationships = ReadEntry(output, "word/_rels/document.xml.rels");
        Assert.DoesNotContain("rLink", relationships, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("../relative/report", "relative hyperlink targets are not supported")]
    [InlineData("file:///C:/secret/report.docx", "Unsupported hyperlink URI scheme 'file'")]
    [InlineData(@"\\server\share\report.docx", "Unsupported hyperlink URI scheme 'file'")]
    [InlineData(@"C:\secret\report.docx", "Unsupported hyperlink URI scheme 'file'")]
    [InlineData("javascript:alert(1)", "Unsupported hyperlink URI scheme 'javascript'")]
    [InlineData("http://[::1", "malformed URI")]
    public static void CheckHyperlinkTargetRejectsUnsupportedUris(string uri, string expectedMessage)
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(
            """
                    <w:p>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rLink">
                        <w:r><w:t>Link</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/old" TargetMode="External"/>
                </Relationships>
                """);
        using var patch = new StringReader($"""
            docxpatch 1

            op set-hyperlink-target
            target M.L0001
            uri {uri}
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E4205" &&
            diagnostic.Message.Contains(expectedMessage, StringComparison.Ordinal));
    }

    [Fact]
    public static void CheckHyperlinkTargetAcceptsMailtoUris()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(
            """
                    <w:p>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rLink">
                        <w:r><w:t>Mail</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/old" TargetMode="External"/>
                </Relationships>
                """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-hyperlink-target
            target M.L0001
            uri mailto:reviewer@example.test
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success);
    }

    [Fact]
    public static void ApplyCanInsertAndRemoveHyperlinkWhilePreservingText()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>Anchor</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-hyperlink-after
            target M.P0001
            text Docs
            uri https://docs.example/
            end

            op remove-hyperlink
            target M.L0001
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal(new[] { "Anchor", "Docs" }, read.Paragraphs.Select(paragraph => paragraph.Text));
        Assert.Empty(read.Hyperlinks);
        output.Position = 0;
        string relationships = ReadEntry(output, "word/_rels/document.xml.rels");
        Assert.DoesNotContain("https://docs.example/", relationships, StringComparison.Ordinal);
        Assert.DoesNotContain("/hyperlink", relationships, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetHyperlinkTextWithMatchingExpectTextSucceeds()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(
            """
                    <w:p>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rLink">
                        <w:r><w:t>Old link</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/old" TargetMode="External"/>
                </Relationships>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-hyperlink-text
            target M.L0001
            expect-text Old link
            text New link
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Equal("New link", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void CheckSetHyperlinkTextWithMismatchedExpectTextFails()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(
            """
                    <w:p>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rLink">
                        <w:r><w:t>Old link</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/old" TargetMode="External"/>
                </Relationships>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-hyperlink-text
            target M.L0001
            expect-text Stale link
            text New link
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E3201");
    }

    [Fact]
    public static void ApplySetHyperlinkTextIdenticalTextIsNoOp()
    {
        using MemoryStream input = CreateDocxWithBodyAndRelationships(
            """
                    <w:p>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rLink">
                        <w:r><w:t>Old link</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/old" TargetMode="External"/>
                </Relationships>
            """);
        using MemoryStream probe = CreateDocxWithBodyAndRelationships(
            """
                    <w:p>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rLink">
                        <w:r><w:t>Old link</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """,
            """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/old" TargetMode="External"/>
                </Relationships>
            """);
        string before = ReadDocumentXml(probe);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-hyperlink-text
            target M.L0001
            text Old link
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Contains(result.Diagnostics, static d => d.Code == "I0001");
        output.Position = 0;
        Assert.Equal(before, ReadDocumentXml(output));
    }

}
