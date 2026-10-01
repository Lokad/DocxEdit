using System.Text;
using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// D13: named result bindings let later operations address created objects.
[Collection("ConsoleCli")]
public static class PatchAliasTests
{
    [Fact]
    public static void InsertBindsAliasForLaterEdit()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text Drafted
            as sec1
            end

            op replace-text
            target @sec1
            find Drafted
            with Final
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Equal(new[] { "Alpha", "Final" }, new DocxEditor().Read(output).Paragraphs.Select(static p => p.Text).ToArray());
        Assert.Equal("M.P0002", Assert.Single(result.Operations[1].AffectedTargets).Id.ToWireValue());
    }
    [Fact]
    public static void InsertThenEditTrackedInsertionRefusesInSamePatch()
    {
        var options = new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            Author = "Agent",
            TimestampUtc = new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero)
        };
        using MemoryStream input = CreateDocx("Alpha");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text Inserted
            as sec1
            end

            op replace-text
            target @sec1
            find Inserted
            with Edited
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, static d => d.Code == "E4305");
        Assert.Contains("tracked-insertion", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void InsertThenEditTrackedInsertionRefusesUnderRequire()
    {
        var options = new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Require,
            Author = "Agent",
            TimestampUtc = new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero)
        };
        using MemoryStream input = CreateDocx("Alpha");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text Inserted
            as sec1
            end

            op replace-text
            target @sec1
            find Inserted
            with Edited
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, options);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, static d => d.Code == "E6002");
        Assert.Contains("tracked-insertion", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void CommentTargetsInsertedParagraphViaAlias()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text Drafted
            as sec1
            end

            op add-comment
            target @sec1
            text Review drafted
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Equal(new[] { "Alpha", "Drafted" }, new DocxEditor().Read(output).Paragraphs.Select(static p => p.Text).ToArray());
        output.Position = 0;
        DocxCommentThreadSummary thread = Assert.Single(new DocxEditor().Changes(output).CommentSummary);
        Assert.Equal("M.P0002", thread.AnchorTargetId);
    }


    [Fact]
    public static void CommentBindsAliasForReply()
    {
        using MemoryStream input = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="3" w:author="Reviewer" w:initials="RV" w:date="2026-06-07T12:00:00Z">
                    <w:p><w:r><w:t>Old comment</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op add-comment
            target M.P0001
            text Fresh note
            as note1
            end

            op add-comment-reply
            target @note1
            text Thanks
            author Reviewer
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Equal(new[] { "comment:4" }, result.Operations[0].CreatedTargetIds);
        Assert.Equal(new[] { "comment:5" }, result.Operations[1].CreatedTargetIds);
    }

    [Fact]
    public static void BookmarkBindsAliasForReplace()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op add-bookmark
            target M.P0001
            name Mark
            as bm1
            end

            op replace-bookmark-text
            target @bm1
            text Seeded
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Equal("Seeded", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void DuplicateAliasFails()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text First
            as twice
            end

            op insert-after
            target M.P0001
            text Second
            as twice
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.True(result.Operations[0].Success);
        Assert.False(result.Operations[1].Success);
        Assert.Contains(result.Diagnostics, static d => d.Code == "E4205");
    }

    [Fact]
    public static void UnknownAliasFails()
    {
        using MemoryStream input = CreateDocx("Alpha Beta");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target @ghost
            find Alpha
            with Omega
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, static d => d.Code == "E1201");
        Assert.Equal("@ghost", diagnostic.TargetId);
    }

    [Fact]
    public static void ForwardReferenceFails()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target @later
            find Alpha
            with Omega
            end

            op insert-after
            target M.P0001
            text Inserted
            as later
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static d => d.Code == "E1201");
    }

    [Fact]
    public static void AliasWithMultipleTextsFails()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text First
            text Second
            as pair
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static d => d.Code == "E4205");
    }

    [Theory]
    [InlineData("1bad")]
    [InlineData("has space")]
    [InlineData("")]
    public static void InvalidAliasNameFails(string name)
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var patch = new StringReader("docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext Fresh\nas " + name + "\nend\n");

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static d => d.Code == "E4205");
    }

    [Fact]
    public static void LintRejectsUnknownAlias()
    {
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target @ghost
            find Alpha
            with Omega
            end
            """);

        DocxLintResult lint = new DocxEditor().Lint(patch);

        Assert.False(lint.Success);
        DocxDiagnostic diagnostic = Assert.Single(lint.Diagnostics, static d => d.Code == "E1201");
        Assert.Equal("replace-text", diagnostic.HelpTopic);
    }

    [Fact]
    public static void LintRejectsDuplicateAlias()
    {
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text First
            as twice
            end

            op insert-after
            target M.P0001
            text Second
            as twice
            end
            """);

        DocxLintResult lint = new DocxEditor().Lint(patch);

        Assert.False(lint.Success);
        Assert.Contains(lint.Diagnostics, static d => d.Code == "E4205");
    }

    [Fact]
    public static void CheckAndApplyAgreeOnAliases()
    {
        const string patchText = "docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext Drafted\nas sec1\nend\n\nop replace-text\ntarget @sec1\nfind Drafted\nwith Final\nend\n";
        using MemoryStream checkInput = CreateDocx("Alpha");
        DocxCheckResult check = new DocxEditor().Check(checkInput, new StringReader(patchText));
        Assert.True(check.Success, string.Join("|", check.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        using MemoryStream applyInput = CreateDocx("Alpha");
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, new StringReader(patchText), output);
        Assert.True(apply.Success, string.Join("|", apply.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Equal(new[] { "Alpha", "Final" }, new DocxEditor().Read(output).Paragraphs.Select(static p => p.Text).ToArray());
    }

    [Fact]
    public static void DeletedBindingFails()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text Temporary
            as tmp
            end

            op delete-block
            target @tmp
            end

            op replace-text
            target @tmp
            find Temporary
            with Permanent
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.False(result.Success);
        Assert.True(result.Operations[0].Success);
        Assert.True(result.Operations[1].Success);
        Assert.False(result.Operations[2].Success);
        Assert.Contains(result.Operations[2].Diagnostics, static d => d.Code == "E1201");
    }
    [Fact]
    public static void SetStyleAfterInsertViaAlias()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody(
            """
              <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
              <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/></w:style>
            """,
            """
                    <w:p><w:r><w:t>Anchor</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text New section
            as sec1
            end

            op set-style
            target @sec1
            style Heading 2
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Equal(1, CountOccurrences(ReadDocumentXml(output), "w:pStyle w:val=\"Heading2\""));
    }

    [Fact]
    public static void InsertAfterAliasTarget()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text First
            as a1
            end

            op insert-after
            target @a1
            text Second
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        Assert.Equal(new[] { "Alpha", "First", "Second" }, new DocxEditor().Read(output).Paragraphs.Select(static p => p.Text).ToArray());
    }
    [Fact]
    public static void AppendRowBindsAliasForDelete()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr><w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc></w:tr>
                      <w:tr><w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op append-row
            target M.T0001
            cell East
            as rnew
            end

            op delete-row
            target @rnew
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal(2, table.RowCount);
        Assert.Contains(table.Cells, static cell => cell.Text == "North");
        Assert.Contains(table.Cells, static cell => cell.Text == "South");
        Assert.DoesNotContain(table.Cells, static cell => cell.Text == "East");
    }

    [Fact]
    public static void InsertRowBindsAliasForHeader()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-row-after
            target M.T0001.R01
            cell A2
            cell B2
            as rnew
            end

            op set-row-header
            target @rnew
            header true
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        output.Position = 0;
        Assert.Contains("<w:tblHeader", ReadDocumentXml(output), StringComparison.Ordinal);
    }

    [Fact]
    public static void ParagraphAliasAsRowTargetFails()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text Drafted
            as sec1
            end

            op delete-row
            target @sec1
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.False(result.Success);
        Assert.True(result.Operations[0].Success);
        Assert.False(result.Operations[1].Success);
        Assert.Contains(result.Operations[1].Diagnostics, static d => d.Code == "E1201");
    }
    [Fact]
    public static void HyperlinkInsertBindsAliasForLaterEdit()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-hyperlink-after
            target M.P0001
            text Click here
            uri https://example.test/new
            as hl1
            end

            op insert-after
            target @hl1
            text After link
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.Equal(new[] { "M.P0002" }, result.Operations[0].CreatedTargetIds);
        output.Position = 0;
        Assert.Equal(new[] { "Alpha", "Click here", "After link" }, new DocxEditor().Read(output).Paragraphs.Select(static p => p.Text).ToArray());
    }
    [Fact]
    public static void InsertRowAfterAliasTarget()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1
            op append-row
            target M.T0001
            cell C1
            cell C2
            as r1
            end
            op insert-row-after
            target @r1
            cell D1
            cell D2
            end
            """);
        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal(3, table.RowCount);
        Assert.Equal(new[] { "A1", "B1", "C1", "C2", "D1", "D2" }, table.Cells.Select(static cell => cell.Text).ToArray());
    }

    [Fact]
    public static void HyperlinkTextViaAlias()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1
            op insert-hyperlink-after
            target M.P0001
            text Click here
            uri https://example.test/new
            as hl1
            end
            op set-hyperlink-text
            target @hl1
            expect-text Click here
            text Follow this link
            end
            """);
        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        output.Position = 0;
        Assert.Equal(new[] { "Alpha", "Follow this link" }, new DocxEditor().Read(output).Paragraphs.Select(static paragraph => paragraph.Text).ToArray());
    }
    [Fact]
    public static void RemovedHyperlinkAliasRefuses()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1
            op insert-hyperlink-after
            target M.P0001
            text Click here
            uri https://example.test/new
            as hl1
            end
            op remove-hyperlink
            target @hl1
            end
            op set-hyperlink-text
            target @hl1
            text Gone
            end
            """);
        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
        Assert.False(result.Success);
        Assert.Contains(result.Operations[2].Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }

    [Fact]
    public static void ImageAltViaAlias()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", CreatePngBytes(4, 3), null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1
            op insert-image-after
            target M.P0001
            asset chart.png
            alt Old image
            as im1
            end
            op set-image-alt
            target @im1
            alt New image
            end
            """);
        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("descr=\"New image\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("descr=\"Old image\"", xml, StringComparison.Ordinal);
    }
    [Fact]
    public static void DeletedImageAliasRefuses()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", CreatePngBytes(4, 3), null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1
            op insert-image-after
            target M.P0001
            asset chart.png
            as im1
            end
            op delete-image
            target @im1
            end
            op set-image-alt
            target @im1
            alt Gone
            end
            """);
        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });
        Assert.False(result.Success);
        Assert.Contains(result.Operations[2].Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }

    [Fact]
    public static void CommentSetAndResolveViaAlias()
    {
        using MemoryStream input = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="3" w:author="Reviewer" w:initials="RV" w:date="2026-06-07T12:00:00Z">
                    <w:p><w:r><w:t>Old comment</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1
            op add-comment
            target M.P0001
            text Fresh note
            as note1
            end
            op set-comment-text
            target @note1
            text Edited note
            end
            op resolve-comment
            target @note1
            end
            """);
        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        Assert.Equal(new[] { "comment:4" }, result.Operations[0].CreatedTargetIds);
        output.Position = 0;
        string xml = Encoding.UTF8.GetString(ReadEntryBytes(output, "word/comments.xml"));
        Assert.Contains("Edited note", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("Fresh note", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void BookmarkRenameAndDeleteViaAlias()
    {
        using MemoryStream input = CreateDocx("Alpha");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1
            op add-bookmark
            target M.P0001
            name Mark
            as bm1
            end
            op rename-bookmark
            target @bm1
            name Renamed
            end
            op delete-bookmark
            target @bm1
            end
            """);
        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        output.Position = 0;
        Assert.Equal(new[] { "Alpha" }, new DocxEditor().Read(output).Paragraphs.Select(static paragraph => paragraph.Text).ToArray());
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.DoesNotContain("bookmarkStart", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("bookmarkEnd", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void RowAliasAsCellTargetFails()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1
            op append-row
            target M.T0001
            cell C1
            cell C2
            as r1
            end
            op set-cell
            target @r1
            text X
            end
            """);
        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
        Assert.False(result.Success);
        Assert.True(result.Operations[0].Success);
        Assert.False(result.Operations[1].Success);
        Assert.Contains(result.Operations[1].Diagnostics, static diagnostic => diagnostic.Code == "E1201");
    }
    [Fact]
    public static void AliasComposeExplicitBindsAndEdits()
    {
        const string body = "<w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p>";
        using MemoryStream input = CreateDocxWithBody(body);
        string patchText = "docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext Inserted\nas added\nend\n\nop replace-paragraph\ntarget @added\ntext CHANGED\nend\n";
        DocxCheckResult check = new DocxEditor().Check(input, new StringReader(patchText));
        Assert.True(check.Success);
        Assert.Equal("M.P0002", Assert.Single(check.Operations[0].CreatedTargetIds));
        using MemoryStream applyInput = CreateDocxWithBody(body);
        using var output = new MemoryStream();
        using var applyPatch = new StringReader(patchText);
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, applyPatch, output);
        Assert.True(apply.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal("First", read.Paragraphs[0].Text);
        Assert.Equal("CHANGED", read.Paragraphs[1].Text);
        Assert.Equal("Second", read.Paragraphs[2].Text);
    }

    [Fact]
    public static void AliasCompositionBindsAndEditsInHeader()
    {
        using MemoryStream input = CreateDocxWithHeaderFooter("Head", "Foot");
        string patchText = "docxpatch 1\n\nop insert-after\ntarget H001.P0001\ntext Inserted\nas added\nend\n\nop replace-paragraph\ntarget @added\ntext CHANGED\nend\n";
        DocxCheckResult check = new DocxEditor().Check(input, new StringReader(patchText));
        Assert.True(check.Success, string.Join("|", check.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        Assert.Equal("H001.P0002", Assert.Single(check.Operations[0].CreatedTargetIds));
        using MemoryStream applyInput = CreateDocxWithHeaderFooter("Head", "Foot");
        using var output = new MemoryStream();
        using var applyPatch = new StringReader(patchText);
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, applyPatch, output);
        Assert.True(apply.Success, string.Join("|", apply.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output, new DocxReadOptions { IncludeHeadersFooters = true });
        Assert.Equal("Main text", read.Paragraphs[0].Text);
        Assert.Equal("Head", read.Paragraphs[1].Text);
        Assert.Equal("CHANGED", read.Paragraphs[2].Text);
        Assert.Equal("H001.P0002", read.Paragraphs[2].Id.ToWireValue());
        Assert.Equal("Foot", read.Paragraphs[3].Text);
    }

    [Fact]
    public static void AliasRebaseReportsFinalId()
    {
        const string body = "<w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p>";
        using MemoryStream input = CreateDocxWithBody(body);
        string patchText = "docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext Inserted\nas added\nend\n\nop replace-paragraph\ntarget @added\ntext CHANGED\nend\n\nop insert-before\ntarget M.P0001\ntext Earlier\nend\n";
        DocxCheckResult check = new DocxEditor().Check(input, new StringReader(patchText));
        Assert.True(check.Success);
        DocxPatchAffectedTarget affected = Assert.Single(check.Operations[1].AffectedTargets);
        Assert.Equal("M.P0003", affected.Id.ToWireValue());
        Assert.Equal("operation-time", affected.Coordinate);
        Assert.Equal("M.P0003", affected.FinalId?.ToWireValue());
        using MemoryStream applyInput = CreateDocxWithBody(body);
        using var output = new MemoryStream();
        using var applyPatch = new StringReader(patchText);
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, applyPatch, output);
        Assert.True(apply.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal(4, read.Paragraphs.Count);
        Assert.Equal("Earlier", read.Paragraphs[0].Text);
        Assert.Equal("First", read.Paragraphs[1].Text);
        Assert.Equal("CHANGED", read.Paragraphs[2].Text);
        Assert.Equal("Second", read.Paragraphs[3].Text);
    }
    [Fact]
    public static void AliasPreviewBindsAtExecution()
    {
        const string body = "<w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p>";
        using MemoryStream input = CreateDocxWithBody(body);
        var options = new DocxEditOptions { MaxPreviewChars = 100 };
        string patchText = "docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext Inserted\nas added\nend\n\nop replace-paragraph\ntarget @added\ntext CHANGED\nend\n";
        DocxCheckResult check = new DocxEditor().Check(input, new StringReader(patchText), options);
        Assert.True(check.Success);
        Assert.Equal("Inserted", check.Operations[1].PreviewBefore);
        Assert.Equal("CHANGED", check.Operations[1].PreviewAfter);
        Assert.Equal("M.P0002", Assert.Single(check.Operations[0].CreatedTargetIds));
        using MemoryStream applyInput = CreateDocxWithBody(body);
        using var output = new MemoryStream();
        using var applyPatch = new StringReader(patchText);
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, applyPatch, output);
        Assert.True(apply.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal(3, read.Paragraphs.Count);
        Assert.Equal("First", read.Paragraphs[0].Text);
        Assert.Equal("CHANGED", read.Paragraphs[1].Text);
        Assert.Equal("Second", read.Paragraphs[2].Text);
    }




}
