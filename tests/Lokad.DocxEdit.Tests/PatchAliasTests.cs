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
}
