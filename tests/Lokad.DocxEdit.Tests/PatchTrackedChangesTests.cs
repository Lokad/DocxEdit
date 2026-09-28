using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml.Linq;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

public static class PatchTrackedChangesTests
{

    [Fact]
    public static void CheckAndApplyEchoEffectiveProvenance()
    {
        var options = new DocxEditOptions
        {
            Author = "Agent",
            TimestampUtc = new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero)
        };
        string expectedVersion = typeof(DocxEditor).Assembly.GetName().Version?.ToString() ?? "unknown";

        using MemoryStream checkInput = CreateDocx("Revenue increased by 8.4%.");
        using var checkPatch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind 8.4%\nwith 9.1%\nend\n");
        DocxCheckResult check = new DocxEditor().Check(checkInput, checkPatch, options);

        Assert.True(check.Success);
        Assert.Equal("Agent", check.Author);
        Assert.Equal(options.TimestampUtc, check.TimestampUtc);
        Assert.Equal(expectedVersion, check.ToolVersion);

        using MemoryStream applyInput = CreateDocx("Revenue increased by 8.4%.");
        using var applyPatch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind 8.4%\nwith 9.1%\nend\n");
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, applyPatch, output, options);

        Assert.True(apply.Success);
        Assert.Equal("Agent", apply.Author);
        Assert.Equal(options.TimestampUtc, apply.TimestampUtc);
        Assert.Equal(expectedVersion, apply.ToolVersion);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestWarnsWhenFieldsRequireWordSideRefresh()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>Revenue increased.</w:t></w:r></w:p>
                    <w:p>
                      <w:fldSimple w:instr=" DATE ">
                        <w:r><w:t>June 12</w:t></w:r>
                      </w:fldSimple>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find increased
            with rose
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "W5103" &&
            diagnostic.Feature == "field" &&
            diagnostic.Fallback == "word-refresh-required");
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:del", xml, StringComparison.Ordinal);
        Assert.Contains("<w:ins", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains("updateFields", ReadEntry(output, "word/settings.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckTrackChangesRequireAllowsSetSimpleFieldResult()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:fldSimple w:instr=" REF ClientName \h ">
                        <w:r><w:t>Old cached result</w:t></w:r>
                      </w:fldSimple>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-field-result
            target M.F0001
            expect-result Old cached result
            text New cached result
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
    }

    [Fact]
    public static void ApplyTrackChangesSuggestGeneratesRevisionMarkupForSetSimpleFieldResult()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:fldSimple w:instr=" REF ClientName \h ">
                        <w:r><w:t>Old cached result</w:t></w:r>
                      </w:fldSimple>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-field-result
            target M.F0001
            expect-result Old cached result
            text New cached result
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code is "W4001" or "W4002");
        Assert.Equal(["1", "2"], Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:fldSimple w:instr=\" REF ClientName \\h \">", xml, StringComparison.Ordinal);
        Assert.Contains("<w:delText>Old cached result</w:delText>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>New cached result</w:t>", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Equal("New cached result", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        output.Position = 0;
        Assert.Equal("Old cached result", Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { TextView = DocxTextView.Original }).Paragraphs).Text);
        output.Position = 0;
        Assert.Equal("[-Old cached result-][+New cached result+]", Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { TextView = DocxTextView.Markup }).Paragraphs).Text);
        output.Position = 0;
        DocxFieldInfo field = Assert.Single(new DocxEditor().Read(output).Fields);
        Assert.Equal("REF ClientName \\h", field.Code);
        Assert.Equal("New cached result", field.CachedResultText);
        output.Position = 0;
        Assert.False(EntryExists(output, "word/settings.xml"));
    }

    [Fact]
    public static void CheckTrackChangesSuggestSetFieldResultFallsBackForComplexField()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> PAGE </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:t>1</w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-field-result
            target M.F0001
            text 2
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        DocxDiagnostic warning = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "W4002");
        Assert.Equal("M.F0001", warning.TargetId);
        Assert.Contains("tracked complex-field result replacement is not modeled", warning.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckTrackChangesRequireRejectsComplexFieldResultReplacement()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> PAGE </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:t>1</w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-field-result
            target M.F0001
            text 2
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
        Assert.Equal("M.F0001", diagnostic.TargetId);
        Assert.Contains("tracked complex-field result replacement is not modeled", diagnostic.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("set-field-dirty", "target M.F0001\ndirty true")]
    [InlineData("set-field-lock", "target M.F0001\nlocked true")]
    [InlineData("set-field-code", "target M.F0001\ncode REF OtherBookmark \\h")]
    [InlineData("refresh-field-result", "target M.F0001")]
    public static void CheckTrackChangesRequireRejectsFieldOperationsAsPreserveOnly(string operationName, string operationFields)
    {
        using MemoryStream input = CreateDocxWithRefField();
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            {operationFields}
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6001");
        Assert.Contains($"operation '{operationName}'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("catalog support is 'preserve-only'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-no-revision-representation", diagnostic.Feature);
        Assert.Equal("require-failed", diagnostic.Fallback);
    }

    [Theory]
    [InlineData("set-field-dirty", "target M.F0001\ndirty true")]
    [InlineData("set-field-lock", "target M.F0001\nlocked true")]
    [InlineData("set-field-code", "target M.F0001\ncode REF OtherBookmark \\h")]
    [InlineData("refresh-field-result", "target M.F0001")]
    public static void ApplyTrackChangesSuggestWarnsForFieldOperationsAsPreserveOnly(string operationName, string operationFields)
    {
        using MemoryStream input = CreateDocxWithRefField();
        using var output = new MemoryStream();
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            {operationFields}
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(
            input,
            patch,
            output,
            new DocxEditOptions
            {
                TrackChanges = TrackChangesMode.Suggest,
                MarkFieldsDirtyWhenEditing = false
            });

        Assert.True(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "W4001");
        Assert.Contains($"operation '{operationName}'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("catalog support is 'preserve-only'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-no-revision-representation", diagnostic.Feature);
        Assert.Equal("direct-edit-preserve-existing-revisions", diagnostic.Fallback);
        Assert.Empty(Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:fldSimple", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:ins", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:del", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckTrackChangesRequireFailsUnsupportedOperations()
    {
        using MemoryStream input = CreateDocx("Revenue increased.");
        using var patch = new StringReader("""
            docxpatch 1

            op remove-hyperlink
            target M.L0001
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6001");
        Assert.Contains("catalog support is 'preserve-only'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("does not generate new revision markup", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("M.L0001", diagnostic.TargetId);
        Assert.Equal("track-changes-no-revision-representation", diagnostic.Feature);
        Assert.Equal("require-failed", diagnostic.Fallback);
    }

    [Fact]
    public static void CheckTrackChangesRequireAllowsReplaceText()
    {
        using MemoryStream input = CreateDocx("Revenue increased.");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find increased
            with rose
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "E6001");
    }

    [Fact]
    public static void CheckTrackChangesRequireRejectsEmptyRevisionAuthor()
    {
        using MemoryStream input = CreateDocx("Revenue increased.");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find increased
            with rose
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Require,
            Author = " "
        });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6003");
        Assert.Equal("track-changes-revision-metadata", diagnostic.Feature);
        Assert.Equal("no-output-written", diagnostic.Fallback);
        Assert.Empty(result.Operations);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestRejectsEmptyRevisionAuthorBeforeWritingOutput()
    {
        using MemoryStream input = CreateDocx("Revenue increased.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find increased
            with rose
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            Author = ""
        });

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E6003");
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public static void CheckTrackChangesRequireMatchesCatalogSupportForEveryPatchOperation()
    {
            string[] annotationOperations = ["add-comment", "resolve-comment", "reopen-comment", "delete-comment", "add-comment-reply", "delete-comment-reply"];
        foreach (DocxPatchOperationInfo operation in DocxHelp.Catalog.PatchOperations)
        {
            using MemoryStream input = CreateDocx("Anchor paragraph.");
            using var patch = new StringReader($"""
                docxpatch 1

                op {operation.Name}
                target M.P0001
                end
                """);

            DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

            bool hasUnsupportedOperationDiagnostic = result.Diagnostics.Any(diagnostic => diagnostic.Code == "E6001");
            Assert.Equal(!operation.GeneratesTrackedChanges && !annotationOperations.Contains(operation.Name), hasUnsupportedOperationDiagnostic);
            if (hasUnsupportedOperationDiagnostic)
            {
                DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6001");
                Assert.Contains($"operation '{operation.Name}'", diagnostic.Message, StringComparison.Ordinal);
                Assert.Contains($"catalog support is '{operation.TrackChangesSupport}'", diagnostic.Message, StringComparison.Ordinal);
            }
        }
    }

    [Fact]
    public static void CheckTrackChangesSuggestMatchesCatalogSupportForEveryPatchOperation()
    {
        foreach (DocxPatchOperationInfo operation in DocxHelp.Catalog.PatchOperations)
        {
            using MemoryStream input = CreateDocx("Anchor paragraph.");
            using var patch = new StringReader($"""
                docxpatch 1

                op {operation.Name}
                target M.P0001
                end
                """);

            DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

            bool hasUnsupportedOperationWarning = result.Diagnostics.Any(diagnostic => diagnostic.Code == "W4001");
            Assert.Equal(!operation.GeneratesTrackedChanges, hasUnsupportedOperationWarning);
            if (hasUnsupportedOperationWarning)
            {
                DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "W4001");
                Assert.Contains($"operation '{operation.Name}'", diagnostic.Message, StringComparison.Ordinal);
                Assert.Contains($"catalog support is '{operation.TrackChangesSupport}'", diagnostic.Message, StringComparison.Ordinal);
            }
        }
    }

    [Fact]
    public static void CheckTrackChangesRequireRejectsComplexReplaceTextMarkup()
    {
        using MemoryStream input = CreateDocx("Revenue increased.");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find increased
            with <<<
            rose
            sharply
            >>>
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
        Assert.Contains("operation 'replace-text'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("target 'M.P0001'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("catalog support is 'tracked-simple'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("replacement contains tabs or line breaks", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-unsupported-target-shape", diagnostic.Feature);
        Assert.Equal("require-failed", diagnostic.Fallback);
    }

    [Fact]
    public static void CheckTrackChangesRequireRejectsMatchedTabs()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:t>Revenue</w:t><w:tab/><w:t>increased</w:t></w:r>
                    </w:p>
            """);
        using var patch = new StringReader("docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind <<<\nRevenue\tincreased\n>>>\nwith Revenue rose\nend\n");

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
    }

    [Theory]
    [InlineData("softHyphen", "<w:r><w:t>A</w:t><w:softHyphen/><w:t>B</w:t></w:r>")]
    [InlineData("sym", "<w:r><w:t>A</w:t><w:sym w:font=\"Wingdings\" w:char=\"F0FC\"/><w:t>B</w:t></w:r>")]
    [InlineData("cr", "<w:r><w:t>A</w:t><w:cr/><w:t>B</w:t></w:r>")]
    public static void CheckTrackChangesRequireRejectsNonTextRunContent(string unsupportedContent, string runXml)
    {
        using MemoryStream input = CreateDocxWithBody($"""
                    <w:p>
                      {runXml}
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find AB
            with AC
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
        Assert.Contains($"unsupported run content '{unsupportedContent}'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-unsupported-target-shape", diagnostic.Feature);
        Assert.Equal("require-failed", diagnostic.Fallback);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestFallsBackAndPreservesSoftHyphenRunContent()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:t>A</w:t><w:softHyphen/><w:t>B</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find AB
            with AC
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "W4002");
        Assert.Contains("unsupported run content 'softHyphen'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Empty(Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:softHyphen", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:del", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:ins", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Equal("AC", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void CheckTrackChangesRequireAllowsUniformSpanInMixedParagraph()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:rPr><w:b/></w:rPr><w:t>Revenue </w:t></w:r>
                      <w:r><w:rPr><w:i/></w:rPr><w:t>increased</w:t></w:r>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find increased
            with rose
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void CheckTrackChangesRequireRejectsSpanAcrossMixedRunFormatting()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:rPr><w:b/></w:rPr><w:t>Revenue </w:t></w:r>
                      <w:r><w:rPr><w:i/></w:rPr><w:t>increased</w:t></w:r>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Revenue increased
            with Revenue rose
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E6002" && diagnostic.Message.Contains("mixed direct run formatting", StringComparison.Ordinal));
    }

    [Fact]
    public static void ApplyTrackChangesSuggestReplacesAcrossRunsWithEquivalentDirectFormatting()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>Revenue</w:t></w:r>
                      <w:r><w:rPr><w:i/><w:b/></w:rPr><w:t> </w:t></w:r>
                      <w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>increased</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Revenue increased
            with Revenue rose
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "W4002");
        Assert.Equal(["1", "2"], Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:delText>Revenue increased</w:delText>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>Revenue rose</w:t>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:b", xml, StringComparison.Ordinal);
        Assert.Contains("<w:i", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestGeneratesRevisionMarkupForReplaceText()
    {
        using MemoryStream input = CreateDocx("Revenue increased.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find increased
            with rose
            end
            """);

        DateTimeOffset timestamp = DateTimeOffset.Parse("2026-06-08T12:00:00Z").ToUniversalTime();
        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            Author = "Reviewer",
            TimestampUtc = timestamp,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "W4001");
        DocxPatchOperationReport operation = Assert.Single(result.Operations);
        Assert.Equal(["1", "2"], operation.GeneratedRevisionIds);
        Assert.Contains("generated-revision-ids=1,2", DocxTextRenderer.RenderOperationSummary(result.Operations), StringComparison.Ordinal);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:del", xml, StringComparison.Ordinal);
        Assert.Contains("<w:delText>increased</w:delText>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:ins", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>rose</w:t>", xml, StringComparison.Ordinal);
        Assert.Contains("w:author=\"Reviewer\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:id=\"1\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:id=\"2\"", xml, StringComparison.Ordinal);

        output.Position = 0;
        Assert.Equal("Revenue rose.", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        output.Position = 0;
        Assert.Equal("Revenue increased.", Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { TextView = DocxTextView.Original }).Paragraphs).Text);
        output.Position = 0;
        Assert.Equal("Revenue [-increased-][+rose+].", Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { TextView = DocxTextView.Markup }).Paragraphs).Text);

        output.Position = 0;
        DocxChangesResult changes = new DocxEditor().Changes(output);
        DocxChangeInfo deletion = Assert.Single(changes.Changes, change => change.Type == "deleted-run");
        DocxChangeInfo insertion = Assert.Single(changes.Changes, change => change.Type == "inserted-run");
        Assert.Equal("Reviewer", deletion.Author);
        Assert.Equal(timestamp, deletion.TimestampUtc);
        Assert.Equal("1", deletion.RevisionId);
        Assert.Equal("2", insertion.RevisionId);
        output.Position = 0;
        Assert.True(new DocxEditor().Validate(output).Success);
    }

    [Fact]
    public static void ValidateReportsMalformedRevisionMarkup()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:ins w:author="">
                        <w:r><w:t>Bad insertion</w:t></w:r>
                      </w:ins>
                      <w:r><w:delText>Orphan delete text</w:delText></w:r>
                    </w:p>
                    <w:p>
                      <w:pPr>
                        <w:pPrChange w:id="4" w:author="Reviewer" w:date="2026-06-08T12:00:00Z"/>
                      </w:pPr>
                    </w:p>
                    <w:sectPr>
                      <w:sectPrChange w:id="8" w:author="Reviewer" w:date="2026-06-08T12:00:00Z"/>
                    </w:sectPr>
                    <w:tbl>
                      <w:tblPr>
                        <w:tblPrChange w:id="5" w:author="Reviewer" w:date="2026-06-08T12:00:00Z"/>
                      </w:tblPr>
                      <w:tr>
                        <w:trPr>
                          <w:trPrChange w:id="6" w:author="Reviewer" w:date="2026-06-08T12:00:00Z"/>
                        </w:trPr>
                        <w:tc>
                          <w:tcPr>
                            <w:tcPrChange w:id="7" w:author="Reviewer" w:date="2026-06-08T12:00:00Z"/>
                          </w:tcPr>
                          <w:p><w:r><w:t>Cell</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """);

        DocxValidateResult result = new DocxEditor().Validate(input);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E9121" && diagnostic.Message.Contains("w:id", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E9121" && diagnostic.Message.Contains("w:author", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E9121" && diagnostic.Message.Contains("w:delText", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E9121" && diagnostic.Message.Contains("w:pPrChange", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E9121" && diagnostic.Message.Contains("w:tblPrChange", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E9121" && diagnostic.Message.Contains("w:trPrChange", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E9121" && diagnostic.Message.Contains("w:tcPrChange", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E9121" && diagnostic.Message.Contains("w:sectPrChange", StringComparison.Ordinal));
    }

    [Fact]
    public static void ApplyTrackChangesSuggestNormalizesRevisionAuthorAndTimestamp()
    {
        using MemoryStream input = CreateDocx("Revenue increased.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find increased
            with rose
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            Author = " Agent ",
            TimestampUtc = DateTimeOffset.Parse("2026-06-08T14:00:00+02:00"),
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("w:author=\"Agent\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:date=\"2026-06-08T12:00:00.0000000+00:00\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyReportsEffectiveRevisionMetadata()
    {
        using MemoryStream input = CreateDocx("Revenue increased.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find increased
            with rose
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            Author = " Agent ",
            TimestampUtc = DateTimeOffset.Parse("2026-06-08T14:00:00+02:00"),
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        Assert.Equal("Agent", result.Author);
        Assert.Equal(TimeSpan.Zero, result.TimestampUtc.Offset);
        Assert.Equal(new DateTimeOffset(2026, 6, 8, 12, 0, 0, TimeSpan.Zero), result.TimestampUtc);
    }

    [Fact]
    public static void CheckReportsEffectiveRevisionMetadata()
    {
        using MemoryStream input = CreateDocx("Revenue increased.");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find increased
            with rose
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            Author = " Agent ",
            TimestampUtc = DateTimeOffset.Parse("2026-06-08T14:00:00+02:00")
        });

        Assert.True(result.Success);
        Assert.Equal("Agent", result.Author);
        Assert.Equal(TimeSpan.Zero, result.TimestampUtc.Offset);
        Assert.Equal(new DateTimeOffset(2026, 6, 8, 12, 0, 0, TimeSpan.Zero), result.TimestampUtc);
    }

    [Fact]
    public static void FailedCheckReportsEffectiveRevisionMetadata()
    {
        using MemoryStream input = CreateDocx("Revenue increased.");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P9999
            find increased
            with rose
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions
        {
            Author = " Agent ",
            TimestampUtc = DateTimeOffset.Parse("2026-06-08T14:00:00+02:00")
        });

        Assert.False(result.Success);
        Assert.Equal("Agent", result.Author);
        Assert.Equal(TimeSpan.Zero, result.TimestampUtc.Offset);
        Assert.Equal(new DateTimeOffset(2026, 6, 8, 12, 0, 0, TimeSpan.Zero), result.TimestampUtc);
    }

    [Fact]
    public static void MixedCommentAndTrackedOperationsKeepWordIdsUnique()
    {
        using MemoryStream input = CreateDocxWithBody("""
            <w:p><w:r><w:t>Anchor paragraph.</w:t></w:r></w:p>
            <w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Anchor
            with Anchored
            end

            op add-comment
            target M.P0001
            text First note
            end

            op add-comment
            target M.P0001
            text Second note
            end

            op add-comment
            target M.P0001
            text Third note
            end

            op add-comment
            target M.P0001
            text Fourth note
            end

            op replace-text
            target M.P0002
            find Second
            with 2nd
            end
            """);

        var options = new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest };
        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, options);

        Assert.True(result.Success);
        Assert.Equal(6, result.Operations.Count);
        Assert.NotEmpty(result.Operations[0].GeneratedRevisionIds);
        Assert.All(result.Operations.Skip(1).Take(4), operation => Assert.Empty(operation.GeneratedRevisionIds));
        Assert.NotEmpty(result.Operations[5].GeneratedRevisionIds);

        output.Position = 0;
        var revisionIds = new List<string>();
        var commentIds = new List<string>();
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        using (var archive = new ZipArchive(output, ZipArchiveMode.Read, leaveOpen: true))
        {
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                if (!entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                using Stream entryStream = entry.Open();
                XDocument document = XDocument.Load(entryStream);
                foreach (XElement element in document.Descendants())
                {
                    if (element.Name != w + "ins" && element.Name != w + "del" && element.Name != w + "comment")
                    {
                        continue;
                    }
                    string? id = (string?)element.Attribute(w + "id");
                    if (id is null)
                    {
                        continue;
                    }
                    if (element.Name == w + "comment")
                    {
                        commentIds.Add(id);
                    }
                    else
                    {
                        revisionIds.Add(id);
                    }
                }
            }
        }

        Assert.NotEmpty(revisionIds);
        Assert.NotEmpty(commentIds);
        Assert.Equal(revisionIds.Count, new HashSet<string>(revisionIds).Count);
        // Comment IDs number independently from the zero-based comment space, so they may
        // overlap earlier revision IDs; revision IDs must instead stay above every w:id the
        // revision scan has seen, including comments (guaranteed here by rescanning after
        // comment operations).
        int maxCommentId = commentIds.Select(int.Parse).Max();
        Assert.All(result.Operations[5].GeneratedRevisionIds, id => Assert.True(int.Parse(id) > maxCommentId));
    }

    [Fact]
    public static void ApplyTrackChangesSuggestGeneratesRevisionMarkupForMultipleAdjacentRunMatches()
    {
        using MemoryStream input = CreateDocxWithRuns("foo ", "foo");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find foo
            with bar
            occurrence all
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Equal(2, CountOccurrences(xml, "<w:del "));
        Assert.Equal(2, CountOccurrences(xml, "<w:ins "));
        output.Position = 0;
        Assert.Equal("bar bar", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestHandlesPunctuationBoundariesAndRepeatedOccurrences()
    {
        using MemoryStream input = CreateDocx("alpha, alpha; alpha.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find alpha
            with beta
            occurrence all
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Equal(3, CountOccurrences(xml, "<w:del "));
        Assert.Equal(3, CountOccurrences(xml, "<w:delText>alpha</w:delText>"));
        Assert.Equal(3, CountOccurrences(xml, "<w:ins "));
        Assert.Equal(3, CountOccurrences(xml, "<w:t>beta</w:t>"));
        output.Position = 0;
        Assert.Equal("beta, beta; beta.", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestReplacesHeaderAndFooterParagraphText()
    {
        using MemoryStream input = CreateDocxWithHeaderFooter("Header increased.", "Footer increased.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target H001.P0001
            find increased
            with rose
            end

            op replace-text
            target F001.P0001
            find increased
            with fell
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            Author = "Agent",
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        Assert.Equal(["1", "2"], result.Operations[0].GeneratedRevisionIds);
        Assert.Equal(["3", "4"], result.Operations[1].GeneratedRevisionIds);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output, new DocxReadOptions { IncludeHeadersFooters = true });
        Assert.Contains(read.Paragraphs, paragraph => paragraph.Id.ToWireValue() == "H001.P0001" && paragraph.Text == "Header rose.");
        Assert.Contains(read.Paragraphs, paragraph => paragraph.Id.ToWireValue() == "F001.P0001" && paragraph.Text == "Footer fell.");
        output.Position = 0;
        Assert.Contains("<w:delText>increased</w:delText>", ReadEntry(output, "word/header1.xml"), StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains("<w:delText>increased</w:delText>", ReadEntry(output, "word/footer1.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestPreservesListParagraphProperties()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:pPr>
                        <w:numPr>
                          <w:ilvl w:val="0"/>
                          <w:numId w:val="9"/>
                        </w:numPr>
                      </w:pPr>
                      <w:r><w:t>Revenue increased.</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find increased
            with rose
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:numPr>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:numId w:val=\"9\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:delText>increased</w:delText>", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Equal("Revenue rose.", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestTracksUniformSpanInMixedParagraph()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:rPr><w:b/></w:rPr><w:t>Revenue </w:t></w:r>
                      <w:r><w:rPr><w:i/></w:rPr><w:t>increased</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find increased
            with rose
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "W4002");
        Assert.Equal(["1", "2"], Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:delText>increased</w:delText>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>rose</w:t>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:b", xml, StringComparison.Ordinal);
        Assert.Contains("<w:i", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Equal("Revenue rose", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestWarnsAndAppliesUnsupportedOperationsDirectly()
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
                        <w:sdtContent>
                          <w:r><w:t>Unchecked</w:t></w:r>
                        </w:sdtContent>
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

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "W4001" && diagnostic.Severity == DocxSeverity.Warning);
        Assert.Contains("catalog support is 'preserve-only'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("apply operation 'set-content-control-checkbox' directly", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-no-revision-representation", diagnostic.Feature);
        Assert.Equal("direct-edit-preserve-existing-revisions", diagnostic.Fallback);
        output.Position = 0;
        Assert.Equal(char.ConvertFromUtf32(0x2612), Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestGeneratesRevisionMarkupForReplaceParagraph()
    {
        using MemoryStream input = CreateDocx("Revenue increased.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-paragraph
            target M.P0001
            text Revenue rose.
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            Author = "Reviewer",
            TimestampUtc = DateTimeOffset.Parse("2026-06-08T12:00:00Z").ToUniversalTime(),
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:del", xml, StringComparison.Ordinal);
        Assert.Contains("<w:delText>Revenue increased.</w:delText>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:ins", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>Revenue rose.</w:t>", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Equal("Revenue rose.", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        output.Position = 0;
        Assert.Equal("Revenue increased.", Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { TextView = DocxTextView.Original }).Paragraphs).Text);
    }

    [Fact]
    public static void CheckTrackChangesRequireAllowsReplaceParagraphTextAndStyleCombination()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/></w:style>
            """, """
              <w:p><w:r><w:t>Revenue increased.</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-paragraph
            target M.P0001
            text Revenue rose.
            style Heading1
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
    }

    [Fact]
    public static void ApplyTrackChangesSuggestGeneratesRevisionMarkupForReplaceParagraphTextAndStyleCombination()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/></w:style>
            """, """
              <w:p><w:r><w:t>Revenue increased.</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-paragraph
            target M.P0001
            text Revenue rose.
            style Heading1
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "W4002");
        Assert.Equal(["1", "2", "3"], Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:del", xml, StringComparison.Ordinal);
        Assert.Contains("<w:delText>Revenue increased.</w:delText>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:ins", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>Revenue rose.</w:t>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:pStyle w:val=\"Heading1\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:pPrChange w:id=\"3\"", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Equal("Revenue rose.", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestGeneratesRevisionMarkupForInsertedParagraph()
    {
        using MemoryStream input = CreateDocx("One");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text Two
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:ins", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>Two</w:t>", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Equal(new[] { "One", "Two" }, new DocxEditor().Read(output).Paragraphs.Select(paragraph => paragraph.Text));
    }

    [Fact]
    public static void ApplyTrackChangesSuggestInsertsParagraphsInHeadersAndFooters()
    {
        using MemoryStream input = CreateDocxWithHeaderFooter("Header one", "Footer one");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target H001.P0001
            text Header two
            end

            op insert-before
            target F001.P0001
            text Footer zero
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            Author = "Agent",
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        Assert.Equal(["1"], result.Operations[0].GeneratedRevisionIds);
        Assert.Equal(["2"], result.Operations[1].GeneratedRevisionIds);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output, new DocxReadOptions { IncludeHeadersFooters = true });
        Assert.Contains(read.Paragraphs, paragraph => paragraph.Id.ToWireValue() == "H001.P0002" && paragraph.Text == "Header two");
        Assert.Contains(read.Paragraphs, paragraph => paragraph.Id.ToWireValue() == "F001.P0001" && paragraph.Text == "Footer zero");
        output.Position = 0;
        Assert.Contains("<w:ins w:id=\"1\" w:author=\"Agent\"", ReadEntry(output, "word/header1.xml"), StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains("<w:ins w:id=\"2\" w:author=\"Agent\"", ReadEntry(output, "word/footer1.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestCopiesListParagraphPropertiesForInsertedParagraph()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:pPr>
                        <w:numPr>
                          <w:ilvl w:val="0"/>
                          <w:numId w:val="9"/>
                        </w:numPr>
                      </w:pPr>
                      <w:r><w:t>Item one</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            copy-paragraph-properties true
            text Item two
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Equal(2, CountOccurrences(xml, "<w:numPr>"));
        Assert.Contains("<w:ins w:id=\"1\"", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Equal(new[] { "Item one", "Item two" }, new DocxEditor().Read(output).Paragraphs.Select(paragraph => paragraph.Text));
    }

    [Fact]
    public static void ApplyTrackChangesSuggestSanitizesCopiedParagraphPropertiesForInsertedParagraph()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:pPr>
                        <w:pStyle w:val="BodyText"/>
                        <w:numPr>
                          <w:ilvl w:val="0"/>
                          <w:numId w:val="9"/>
                        </w:numPr>
                        <w:pPrChange w:id="4" w:author="Reviewer" w:date="2026-06-08T12:00:00Z">
                          <w:pPr><w:pStyle w:val="OldStyle"/></w:pPr>
                        </w:pPrChange>
                        <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                      </w:pPr>
                      <w:r><w:t>Item one</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            copy-paragraph-properties true
            text Item two
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Equal(2, CountOccurrences(xml, "w:pStyle w:val=\"BodyText\""));
        Assert.Equal(2, CountOccurrences(xml, "<w:numPr>"));
        Assert.Equal(1, CountOccurrences(xml, "<w:pPrChange "));
        Assert.Equal(1, CountOccurrences(xml, "<w:sectPr>"));
        Assert.Contains("<w:ins w:id=\"5\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestInsertsListContinuationWithCopiedNumberingProperties()
    {
        using MemoryStream input = CreateDocxWithNumbering(
            """
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr>
                      <w:r><w:t>Item one</w:t></w:r>
                    </w:p>
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr>
                      <w:r><w:t>Item three</w:t></w:r>
                    </w:p>
            """,
            SimpleDecimalNumberingXml());
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            copy-paragraph-properties true
            text Item two
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        Assert.Equal(["1"], Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal(["Item one", "Item two", "Item three"], read.Paragraphs.Select(paragraph => paragraph.Text).ToArray());
        Assert.Equal(["1.", "2.", "3."], read.Paragraphs.Select(paragraph => paragraph.List?.LabelText ?? string.Empty).ToArray());
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Equal(3, CountOccurrences(xml, "<w:numPr>"));
        Assert.Contains("<w:ins w:id=\"1\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestInsertsListContinuationWithRestartOverride()
    {
        using MemoryStream input = CreateDocxWithNumbering(
            """
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr>
                      <w:r><w:t>Default item</w:t></w:r>
                    </w:p>
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="10"/></w:numPr></w:pPr>
                      <w:r><w:t>Restart item</w:t></w:r>
                    </w:p>
            """,
            """
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="7">
                    <w:lvl w:ilvl="0">
                      <w:start w:val="1"/>
                      <w:numFmt w:val="decimal"/>
                      <w:lvlText w:val="%1."/>
                    </w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="9"><w:abstractNumId w:val="7"/></w:num>
                  <w:num w:numId="10">
                    <w:abstractNumId w:val="7"/>
                    <w:lvlOverride w:ilvl="0"><w:startOverride w:val="7"/></w:lvlOverride>
                  </w:num>
                </w:numbering>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0002
            copy-paragraph-properties true
            text Restart continuation
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal(["Default item", "Restart item", "Restart continuation"], read.Paragraphs.Select(paragraph => paragraph.Text).ToArray());
        Assert.Equal(["1.", "7.", "8."], read.Paragraphs.Select(paragraph => paragraph.List?.LabelText ?? string.Empty).ToArray());
        Assert.Equal(["9", "10", "10"], read.Paragraphs.Select(paragraph => paragraph.List?.NumberingId ?? string.Empty).ToArray());
        output.Position = 0;
        Assert.Contains("<w:ins w:id=\"1\"", ReadDocumentXml(output), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyTrackedInsertionPreservesExistingTrackedChangesAndComments()
    {
        using MemoryStream input = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:ins w:id="9" w:author="Existing" w:date="2026-06-01T12:00:00Z">
                        <w:r><w:t>Existing insertion </w:t></w:r>
                      </w:ins>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="3" w:author="Commenter">
                    <w:p><w:r><w:t>Existing comment</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.P0001
            text New tracked paragraph
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            Author = "Agent",
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        output.Position = 0;
        string documentXml = ReadDocumentXml(output);
        Assert.Contains("w:id=\"9\" w:author=\"Existing\"", documentXml, StringComparison.Ordinal);
        Assert.Contains("commentRangeStart w:id=\"3\"", documentXml, StringComparison.Ordinal);
        Assert.Contains("commentReference w:id=\"3\"", documentXml, StringComparison.Ordinal);
        Assert.Contains("w:id=\"10\" w:author=\"Agent\"", documentXml, StringComparison.Ordinal);
        Assert.Contains("New tracked paragraph", documentXml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains("Existing comment", ReadEntry(output, "word/comments.xml"), StringComparison.Ordinal);
        output.Position = 0;
        DocxChangesResult changes = new DocxEditor().Changes(output);
        Assert.Contains(changes.Changes, change => change.Type == "inserted-run" && change.RevisionId == "9" && change.Author == "Existing");
        Assert.Contains(changes.Changes, change => change.Type == "inserted-run" && change.RevisionId == "10" && change.Author == "Agent");
        Assert.Contains(changes.CommentSummary, summary => summary.CommentId == "3");
    }

    [Fact]
    public static void ApplyTrackChangesSuggestAllocatesRevisionIdsAcrossWordStories()
    {
        using MemoryStream input = CreateDocxWithRevisionIdsAcrossStories();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0002
            find increased
            with rose
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            Author = "Agent",
            TimestampUtc = DateTimeOffset.Parse("2026-06-08T12:00:00Z").ToUniversalTime(),
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        output.Position = 0;
        string documentXml = ReadDocumentXml(output);
        Assert.Contains("w:id=\"31\" w:author=\"Agent\"", documentXml, StringComparison.Ordinal);
        Assert.Contains("w:id=\"32\" w:author=\"Agent\"", documentXml, StringComparison.Ordinal);
        Assert.Contains("w:id=\"9\" w:author=\"Main\"", documentXml, StringComparison.Ordinal);
        Assert.Contains("w:id=\"21\" w:author=\"Table\"", documentXml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains("w:id=\"12\" w:author=\"Header\"", ReadEntry(output, "word/header1.xml"), StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains("w:id=\"17\" w:author=\"Footer\"", ReadEntry(output, "word/footer1.xml"), StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains("w:id=\"30\" w:author=\"Comment\"", ReadEntry(output, "word/comments.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestGeneratesRevisionMarkupForDeletedParagraph()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>One</w:t></w:r></w:p>
                    <w:p><w:r><w:t>Two</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-block
            target M.P0002
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:del", xml, StringComparison.Ordinal);
        Assert.Contains("<w:delText>Two</w:delText>", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains(new DocxEditor().Changes(output).Changes, change => change.Type == "deleted-run" && change.TargetId == "M.P0002");
    }

    [Fact]
    public static void CheckTrackChangesRequireRejectsDeletingSectionBoundaryParagraph()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:pPr>
                        <w:sectPr>
                          <w:pgSz w:w="12240" w:h="15840"/>
                        </w:sectPr>
                      </w:pPr>
                      <w:r><w:t>Section boundary</w:t></w:r>
                    </w:p>
                    <w:p><w:r><w:t>Next section</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op delete-block
            target M.P0001
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
        Assert.Contains("operation 'delete-block'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("target 'M.P0001'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("paragraph contains section properties", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-unsupported-target-shape", diagnostic.Feature);
        Assert.Equal("require-failed", diagnostic.Fallback);
    }

    [Fact]
    public static void CheckTrackChangesRequireRejectsDeletingTableBlockAsUnsupportedShape()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op delete-block
            target M.T0001
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
        Assert.Contains("tracked block deletion is supported only for paragraph targets", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-unsupported-target-shape", diagnostic.Feature);
        Assert.Equal("require-failed", diagnostic.Fallback);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestDeletesTableBlockDirectlyWithWarning()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:p><w:r><w:t>After</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-block
            target M.T0001
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "W4002");
        Assert.Contains("tracked block deletion is supported only for paragraph targets", diagnostic.Message, StringComparison.Ordinal);
        Assert.Empty(Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Empty(read.Tables);
        Assert.Equal("After", Assert.Single(read.Paragraphs).Text);
    }

    [Fact]
    public static void CheckTrackChangesRequireRejectsDeletingImageParagraph()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:drawing/></w:r>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op delete-block
            target M.P0001
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
        Assert.Contains("protected OOXML boundary 'drawing'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-unsupported-target-shape", diagnostic.Feature);
        Assert.Equal("require-failed", diagnostic.Fallback);
    }

    [Theory]
    [InlineData("field", """
                    <w:p>
                      <w:fldSimple w:instr=" DATE ">
                        <w:r><w:t>June 2026</w:t></w:r>
                      </w:fldSimple>
                    </w:p>
        """)]
    [InlineData("bookmark", """
                    <w:p>
                      <w:bookmarkStart w:id="4" w:name="ClientName"/>
                      <w:r><w:t>Client</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                    </w:p>
        """)]
    [InlineData("comment", """
                    <w:p>
                      <w:commentRangeStart w:id="3"/>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
        """)]
    [InlineData("hyperlink", """
                    <w:p>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rLink">
                        <w:r><w:t>Link text</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
        """)]
    public static void CheckTrackChangesRequireRejectsDeletingProtectedParagraphContent(string protectedFeature, string bodyXml)
    {
        using MemoryStream input = CreateDocxWithBody(bodyXml);
        using var patch = new StringReader("""
            docxpatch 1

            op delete-block
            target M.P0001
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
        Assert.Contains("operation 'delete-block'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("target 'M.P0001'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains($"protected OOXML boundary '{protectedFeature}'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-unsupported-target-shape", diagnostic.Feature);
        Assert.Equal("require-failed", diagnostic.Fallback);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestGeneratesParagraphPropertyChangeForStyle()
    {
        using MemoryStream input = CreateDocxWithStyles("""
                  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
                  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/></w:style>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-style
            target M.P0001
            style Heading2
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            Author = "Reviewer",
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("w:pStyle w:val=\"Heading2\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:pPrChange", xml, StringComparison.Ordinal);
        Assert.Contains("w:author=\"Reviewer\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestCoversParagraphStylePropertyShapes()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody(
            """
                  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
                  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/></w:style>
            """,
            """
                    <w:p><w:r><w:t>No properties</w:t></w:r></w:p>
                    <w:p>
                      <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
                      <w:r><w:t>Existing style</w:t></w:r>
                    </w:p>
                    <w:p>
                      <w:pPr>
                        <w:spacing w:before="120"/>
                        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr>
                      </w:pPr>
                      <w:r><w:t>Complex properties</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-style
            target M.P0001
            style Heading2
            end

            op set-style
            target M.P0002
            style Heading2
            end

            op set-style
            target M.P0003
            style Heading2
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            Author = "Reviewer",
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        Assert.Equal(["1"], result.Operations[0].GeneratedRevisionIds);
        Assert.Equal(["2"], result.Operations[1].GeneratedRevisionIds);
        Assert.Equal(["3"], result.Operations[2].GeneratedRevisionIds);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Equal(3, CountOccurrences(xml, "<w:pPrChange "));
        Assert.Equal(3, CountOccurrences(xml, "w:pStyle w:val=\"Heading2\""));
        Assert.Contains("<w:spacing w:before=\"120\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:numPr>", xml, StringComparison.Ordinal);
        output.Position = 0;
        DocxChangesResult changes = new DocxEditor().Changes(output);
        Assert.Contains(changes.Summary, summary => summary.Type == "paragraph-properties-change" && summary.Count == 3);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestGeneratesRevisionMarkupForSetCell()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Old</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-cell
            target M.T0001.R01.C01
            text New
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:del", xml, StringComparison.Ordinal);
        Assert.Contains("<w:delText>Old</w:delText>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:ins", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>New</w:t>", xml, StringComparison.Ordinal);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal("New", Assert.Single(table.Cells).Text);
        output.Position = 0;
        DocxTableInfo originalTable = Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { TextView = DocxTextView.Original }).Tables);
        Assert.Equal("Old", Assert.Single(originalTable.Cells).Text);
        output.Position = 0;
        DocxTableInfo markupTable = Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { TextView = DocxTextView.Markup }).Tables);
        Assert.Equal("[-Old-][+New+]", Assert.Single(markupTable.Cells).Text);
    }

    [Fact]
    public static void CheckTrackChangesRequireAllowsSetCellWithMultipleCompatibleParagraphs()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:p><w:r><w:t>One</w:t></w:r></w:p>
                          <w:p><w:r><w:t>Two</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-cell
            target M.T0001.R01.C01
            expect-text OneTwo
            text Replacement
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
    }

    [Fact]
    public static void ApplyTrackChangesSuggestSetCellSupportsMultipleCompatibleParagraphs()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>
                          <w:p>
                            <w:pPr><w:pStyle w:val="TableBody"/></w:pPr>
                            <w:r><w:t>One</w:t></w:r>
                          </w:p>
                          <w:p><w:r><w:t>Two</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-cell
            target M.T0001.R01.C01
            expect-text OneTwo
            text Replacement
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "W4002");
        Assert.Equal(["1", "2", "3"], Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Equal(2, CountOccurrences(xml, "<w:del "));
        Assert.Equal(1, CountOccurrences(xml, "<w:ins "));
        Assert.Contains("<w:t>Replacement</w:t>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:delText>One</w:delText>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:delText>Two</w:delText>", xml, StringComparison.Ordinal);
        Assert.Contains("w:val=\"TableBody\"", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Equal("Replacement", Assert.Single(new DocxEditor().Read(output).Tables).Cells[0].Text);
        output.Position = 0;
        Assert.Equal("OneTwo", Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { TextView = DocxTextView.Original }).Tables).Cells[0].Text);
        output.Position = 0;
        Assert.Equal("[-One-][+Replacement+][-Two-]", Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { TextView = DocxTextView.Markup }).Tables).Cells[0].Text);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestSetCellPreservesHorizontalMerge()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
                          <w:p><w:r><w:t>Old merged</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-cell
            target M.T0001.R01.C01
            expect-text Old merged
            text New merged
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:gridSpan w:val=\"2\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:delText>Old merged</w:delText>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>New merged</w:t>", xml, StringComparison.Ordinal);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal(2, table.ColumnCount);
        Assert.Equal("New merged", Assert.Single(table.Cells).Text);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestSetCellPreservesVerticalMergeRoot()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
                          <w:p><w:r><w:t>Old root</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:vMerge/></w:tcPr>
                          <w:p><w:r><w:t></w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-cell
            target M.T0001.R01.C01
            expect-text Old root
            text New root
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:vMerge w:val=\"restart\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:delText>Old root</w:delText>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>New root</w:t>", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckTrackChangesRequireRejectsProtectedBoundaries()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rLink">
                        <w:r><w:t>Link text</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Link
            with Anchor
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
        Assert.Contains("protected OOXML boundary 'hyperlink'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-unsupported-target-shape", diagnostic.Feature);
        Assert.Equal("require-failed", diagnostic.Fallback);
    }

    [Fact]
    public static void CheckTrackChangesRequireTracksTextInsideCommentRange()
    {
        using MemoryStream input = CreateDocxWithCommentAnchoredParagraph();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Commented
            with Updated
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));

        using MemoryStream applyInput = CreateDocxWithCommentAnchoredParagraph();
        using var output = new MemoryStream();
        using var applyPatch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Commented
            with Updated
            end
            """);

        DocxApplyResult apply = new DocxEditor().Apply(applyInput, applyPatch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });
        Assert.True(apply.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:delText>Commented</w:delText>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>Updated</w:t>", xml, StringComparison.Ordinal);
        Assert.Contains("commentRangeStart", xml, StringComparison.Ordinal);
        Assert.Contains("commentRangeEnd", xml, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("simple-field", """
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:fldSimple w:instr=" DATE ">
                        <w:r><w:t>June 12</w:t></w:r>
                      </w:fldSimple>
                      <w:r><w:t> after</w:t></w:r>
                    </w:p>
        """)]
    [InlineData("complex-field", """
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> PAGE </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:t>1</w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                      <w:r><w:t> after</w:t></w:r>
                    </w:p>
        """)]
    public static void CheckTrackChangesRequireAllowsTextReplacementAdjacentToFieldBoundaries(string shape, string bodyXml)
    {
        _ = shape;
        using MemoryStream input = CreateDocxWithBody(bodyXml);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Before
            with After
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void CheckTrackChangesRequireRejectsTextReplacementInsideSimpleField()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:fldSimple w:instr=" DATE ">
                        <w:r><w:t>June 12</w:t></w:r>
                      </w:fldSimple>
                      <w:r><w:t> after</w:t></w:r>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find June
            with July
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
        Assert.Contains("protected OOXML boundary 'field'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-unsupported-target-shape", diagnostic.Feature);
        Assert.Equal("require-failed", diagnostic.Fallback);
    }

    [Theory]
    [InlineData("adjacent-insertion")]
    [InlineData("adjacent-deletion")]
    public static void CheckTrackChangesRequireAllowsTextReplacementBesideExistingRevisions(string shape)
    {
        string bodyXml = shape switch
        {
            "adjacent-insertion" => """
                    <w:p>
                      <w:ins w:id="1" w:author="A" w:date="2026-06-01T00:00:00Z">
                        <w:r><w:t>Inserted </w:t></w:r>
                      </w:ins>
                      <w:r><w:t>plain text</w:t></w:r>
                    </w:p>
                """,
            _ => """
                    <w:p>
                      <w:del w:id="1" w:author="A" w:date="2026-06-01T00:00:00Z">
                        <w:r><w:delText>Removed </w:delText></w:r>
                      </w:del>
                      <w:r><w:t>plain text</w:t></w:r>
                    </w:p>
                """,
        };
        using MemoryStream input = CreateDocxWithBody(bodyXml);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find plain
            with edited
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
    }

    [Fact]
    public static void CheckTrackChangesRequireRejectsTextReplacementInsideInsertion()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:ins w:id="1" w:author="A" w:date="2026-06-01T00:00:00Z">
                        <w:r><w:t>Inserted </w:t></w:r>
                      </w:ins>
                      <w:r><w:t>plain text</w:t></w:r>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Inserted
            with edited
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
        Assert.Contains("tracked-insertion", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-unsupported-target-shape", diagnostic.Feature);
        Assert.Equal("require-failed", diagnostic.Fallback);
    }

    [Fact]
    public static void CheckTrackChangesRequireReportsNotFoundForTextOnlyInsideTrackedDeletion()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:del w:id="1" w:author="A" w:date="2026-06-01T00:00:00Z">
                        <w:r><w:delText>Removed </w:delText></w:r>
                      </w:del>
                      <w:r><w:t>plain text</w:t></w:r>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Removed
            with edited
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Code == "E4203");
    }

    [Fact]
    public static void CheckTrackChangesRequireAllowsSetHyperlinkText()
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
            text New link
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
    }

    [Fact]
    public static void ApplyTrackChangesSuggestGeneratesRevisionMarkupForSetHyperlinkText()
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
            text New link
            end
            """);

        DateTimeOffset timestamp = DateTimeOffset.Parse("2026-06-08T12:00:00Z").ToUniversalTime();
        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            Author = "Reviewer",
            TimestampUtc = timestamp,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "W4002");
        Assert.Equal(["1", "2"], Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:hyperlink", xml, StringComparison.Ordinal);
        Assert.Contains("r:id=\"rLink\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:delText>Old link</w:delText>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>New link</w:t>", xml, StringComparison.Ordinal);
        Assert.Contains("w:author=\"Reviewer\"", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Equal("New link", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        output.Position = 0;
        Assert.Equal("Old link", Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { TextView = DocxTextView.Original }).Paragraphs).Text);
        output.Position = 0;
        Assert.Equal("[-Old link-][+New link+]", Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { TextView = DocxTextView.Markup }).Paragraphs).Text);
        output.Position = 0;
        DocxHyperlinkInfo hyperlink = Assert.Single(new DocxEditor().Read(output).Hyperlinks);
        Assert.Equal("https://example.test/old", hyperlink.Uri);
        Assert.Equal(8, hyperlink.DisplayTextLength);
        output.Position = 0;
        DocxChangesResult changes = new DocxEditor().Changes(output);
        Assert.Contains(changes.Changes, change => change.Type == "deleted-run" && change.RevisionId == "1" && change.Author == "Reviewer");
        Assert.Contains(changes.Changes, change => change.Type == "inserted-run" && change.RevisionId == "2" && change.Author == "Reviewer");
    }

    [Fact]
    public static void ApplyTrackChangesSuggestGeneratesRevisionMarkupForHeaderAndFooterHyperlinkText()
    {
        using MemoryStream input = CreateDocxWithHeaderFooterContent(
            """
                  <w:p>
                    <w:hyperlink w:anchor="HeaderAnchor">
                      <w:r><w:t>Header link</w:t></w:r>
                    </w:hyperlink>
                  </w:p>
            """,
            """
                  <w:p>
                    <w:hyperlink w:anchor="FooterAnchor">
                      <w:r><w:t>Footer link</w:t></w:r>
                    </w:hyperlink>
                  </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-hyperlink-text
            target H001.L0001
            text Header new
            end

            op set-hyperlink-text
            target F001.L0001
            text Footer new
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        Assert.Equal(["1", "2"], result.Operations[0].GeneratedRevisionIds);
        Assert.Equal(["3", "4"], result.Operations[1].GeneratedRevisionIds);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output, new DocxReadOptions { IncludeHeadersFooters = true });
        Assert.Contains(read.Paragraphs, paragraph => paragraph.Story == "header[1]" && paragraph.Text == "Header new");
        Assert.Contains(read.Paragraphs, paragraph => paragraph.Story == "footer[1]" && paragraph.Text == "Footer new");
        Assert.Contains(read.Hyperlinks, hyperlink => hyperlink.Story == "header[1]" && hyperlink.Anchor == "HeaderAnchor");
        Assert.Contains(read.Hyperlinks, hyperlink => hyperlink.Story == "footer[1]" && hyperlink.Anchor == "FooterAnchor");
        output.Position = 0;
        Assert.Contains("<w:delText>Header link</w:delText>", ReadEntry(output, "word/header1.xml"), StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains("<w:delText>Footer link</w:delText>", ReadEntry(output, "word/footer1.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckTrackChangesRequireRejectsSetHyperlinkTextWithExistingRevision()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:hyperlink w:anchor="Anchor">
                        <w:ins w:id="9" w:author="Reviewer" w:date="2026-06-01T00:00:00Z">
                          <w:r><w:t>Old link</w:t></w:r>
                        </w:ins>
                      </w:hyperlink>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-hyperlink-text
            target M.L0001
            text New link
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
        Assert.Contains("protected OOXML boundary 'tracked-insertion'", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestSetHyperlinkAnchorRemovesUnusedRelationship()
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

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "W4001");
        Assert.Contains("operation 'set-hyperlink-target'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("catalog support is 'preserve-only'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Empty(Assert.Single(result.Operations).GeneratedRevisionIds);
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
    public static void ApplyTrackChangesSuggestRemoveHyperlinkPreservesRelationshipStillInUse()
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

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "W4001");
        Assert.Contains("operation 'remove-hyperlink'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("catalog support is 'preserve-only'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Empty(Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal("First Second", Assert.Single(read.Paragraphs).Text);
        Assert.Single(read.Hyperlinks);
        output.Position = 0;
        string relationships = ReadEntry(output, "word/_rels/document.xml.rels");
        Assert.Contains("Id=\"rShared\"", relationships, StringComparison.Ordinal);
        Assert.Contains("Target=\"https://example.test/shared\"", relationships, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckTrackChangesRequireAllowsInsertHyperlinkAfter()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>Anchor</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op insert-hyperlink-after
            target M.P0001
            text Docs
            uri https://docs.example/
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code is "E6001" or "E6002");
    }

    [Fact]
    public static void ApplyTrackChangesSuggestGeneratesRevisionMarkupForInsertHyperlinkAfter()
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
            tooltip Documentation
            target-frame _blank
            history true
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code is "W4001" or "W4002");
        Assert.Equal(["1"], Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal(["Anchor", "Docs"], read.Paragraphs.Select(paragraph => paragraph.Text).ToArray());
        DocxHyperlinkInfo hyperlink = Assert.Single(read.Hyperlinks);
        Assert.Equal(4, hyperlink.DisplayTextLength);
        Assert.Equal("https://docs.example/", hyperlink.Uri);
        Assert.Equal("Documentation", hyperlink.Tooltip);
        Assert.Equal("_blank", hyperlink.TargetFrame);
        Assert.Equal(true, hyperlink.History);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:hyperlink", xml, StringComparison.Ordinal);
        Assert.Contains("<w:ins", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>Docs</w:t>", xml, StringComparison.Ordinal);
        Assert.Contains("w:tooltip=\"Documentation\"", xml, StringComparison.Ordinal);
        output.Position = 0;
        string relationships = ReadEntry(output, "word/_rels/document.xml.rels");
        Assert.Contains("Target=\"https://docs.example/\"", relationships, StringComparison.Ordinal);
        Assert.Contains("TargetMode=\"External\"", relationships, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Equal(["Anchor", ""], new DocxEditor().Read(output, new DocxReadOptions { TextView = DocxTextView.Original }).Paragraphs.Select(paragraph => paragraph.Text).ToArray());
        output.Position = 0;
        Assert.Equal(["Anchor", "[+Docs+]"], new DocxEditor().Read(output, new DocxReadOptions { TextView = DocxTextView.Markup }).Paragraphs.Select(paragraph => paragraph.Text).ToArray());
    }

    [Fact]
    public static void CheckTrackChangesRequireRejectsInsertHyperlinkAfterWithLineBreak()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>Anchor</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op insert-hyperlink-after
            target M.P0001
            text <<<
            Docs
            More
            >>>
            uri https://docs.example/
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
        Assert.Contains("inserted hyperlink text contains tabs or line breaks", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestSetContentControlTextPreservesSdtProperties()
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

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code is "W4001" or "W4002");
        Assert.Equal(["1", "2"], Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal("New Client", Assert.Single(read.Paragraphs).Text);
        DocxContentControlInfo control = Assert.Single(read.ContentControls);
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
        Assert.Contains("w:alias w:val=\"Client Name\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:lock w:val=\"unlocked\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:dataBinding w:xpath=\"/root/client\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:delText>Old Client</w:delText>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>New Client</w:t>", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Equal("Old Client", Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { TextView = DocxTextView.Original }).Paragraphs).Text);
        output.Position = 0;
        Assert.Equal("[-Old Client-][+New Client+]", Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { TextView = DocxTextView.Markup }).Paragraphs).Text);
    }

    [Fact]
    public static void CheckTrackChangesRequireAllowsPlainTextContentControlText()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:text/>
                          <w:tag w:val="client-name"/>
                        </w:sdtPr>
                        <w:sdtContent>
                          <w:r><w:t>Old Client</w:t></w:r>
                        </w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-content-control-text
            target M.CC0001
            text New Client
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
    }

    [Fact]
    public static void CheckTrackChangesRequireAllowsRichTextContentControlText()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:richText/>
                          <w:tag w:val="client-name"/>
                        </w:sdtPr>
                        <w:sdtContent>
                          <w:p><w:r><w:t>Old Client</w:t></w:r></w:p>
                        </w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-content-control-text
            target M.CC0001
            expect-text Old Client
            text New Client
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
    }

    [Fact]
    public static void ApplyTrackChangesSuggestSetRichTextContentControlTextPreservesParagraphs()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:id w:val="77"/>
                          <w:tag w:val="summary"/>
                          <w:alias w:val="Summary"/>
                          <w:richText/>
                        </w:sdtPr>
                        <w:sdtContent>
                          <w:p>
                            <w:pPr><w:pStyle w:val="BodyText"/></w:pPr>
                            <w:r><w:t>One</w:t></w:r>
                          </w:p>
                          <w:p><w:r><w:t>Two</w:t></w:r></w:p>
                        </w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-content-control-text
            target M.CC0001
            expect-text OneTwo
            text Replacement
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code is "W4001" or "W4002");
        Assert.Equal(["1", "2", "3"], Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:sdt>", xml, StringComparison.Ordinal);
        Assert.Contains("w:id w:val=\"77\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:tag w:val=\"summary\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:alias w:val=\"Summary\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:richText", xml, StringComparison.Ordinal);
        Assert.Contains("w:val=\"BodyText\"", xml, StringComparison.Ordinal);
        Assert.Equal(2, CountOccurrences(xml, "<w:del "));
        Assert.Equal(1, CountOccurrences(xml, "<w:ins "));
        Assert.Contains("<w:delText>One</w:delText>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:delText>Two</w:delText>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>Replacement</w:t>", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Equal("Replacement", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        output.Position = 0;
        Assert.Equal("OneTwo", Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { TextView = DocxTextView.Original }).Paragraphs).Text);
        output.Position = 0;
        Assert.Equal("[-One-][+Replacement+][-Two-]", Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { TextView = DocxTextView.Markup }).Paragraphs).Text);
    }

    [Theory]
    [InlineData("set-content-control-checkbox", "<w:checkBox><w:checked w:val=\"0\"/><w:checkedState w:val=\"2612\"/><w:uncheckedState w:val=\"2610\"/></w:checkBox>", "Unchecked", "checked true")]
    [InlineData("set-content-control-choice", "<w:dropDownList><w:listItem w:displayText=\"North\" w:value=\"north\"/><w:listItem w:displayText=\"South\" w:value=\"south\"/></w:dropDownList>", "North", "value south")]
    [InlineData("set-content-control-date", "<w:date><w:dateFormat w:val=\"yyyy-MM-dd\"/><w:fullDate w:val=\"2026-06-12T00:00:00Z\"/></w:date>", "2026-06-12", "value 2026-07-01T00:00:00Z\ndisplay-text 2026-07-01")]
    public static void CheckTrackChangesRequireRejectsContentControlStateUpdatesAsPreserveOnly(
        string operationName,
        string kindXml,
        string contentText,
        string operationFields)
    {
        using MemoryStream input = CreateDocxWithBody($"""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          {kindXml}
                          <w:tag w:val="state"/>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>{contentText}</w:t></w:r></w:sdtContent>
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

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6001");
        Assert.Contains($"operation '{operationName}'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("catalog support is 'preserve-only'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-no-revision-representation", diagnostic.Feature);
        Assert.Equal("require-failed", diagnostic.Fallback);
    }

    [Theory]
    [InlineData("set-content-control-checkbox", "<w:checkBox><w:checked w:val=\"0\"/><w:checkedState w:val=\"2612\"/><w:uncheckedState w:val=\"2610\"/></w:checkBox>", "Unchecked", "checked true")]
    [InlineData("set-content-control-choice", "<w:dropDownList><w:listItem w:displayText=\"North\" w:value=\"north\"/><w:listItem w:displayText=\"South\" w:value=\"south\"/></w:dropDownList>", "North", "value south")]
    [InlineData("set-content-control-date", "<w:date><w:dateFormat w:val=\"yyyy-MM-dd\"/><w:fullDate w:val=\"2026-06-12T00:00:00Z\"/></w:date>", "2026-06-12", "value 2026-07-01T00:00:00Z\ndisplay-text 2026-07-01")]
    public static void ApplyTrackChangesSuggestWarnsForContentControlStateUpdatesAsPreserveOnly(
        string operationName,
        string kindXml,
        string contentText,
        string operationFields)
    {
        using MemoryStream input = CreateDocxWithBody($"""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          {kindXml}
                          <w:tag w:val="state"/>
                        </w:sdtPr>
                        <w:sdtContent><w:r><w:t>{contentText}</w:t></w:r></w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            target M.CC0001
            {operationFields}
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "W4001");
        Assert.Contains($"operation '{operationName}'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("catalog support is 'preserve-only'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-no-revision-representation", diagnostic.Feature);
        Assert.Equal("direct-edit-preserve-existing-revisions", diagnostic.Fallback);
        Assert.Empty(Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:sdt>", xml, StringComparison.Ordinal);
        Assert.Contains("w:tag w:val=\"state\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:ins", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:del", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckTrackChangesRequireAllowsReplaceBookmarkTextSameParagraph()
    {
        using MemoryStream input = CreateDocxWithSingleBookmark();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-bookmark-text
            target M.B0001
            text New Client
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code is "E6001" or "E6002");
    }

    [Fact]
    public static void ApplyTrackChangesSuggestGeneratesRevisionMarkupForReplaceBookmarkText()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:bookmarkStart w:id="4" w:name="ClientName"/>
                      <w:r><w:t>Old </w:t></w:r>
                      <w:r><w:t>Client</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-bookmark-text
            target M.B0001
            text New Client
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code is "W4001" or "W4002");
        Assert.Equal(["5", "6"], Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("w:bookmarkStart w:id=\"4\" w:name=\"ClientName\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:bookmarkEnd w:id=\"4\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:delText>Old Client</w:delText>", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>New Client</w:t>", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Equal("Before New Client After", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        output.Position = 0;
        Assert.Equal("Before Old Client After", Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { TextView = DocxTextView.Original }).Paragraphs).Text);
        output.Position = 0;
        Assert.Equal("Before [-Old Client-][+New Client+] After", Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { TextView = DocxTextView.Markup }).Paragraphs).Text);
        output.Position = 0;
        DocxChangesResult changes = new DocxEditor().Changes(output);
        Assert.Contains(changes.Changes, change => change.Type == "deleted-run" && change.RevisionId == "5");
        Assert.Contains(changes.Changes, change => change.Type == "inserted-run" && change.RevisionId == "6");
    }

    [Fact]
    public static void CheckTrackChangesRequireRejectsReplaceBookmarkTextMultiParagraph()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="4" w:name="ClientName"/>
                      <w:r><w:t>Old first</w:t></w:r>
                    </w:p>
                    <w:p><w:r><w:t>Old middle</w:t></w:r></w:p>
                    <w:p>
                      <w:r><w:t>Old last</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-bookmark-text
            target M.B0001
            text New Client
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
        Assert.Contains("same-paragraph bookmark ranges", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-unsupported-target-shape", diagnostic.Feature);
        Assert.Equal("require-failed", diagnostic.Fallback);
    }

    [Theory]
    [InlineData("add-bookmark", """
        target M.P0002
        name AddedBookmark
        """)]
    [InlineData("rename-bookmark", """
        target M.B0001
        name NewBookmark
        """)]
    [InlineData("delete-bookmark", """
        target M.B0001
        """)]
    public static void CheckTrackChangesRequireRejectsBookmarkOperationsAsPreserveOnly(string operationName, string operationFields)
    {
        using MemoryStream input = CreateDocxWithSingleBookmark();
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            {operationFields}
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6001");
        Assert.Contains($"operation '{operationName}'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("catalog support is 'preserve-only'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-no-revision-representation", diagnostic.Feature);
        Assert.Equal("require-failed", diagnostic.Fallback);
    }

    [Theory]
    [InlineData("add-bookmark", """
        target M.P0002
        name AddedBookmark
        """, "w:name=\"AddedBookmark\"", null)]
    [InlineData("rename-bookmark", """
        target M.B0001
        name NewBookmark
        """, "w:name=\"NewBookmark\"", "w:name=\"ClientName\"")]
    [InlineData("delete-bookmark", """
        target M.B0001
        """, "Old Client", "bookmarkStart")]
    public static void ApplyTrackChangesSuggestWarnsForBookmarkOperationsAsPreserveOnly(
        string operationName,
        string operationFields,
        string expectedXml,
        string? unexpectedXml)
    {
        using MemoryStream input = CreateDocxWithSingleBookmark();
        using var output = new MemoryStream();
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            {operationFields}
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "W4001");
        Assert.Contains($"operation '{operationName}'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("catalog support is 'preserve-only'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-no-revision-representation", diagnostic.Feature);
        Assert.Equal("direct-edit-preserve-existing-revisions", diagnostic.Fallback);
        Assert.Empty(Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains(expectedXml, xml, StringComparison.Ordinal);
        if (unexpectedXml is not null)
        {
            Assert.DoesNotContain(unexpectedXml, xml, StringComparison.Ordinal);
        }

        Assert.DoesNotContain("<w:ins", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:del", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestTreatsCommentCreationAsPreserveOnly()
    {
        using MemoryStream input = CreateDocx("Anchor paragraph.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op add-comment
            target M.P0001
            text Review note
            author Reviewer
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "W4001");
        Assert.Contains("catalog support is 'preserve-only'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-no-revision-representation", diagnostic.Feature);
        Assert.Equal("direct-edit-preserve-existing-revisions", diagnostic.Fallback);
        Assert.Empty(Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        Assert.Contains("w:comment w:id=\"0\"", ReadEntry(output, "word/comments.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyTrackChangesRequirePermitsCommentCreationAsAnnotation()
    {
        using MemoryStream input = CreateDocx("Anchor paragraph.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op add-comment
            target M.P0001
            text Review note
            author Reviewer
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.DoesNotContain(result.Diagnostics, static diagnostic => diagnostic.Code == "E6001");
        Assert.True(Assert.Single(result.Operations).Success);
        Assert.Empty(Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        Assert.Contains("Review note", ReadEntry(output, "word/comments.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckTrackChangesRequireAllowsSetCommentText()
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
                  <w:comment w:id="3" w:author="Reviewer">
                    <w:p><w:r><w:t>Old comment</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-comment-text
            target comment:3
            text Updated comment
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code is "E6001" or "E6002");
    }

    [Fact]
    public static void CheckTrackChangesRequireRejectsProtectedSetCommentText()
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
                  <w:comment w:id="3" w:author="Reviewer">
                    <w:p>
                      <w:fldSimple w:instr=" DATE ">
                        <w:r><w:t>Old comment</w:t></w:r>
                      </w:fldSimple>
                    </w:p>
                  </w:comment>
                </w:comments>
                """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-comment-text
            target comment:3
            text Updated comment
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
        Assert.Contains("comment body contains field content", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-unsupported-target-shape", diagnostic.Feature);
        Assert.Equal("require-failed", diagnostic.Fallback);
    }

    [Theory]
    [InlineData("resolve-comment", "")]
    [InlineData("reopen-comment", "")]
    [InlineData("delete-comment", "")]
    public static void ApplyTrackChangesSuggestTreatsCommentMutationsAsPreserveOnly(string operationName, string extraFields)
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
                <w:comments
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w:comment w:id="3" w:author="Reviewer">
                    <w:p w15:paraId="00ABCDEF"><w:r><w:t>Old comment</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """,
            """
                <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w15:commentEx w15:paraId="00ABCDEF" w15:done="1"/>
                </w15:commentsEx>
                """, null);
        using var output = new MemoryStream();
        using var patch = new StringReader(BuildCommentPatch(operationName, extraFields));

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "W4001");
        Assert.Contains($"operation '{operationName}'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("catalog support is 'preserve-only'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-no-revision-representation", diagnostic.Feature);
        Assert.Equal("direct-edit-preserve-existing-revisions", diagnostic.Fallback);
        Assert.Empty(Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        Assert.DoesNotContain(
            new DocxEditor().Changes(output).Changes,
            change => change.Type is "insertion" or "deletion" or "paragraph-properties-change");

        output.Position = 0;
        switch (operationName)
        {
            case "resolve-comment":
                Assert.True(Assert.Single(new DocxEditor().Changes(output).CommentSummary).Resolved);
                break;
            case "reopen-comment":
                Assert.False(Assert.Single(new DocxEditor().Changes(output).CommentSummary).Resolved);
                break;
            case "delete-comment":
                Assert.Empty(new DocxEditor().Changes(output).Changes);
                break;
        }
    }

    [Theory]
    [InlineData("resolve-comment", "")]
    [InlineData("reopen-comment", "")]
    [InlineData("delete-comment", "")]
    public static void ApplyTrackChangesRequirePermitsCommentMutationsAsAnnotation(string operationName, string extraFields)
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
                  <w:comment w:id="3" w:author="Reviewer">
                    <w:p><w:r><w:t>Old comment</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader(BuildCommentPatch(operationName, extraFields));

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static d => d.Code + ":" + d.Message)));
        Assert.DoesNotContain(result.Diagnostics, static diagnostic => diagnostic.Code == "E6001");
        Assert.True(Assert.Single(result.Operations).Success);
        Assert.Empty(Assert.Single(result.Operations).GeneratedRevisionIds);
        Assert.NotEqual(0, output.Length);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestSetCommentTextKeepsDocumentMarkupAndTracksCommentBody()
    {
        using MemoryStream input = CreateDocxWithBodyAndComments(
            """
                    <w:p>
                      <w:ins w:id="9" w:author="Existing" w:date="2026-06-01T12:00:00Z">
                        <w:r><w:t>Existing insertion </w:t></w:r>
                      </w:ins>
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:commentRangeStart w:id="3"/>
                      <w:commentRangeEnd w:id="3"/>
                      <w:r><w:commentReference w:id="3"/></w:r>
                    </w:p>
            """,
            """
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="3" w:author="Reviewer">
                    <w:p><w:r><w:t>Old comment</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        string documentXmlBefore = ReadDocumentXml(input);
        input.Position = 0;
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-comment-text
            target comment:3
            text Updated comment
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code is "W4001" or "W4002");
        Assert.Equal(["10", "11"], Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        Assert.Equal(documentXmlBefore, ReadDocumentXml(output));
        output.Position = 0;
        string commentsXml = ReadEntry(output, "word/comments.xml");
        Assert.Contains("<w:delText>Old comment</w:delText>", commentsXml, StringComparison.Ordinal);
        Assert.Contains("Updated comment", commentsXml, StringComparison.Ordinal);
        Assert.Contains("w:id=\"10\"", commentsXml, StringComparison.Ordinal);
        Assert.Contains("w:id=\"11\"", commentsXml, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("set-table-metadata", """
        target M.T0001
        caption Updated caption
        """)]
    public static void CheckTrackChangesRequireRejectsTablePropertyOperationsAsPreserveOnly(string operationName, string operationFields)
    {
        using MemoryStream input = CreateDocxWithSimpleTwoByTwoTable();
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            {operationFields}
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6001");
        Assert.Contains($"operation '{operationName}'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("catalog support is 'preserve-only'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-no-revision-representation", diagnostic.Feature);
        Assert.Equal("require-failed", diagnostic.Fallback);
    }

    [Fact]
    public static void CheckTrackChangesRequireAllowsTableCellAndRowPropertyRevisions()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>
            """, """
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Profit</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-table-style
            target M.T0001
            style TableGrid
            end

            op set-cell-shading
            target M.T0001.R01.C01
            fill A1B2C3
            end

            op set-row-header
            target M.T0001.R01
            header true
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code is "E6001" or "E6002");
    }

    [Fact]
    public static void ApplyTrackChangesSuggestGeneratesTableCellAndRowPropertyRevisions()
    {
        using MemoryStream input = CreateDocxWithStylesAndBody("""
              <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>
            """, """
                    <w:tbl>
                      <w:tblPr>
                        <w:tblStyle w:val="OldStyle"/>
                      </w:tblPr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Profit</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-table-style
            target M.T0001
            expect-style OldStyle
            style TableGrid
            end

            op set-cell-shading
            target M.T0001.R01.C01
            expect-fill none
            fill A1B2C3
            end

            op set-row-header
            target M.T0001.R01
            expect-header false
            header true
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code is "W4001" or "W4002");
        Assert.Equal(["1"], result.Operations[0].GeneratedRevisionIds);
        Assert.Equal(["2"], result.Operations[1].GeneratedRevisionIds);
        Assert.Equal(["3"], result.Operations[2].GeneratedRevisionIds);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:tblStyle w:val=\"TableGrid\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:tblPrChange w:id=\"1\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:tblStyle w:val=\"OldStyle\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:shd w:val=\"clear\" w:fill=\"A1B2C3\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:tcPrChange w:id=\"2\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:tblHeader", xml, StringComparison.Ordinal);
        Assert.Contains("<w:trPrChange w:id=\"3\"", xml, StringComparison.Ordinal);
        output.Position = 0;
        DocxChangesResult changes = new DocxEditor().Changes(output);
        Assert.Contains(changes.Summary, summary => summary.Type == "table-properties-change" && summary.Count == 1);
        Assert.Contains(changes.Summary, summary => summary.Type == "cell-properties-change" && summary.Count == 1);
        Assert.Contains(changes.Summary, summary => summary.Type == "row-properties-change" && summary.Count == 1);
    }

    [Theory]
    [InlineData("set-table-metadata", """
        target M.T0001
        caption Updated caption
        """, "<w:tblCaption w:val=\"Updated caption\"")]
    public static void ApplyTrackChangesSuggestWarnsForTablePropertyOperationsAsPreserveOnly(
        string operationName,
        string operationFields,
        string expectedXml)
    {
        using MemoryStream input = CreateDocxWithSimpleTwoByTwoTable();
        using var output = new MemoryStream();
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            {operationFields}
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "W4001");
        Assert.Contains($"operation '{operationName}'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("catalog support is 'preserve-only'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-no-revision-representation", diagnostic.Feature);
        Assert.Equal("direct-edit-preserve-existing-revisions", diagnostic.Fallback);
        Assert.Empty(Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains(expectedXml, xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:ins", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:del", xml, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("replace-image", """
        target M.I0001
        asset chart.png
        """)]
    [InlineData("insert-image-after", """
        target M.P0001
        asset chart.png
        """)]
    [InlineData("set-image-alt", """
        target M.I0001
        alt Updated chart
        """)]
    [InlineData("set-image-metadata", """
        target M.I0001
        alt Updated chart
        title Revenue chart
        name Revenue picture
        """)]
    [InlineData("set-image-size", """
        target M.I0001
        width 2in
        """)]
    [InlineData("set-image-wrap", """
        target M.I0001
        mode top-bottom
        """)]
    [InlineData("set-image-position", """
        target M.I0001
        horizontal-relative page
        horizontal-offset 1in
        """)]
    [InlineData("set-image-crop", """
        target M.I0001
        left-percent 10
        """)]
    [InlineData("delete-image", """
        target M.I0001
        """)]
    public static void CheckTrackChangesRequireRejectsEveryImageOperationAsPreserveOnly(string operationName, string operationFields)
    {
        using MemoryStream input = CreateDocxWithAnchoredImage();
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            {operationFields}
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6001");
        Assert.Contains($"operation '{operationName}'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("catalog support is 'preserve-only'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-no-revision-representation", diagnostic.Feature);
        Assert.Equal("require-failed", diagnostic.Fallback);
    }

    [Theory]
    [InlineData("replace-image", """
        target M.I0001
        asset chart.png
        """)]
    [InlineData("insert-image-after", """
        target M.P0001
        asset chart.png
        """)]
    [InlineData("set-image-alt", """
        target M.I0001
        alt Updated chart
        """)]
    [InlineData("set-image-metadata", """
        target M.I0001
        alt Updated chart
        title Revenue chart
        name Revenue picture
        """)]
    [InlineData("set-image-size", """
        target M.I0001
        width 2in
        """)]
    [InlineData("set-image-wrap", """
        target M.I0001
        mode top-bottom
        """)]
    [InlineData("set-image-position", """
        target M.I0001
        horizontal-relative page
        horizontal-offset 1in
        """)]
    [InlineData("set-image-crop", """
        target M.I0001
        left-percent 10
        """)]
    [InlineData("delete-image", """
        target M.I0001
        """)]
    public static void ApplyTrackChangesSuggestWarnsForEveryImageOperationAsPreserveOnly(string operationName, string operationFields)
    {
        using MemoryStream input = CreateDocxWithAnchoredImage();
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", CreatePngBytes(4, 3), null, "chart.png");
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            {operationFields}
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(
            input,
            patch,
            output,
            new DocxEditOptions
            {
                TrackChanges = TrackChangesMode.Suggest,
                AssetProvider = assets,
                MarkFieldsDirtyWhenEditing = false
            });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "W4001");
        Assert.Contains($"operation '{operationName}'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("catalog support is 'preserve-only'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("track-changes-no-revision-representation", diagnostic.Feature);
        Assert.Equal("direct-edit-preserve-existing-revisions", diagnostic.Fallback);
        Assert.Empty(Assert.Single(result.Operations).GeneratedRevisionIds);
        Assert.True(output.Length > 0);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestImageOperationKeepsUnrelatedRevisionMarkup()
    {
        using MemoryStream input = CreateDocxWithRevisionAndImage();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-alt
            target M.I0001
            alt Updated chart
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
        {
            TrackChanges = TrackChangesMode.Suggest,
            MarkFieldsDirtyWhenEditing = false
        });

        Assert.True(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W4001");
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("descr=\"Updated chart\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:id=\"9\" w:author=\"Existing\"", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains(new DocxEditor().Changes(output).Changes, change => change.Type == "inserted-run" && change.RevisionId == "9");
    }

    [Theory]
    [InlineData("append-row", """
        target M.T0001
        cell East
        cell Margin
        """)]
    [InlineData("insert-row-before", """
        target M.T0001.R02
        cell East
        cell Margin
        """)]
    [InlineData("insert-row-after", """
        target M.T0001.R01
        cell East
        cell Margin
        """)]
    [InlineData("delete-row", """
        target M.T0001.R02
        """)]
    public static void CheckTrackChangesRequireAllowsRowStructureRevisions(string operationName, string operationFields)
    {
        using MemoryStream input = CreateDocxWithSimpleTwoByTwoTable();
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            {operationFields}
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code is "E6001" or "E6002");
    }

    [Theory]
    [InlineData("append-row", """
        target M.T0001
        cell East
        cell Margin
        """, 3, 3, "<w:ins w:id=\"1\"", "row-inserted", "[+East+]")]
    [InlineData("insert-row-before", """
        target M.T0001.R02
        cell East
        cell Margin
        """, 3, 3, "<w:ins w:id=\"1\"", "row-inserted", "[+East+]")]
    [InlineData("insert-row-after", """
        target M.T0001.R01
        cell East
        cell Margin
        """, 3, 3, "<w:ins w:id=\"1\"", "row-inserted", "[+East+]")]
    [InlineData("delete-row", """
        target M.T0001.R02
        """, 1, 2, "<w:del w:id=\"1\"", "row-deleted", "[-South-]")]
    public static void ApplyTrackChangesSuggestGeneratesRowStructureRevisions(
        string operationName,
        string operationFields,
        int expectedFinalRows,
        int expectedMarkupRows,
        string expectedRevisionXml,
        string expectedChangeType,
        string expectedMarkupText)
    {
        using MemoryStream input = CreateDocxWithSimpleTwoByTwoTable();
        using var output = new MemoryStream();
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            {operationFields}
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code is "W4001" or "W4002");
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.Equal(["1"], report.GeneratedRevisionIds);
        Assert.NotEmpty(report.AffectedTargets);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal(expectedFinalRows, table.RowCount);
        output.Position = 0;
        DocxTableInfo originalTable = Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { TextView = DocxTextView.Original }).Tables);
        Assert.Equal(2, originalTable.RowCount);
        output.Position = 0;
        DocxTableInfo markupTable = Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { TextView = DocxTextView.Markup }).Tables);
        Assert.Equal(expectedMarkupRows, markupTable.RowCount);
        Assert.Contains(markupTable.Cells, cell => cell.Text == expectedMarkupText);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains(expectedRevisionXml, xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains(new DocxEditor().Changes(output).Summary, summary => summary.Type == expectedChangeType && summary.Count == 1);
    }

    [Theory]
    [InlineData("insert-row-before", """
        target M.T0001.R02
        force true
        cell East
        """)]
    [InlineData("delete-row", """
        target M.T0001.R01
        force true
        """)]
    public static void CheckTrackChangesRequireRejectsForcedRowStructureRevisions(string operationName, string operationFields)
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
                          <w:p><w:r><w:t>South total</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            {operationFields}
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
        Assert.Contains("tracked row operations do not support force true", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckTrackChangesRequireRejectsVisualGridRowStructureRevisions()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
                          <w:p><w:r><w:t>North total</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Profit</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op insert-row-before
            target M.T0001.R01
            cell Inserted total
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { TrackChanges = TrackChangesMode.Require });

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
        Assert.Contains("tracked row operations require a simple rectangular table", diagnostic.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("insert-row-before", """
        target M.T0001.R02
        force true
        cell East
        """)]
    [InlineData("delete-row", """
        target M.T0001.R01
        force true
        """)]
    public static void ApplyTrackChangesSuggestFallsBackForForcedRowStructureRevisions(string operationName, string operationFields)
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
                          <w:p><w:r><w:t>South total</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            {operationFields}
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "W4002");
        Assert.Contains("tracked row operations do not support force true", diagnostic.Message, StringComparison.Ordinal);
        Assert.Empty(Assert.Single(result.Operations).GeneratedRevisionIds);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.DoesNotContain("<w:ins w:id=", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:del w:id=", xml, StringComparison.Ordinal);
    }
}
