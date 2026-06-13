using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace DocxEdit.Tests;

public static class PatchApplyTests
{
    [Fact]
    public static void CheckReplaceTextSucceedsForMatchingParagraph()
    {
        using MemoryStream input = CreateDocx("Revenue increased by 8.4%.");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            expect-text <<<
            Revenue increased by 8.4%.
            >>>
            find <<<
            8.4%
            >>>
            with <<<
            9.1%
            >>>
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success);
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        Assert.True(report.Success);
    }

    [Fact]
    public static void CheckReplaceTextFailsWhenGuardDoesNotMatch()
    {
        using MemoryStream input = CreateDocx("Revenue increased by 9.1%.");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            expect-text <<<
            Revenue increased by 8.4%.
            >>>
            find <<<
            8.4%
            >>>
            with <<<
            9.1%
            >>>
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E3201");
        Assert.False(Assert.Single(result.Operations).Success);
    }

    [Fact]
    public static void ApplyReplaceTextWritesNewDocumentWithoutMutatingInput()
    {
        using MemoryStream input = CreateDocx("Revenue increased by 8.4%.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            expect-text <<<
            Revenue increased by 8.4%.
            >>>
            find <<<
            8.4%
            >>>
            with <<<
            9.1%
            >>>
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        input.Position = 0;
        output.Position = 0;
        Assert.Equal("Revenue increased by 8.4%.", Assert.Single(new DocxEditor().Read(input).Paragraphs).Text);
        Assert.Equal("Revenue increased by 9.1%.", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void ApplyWorksWithMemoryStreamInputAndOutput()
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

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("Revenue rose.", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void ApplyMarksFieldsDirtyByDefault()
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

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        string settings = ReadEntry(output, "word/settings.xml");
        Assert.Contains("updateFields", settings, StringComparison.Ordinal);
        Assert.Contains("w:val=\"true\"", settings, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains("relationships/settings", ReadEntry(output, "word/_rels/document.xml.rels"), StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains("wordprocessingml.settings+xml", ReadEntry(output, "[Content_Types].xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyWarnsWhenFieldsRequireWordSideRefresh()
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

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "W5103" &&
            diagnostic.Feature == "field" &&
            diagnostic.Fallback == "word-refresh-required");
        output.Position = 0;
        Assert.Contains("updateFields", ReadEntry(output, "word/settings.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyCanSkipMarkingFieldsDirty()
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

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { MarkFieldsDirtyWhenEditing = false });

        Assert.True(result.Success);
        output.Position = 0;
        Assert.False(EntryExists(output, "word/settings.xml"));
    }

    [Fact]
    public static void ApplySetFieldFlagsUpdatesSimpleAndComplexFields()
    {
        using MemoryStream input = CreateDocxWithBody("""
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
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-field-dirty
            target M.F0001
            dirty true
            end

            op set-field-lock
            target M.F0002
            locked true
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { MarkFieldsDirtyWhenEditing = false });

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.True(read.Fields[0].IsDirty);
        Assert.False(read.Fields[0].IsLocked);
        Assert.False(read.Fields[1].IsDirty);
        Assert.True(read.Fields[1].IsLocked);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:fldSimple w:instr=\" DATE \" w:dirty=\"true\" w:fldLock=\"0\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:fldChar w:fldCharType=\"begin\" w:dirty=\"false\" w:fldLock=\"true\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetFieldFlagsCanTargetAllFields()
    {
        using MemoryStream input = CreateDocxWithBody("""
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
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-field-lock
            target all
            locked true
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { MarkFieldsDirtyWhenEditing = false });

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.All(read.Fields, field => Assert.True(field.IsLocked));
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:fldSimple w:instr=\" DATE \" w:dirty=\"false\" w:fldLock=\"true\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:fldChar w:fldCharType=\"begin\" w:dirty=\"false\" w:fldLock=\"true\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetSimpleFieldCodeMarksFieldDirty()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:fldSimple w:instr=" REF OldBookmark \h " w:dirty="false">
                        <w:r><w:t>Old result</w:t></w:r>
                      </w:fldSimple>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-field-code
            target M.F0001
            expect-code REF OldBookmark \h
            code REF NewBookmark \h
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { MarkFieldsDirtyWhenEditing = false });

        Assert.True(result.Success);
        output.Position = 0;
        DocxFieldInfo field = Assert.Single(new DocxEditor().Read(output).Fields);
        Assert.Equal("REF NewBookmark \\h", field.Code);
        Assert.True(field.IsDirty);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("w:instr=\"REF NewBookmark \\h\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:dirty=\"true\"", xml, StringComparison.Ordinal);
        Assert.Contains("Old result", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetSimpleFieldResultPreservesCodeAndDoesNotMarkFieldsDirty()
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

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("New cached result", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        output.Position = 0;
        DocxFieldInfo field = Assert.Single(new DocxEditor().Read(output).Fields);
        Assert.Equal("REF ClientName \\h", field.Code);
        Assert.Equal("New cached result", field.CachedResultText);
        Assert.Equal(17, field.ResultTextLength);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("w:instr=\" REF ClientName \\h \"", xml, StringComparison.Ordinal);
        Assert.Contains("New cached result", xml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.False(EntryExists(output, "word/settings.xml"));
    }

    [Fact]
    public static void ApplyRefreshSimpleRefFieldResultFromBookmark()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="1" w:name="ClientName"/>
                      <w:r><w:t>Acme Corp</w:t></w:r>
                      <w:bookmarkEnd w:id="1"/>
                    </w:p>
                    <w:p>
                      <w:fldSimple w:instr=" REF ClientName \h " w:dirty="true">
                        <w:r><w:t>Old cached result</w:t></w:r>
                      </w:fldSimple>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op refresh-field-result
            target M.F0001
            expect-result Old cached result
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        DocxFieldInfo field = Assert.Single(read.Fields);
        Assert.Equal("REF ClientName \\h", field.Code);
        Assert.Equal("ClientName", Assert.Single(field.BookmarkDependencies));
        Assert.Equal("Acme Corp", field.CachedResultText);
        Assert.Equal(9, field.ResultTextLength);
        Assert.Null(field.IsDirty);
        Assert.Equal("Acme Corp", read.Paragraphs[1].Text);
        output.Position = 0;
        Assert.False(EntryExists(output, "word/settings.xml"));
    }

    [Fact]
    public static void CheckRefreshFieldResultRejectsAmbiguousBookmark()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="1" w:name="ClientName"/>
                      <w:r><w:t>Acme</w:t></w:r>
                      <w:bookmarkEnd w:id="1"/>
                    </w:p>
                    <w:p>
                      <w:bookmarkStart w:id="2" w:name="ClientName"/>
                      <w:r><w:t>Contoso</w:t></w:r>
                      <w:bookmarkEnd w:id="2"/>
                    </w:p>
                    <w:p>
                      <w:fldSimple w:instr=" REF ClientName \h ">
                        <w:r><w:t>Old cached result</w:t></w:r>
                      </w:fldSimple>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op refresh-field-result
            target M.F0001
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E4313" &&
            diagnostic.Message.Contains("ambiguous", StringComparison.Ordinal));
    }

    [Fact]
    public static void CheckSetFieldResultRejectsComplexFields()
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

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E4313");
    }

    [Fact]
    public static void CheckSetFieldCodeRejectsMismatchedGuard()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:fldSimple w:instr=" DATE ">
                        <w:r><w:t>June 12</w:t></w:r>
                      </w:fldSimple>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-field-code
            target M.F0001
            expect-code REF Missing
            code REF Present
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E3201");
    }

    [Fact]
    public static void ReadWorksWithNonSeekableInputStream()
    {
        using MemoryStream seekable = CreateDocx("Revenue increased.");
        using var nonSeekable = new NonSeekableReadStream(seekable.ToArray());

        DocxReadResult result = new DocxEditor().Read(nonSeekable);

        Assert.True(result.Success);
        Assert.Equal("Revenue increased.", Assert.Single(result.Paragraphs).Text);
    }

    [Fact]
    public static void ApplyReplaceTextWorksAcrossAdjacentRuns()
    {
        using MemoryStream input = CreateDocxWithRuns("Revenue ", "increased");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            expect-text <<<
            Revenue increased
            >>>
            find <<<
            Revenue increased
            >>>
            with <<<
            Revenue rose
            >>>
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("Revenue rose", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
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
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E6002");
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

    [Fact]
    public static void CheckTrackChangesRequireRejectsMixedRunFormatting()
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

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E6002" && diagnostic.Message.Contains("mixed direct run formatting", StringComparison.Ordinal));
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
        DocxChangesResult changes = new DocxEditor().Changes(output);
        DocxChangeInfo deletion = Assert.Single(changes.Changes, change => change.Type == "deleted-run");
        DocxChangeInfo insertion = Assert.Single(changes.Changes, change => change.Type == "inserted-run");
        Assert.Equal("Reviewer", deletion.Author);
        Assert.Equal(timestamp, deletion.TimestampUtc);
        Assert.Equal("1", deletion.RevisionId);
        Assert.Equal("2", insertion.RevisionId);
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
    public static void ApplyTrackChangesSuggestFallsBackForMixedRunFormatting()
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
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "W4002");
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.DoesNotContain("<w:del ", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:ins ", xml, StringComparison.Ordinal);
        Assert.Contains("<w:i", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>rose</w:t>", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyTrackChangesSuggestWarnsAndAppliesUnsupportedOperationsDirectly()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:sdt>
                        <w:sdtPr>
                          <w:text/>
                          <w:tag w:val="client-name"/>
                        </w:sdtPr>
                        <w:sdtContent>
                          <w:r><w:t>Client</w:t></w:r>
                        </w:sdtContent>
                      </w:sdt>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-content-control-text
            target M.CC0001
            text Customer
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { TrackChanges = TrackChangesMode.Suggest });

        Assert.True(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "W4001" && diagnostic.Severity == DocxSeverity.Warning);
        Assert.Contains("catalog support is 'preserve-only'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("apply operation 'set-content-control-text' directly", diagnostic.Message, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Equal("Customer", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
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
    }

    [Fact]
    public static void ApplyReplaceTextFailsWhenFindTextIsMissing()
    {
        using MemoryStream input = CreateDocx("Revenue increased by 8.4%.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find missing
            with value
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E4203");
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public static void ApplyReplaceTextPreservesXmlSpaceWhenReplacementRequiresIt()
    {
        using MemoryStream input = CreateDocx("Revenue increased by 8.4%.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find 8.4%
            with <<<
             9.1% 
            >>>
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("Revenue increased by  9.1% .", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        output.Position = 0;
        Assert.Contains("xml:space=\"preserve\"", ReadDocumentXml(output), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyReplaceTextHonorsOccurrence()
    {
        using MemoryStream input = CreateDocx("foo foo foo");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find foo
            with bar
            occurrence 2
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("foo bar foo", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void ApplyReplaceTextPreservesTargetRunProperties()
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

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("Revenue rose", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:b", xml, StringComparison.Ordinal);
        Assert.Contains("<w:i", xml, StringComparison.Ordinal);
        Assert.Contains("<w:t>rose</w:t>", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckReplaceTextRejectsProtectedBoundaries()
    {
        using MemoryStream hyperlinkInput = CreateDocxWithBody("""
                    <w:p>
                      <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rLink">
                        <w:r><w:t>Link text</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
            """);
        using var hyperlinkPatch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Link
            with Anchor
            end
            """);

        DocxCheckResult hyperlinkResult = new DocxEditor().Check(hyperlinkInput, hyperlinkPatch);

        Assert.False(hyperlinkResult.Success);
        Assert.Contains(hyperlinkResult.Diagnostics, diagnostic => diagnostic.Code == "E4305");

        using MemoryStream revisionInput = CreateDocxWithBody("""
                    <w:p>
                      <w:ins w:id="1" w:author="A">
                        <w:r><w:t>Inserted</w:t></w:r>
                      </w:ins>
                    </w:p>
            """);
        using var revisionPatch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Inserted
            with Edited
            end
            """);

        DocxCheckResult revisionResult = new DocxEditor().Check(revisionInput, revisionPatch);

        Assert.False(revisionResult.Success);
        Assert.Contains(revisionResult.Diagnostics, diagnostic => diagnostic.Code == "E4305");
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
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E4305");
    }

    [Fact]
    public static void ApplyPreservesUnknownPartsAndUnrelatedMedia()
    {
        using MemoryStream input = CreateDocxWithBody(
            """
                    <w:p><w:r><w:t>Before</w:t></w:r></w:p>
            """,
            archive =>
            {
                AddEntry(archive, "custom/data.bin", "opaque");
                AddEntry(archive, "word/media/unrelated.png", "unrelated");
            });
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Before
            with After
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("opaque", ReadEntry(output, "custom/data.bin"));
        output.Position = 0;
        Assert.Equal("unrelated", ReadEntry(output, "word/media/unrelated.png"));
    }

    [Fact]
    public static void ApplyReplaceParagraphUpdatesTextAndStyle()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
                      <w:r><w:t>Old heading</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-paragraph
            target M.P0001
            expect-text Old heading
            style Heading2
            text New heading
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("New heading", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        output.Position = 0;
        Assert.Contains("w:val=\"Heading2\"", ReadDocumentXml(output), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyReplaceTextCanTargetHeadingByTextAndLevel()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
                      <w:r><w:t>Old heading</w:t></w:r>
                    </w:p>
                    <w:p><w:r><w:t>Body</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target heading:2:"Old heading"
            find Old
            with New
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("New heading", new DocxEditor().Read(output).Paragraphs[0].Text);
    }

    [Fact]
    public static void ApplyInsertAfterCanTargetParagraphContainingText()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
                    <w:p><w:r><w:t>Needle paragraph</w:t></w:r></w:p>
                    <w:p><w:r><w:t>Omega</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target text:"Needle"
            text Inserted
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal(
            new[] { "Alpha", "Needle paragraph", "Inserted", "Omega" },
            new DocxEditor().Read(output).Paragraphs.Select(paragraph => paragraph.Text));
    }

    [Fact]
    public static void ApplyTextEditAndInsertAfterPreserveListParagraphProperties()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:pPr>
                        <w:pStyle w:val="ListParagraph"/>
                        <w:numPr>
                          <w:ilvl w:val="1"/>
                          <w:numId w:val="9"/>
                        </w:numPr>
                      </w:pPr>
                      <w:r><w:t>First item</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find First
            with Updated
            end

            op insert-after
            target M.P0001
            copy-paragraph-properties true
            text Second item
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal(new[] { "Updated item", "Second item" }, new DocxEditor().Read(output).Paragraphs.Select(paragraph => paragraph.Text));
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Equal(2, CountOccurrences(xml, "w:pStyle w:val=\"ListParagraph\""));
        Assert.Equal(2, CountOccurrences(xml, "w:numId w:val=\"9\""));
        Assert.Equal(2, CountOccurrences(xml, "w:ilvl w:val=\"1\""));
    }

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

    [Theory]
    [InlineData("../relative/report", "relative hyperlink targets are not supported")]
    [InlineData("file:///C:/secret/report.docx", "Unsupported hyperlink URI scheme 'file'")]
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
    public static void CheckTextSelectorRejectsAmbiguousMatches()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>Revenue north</w:t></w:r></w:p>
                    <w:p><w:r><w:t>Revenue south</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target text:Revenue
            find Revenue
            with Sales
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E1202");
    }

    [Fact]
    public static void CheckTextSelectorSuggestsNearbyTargetsWithoutLeakingText()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
                      <w:r><w:t>Private heading text</w:t></w:r>
                    </w:p>
                    <w:p><w:r><w:t>Private paragraph text</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target text:"Missing"
            find Missing
            with Updated
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal("E1201", diagnostic.Code);
        Assert.Contains("Nearby paragraphs: M.P0001, M.P0002", diagnostic.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("Private", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyInsertAfterCanTargetBookmarkSelectorAndPreserveBookmark()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="1" w:name="ReviewPoint"/>
                      <w:r><w:t>Bookmarked paragraph</w:t></w:r>
                      <w:bookmarkEnd w:id="1"/>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target bookmark:"ReviewPoint"
            text Inserted
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal(new[] { "Bookmarked paragraph", "Inserted" }, new DocxEditor().Read(output).Paragraphs.Select(paragraph => paragraph.Text));
        output.Position = 0;
        Assert.Contains("w:name=\"ReviewPoint\"", ReadDocumentXml(output), StringComparison.Ordinal);
    }

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

    [Fact]
    public static void CheckRepeatingSectionItemOperationsFailWithExplicitUnsupportedDiagnostic()
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
        using var patch = new StringReader("""
            docxpatch 1

            op add-repeating-section-item
            target M.CC0001
            index 1
            text Added
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
    public static void ApplyAddBookmarkWrapsParagraphContent()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:pPr><w:pStyle w:val="BodyText"/></w:pPr>
                      <w:r><w:t>Client paragraph</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op add-bookmark
            target M.P0001
            expect-text Client paragraph
            name ClientParagraph
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        DocxBookmarkInfo bookmark = Assert.Single(read.Bookmarks);
        Assert.Equal("ClientParagraph", bookmark.Name);
        Assert.Equal("M.P0001", bookmark.StartTargetId);
        Assert.Equal("M.P0001", bookmark.EndTargetId);
        Assert.Equal("Client paragraph", Assert.Single(read.Paragraphs).Text);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:pPr><w:pStyle w:val=\"BodyText\" /></w:pPr><w:bookmarkStart", xml, StringComparison.Ordinal);
        Assert.Contains("w:name=\"ClientParagraph\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:bookmarkEnd", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckAddBookmarkRejectsDuplicateName()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="4" w:name="ClientParagraph"/>
                      <w:r><w:t>First</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                    </w:p>
                    <w:p><w:r><w:t>Second</w:t></w:r></w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op add-bookmark
            target M.P0002
            name ClientParagraph
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E4311" &&
            diagnostic.Message.Contains("already exists", StringComparison.Ordinal));
    }

    [Fact]
    public static void ApplyReplaceBookmarkTextPreservesMarkers()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:bookmarkStart w:id="4" w:name="ClientName"/>
                      <w:r><w:t>Old Client</w:t></w:r>
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

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("Before New Client After", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("w:bookmarkStart w:id=\"4\" w:name=\"ClientName\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:bookmarkEnd w:id=\"4\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("Old Client", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyReplaceBookmarkTextHandlesMultiRunSameParagraphRange()
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

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("Before New Client After", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("w:bookmarkStart w:id=\"4\" w:name=\"ClientName\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:bookmarkEnd w:id=\"4\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("Old ", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:r><w:t>Client</w:t></w:r>", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyRenameBookmarkUpdatesNameAndInternalHyperlinkAnchors()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="4" w:name="OldName"/>
                      <w:r><w:t>Target</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                    </w:p>
                    <w:p>
                      <w:hyperlink w:anchor="OldName"><w:r><w:t>Jump</w:t></w:r></w:hyperlink>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op rename-bookmark
            target M.B0001
            name NewName
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        DocxBookmarkInfo bookmark = Assert.Single(read.Bookmarks);
        Assert.Equal("NewName", bookmark.Name);
        DocxHyperlinkInfo hyperlink = Assert.Single(read.Hyperlinks);
        Assert.Equal("NewName", hyperlink.Anchor);
        Assert.False(hyperlink.IsAnchorMissing);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("w:name=\"NewName\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:anchor=\"NewName\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("OldName", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyDeleteBookmarkRemovesMarkersAndPreservesText()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:bookmarkStart w:id="4" w:name="ClientName"/>
                      <w:r><w:t>Bookmarked</w:t></w:r>
                      <w:bookmarkEnd w:id="4"/>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-bookmark
            target M.B0001
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Empty(read.Bookmarks);
        Assert.Equal("Before Bookmarked After", Assert.Single(read.Paragraphs).Text);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.DoesNotContain("bookmarkStart", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("bookmarkEnd", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckRenameBookmarkRejectsDuplicateName()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:bookmarkStart w:id="1" w:name="First"/>
                      <w:r><w:t>First</w:t></w:r>
                      <w:bookmarkEnd w:id="1"/>
                    </w:p>
                    <w:p>
                      <w:bookmarkStart w:id="2" w:name="Second"/>
                      <w:r><w:t>Second</w:t></w:r>
                      <w:bookmarkEnd w:id="2"/>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op rename-bookmark
            target M.B0001
            name Second
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E4311");
    }

    [Fact]
    public static void ApplyAddCommentCreatesCommentsPartAndAnchorsParagraph()
    {
        using MemoryStream input = CreateDocx("Anchor paragraph.");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op add-comment
            target M.P0001
            expect-text Anchor paragraph.
            text Review note
            author Reviewer
            initials RV
            date 2026-06-07T12:00:00Z
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxChangeInfo comment = Assert.Single(
            new DocxEditor().Changes(output, new DocxChangesOptions { IncludeCommentText = true }).Changes,
            change => change.Type == "comment");
        Assert.Equal("0", comment.CommentId);
        Assert.Equal("Reviewer", comment.CommentAuthor);
        Assert.Equal("RV", comment.CommentInitials);
        Assert.Equal("Review note", comment.CommentTextSnippet);
        Assert.Equal("M.P0001", comment.CommentAnchorTargetId);
        Assert.Equal("M.P0001", comment.CommentReferenceTargetId);

        output.Position = 0;
        string documentXml = ReadDocumentXml(output);
        Assert.Contains("<w:commentRangeStart w:id=\"0\"", documentXml, StringComparison.Ordinal);
        Assert.Contains("<w:commentRangeEnd w:id=\"0\"", documentXml, StringComparison.Ordinal);
        Assert.Contains("<w:commentReference w:id=\"0\"", documentXml, StringComparison.Ordinal);

        output.Position = 0;
        string commentsXml = ReadEntry(output, "word/comments.xml");
        Assert.Contains("w:comment w:id=\"0\"", commentsXml, StringComparison.Ordinal);
        Assert.Contains("w:author=\"Reviewer\"", commentsXml, StringComparison.Ordinal);
        Assert.Contains("Review note", commentsXml, StringComparison.Ordinal);

        output.Position = 0;
        Assert.Contains("relationships/comments", ReadEntry(output, "word/_rels/document.xml.rels"), StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains("wordprocessingml.comments+xml", ReadEntry(output, "[Content_Types].xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyAddCommentAllocatesNextExistingCommentId()
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
                    <w:p><w:r><w:t>Existing comment</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op add-comment
            target M.P0001
            text Follow-up note
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(
            input,
            patch,
            output,
            new DocxEditOptions
            {
                Author = "Second Reviewer",
                TimestampUtc = DateTimeOffset.Parse("2026-06-08T09:30:00Z").ToUniversalTime()
            });

        Assert.True(result.Success);
        output.Position = 0;
        IReadOnlyList<DocxCommentThreadSummary> summaries = new DocxEditor()
            .Changes(output, new DocxChangesOptions { IncludeCommentText = true })
            .CommentSummary;
        Assert.Contains(summaries, summary => summary.CommentId == "3" && summary.TextSnippet == "Existing comment");
        Assert.Contains(summaries, summary => summary.CommentId == "4" && summary.TextSnippet == "Follow-up note");

        output.Position = 0;
        string commentsXml = ReadEntry(output, "word/comments.xml");
        Assert.Contains("w:comment w:id=\"3\"", commentsXml, StringComparison.Ordinal);
        Assert.Contains("w:comment w:id=\"4\"", commentsXml, StringComparison.Ordinal);
        Assert.Contains("w:author=\"Second Reviewer\"", commentsXml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetCommentTextPreservesCommentMetadata()
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

            op set-comment-text
            target comment:3
            text Updated comment
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxChangeInfo comment = Assert.Single(
            new DocxEditor().Changes(output, new DocxChangesOptions { IncludeCommentText = true }).Changes,
            change => change.Type == "comment");
        Assert.Equal("3", comment.CommentId);
        Assert.Equal("Reviewer", comment.CommentAuthor);
        Assert.Equal("RV", comment.CommentInitials);
        Assert.Equal("Updated comment", comment.CommentTextSnippet);
        output.Position = 0;
        string commentsXml = ReadEntry(output, "word/comments.xml");
        Assert.Contains("w:author=\"Reviewer\"", commentsXml, StringComparison.Ordinal);
        Assert.Contains("Updated comment", commentsXml, StringComparison.Ordinal);
        Assert.DoesNotContain("Old comment", commentsXml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyResolveAndReopenCommentUpdatesExtendedMetadata()
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
                    <w:p w15:paraId="00ABCDEF"><w:r><w:t>Comment body</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """,
            """
                <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w15:commentEx w15:paraId="00ABCDEF" w15:done="0"/>
                </w15:commentsEx>
                """);
        using var resolvedOutput = new MemoryStream();
        using var resolvePatch = new StringReader("""
            docxpatch 1

            op resolve-comment
            target comment:3
            end
            """);

        DocxApplyResult resolveResult = new DocxEditor().Apply(input, resolvePatch, resolvedOutput);

        Assert.True(resolveResult.Success);
        resolvedOutput.Position = 0;
        Assert.True(Assert.Single(new DocxEditor().Changes(resolvedOutput).CommentSummary).Resolved);
        resolvedOutput.Position = 0;
        Assert.Contains("w15:done=\"1\"", ReadEntry(resolvedOutput, "word/commentsExtended.xml"), StringComparison.Ordinal);

        resolvedOutput.Position = 0;
        using var reopenedOutput = new MemoryStream();
        using var reopenPatch = new StringReader("""
            docxpatch 1

            op reopen-comment
            target C001.C0001
            end
            """);

        DocxApplyResult reopenResult = new DocxEditor().Apply(resolvedOutput, reopenPatch, reopenedOutput);

        Assert.True(reopenResult.Success);
        reopenedOutput.Position = 0;
        Assert.False(Assert.Single(new DocxEditor().Changes(reopenedOutput).CommentSummary).Resolved);
        reopenedOutput.Position = 0;
        Assert.Contains("w15:done=\"0\"", ReadEntry(reopenedOutput, "word/commentsExtended.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyResolveCommentCreatesExtendedMetadataWhenMissing()
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
                    <w:p><w:r><w:t>Comment body</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op resolve-comment
            target comment:3
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxCommentThreadSummary summary = Assert.Single(new DocxEditor().Changes(output).CommentSummary);
        Assert.Equal("3", summary.CommentId);
        Assert.True(summary.Resolved);
        Assert.Equal("00000001", summary.ParaId);

        output.Position = 0;
        string commentsXml = ReadEntry(output, "word/comments.xml");
        Assert.Contains("w15:paraId=\"00000001\"", commentsXml, StringComparison.Ordinal);

        output.Position = 0;
        string commentsExtendedXml = ReadEntry(output, "word/commentsExtended.xml");
        Assert.Contains("w15:commentEx w15:paraId=\"00000001\" w15:done=\"1\"", commentsExtendedXml, StringComparison.Ordinal);

        output.Position = 0;
        Assert.Contains("relationships/commentsExtended", ReadEntry(output, "word/_rels/document.xml.rels"), StringComparison.Ordinal);
        output.Position = 0;
        Assert.Contains("vnd.ms-word.commentsExtended+xml", ReadEntry(output, "[Content_Types].xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyReopenCommentCreatesUnresolvedExtendedMetadataWhenMissing()
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
                    <w:p><w:r><w:t>Comment body</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op reopen-comment
            target comment:3
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxCommentThreadSummary summary = Assert.Single(new DocxEditor().Changes(output).CommentSummary);
        Assert.False(summary.Resolved);
        output.Position = 0;
        Assert.Contains("w15:done=\"0\"", ReadEntry(output, "word/commentsExtended.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyDeleteCommentRemovesExtendedMetadata()
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
                    <w:p w15:paraId="00ABCDEF"><w:r><w:t>Comment body</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """,
            """
                <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w15:commentEx w15:paraId="00ABCDEF" w15:done="1"/>
                </w15:commentsEx>
                """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-comment
            target comment:3
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Empty(new DocxEditor().Changes(output).Changes);
        output.Position = 0;
        string commentsExtendedXml = ReadEntry(output, "word/commentsExtended.xml");
        Assert.DoesNotContain("00ABCDEF", commentsExtendedXml, StringComparison.Ordinal);
        Assert.DoesNotContain("commentEx", commentsExtendedXml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyDeleteCommentRemovesBodyAndAnchors()
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

            op delete-comment
            target C001.C0001
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Empty(new DocxEditor().Changes(output).Changes);
        output.Position = 0;
        string documentXml = ReadDocumentXml(output);
        Assert.DoesNotContain("commentRangeStart", documentXml, StringComparison.Ordinal);
        Assert.DoesNotContain("commentRangeEnd", documentXml, StringComparison.Ordinal);
        Assert.DoesNotContain("commentReference", documentXml, StringComparison.Ordinal);
        output.Position = 0;
        string commentsXml = ReadEntry(output, "word/comments.xml");
        Assert.DoesNotContain("w:comment w:id=\"3\"", commentsXml, StringComparison.Ordinal);
        output.Position = 0;
        Assert.Equal("Commented", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
    }

    [Fact]
    public static void ApplyDeleteCommentRemovesHeaderAnchorsAndExtendedMetadata()
    {
        using MemoryStream input = CreateDocxWithHeaderComment();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-comment
            target comment:3
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        string headerXml = ReadEntry(output, "word/header1.xml");
        Assert.DoesNotContain("commentRangeStart", headerXml, StringComparison.Ordinal);
        Assert.DoesNotContain("commentRangeEnd", headerXml, StringComparison.Ordinal);
        Assert.DoesNotContain("commentReference", headerXml, StringComparison.Ordinal);
        Assert.Contains("Header text", headerXml, StringComparison.Ordinal);
        output.Position = 0;
        string commentsExtendedXml = ReadEntry(output, "word/commentsExtended.xml");
        Assert.DoesNotContain("00ABCDEF", commentsExtendedXml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckCommentReplyOperationsFailWithExplicitUnsupportedDiagnostic()
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
                    <w:p><w:r><w:t>Comment body</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
        using var patch = new StringReader("""
            docxpatch 1

            op add-comment-reply
            target comment:3
            text Reply
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E4314" &&
            diagnostic.TargetId == "comment:3" &&
            diagnostic.Message.Contains("threaded comment replies", StringComparison.Ordinal));
    }

    [Fact]
    public static void ApplyInsertBeforeParagraphAddsParagraphAtTargetPosition()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>One</w:t></w:r></w:p>
                    <w:p><w:r><w:t>Two</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-before
            target M.P0002
            text Inserted
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal(new[] { "One", "Inserted", "Two" }, new DocxEditor().Read(output).Paragraphs.Select(paragraph => paragraph.Text));
    }

    [Fact]
    public static void ApplyInsertAfterTableAddsParagraphAfterTable()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>One</w:t></w:r></w:p>
                    <w:tbl>
                      <w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:p><w:r><w:t>Two</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target M.T0001
            style Normal
            text Inserted
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal(new[] { "One", "Inserted", "Two" }, read.Paragraphs.Select(paragraph => paragraph.Text));
        output.Position = 0;
        Assert.Contains("w:val=\"Normal\"", ReadDocumentXml(output), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyDeleteBlockRemovesParagraphOrTable()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>One</w:t></w:r></w:p>
                    <w:tbl>
                      <w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:p><w:r><w:t>Two</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-block
            target M.T0001
            end

            op delete-block
            target M.P0001
            expect-text One
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Empty(read.Tables);
        Assert.Equal("Two", Assert.Single(read.Paragraphs).Text);
    }

    [Fact]
    public static void ApplySetSectionColumnsAndOrientation()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>One</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:pgSz w:w="12240" w:h="15840"/>
                      <w:cols w:num="1"/>
                    </w:sectPr>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-section-columns
            target M.S0001
            count 2
            end

            op set-section-orientation
            target M.S0001
            orientation landscape
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxSectionInfo section = Assert.Single(new DocxEditor().Read(output).Sections);
        Assert.Equal(2, section.Columns);
        Assert.Equal("landscape", section.Orientation);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("w:num=\"2\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:orient=\"landscape\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:w=\"15840\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:h=\"12240\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckSetSectionRejectsFailedGuards()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>One</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:pgSz w:w="12240" w:h="15840"/>
                      <w:cols w:num="1"/>
                    </w:sectPr>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-section-orientation
            target M.S0001
            expect-columns 2
            orientation landscape
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E3201" && diagnostic.TargetId == "M.S0001");
    }

    [Fact]
    public static void ApplySetStyleResolvesParagraphStyleByDisplayName()
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
            style Heading 2
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Contains("w:val=\"Heading2\"", ReadDocumentXml(output), StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckSetStyleFailsWhenRequestedStyleIsMissing()
    {
        using MemoryStream input = CreateDocxWithStyles("""
              <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-style
            target M.P0001
            style Missing Style
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E7101");
    }

    [Fact]
    public static void CheckSetStyleFailsWhenDisplayNameIsAmbiguous()
    {
        using MemoryStream input = CreateDocxWithStyles("""
              <w:style w:type="paragraph" w:styleId="Custom1"><w:name w:val="Custom"/></w:style>
              <w:style w:type="paragraph" w:styleId="Custom2"><w:name w:val="Custom"/></w:style>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-style
            target M.P0001
            style Custom
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E7102");
    }

    [Fact]
    public static void ApplyReplaceTextCanEditHeaderParagraph()
    {
        using MemoryStream input = CreateDocxWithHeaderFooter("Header text", "Footer text");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target H001.P0001
            find Header
            with Confidential
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output, new DocxReadOptions { IncludeHeadersFooters = true });
        Assert.Contains(read.Paragraphs, paragraph => paragraph.Id == "H001.P0001" && paragraph.Text == "Confidential text");
    }

    [Fact]
    public static void ApplyReplaceParagraphCanEditFooterParagraph()
    {
        using MemoryStream input = CreateDocxWithHeaderFooter("Header text", "Footer text");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-paragraph
            target F001.P0001
            expect-text Footer text
            text Page 2
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output, new DocxReadOptions { IncludeHeadersFooters = true });
        Assert.Contains(read.Paragraphs, paragraph => paragraph.Id == "F001.P0001" && paragraph.Text == "Page 2");
    }

    [Fact]
    public static void ApplyBlockOperationsCanEditHeaderAndFooterStories()
    {
        using MemoryStream input = CreateDocxWithHeaderFooter("Header text", "Footer text");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-after
            target H001.P0001
            text Header detail
            end

            op insert-after
            target F001.P0001
            text Footer detail
            end

            op delete-block
            target F001.P0001
            expect-text Footer text
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output, new DocxReadOptions { IncludeHeadersFooters = true });
        Assert.Contains(read.Paragraphs, paragraph => paragraph.Id == "H001.P0001" && paragraph.Text == "Header text");
        Assert.Contains(read.Paragraphs, paragraph => paragraph.Id == "H001.P0002" && paragraph.Text == "Header detail");
        Assert.Contains(read.Paragraphs, paragraph => paragraph.Id == "F001.P0001" && paragraph.Text == "Footer detail");
        Assert.DoesNotContain(read.Paragraphs, paragraph => paragraph.Story == "footer[1]" && paragraph.Text == "Footer text");
    }

    [Fact]
    public static void ApplySetStyleCanEditHeaderParagraph()
    {
        using MemoryStream input = CreateDocxWithHeaderFooterAndStyles();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-style
            target H001.P0001
            style Heading 2
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Contains("w:val=\"Heading2\"", ReadEntry(output, "word/header1.xml"), StringComparison.Ordinal);
        output.Position = 0;
        DocxParagraphInfo paragraph = new DocxEditor().Read(output, new DocxReadOptions { IncludeHeadersFooters = true }).Paragraphs.Single(paragraph => paragraph.Id == "H001.P0001");
        Assert.Equal(2, paragraph.HeadingLevel);
    }

    [Fact]
    public static void ApplySetCellCanEditHeaderTableAndUseItAsBlockAnchor()
    {
        using MemoryStream input = CreateDocxWithHeaderFooterContent(
            """
                  <w:tbl>
                    <w:tr><w:tc><w:p><w:r><w:t>Old</w:t></w:r></w:p></w:tc></w:tr>
                  </w:tbl>
            """,
            """
                  <w:p><w:r><w:t>Footer text</w:t></w:r></w:p>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-cell
            target H001.T0001.R01.C01
            expect-text Old
            text New
            end

            op insert-after
            target H001.T0001
            text After header table
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output, new DocxReadOptions { IncludeHeadersFooters = true });
        DocxTableInfo table = Assert.Single(read.Tables, table => table.Story == "header[1]");
        Assert.Equal("New", Assert.Single(table.Cells).Text);
        Assert.Contains(read.Paragraphs, paragraph => paragraph.Id == "H001.P0001" && paragraph.Text == "After header table");
    }

    [Fact]
    public static void ApplyRowOperationsCanEditFooterTable()
    {
        using MemoryStream input = CreateDocxWithHeaderFooterContent(
            """
                  <w:p><w:r><w:t>Header text</w:t></w:r></w:p>
            """,
            """
                  <w:tbl>
                    <w:tr><w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc></w:tr>
                    <w:tr><w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc></w:tr>
                  </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op append-row
            target F001.T0001
            cell East
            end

            op insert-row-before
            target F001.T0001.R02
            cell Central
            end

            op delete-row
            target F001.T0001.R01
            expect-contains North
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { IncludeHeadersFooters = true }).Tables, table => table.Story == "footer[1]");
        Assert.Equal(3, table.RowCount);
        Assert.Equal("Central", table.Cells.Single(cell => cell.RowIndex == 1).Text);
        Assert.Equal("South", table.Cells.Single(cell => cell.RowIndex == 2).Text);
        Assert.Equal("East", table.Cells.Single(cell => cell.RowIndex == 3).Text);
    }

    [Fact]
    public static void ApplySetCellPreservesCellPropertiesAndDoesNotMutateInput()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>
                          <w:p>
                            <w:pPr><w:pStyle w:val="TableBody"/></w:pPr>
                            <w:r><w:t>Old</w:t></w:r>
                          </w:p>
                        </w:tc>
                        <w:tc><w:p><w:r><w:t>Other</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-cell
            target M.T0001.R01.C01
            expect-text Old
            text New
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        input.Position = 0;
        output.Position = 0;
        Assert.Equal("Old", new DocxEditor().Read(input).Tables[0].Cells[0].Text);
        Assert.Equal("New", new DocxEditor().Read(output).Tables[0].Cells[0].Text);

        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:tcW", xml, StringComparison.Ordinal);
        Assert.Contains("w:val=\"TableBody\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckSetCellRejectsComplexCellWithoutForce()
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
            text Replacement
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E4302");
    }

    [Fact]
    public static void ApplyAppendRowClonesRowShapeAndCellProperties()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>
                      <w:tr>
                        <w:trPr><w:trHeight w:val="240"/></w:trPr>
                        <w:tc>
                          <w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>
                          <w:p>
                            <w:pPr><w:pStyle w:val="TableRegion"/></w:pPr>
                            <w:r><w:t>North</w:t></w:r>
                          </w:p>
                        </w:tc>
                        <w:tc>
                          <w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>
                          <w:p><w:r><w:t>Revenue</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op append-row
            target M.T0001
            cell South
            cell Profit
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal(2, table.RowCount);
        Assert.Equal("South", table.Cells.Single(cell => cell.RowIndex == 2 && cell.ColumnIndex == 1).Text);
        Assert.Equal("Profit", table.Cells.Single(cell => cell.RowIndex == 2 && cell.ColumnIndex == 2).Text);

        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Equal(2, CountOccurrences(xml, "<w:trHeight"));
        Assert.Equal(4, CountOccurrences(xml, "<w:tcW"));
        Assert.Equal(2, CountOccurrences(xml, "w:val=\"TableRegion\""));
    }

    [Fact]
    public static void ApplySetTableStyleAndRowHeaderFlag()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-table-style
            target M.T0001
            style TableGrid
            end

            op set-row-header
            target M.T0001.R01
            expect-header false
            header true
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal("TableGrid", table.StyleId);
        Assert.True(table.HasHeaderRow);
        Assert.True(table.Rows[0].IsHeader);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:tblStyle w:val=\"TableGrid\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:tblHeader", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetTableMetadataUpdatesAndClearsCaptionDescription()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tblPr>
                        <w:tblCaption w:val="Old caption"/>
                        <w:tblDescription w:val="Old description"/>
                      </w:tblPr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-table-metadata
            target M.T0001
            expect-caption Old caption
            expect-description Old description
            caption Revenue table
            description <<<
            >>>
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal("Revenue table", table.Caption);
        Assert.Null(table.Description);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<w:tblCaption w:val=\"Revenue table\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("tblDescription", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckTablePropertyGuardsRejectMismatches()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tblPr>
                        <w:tblStyle w:val="ExistingStyle"/>
                        <w:tblCaption w:val="Existing caption"/>
                      </w:tblPr>
                      <w:tr>
                        <w:trPr><w:tblHeader/></w:trPr>
                        <w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-table-style
            target M.T0001
            expect-style OtherStyle
            style TableGrid
            end

            op set-row-header
            target M.T0001.R01
            expect-header false
            header false
            end

            op set-table-metadata
            target M.T0001
            expect-caption Other caption
            caption Updated caption
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E3201" && diagnostic.TargetId == "M.T0001");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E3201" && diagnostic.TargetId == "M.T0001.R01");
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E3201" &&
            diagnostic.TargetId == "M.T0001" &&
            diagnostic.Message.Contains("caption", StringComparison.Ordinal));
    }

    [Fact]
    public static void CheckAppendRowRejectsCellCountMismatch()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op append-row
            target M.T0001
            cell South
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E4303");
    }

    [Fact]
    public static void CheckAppendRowRejectsFailedShapeGuards()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op append-row
            target M.T0001
            expect-row-count 2
            expect-column-count 2
            cell South
            cell Profit
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E3201" && diagnostic.TargetId == "M.T0001");
    }

    [Fact]
    public static void ApplyReplaceImageUsesAssetProviderForPng()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", "new-png", null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-image
            target M.I0001
            asset chart.png
            alt Updated chart
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("new-png", ReadEntry(output, "word/media/image1.png"));
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("descr=\"Updated chart\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("descr=\"Old chart\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyReplaceImagePreservesAnchoredLayoutAndCropMetadata()
    {
        using MemoryStream input = CreateDocxWithAnchoredImage();
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", "new-png", null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-image
            target M.I0001
            asset chart.png
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("new-png", ReadEntry(output, "word/media/image1.png"));
        output.Position = 0;
        DocxImageInfo image = Assert.Single(new DocxEditor().Read(output).Images);
        Assert.Equal("anchor", image.LayoutKind);
        Assert.Equal(914400, image.WidthEmu);
        Assert.Equal(457200, image.HeightEmu);
        Assert.Equal("wrapSquare", image.WrapMode);
        Assert.Equal(10L, image.WrapDistanceTopEmu);
        Assert.Equal(20L, image.WrapDistanceBottomEmu);
        Assert.Equal(30L, image.WrapDistanceLeftEmu);
        Assert.Equal(40L, image.WrapDistanceRightEmu);
        Assert.Equal("column", image.HorizontalPositionRelativeFrom);
        Assert.Equal(123L, image.HorizontalPositionOffsetEmu);
        Assert.Equal("paragraph", image.VerticalPositionRelativeFrom);
        Assert.Equal("top", image.VerticalPositionAlign);
        Assert.Equal(1m, image.CropLeftPercent);
        Assert.Equal(2m, image.CropTopPercent);
        Assert.Equal(3m, image.CropRightPercent);
        Assert.Equal(4m, image.CropBottomPercent);
    }

    [Fact]
    public static void CheckReplaceImageRejectsFailedContentTypeGuard()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        var assets = new MemoryAssetProvider("chart.png", "new-png", null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-image
            target M.I0001
            expect-content-type image/jpeg
            asset chart.png
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { AssetProvider = assets });

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E3201" && diagnostic.TargetId == "M.I0001");
    }

    [Fact]
    public static void ApplyReplaceImageUsesAssetProviderForJpeg()
    {
        using MemoryStream input = CreateDocxWithImage("jpeg", "image/jpeg", "old-jpeg");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("photo.jpeg", "new-jpeg", null, "photo.jpeg");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-image
            target M.I0001
            asset photo.jpeg
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("new-jpeg", ReadEntry(output, "word/media/image1.jpeg"));
    }

    [Fact]
    public static void CheckReplaceImageFailsWithoutAssetProvider()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-image
            target M.I0001
            asset chart.png
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E5201");
    }

    [Fact]
    public static void CheckReplaceImageRejectsUnsupportedAssetType()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        var assets = new MemoryAssetProvider("asset.bin", "not-image", null, "asset.bin");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-image
            target M.I0001
            asset asset.bin
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch, new DocxEditOptions { AssetProvider = assets });

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E5203");
    }

    [Fact]
    public static void ApplySetImageAltUpdatesDrawingProperties()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-alt
            target M.I0001
            alt Updated chart
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("descr=\"Updated chart\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("descr=\"Old chart\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetImageMetadataUpdatesDrawingProperties()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-metadata
            target M.I0001
            alt Updated chart
            title Revenue chart
            name Revenue picture
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxImageInfo image = Assert.Single(new DocxEditor().Read(output).Images);
        Assert.Equal("Updated chart", image.Description);
        Assert.Equal("Revenue chart", image.Title);
        Assert.Equal("Revenue picture", image.Name);
    }

    [Fact]
    public static void ApplySetImageSizePreservesAspectWhenOnlyWidthIsGiven()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-size
            target M.I0001
            width 2in
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxImageInfo image = Assert.Single(new DocxEditor().Read(output).Images);
        Assert.Equal(1_828_800, image.WidthEmu);
        Assert.Equal(914_400, image.HeightEmu);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<wp:extent cx=\"1828800\" cy=\"914400\"", xml, StringComparison.Ordinal);
        Assert.Contains("<a:ext cx=\"1828800\" cy=\"914400\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetImageWrapUpdatesAnchorModeAndDistances()
    {
        using MemoryStream input = CreateDocxWithAnchoredImage();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-wrap
            target M.I0001
            mode top-bottom
            dist-top 1pt
            dist-right 2pt
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxImageInfo image = Assert.Single(new DocxEditor().Read(output).Images);
        Assert.Equal("wrapTopAndBottom", image.WrapMode);
        Assert.Equal(12_700, image.WrapDistanceTopEmu);
        Assert.Equal(20, image.WrapDistanceBottomEmu);
        Assert.Equal(30, image.WrapDistanceLeftEmu);
        Assert.Equal(25_400, image.WrapDistanceRightEmu);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<wp:wrapTopAndBottom", xml, StringComparison.Ordinal);
        Assert.Contains("distT=\"12700\"", xml, StringComparison.Ordinal);
        Assert.Contains("distR=\"25400\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("wrapSquare", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckSetImageWrapRejectsInlineImage()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-wrap
            target M.I0001
            mode square
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E5205");
    }

    [Fact]
    public static void ApplySetImagePositionUpdatesAnchorPosition()
    {
        using MemoryStream input = CreateDocxWithAnchoredImage();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-position
            target M.I0001
            horizontal-relative page
            horizontal-offset -0.25in
            vertical-relative paragraph
            vertical-align bottom
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxImageInfo image = Assert.Single(new DocxEditor().Read(output).Images);
        Assert.Equal("page", image.HorizontalPositionRelativeFrom);
        Assert.Equal(-228_600, image.HorizontalPositionOffsetEmu);
        Assert.Null(image.HorizontalPositionAlign);
        Assert.Equal("paragraph", image.VerticalPositionRelativeFrom);
        Assert.Null(image.VerticalPositionOffsetEmu);
        Assert.Equal("bottom", image.VerticalPositionAlign);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<wp:positionH relativeFrom=\"page\"><wp:posOffset>-228600</wp:posOffset></wp:positionH>", xml, StringComparison.Ordinal);
        Assert.Contains("<wp:positionV relativeFrom=\"paragraph\"><wp:align>bottom</wp:align></wp:positionV>", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckSetImagePositionRejectsInlineImage()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-position
            target M.I0001
            horizontal-offset 1in
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E5205");
    }

    [Fact]
    public static void ApplySetImageCropUpdatesSourceRectangle()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-crop
            target M.I0001
            left-percent 12.5
            top-percent 5
            right-percent 2.25
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxImageInfo image = Assert.Single(new DocxEditor().Read(output).Images);
        Assert.Equal(12.5m, image.CropLeftPercent);
        Assert.Equal(5m, image.CropTopPercent);
        Assert.Equal(2.25m, image.CropRightPercent);
        Assert.Null(image.CropBottomPercent);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<a:srcRect l=\"12500\" t=\"5000\" r=\"2250\" />", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckSetImageCropRejectsEmptyHorizontalCrop()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var patch = new StringReader("""
            docxpatch 1

            op set-image-crop
            target M.I0001
            left-percent 60
            right-percent 40
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E5208");
    }

    [Fact]
    public static void ApplyDeleteImageRemovesDrawingAndPreservesParagraph()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-image
            target M.I0001
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Empty(read.Images);
        Assert.Single(read.Paragraphs);
        output.Position = 0;
        Assert.DoesNotContain("<w:drawing>", ReadDocumentXml(output), StringComparison.Ordinal);
        output.Position = 0;
        Assert.False(EntryExists(output, "word/media/image1.png"));
        output.Position = 0;
        string relationships = ReadEntry(output, "word/_rels/document.xml.rels");
        Assert.DoesNotContain("rImage", relationships, StringComparison.Ordinal);
        Assert.DoesNotContain("media/image1.png", relationships, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyImageOperationsCanTargetHeaderAndFooterImages()
    {
        using MemoryStream input = CreateDocxWithHeaderFooterImages();
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("header.png", "new-header", null, "header.png");
        using var patch = new StringReader("""
            docxpatch 1

            op replace-image
            target H001.I0001
            asset header.png
            alt Updated header image
            end

            op delete-image
            target F001.I0001
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        Assert.Equal("new-header", ReadEntry(output, "word/media/header.png"));
        output.Position = 0;
        Assert.False(EntryExists(output, "word/media/footer.png"));
        output.Position = 0;
        Assert.Contains("descr=\"Updated header image\"", ReadEntry(output, "word/header1.xml"), StringComparison.Ordinal);
        output.Position = 0;
        Assert.DoesNotContain("rFooterImage", ReadEntry(output, "word/_rels/footer1.xml.rels"), StringComparison.Ordinal);
        output.Position = 0;
        DocxImageInfo image = Assert.Single(new DocxEditor().Read(output, new DocxReadOptions { IncludeHeadersFooters = true }).Images);
        Assert.Equal("H001.I0001", image.Id);
        Assert.Equal("/word/media/header.png", image.PartName);
    }

    [Fact]
    public static void ApplyInsertImageAfterAddsInlinePngImage()
    {
        using MemoryStream input = CreateDocx("Intro");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", "new-png", null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset chart.png
            width 1in
            alt Inserted chart
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        DocxImageInfo image = Assert.Single(read.Images);
        Assert.Equal("/word/media/image1.png", image.PartName);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(2, read.Paragraphs.Count);
        output.Position = 0;
        Assert.Equal("new-png", ReadEntry(output, "word/media/image1.png"));
        output.Position = 0;
        Assert.Contains("ContentType=\"image/png\"", ReadEntry(output, "[Content_Types].xml"), StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyInsertImageRejectsEmptyImageAssetDuringPostEditValidation()
    {
        using MemoryStream input = CreateDocx("Intro");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", [], null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset chart.png
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E9001" && diagnostic.PartName == "/word/media/image1.png");
    }

    [Fact]
    public static void ApplyGeneratedDocumentPassesStructuralValidation()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p><w:r><w:t>Anchor</w:t></w:r></w:p>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            expect-text Anchor
            find Anchor
            with Updated anchor
            end

            op add-bookmark
            target M.P0001
            expect-text Updated anchor
            name ReviewAnchor
            end

            op set-table-metadata
            target M.T0001
            caption Review table
            description Generated validation fixture
            end

            op append-row
            target M.T0001
            cell South
            cell Profit
            end
            """);

        DocxApplyResult apply = new DocxEditor().Apply(input, patch, output);

        Assert.True(apply.Success);
        output.Position = 0;
        DocxValidateResult validate = new DocxEditor().Validate(output);
        Assert.True(validate.Success, string.Join(Environment.NewLine, validate.Diagnostics.Select(diagnostic => $"{diagnostic.Code}: {diagnostic.Message}")));
        Assert.DoesNotContain(validate.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error);
    }

    [Fact]
    public static void ApplyInsertImageAfterPreservesPngAspectRatioWhenOnlyWidthIsSupplied()
    {
        using MemoryStream input = CreateDocx("Intro");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", CreatePngBytes(200, 100), null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset chart.png
            width 2in
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        (long cx, long cy) = ReadFirstInlineImageExtent(output);
        Assert.Equal(1_828_800, cx);
        Assert.Equal(914_400, cy);
    }

    [Fact]
    public static void ApplyInsertImageAfterInfersPngDimensionsWhenNoneAreSupplied()
    {
        using MemoryStream input = CreateDocx("Intro");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", CreatePngBytes(200, 100), null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset chart.png
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        (long cx, long cy) = ReadFirstInlineImageExtent(output);
        Assert.Equal(1_905_000, cx);
        Assert.Equal(952_500, cy);
    }

    [Fact]
    public static void ApplyInsertImageAfterPreservesJpegAspectRatioWhenOnlyHeightIsSupplied()
    {
        using MemoryStream input = CreateDocx("Intro");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("photo.jpeg", CreateJpegBytes(120, 60), null, "photo.jpeg");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset photo.jpeg
            height 1in
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        (long cx, long cy) = ReadFirstInlineImageExtent(output);
        Assert.Equal(1_828_800, cx);
        Assert.Equal(914_400, cy);
    }

    [Fact]
    public static void ApplyInsertImageAfterAddsInlineJpegImage()
    {
        using MemoryStream input = CreateDocx("Intro");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("photo.jpeg", "new-jpeg", null, "photo.jpeg");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset photo.jpeg
            width 72pt
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        DocxImageInfo image = Assert.Single(new DocxEditor().Read(output).Images);
        Assert.Equal("/word/media/image1.jpeg", image.PartName);
        Assert.Equal("image/jpeg", image.ContentType);
        output.Position = 0;
        Assert.Equal("new-jpeg", ReadEntry(output, "word/media/image1.jpeg"));
    }

    [Fact]
    public static void ApplyInsertImageAfterAllocatesUniqueDocPrId()
    {
        using MemoryStream input = CreateDocxWithImage("png", "image/png", "old-png");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider("chart.png", "new-png", null, "chart.png");
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset chart.png
            width 1in
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("<wp:docPr id=\"1\"", xml, StringComparison.Ordinal);
        Assert.Contains("<wp:docPr id=\"2\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplyInsertImageAfterTwiceRefreshesContentTypesAndRelationships()
    {
        using MemoryStream input = CreateDocx("Intro");
        using var output = new MemoryStream();
        var assets = new MemoryAssetProvider(
            ("chart.png", "new-png", null, "chart.png"),
            ("logo.png", "new-logo", null, "logo.png"));
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset chart.png
            width 1in
            end

            op insert-image-after
            target M.P0001
            asset logo.png
            width 72pt
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { AssetProvider = assets });

        Assert.True(result.Success);
        output.Position = 0;
        DocxReadResult read = new DocxEditor().Read(output);
        Assert.Equal(2, read.Images.Count);
        Assert.Contains(read.Images, image => image.PartName == "/word/media/image1.png" && image.ContentType == "image/png");
        Assert.Contains(read.Images, image => image.PartName == "/word/media/image2.png" && image.ContentType == "image/png");

        output.Position = 0;
        Assert.Equal("new-png", ReadEntry(output, "word/media/image1.png"));
        output.Position = 0;
        Assert.Equal("new-logo", ReadEntry(output, "word/media/image2.png"));

        output.Position = 0;
        string contentTypes = ReadEntry(output, "[Content_Types].xml");
        Assert.Contains("PartName=\"/word/media/image1.png\"", contentTypes, StringComparison.Ordinal);
        Assert.Contains("ContentType=\"image/png\"", contentTypes, StringComparison.Ordinal);
        Assert.Contains("PartName=\"/word/media/image2.png\"", contentTypes, StringComparison.Ordinal);

        output.Position = 0;
        string relationships = ReadEntry(output, "word/_rels/document.xml.rels");
        Assert.Contains("Target=\"media/image1.png\"", relationships, StringComparison.Ordinal);
        Assert.Contains("Target=\"media/image2.png\"", relationships, StringComparison.Ordinal);
        Assert.Equal(2, CountOccurrences(relationships, "relationships/image"));
    }

    [Fact]
    public static void ApplyInsertRowBeforeClonesTargetRowShape()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:trPr><w:trHeight w:val="480"/></w:trPr>
                        <w:tc>
                          <w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>
                          <w:p><w:r><w:t>South</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                          <w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>
                          <w:p><w:r><w:t>Profit</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-row-before
            target M.T0001.R02
            cell East
            cell Margin
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal(3, table.RowCount);
        Assert.Equal("East", table.Cells.Single(cell => cell.RowIndex == 2 && cell.ColumnIndex == 1).Text);
        Assert.Equal("South", table.Cells.Single(cell => cell.RowIndex == 3 && cell.ColumnIndex == 1).Text);

        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Equal(2, CountOccurrences(xml, "<w:trHeight"));
        Assert.Equal(4, CountOccurrences(xml, "<w:tcW"));
    }

    [Fact]
    public static void ApplyInsertRowAfterPlacesRowAfterTarget()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-row-after
            target M.T0001.R01
            cell East
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal("North", table.Cells.Single(cell => cell.RowIndex == 1 && cell.ColumnIndex == 1).Text);
        Assert.Equal("East", table.Cells.Single(cell => cell.RowIndex == 2 && cell.ColumnIndex == 1).Text);
        Assert.Equal("South", table.Cells.Single(cell => cell.RowIndex == 3 && cell.ColumnIndex == 1).Text);
    }

    [Fact]
    public static void ApplyDeleteRowRemovesSelectedRow()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>East</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-row
            target M.T0001.R02
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxTableInfo table = Assert.Single(new DocxEditor().Read(output).Tables);
        Assert.Equal(2, table.RowCount);
        Assert.Equal("North", table.Cells.Single(cell => cell.RowIndex == 1 && cell.ColumnIndex == 1).Text);
        Assert.Equal("East", table.Cells.Single(cell => cell.RowIndex == 2 && cell.ColumnIndex == 1).Text);
    }

    [Theory]
    [InlineData("append-row", "M.T0001")]
    [InlineData("insert-row-before", "M.T0001.R02")]
    [InlineData("delete-row", "M.T0001.R02")]
    public static void CheckRowOperationsRejectVerticalMergeTables(string operationName, string target)
    {
        string cellFields = operationName == "delete-row"
            ? string.Empty
            : """
            cell Inserted
            cell Value
            """;
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
                          <w:p><w:r><w:t>North</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:vMerge/></w:tcPr>
                          <w:p><w:r><w:t>South</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc><w:p><w:r><w:t>Profit</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader($"""
            docxpatch 1

            op {operationName}
            target {target}
            {cellFields}
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "E4301" &&
            diagnostic.Message.Contains("vertical merges", StringComparison.Ordinal));
    }

    [Fact]
    public static void CheckAppendRowReportsAffectedRowAndCells()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op append-row
            target M.T0001
            cell South
            cell Profit
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.True(result.Success);
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        DocxPatchAffectedTarget row = Assert.Single(report.AffectedTargets, target => target.Kind == "row");
        Assert.Equal("M.T0001.R02", row.Id);
        Assert.Equal("append", row.Action);
        Assert.Equal(1, row.RowCountBefore);
        Assert.Equal(2, row.RowCountAfter);
        Assert.Equal(2, row.CellCount);
        Assert.Contains(report.AffectedTargets, target => target.Id == "M.T0001.R02.C01" && target.Kind == "cell" && target.ColumnIndex == 1);
        Assert.Contains(report.AffectedTargets, target => target.Id == "M.T0001.R02.C02" && target.Kind == "cell" && target.ColumnIndex == 2);
    }

    [Fact]
    public static void ApplyDeleteRowReportsAffectedRowAndCells()
    {
        using MemoryStream input = CreateDocxWithBody("""
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
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op delete-row
            target M.T0001.R02
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        DocxPatchOperationReport report = Assert.Single(result.Operations);
        DocxPatchAffectedTarget row = Assert.Single(report.AffectedTargets, target => target.Kind == "row");
        Assert.Equal("M.T0001.R02", row.Id);
        Assert.Equal("delete", row.Action);
        Assert.Equal(2, row.RowCountBefore);
        Assert.Equal(1, row.RowCountAfter);
        Assert.Equal(2, row.CellCount);
        Assert.Contains(report.AffectedTargets, target => target.Id == "M.T0001.R02.C01" && target.Kind == "cell" && target.Action == "delete");
        Assert.Contains(report.AffectedTargets, target => target.Id == "M.T0001.R02.C02" && target.Kind == "cell" && target.Action == "delete");
    }

    [Fact]
    public static void OperationSummaryRendersAffectedTableTargets()
    {
        var operations = new[]
        {
            new DocxPatchOperationReport(1, "append-row", "M.T0001", true, [])
            {
                AffectedTargets =
                [
                    new("M.T0001.R02", "row", "append")
                    {
                        ParentId = "M.T0001",
                        RowIndex = 2,
                        RowCountBefore = 1,
                        RowCountAfter = 2,
                        ColumnCount = 2,
                        CellCount = 2
                    }
                ]
            }
        };

        string text = DocxTextRenderer.RenderOperationSummary(operations);

        Assert.Contains("operation index=1 name=append-row target=M.T0001 success=True", text, StringComparison.Ordinal);
        Assert.Contains("affected id=M.T0001.R02 kind=row action=append parent=M.T0001 row=2 rows-before=1 rows-after=2 columns=2 cells=2", text, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckDeleteRowRejectsLastRow()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Only</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op delete-row
            target M.T0001.R01
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E4304");
    }

    [Fact]
    public static void CheckDeleteRowRejectsFailedExpectContainsGuard()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op delete-row
            target M.T0001.R01
            expect-contains South
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "E3201" && diagnostic.TargetId == "M.T0001.R01");
        Assert.False(Assert.Single(result.Operations).Success);
    }

    private static MemoryStream CreateDocx(string paragraphText)
    {
        return CreateDocxWithRuns(paragraphText);
    }

    private static MemoryStream CreateDocxWithRuns(params string[] runTexts)
    {
        var stream = new MemoryStream();
        string runs = string.Concat(runTexts.Select(text => $"<w:r><w:t>{text}</w:t></w:r>"));
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """);
            AddEntry(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
                """);
            AddEntry(archive, "word/document.xml", $$"""
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>{{runs}}</w:p>
                  </w:body>
                </w:document>
                """);
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithBody(string bodyXml, Action<ZipArchive>? extra = null)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """);
            AddEntry(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
                """);
            string documentXml = """
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>

                """ + bodyXml + """

                  </w:body>
                </w:document>
                """;
            AddEntry(archive, "word/document.xml", documentXml);
            extra?.Invoke(archive);
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithBodyAndRelationships(string bodyXml, string relationshipsXml)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """);
            AddEntry(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/_rels/document.xml.rels", relationshipsXml);
            string documentXml = """
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>

                """ + bodyXml + """

                  </w:body>
                </w:document>
                """;
            AddEntry(archive, "word/document.xml", documentXml);
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithBodyAndComments(
        string bodyXml,
        string commentsXml,
        string? commentsExtendedXml = null)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            string contentTypesXml = commentsExtendedXml is null
                ? """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
                </Types>
                """
                : """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
                  <Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.ms-word.commentsExtended+xml"/>
                </Types>
                """;
            AddEntry(archive, "[Content_Types].xml", contentTypesXml);
            AddEntry(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            string relationshipsXml = commentsExtendedXml is null
                ? """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
                </Relationships>
                """
                : """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
                  <Relationship Id="rCommentsExtended" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>
                </Relationships>
                """;
            AddEntry(archive, "word/_rels/document.xml.rels", relationshipsXml);
            string documentXml = """
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>

                """ + bodyXml + """

                  </w:body>
                </w:document>
                """;
            AddEntry(archive, "word/document.xml", documentXml);
            AddEntry(archive, "word/comments.xml", commentsXml);
            if (commentsExtendedXml is not null)
            {
                AddEntry(archive, "word/commentsExtended.xml", commentsExtendedXml);
            }
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithHeaderComment()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
                  <Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.ms-word.commentsExtended+xml"/>
                </Types>
                """);
            AddEntry(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                  <Relationship Id="rComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
                  <Relationship Id="rCommentsExtended" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/document.xml", """
                <w:document
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>Body</w:t></w:r></w:p>
                    <w:sectPr><w:headerReference w:type="default" r:id="rHeader"/></w:sectPr>
                  </w:body>
                </w:document>
                """);
            AddEntry(archive, "word/header1.xml", """
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:p>
                    <w:commentRangeStart w:id="3"/>
                    <w:r><w:t>Header text</w:t></w:r>
                    <w:commentRangeEnd w:id="3"/>
                    <w:r><w:commentReference w:id="3"/></w:r>
                  </w:p>
                </w:hdr>
                """);
            AddEntry(archive, "word/comments.xml", """
                <w:comments
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w:comment w:id="3" w:author="Reviewer">
                    <w:p w15:paraId="00ABCDEF"><w:r><w:t>Comment body</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);
            AddEntry(archive, "word/commentsExtended.xml", """
                <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w15:commentEx w15:paraId="00ABCDEF" w15:done="1"/>
                </w15:commentsEx>
                """);
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithImage(string extension, string contentType, string mediaBytes)
    {
        var stream = new MemoryStream();
        string partName = $"word/media/image1.{extension}";
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", $$"""
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="{{extension}}" ContentType="{{contentType}}"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """);
            AddEntry(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/_rels/document.xml.rels", $$"""
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.{{extension}}"/>
                </Relationships>
                """);
            AddEntry(archive, "word/document.xml", """
                <w:document
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:drawing>
                          <wp:inline>
                            <wp:extent cx="914400" cy="457200"/>
                            <wp:docPr id="1" name="Picture 1" descr="Old chart"/>
                            <a:graphic>
                              <a:graphicData>
                                <pic:pic>
                                  <pic:blipFill>
                                    <a:blip r:embed="rImage"/>
                                  </pic:blipFill>
                                </pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                    </w:p>
                  </w:body>
                </w:document>
                """);
            AddEntry(archive, partName, mediaBytes);
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithAnchoredImage()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
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
                  <Relationship Id="rImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
                </Relationships>
                """);
            AddEntry(archive, "word/document.xml", """
                <w:document
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:drawing>
                          <wp:anchor distT="10" distB="20" distL="30" distR="40">
                            <wp:positionH relativeFrom="column"><wp:posOffset>123</wp:posOffset></wp:positionH>
                            <wp:positionV relativeFrom="paragraph"><wp:align>top</wp:align></wp:positionV>
                            <wp:extent cx="914400" cy="457200"/>
                            <wp:wrapSquare/>
                            <wp:docPr id="1" name="Picture 1" descr="Old chart"/>
                            <a:graphic>
                              <a:graphicData>
                                <pic:pic>
                                  <pic:blipFill>
                                    <a:blip r:embed="rImage"/>
                                    <a:srcRect l="1000" t="2000" r="3000" b="4000"/>
                                  </pic:blipFill>
                                </pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:anchor>
                        </w:drawing>
                      </w:r>
                    </w:p>
                  </w:body>
                </w:document>
                """);
            AddEntry(archive, "word/media/image1.png", "old-png");
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithHeaderFooter(string headerText, string footerText)
    {
        return CreateDocxWithHeaderFooterContent(
            $$"""
                  <w:p><w:r><w:t>{{headerText}}</w:t></w:r></w:p>
            """,
            $$"""
                  <w:p><w:r><w:t>{{footerText}}</w:t></w:r></w:p>
            """);
    }

    private static MemoryStream CreateDocxWithHeaderFooterContent(string headerContent, string footerContent)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
                </Types>
                """);
            AddEntry(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                  <Relationship Id="rFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/document.xml", """
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>Main text</w:t></w:r></w:p>
                  </w:body>
                </w:document>
                """);
            AddEntry(archive, "word/header1.xml", """
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

                """ + headerContent + """

                </w:hdr>
                """);
            AddEntry(archive, "word/footer1.xml", """
                <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

                """ + footerContent + """

                </w:ftr>
                """);
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithHeaderFooterImages()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
                </Types>
                """);
            AddEntry(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                  <Relationship Id="rFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/_rels/header1.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rHeaderImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/header.png"/>
                </Relationships>
                """);
            AddEntry(archive, "word/_rels/footer1.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rFooterImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/footer.png"/>
                </Relationships>
                """);
            AddEntry(archive, "word/document.xml", """
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>Main text</w:t></w:r></w:p>
                  </w:body>
                </w:document>
                """);
            AddEntry(archive, "word/header1.xml", CreateImageStoryXml("hdr", "rHeaderImage", "Header image"));
            AddEntry(archive, "word/footer1.xml", CreateImageStoryXml("ftr", "rFooterImage", "Footer image"));
            AddEntry(archive, "word/media/header.png", "old-header");
            AddEntry(archive, "word/media/footer.png", "old-footer");
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateDocxWithHeaderFooterAndStyles()
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
                  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
                </Types>
                """);
            AddEntry(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                  <Relationship Id="rFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
                  <Relationship Id="rStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/document.xml", """
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>Main text</w:t></w:r></w:p>
                  </w:body>
                </w:document>
                """);
            AddEntry(archive, "word/header1.xml", """
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:p><w:r><w:t>Header text</w:t></w:r></w:p>
                </w:hdr>
                """);
            AddEntry(archive, "word/footer1.xml", """
                <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:p><w:r><w:t>Footer text</w:t></w:r></w:p>
                </w:ftr>
                """);
            AddEntry(archive, "word/styles.xml", """
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
                  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/></w:style>
                </w:styles>
                """);
        }

        stream.Position = 0;
        return stream;
    }

    private static string CreateImageStoryXml(string rootName, string relationshipId, string description)
    {
        return $$"""
            <w:{{rootName}}
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <w:p>
                <w:r>
                  <w:drawing>
                    <wp:inline>
                      <wp:docPr id="1" name="Picture 1" descr="{{description}}"/>
                      <a:graphic>
                        <a:graphicData>
                          <pic:pic>
                            <pic:blipFill>
                              <a:blip r:embed="{{relationshipId}}"/>
                            </pic:blipFill>
                          </pic:pic>
                        </a:graphicData>
                      </a:graphic>
                    </wp:inline>
                  </w:drawing>
                </w:r>
              </w:p>
            </w:{{rootName}}>
            """;
    }

    private static MemoryStream CreateDocxWithStyles(string styleElements)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
                </Types>
                """);
            AddEntry(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/document.xml", """
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>Styled paragraph</w:t></w:r></w:p>
                  </w:body>
                </w:document>
                """);
            string stylesXml = """
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

                """ + styleElements + """

                </w:styles>
                """;
            AddEntry(archive, "word/styles.xml", stylesXml);
        }

        stream.Position = 0;
        return stream;
    }

    private static string ReadDocumentXml(Stream docx)
    {
        using var archive = new ZipArchive(docx, ZipArchiveMode.Read, leaveOpen: true);
        ZipArchiveEntry entry = archive.GetEntry("word/document.xml")
            ?? throw new InvalidDataException("Missing word/document.xml.");
        using Stream stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    }

    private static (long Cx, long Cy) ReadFirstInlineImageExtent(Stream docx)
    {
        string xml = ReadDocumentXml(docx);
        XElement extent = XDocument.Parse(xml)
            .Descendants(XName.Get("extent", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"))
            .First();
        return ((long)extent.Attribute("cx")!, (long)extent.Attribute("cy")!);
    }

    private static string ReadEntry(Stream docx, string entryName)
    {
        using var archive = new ZipArchive(docx, ZipArchiveMode.Read, leaveOpen: true);
        ZipArchiveEntry entry = archive.GetEntry(entryName)
            ?? throw new InvalidDataException($"Missing {entryName}.");
        using Stream stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    }

    private static bool EntryExists(Stream docx, string entryName)
    {
        using var archive = new ZipArchive(docx, ZipArchiveMode.Read, leaveOpen: true);
        return archive.GetEntry(entryName) is not null;
    }

    private static int CountOccurrences(string text, string value)
    {
        int count = 0;
        int index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += value.Length;
        }

        return count;
    }

    private static void AddEntry(ZipArchive archive, string name, string text)
    {
        ZipArchiveEntry entry = archive.CreateEntry(name);
        using Stream stream = entry.Open();
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static byte[] CreatePngBytes(int width, int height)
    {
        byte[] bytes =
        [
            0x89, 0x50, 0x4E, 0x47,
            0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x0D,
            0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00
        ];
        WriteBigEndianInt32(bytes, 16, width);
        WriteBigEndianInt32(bytes, 20, height);
        return bytes;
    }

    private static byte[] CreateJpegBytes(int width, int height)
    {
        byte[] bytes =
        [
            0xFF, 0xD8,
            0xFF, 0xC0,
            0x00, 0x11,
            0x08,
            0x00, 0x00,
            0x00, 0x00,
            0x03,
            0x01, 0x11, 0x00,
            0x02, 0x11, 0x00,
            0x03, 0x11, 0x00,
            0xFF, 0xD9
        ];
        WriteBigEndianUInt16(bytes, 7, height);
        WriteBigEndianUInt16(bytes, 9, width);
        return bytes;
    }

    private static void WriteBigEndianInt32(byte[] bytes, int offset, int value)
    {
        bytes[offset] = (byte)((value >> 24) & 0xFF);
        bytes[offset + 1] = (byte)((value >> 16) & 0xFF);
        bytes[offset + 2] = (byte)((value >> 8) & 0xFF);
        bytes[offset + 3] = (byte)(value & 0xFF);
    }

    private static void WriteBigEndianUInt16(byte[] bytes, int offset, int value)
    {
        bytes[offset] = (byte)((value >> 8) & 0xFF);
        bytes[offset + 1] = (byte)(value & 0xFF);
    }

    private sealed class MemoryAssetProvider : IDocxAssetProvider
    {
        private readonly Dictionary<string, MemoryAsset> assets;

        public MemoryAssetProvider(string reference, string text, string? contentTypeHint, string? fileNameHint)
            : this((reference, text, contentTypeHint, fileNameHint))
        {
        }

        public MemoryAssetProvider(string reference, byte[] bytes, string? contentTypeHint, string? fileNameHint)
            : this(new MemoryAsset(reference, bytes, contentTypeHint, fileNameHint))
        {
        }

        public MemoryAssetProvider(params (string Reference, string Text, string? ContentTypeHint, string? FileNameHint)[] assets)
            : this(assets.Select(asset => new MemoryAsset(
                asset.Reference,
                Encoding.UTF8.GetBytes(asset.Text),
                asset.ContentTypeHint,
                asset.FileNameHint)).ToArray())
        {
        }

        private MemoryAssetProvider(params MemoryAsset[] assets)
        {
            this.assets = assets.ToDictionary(
                asset => asset.Reference,
                asset => asset,
                StringComparer.Ordinal);
        }

        public bool TryOpen(string requestedReference, out Stream stream, out string? contentTypeHint, out string? fileNameHint)
        {
            if (!assets.TryGetValue(requestedReference, out MemoryAsset? asset))
            {
                stream = Stream.Null;
                contentTypeHint = null;
                fileNameHint = null;
                return false;
            }

            stream = new MemoryStream(asset.Bytes, writable: false);
            contentTypeHint = asset.ContentTypeHint;
            fileNameHint = asset.FileNameHint;
            return true;
        }

        private sealed record MemoryAsset(string Reference, byte[] Bytes, string? ContentTypeHint, string? FileNameHint);
    }

    private sealed class NonSeekableReadStream : MemoryStream
    {
        public NonSeekableReadStream(byte[] bytes)
            : base(bytes, writable: false)
        {
        }

        public override bool CanSeek => false;

        public override long Seek(long offset, SeekOrigin loc)
        {
            throw new NotSupportedException();
        }

        public override long Position
        {
            get => base.Position;
            set => throw new NotSupportedException();
        }

        public override long Length => throw new NotSupportedException();
    }
}
