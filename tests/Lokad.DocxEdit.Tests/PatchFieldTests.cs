using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml.Linq;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

public static class PatchFieldTests
{

    [Fact]
    public static void ApplyMarksFieldsDirtyByDefault()
    {
        using MemoryStream input = CreateDocumentWithField();
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
        using MemoryStream input = CreateDocumentWithField();
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

    [Theory]
    [InlineData(null)]
    [InlineData("false")]
    [InlineData("true")]
    public static void FieldFreeEditPreservesRefreshSettings(string? refresh)
    {
        const string body = "<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>";
        string? relationships = refresh is null ? null : """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
            </Relationships>
            """;
        string settings = "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:updateFields w:val=\"" + refresh + "\"/></w:settings>";
        using MemoryStream input = CreateDocxWithBody(body, relationships,
            refresh is null ? null : archive => AddEntry(archive, "word/settings.xml", settings));
        byte[] original = input.ToArray();
        string originalRelationships = ReadEntry(input, "word/_rels/document.xml.rels");
        input.Position = 0;
        string originalTypes = ReadEntry(input, "[Content_Types].xml");
        input.Position = 0;
        const string patch = "docxpatch 1\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n";
        DocxCheckResult check = new DocxEditor().Check(input, new StringReader(patch));
        using var applyInput = new MemoryStream(original);
        using var output = new MemoryStream();
        DocxApplyResult apply = new DocxEditor().Apply(applyInput, new StringReader(patch), output);

        Assert.True(check.Success);
        Assert.True(apply.Success);
        Assert.Equal(check.Diagnostics, apply.Diagnostics);
        Assert.DoesNotContain(apply.Diagnostics, diagnostic => diagnostic.Code == "W5103");
        output.Position = 0;
        Assert.Equal("Beta", Assert.Single(new DocxEditor().Read(output).Paragraphs).Text);
        output.Position = 0;
        Assert.Equal(refresh is not null, EntryExists(output, "word/settings.xml"));
        if (refresh is not null)
        {
            output.Position = 0;
            Assert.Equal(settings, ReadEntry(output, "word/settings.xml"));
        }
        output.Position = 0;
        Assert.Equal(originalRelationships, ReadEntry(output, "word/_rels/document.xml.rels"));
        output.Position = 0;
        Assert.Equal(originalTypes, ReadEntry(output, "[Content_Types].xml"));
    }

    private static MemoryStream CreateDocumentWithField() => CreateDocxWithBody("""
        <w:p><w:r><w:t>Revenue increased.</w:t></w:r></w:p>
        <w:p><w:fldSimple w:instr=" DATE "><w:r><w:t>June 12</w:t></w:r></w:fldSimple></w:p>
        """);

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
    public static void ApplyRefreshQuoteFieldResultFromLiteralArguments()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:fldSimple w:instr=" QUOTE &quot;Acme Corp&quot; " w:dirty="true">
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
        DocxFieldInfo field = Assert.Single(new DocxEditor().Read(output).Fields);
        Assert.Equal("QUOTE", field.FieldType);
        Assert.Equal(new[] { "Acme Corp" }, field.Arguments);
        Assert.Equal(DocxRefreshPolicy.Literal, field.RefreshPolicy);
        Assert.True(field.CanRefreshDeterministically);
        Assert.Equal("Acme Corp", field.CachedResultText);
        Assert.Null(field.IsDirty);
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

    [Theory]
    [InlineData("TOC \\o \"1-3\"", "TOC", "layout")]
    [InlineData("PAGE", "PAGE", "pagination")]
    [InlineData("DOCPROPERTY Title", "DOCPROPERTY", "document property")]
    [InlineData("MERGEFIELD CustomerName", "MERGEFIELD", "mail merge")]
    [InlineData("IF \"A\" = \"A\" \"Yes\" \"No\"", "IF", "conditional")]
    [InlineData("= 1 + 1", "FORMULA", "formula")]
    [InlineData("DATE", "DATE", "date/time")]
    public static void CheckRefreshFieldResultRejectsNonGoalFieldTypesWithSpecificDiagnostics(
        string fieldCode,
        string expectedType,
        string expectedReason)
    {
        using MemoryStream input = CreateDocxWithBody($"""
                    <w:p>
                      <w:fldSimple w:instr="{SecurityElement.Escape(fieldCode)}">
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
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal("E4313", diagnostic.Code);
        Assert.Equal("M.F0001", diagnostic.TargetId);
        Assert.Contains($"field type '{expectedType}'", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains(expectedReason, diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("QUOTE", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public static void ApplySetFieldResultUpdatesSimpleComplexFieldResult()
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
            expect-result 1
            text 2
            end
            """);
        using var output = new MemoryStream();

        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        DocxFieldInfo field = Assert.Single(new DocxEditor().Read(output).Fields);
        Assert.Equal("complex", field.Kind);
        Assert.Equal("PAGE", field.FieldType);
        Assert.Equal("2", field.CachedResultText);
        output.Position = 0;
        string xml = ReadDocumentXml(output);
        Assert.Contains("w:fldCharType=\"begin\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:fldCharType=\"separate\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:fldCharType=\"end\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:instrText> PAGE </w:instrText>", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:t>1</w:t>", xml, StringComparison.Ordinal);
    }

    [Fact]
    public static void CheckSetFieldResultRejectsNestedComplexFieldResultTopology()
    {
        using MemoryStream input = CreateDocxWithBody("""
                    <w:p>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> IF </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> DATE </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:t>June 14</w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                    </w:p>
            """);
        using var patch = new StringReader("""
            docxpatch 1

            op set-field-result
            target M.F0002
            text Replacement
            end
            """);

        DocxCheckResult result = new DocxEditor().Check(input, patch);

        Assert.False(result.Success);
        DocxDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal("E4313", diagnostic.Code);
        Assert.Contains("nested complex field", diagnostic.Message, StringComparison.Ordinal);
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
}
