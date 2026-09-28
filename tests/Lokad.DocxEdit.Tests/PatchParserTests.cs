namespace Lokad.DocxEdit.Tests;

public static class PatchParserTests
{
    [Fact]
    public static void ParsePatchRejectsMissingPreambleWithLineAndColumn()
    {
        var editor = new DocxEditor();

        DocxPatch patch = editor.ParsePatch(new StringReader("""

            op replace-text
            end
            """));

        Assert.False(patch.Success);
        DocxDiagnostic diagnostic = Assert.Single(patch.Diagnostics);
        Assert.Equal("E2002", diagnostic.Code);
        Assert.Equal(2, diagnostic.Line);
        Assert.Equal(1, diagnostic.Column);
    }

    [Fact]
    public static void ParsePatchRejectsExpectHash()
    {
        var editor = new DocxEditor();

        DocxPatch patch = editor.ParsePatch(new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            expect-hash abc
            end
            """));

        Assert.False(patch.Success);
        DocxDiagnostic diagnostic = Assert.Single(patch.Diagnostics);
        Assert.Equal("E2004", diagnostic.Code);
        Assert.Equal(5, diagnostic.Line);
        Assert.Equal(1, diagnostic.Column);
    }

    [Fact]
    public static void ParsePatchRejectsUnterminatedHeredoc()
    {
        var editor = new DocxEditor();

        DocxPatch patch = editor.ParsePatch(new StringReader("""
            docxpatch 1

            op replace-text
            with <<<
            hello
            end
            """));

        Assert.False(patch.Success);
        DocxDiagnostic diagnostic = Assert.Single(patch.Diagnostics);
        Assert.Equal("E2008", diagnostic.Code);
        Assert.Equal(5, diagnostic.Line);
        Assert.Equal(1, diagnostic.Column);
    }

    [Fact]
    public static void ParsePatchParsesOperationWithHeredoc()
    {
        var editor = new DocxEditor();

        DocxPatch patch = editor.ParsePatch(new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            with <<<
            hello
            world
            >>>
            end
            """));

        Assert.True(patch.Success);
        Assert.Single(patch.Operations);
        Assert.Equal("replace-text", patch.Operations[0].OperationName);
        Assert.Equal("hello\nworld", patch.Operations[0].Fields["with"]);
    }

    [Fact]
    public static void ParsePatchDecodesQuotedFieldValueEscapes()
    {
        var editor = new DocxEditor();

        DocxPatch patch = editor.ParsePatch(new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find A
            with "x\ny\tz\\w\"v"
            end
            """));

        Assert.True(patch.Success);
        Assert.Equal("M.P0001", patch.Operations[0].Fields["target"]);
        Assert.Equal("x\ny\tz\\w\"v", patch.Operations[0].Fields["with"]);
    }

    [Fact]
    public static void ParsePatchLeavesBareAndUnbalancedValuesVerbatim()
    {
        var editor = new DocxEditor();

        DocxPatch bare = editor.ParsePatch(new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find A
            with C:\new
            end
            """));
        Assert.True(bare.Success);
        Assert.Equal(@"C:\new", bare.Operations[0].Fields["with"]);

        DocxPatch unbalanced = editor.ParsePatch(new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find A
            with "abc
            end
            """));
        Assert.True(unbalanced.Success);
        Assert.Equal("\"abc", unbalanced.Operations[0].Fields["with"]);

        DocxPatch empty = editor.ParsePatch(new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find A
            with ""
            end
            """));
        Assert.True(empty.Success);
        Assert.Equal("", empty.Operations[0].Fields["with"]);
    }

    [Fact]
    public static void ParsePatchPreservesRepeatedFieldsInOrder()
    {
        var editor = new DocxEditor();

        DocxPatch patch = editor.ParsePatch(new StringReader("""
            docxpatch 1

            op append-row
            target M.T0001
            cell North
            cell Revenue
            end
            """));

        Assert.True(patch.Success);
        DocxPatchOperation operation = Assert.Single(patch.Operations);
        Assert.Equal(new[] { "target", "cell", "cell" }, operation.FieldValues.Select(field => field.Name));
        Assert.Equal(new[] { "M.T0001", "North", "Revenue" }, operation.FieldValues.Select(field => field.Value));
        Assert.Equal(new[] { 4, 5, 6 }, operation.FieldValues.Select(field => field.Line));
        Assert.Equal("Revenue", operation.Fields["cell"]);
    }

    [Fact]
    public static void ParsePatchRejectsUnknownOperationName()
    {
        var editor = new DocxEditor();

        DocxPatch patch = editor.ParsePatch(new StringReader("""
            docxpatch 1

            op update-widget
            target M.P0001
            end
            """));

        Assert.False(patch.Success);
        DocxDiagnostic diagnostic = Assert.Single(patch.Diagnostics);
        Assert.Equal("E2010", diagnostic.Code);
        Assert.Equal(3, diagnostic.Line);
        Assert.Equal(4, diagnostic.Column);
    }

    [Fact]
    public static void ParsePatchRejectsUnknownFieldForKnownOperation()
    {
        var editor = new DocxEditor();

        DocxPatch patch = editor.ParsePatch(new StringReader("""
            docxpatch 1

            op set-cell
            target M.T0001.R01.C01
            color blue
            text New
            end
            """));

        Assert.False(patch.Success);
        DocxDiagnostic diagnostic = Assert.Single(patch.Diagnostics);
        Assert.Equal("E2011", diagnostic.Code);
        Assert.Equal(5, diagnostic.Line);
        Assert.Equal(1, diagnostic.Column);
    }

    [Fact]
    public static void ParsePatchRejectsInvalidBooleanField()
    {
        var editor = new DocxEditor();

        DocxPatch patch = editor.ParsePatch(new StringReader("""
            docxpatch 1

            op set-cell
            target M.T0001.R01.C01
            force yes
            text New
            end
            """));

        Assert.False(patch.Success);
        Assert.Contains(patch.Diagnostics, diagnostic => diagnostic.Code == "E2012");
    }

    [Fact]
    public static void ParsePatchRejectsInvalidIntegerField()
    {
        var editor = new DocxEditor();

        DocxPatch patch = editor.ParsePatch(new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            occurrence second
            find old
            with new
            end
            """));

        Assert.False(patch.Success);
        Assert.Contains(patch.Diagnostics, diagnostic => diagnostic.Code == "E2013");
    }

    [Theory]
    [InlineData("replace-image", "preserve-size true")]
    [InlineData("insert-image-after", "caption Figure 1")]
    public static void ParsePatchRejectsUnimplementedImageFields(string operationName, string unsupportedField)
    {
        var editor = new DocxEditor();

        DocxPatch patch = editor.ParsePatch(new StringReader($"""
            docxpatch 1

            op {operationName}
            target M.I0001
            asset chart.png
            {unsupportedField}
            end
            """));

        Assert.False(patch.Success);
        Assert.Contains(patch.Diagnostics, diagnostic => diagnostic.Code == "E2011");
    }

    [Fact]
    public static void ParsePatchKeepsExpectHashPrefixedHeredocTextLiteral()
    {
        var editor = new DocxEditor();

        DocxPatch patch = editor.ParsePatch(new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            with <<<
            expect-hash abc
            >>>
            end
            """));

        Assert.True(patch.Success);
        Assert.Equal("expect-hash abc", patch.Operations[0].Fields["with"]);
    }

    [Fact]
    public static void ParsePatchRecordsHeredocFieldStartLine()
    {
        var editor = new DocxEditor();

        DocxPatch patch = editor.ParsePatch(new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            with <<<
            hello
            world
            >>>
            end
            """));

        Assert.True(patch.Success);
        DocxPatchOperation operation = Assert.Single(patch.Operations);
        Assert.Equal("target", operation.FieldValues[0].Name);
        Assert.Equal(4, operation.FieldValues[0].Line);
        Assert.Equal("with", operation.FieldValues[1].Name);
        Assert.Equal(5, operation.FieldValues[1].Line);
    }

    [Fact]
    public static void ParsePatchReportsHeredocValueErrorAtFieldStartLine()
    {
        var editor = new DocxEditor();

        DocxPatch patch = editor.ParsePatch(new StringReader("""
            docxpatch 1

            op append-row
            target M.T0001
            expect-row-count <<<
            two
            >>>
            end
            """));

        Assert.False(patch.Success);
        DocxDiagnostic diagnostic = Assert.Single(patch.Diagnostics);
        Assert.Equal("E2013", diagnostic.Code);
        Assert.Equal(5, diagnostic.Line);
        Assert.Equal(1, diagnostic.Column);
    }

    [Fact]
    public static void ParsePatchKeepsPerOccurrenceLinesForRepeatedFields()
    {
        var editor = new DocxEditor();

        DocxPatch patch = editor.ParsePatch(new StringReader("""
            docxpatch 1

            op append-row
            target M.T0001
            cell A
            cell <<<
            B
            >>>
            end
            """));

        Assert.True(patch.Success);
        DocxPatchOperation operation = Assert.Single(patch.Operations);
        Assert.Equal("B", operation.Fields["cell"]);
        Assert.Equal(3, operation.FieldValues.Count);
        Assert.Equal(5, operation.FieldValues[1].Line);
        Assert.Equal(6, operation.FieldValues[2].Line);
    }


    [Fact]
    public static void ParsePatchRejectsRepeatedTargetField()
    {
        var editor = new DocxEditor();

        DocxPatch patch = editor.ParsePatch(new StringReader("""
            docxpatch 1

            op replace-paragraph
            target M.P0001
            target M.P0002
            text Changed
            end
            """));

        Assert.False(patch.Success);
        DocxDiagnostic diagnostic = Assert.Single(patch.Diagnostics);
        Assert.Equal("E2015", diagnostic.Code);
        Assert.Equal(5, diagnostic.Line);
        Assert.Equal(1, diagnostic.Column);
    }

    [Fact]
    public static void ParsePatchRejectsRepeatedFindField()
    {
        var editor = new DocxEditor();

        DocxPatch patch = editor.ParsePatch(new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            find Alpha
            find Beta
            with Omega
            end
            """));

        Assert.False(patch.Success);
        DocxDiagnostic diagnostic = Assert.Single(patch.Diagnostics);
        Assert.Equal("E2015", diagnostic.Code);
        Assert.Equal(6, diagnostic.Line);
        Assert.Equal(1, diagnostic.Column);
    }

    [Fact]
    public static void ParsePatchRejectsRepeatedGuardField()
    {
        var editor = new DocxEditor();

        DocxPatch patch = editor.ParsePatch(new StringReader("""
            docxpatch 1

            op replace-text
            target M.P0001
            expect-text Alpha
            expect-text Beta
            find Alpha
            with Omega
            end
            """));

        Assert.False(patch.Success);
        DocxDiagnostic diagnostic = Assert.Single(patch.Diagnostics);
        Assert.Equal("E2015", diagnostic.Code);
        Assert.Equal(6, diagnostic.Line);
        Assert.Equal(1, diagnostic.Column);
    }

}
