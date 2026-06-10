namespace DocxEdit.Tests;

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
        Assert.Contains(patch.Diagnostics, diagnostic => diagnostic.Code == "E2004");
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
}
