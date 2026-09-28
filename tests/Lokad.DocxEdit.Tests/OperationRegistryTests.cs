using System.IO.Compression;
using System.Text;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

public static class OperationRegistryTests
{
    [Fact]
    public static void CatalogAndParserAgreeOnEveryField()
    {
        var editor = new DocxEditor();

        foreach (DocxPatchOperationInfo operation in DocxHelp.Catalog.PatchOperations)
        {
            DocxPatch known = editor.ParsePatch(new StringReader($"""
                docxpatch 1

                op {operation.Name}
                end
                """));
            Assert.DoesNotContain(known.Diagnostics, diagnostic => diagnostic.Code == "E2010");

            var catalogued = new HashSet<string>(
                operation.RequiredFields
                    .Concat(operation.RequiredAlternatives.SelectMany(group => group))
                    .Concat(operation.OptionalFields)
                    .Concat(operation.RepeatableFields),
                StringComparer.Ordinal);
            Assert.NotEmpty(catalogued);
            Assert.All(catalogued, field => Assert.DoesNotContain(" ", field));

            foreach (string field in catalogued)
            {
                DocxPatch parsed = editor.ParsePatch(new StringReader($"""
                    docxpatch 1

                    op {operation.Name}
                    {field} v
                    end
                    """));
                Assert.DoesNotContain(parsed.Diagnostics, diagnostic => diagnostic.Code == "E2010");
                Assert.DoesNotContain(parsed.Diagnostics, diagnostic => diagnostic.Code == "E2011");

                if (parsed.Success)
                {
                    DocxPatchOperation parsedOperation = Assert.Single(parsed.Operations);
                    DocxPatchField parsedField = Assert.Single(parsedOperation.FieldValues);
                    Assert.Equal(field, parsedField.Name);
                    Assert.Equal(4, parsedField.Line);
                }
                else
                {
                    DocxDiagnostic valueError = Assert.Single(
                        parsed.Diagnostics,
                        diagnostic => diagnostic.Code == "E2012" || diagnostic.Code == "E2013");
                    Assert.Equal(4, valueError.Line);
                }
            }
        }

        DocxPatch unknown = editor.ParsePatch(new StringReader("""
            docxpatch 1

            op no-such-operation
            end
            """));
        Assert.Contains(unknown.Diagnostics, diagnostic => diagnostic.Code == "E2010");
    }

    [Fact]
    public static void ApplyAppendRowMarksFieldsDirty()
    {
        using MemoryStream stream = CreateDocxWithTable();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op append-row
            target M.T0001
            cell East
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(stream, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        string settings = ReadEntry(output, "word/settings.xml");
        Assert.Contains("updateFields", settings, StringComparison.Ordinal);
    }

    private static MemoryStream CreateDocxWithTable()
    {
        return CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", """
                <?xml version="1.0" encoding="utf-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """);
            AddEntry(archive, "_rels/.rels", """
                <?xml version="1.0" encoding="utf-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            AddEntry(archive, "word/document.xml", """
                <?xml version="1.0" encoding="utf-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tr><w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                  </w:body>
                </w:document>
                """);
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <?xml version="1.0" encoding="utf-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
                """);
        });
    }

    private static MemoryStream CreatePackage(Action<ZipArchive> configure)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            configure(archive);
        }

        stream.Position = 0;
        return stream;
    }



    [Fact]
    public static void PatchOperationTablesRenderStructuredFields()
    {
        string tables = DocxHelp.RenderPatchOperationTables();

        Assert.Contains("`cell`+", tables, StringComparison.Ordinal);
        Assert.DoesNotContain("target plus", tables, StringComparison.Ordinal);
        Assert.DoesNotContain("repeated cell", tables, StringComparison.Ordinal);
        Assert.DoesNotContain("plus one crop", tables, StringComparison.Ordinal);
    }

    [Fact]
    public static void PatchHelpListsStructuredAlternatives()
    {
        Assert.True(DocxHelp.TryRenderTopic("patch", out string patchHelp));
        Assert.Contains("left-percent|top-percent|right-percent|bottom-percent", patchHelp, StringComparison.Ordinal);
        Assert.DoesNotContain("target plus one crop percentage", patchHelp, StringComparison.Ordinal);
    }

    [Fact]
    public static void CatalogExposesEmptyAllowedFields()
    {
        IReadOnlyDictionary<string, DocxPatchOperationInfo> catalogued = DocxHelp.Catalog.PatchOperations.ToDictionary(static operation => operation.Name);

        Assert.Equal(["with"], catalogued["replace-text"].EmptyAllowedFields);
        Assert.Equal(["text"], catalogued["replace-paragraph"].EmptyAllowedFields);
        Assert.Equal(["text"], catalogued["set-cell"].EmptyAllowedFields);
        Assert.Empty(catalogued["insert-after"].EmptyAllowedFields);
        Assert.Empty(catalogued["set-style"].EmptyAllowedFields);
        Assert.Empty(catalogued["add-comment"].EmptyAllowedFields);
    }

}
