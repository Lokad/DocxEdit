using System.IO.Compression;
using System.Text;

namespace Lokad.DocxEdit.Tests;

public static class OperationRegistryTests
{
    [Fact]
    public static void CatalogOperationsParseWithoutUnknownOperationOrFieldErrors()
    {
        var editor = new DocxEditor();
        foreach (DocxPatchOperationInfo operation in DocxHelp.Catalog.PatchOperations)
        {
            DocxPatch known = editor.ParsePatch(new StringReader($"docxpatch 1\n\nop {operation.Name}\nend\n"));
            Assert.DoesNotContain(known.Diagnostics, diagnostic => diagnostic.Code == "E2010");

            foreach (string field in operation.RequiredFields.Concat(operation.OptionalFields))
            {
                if (string.IsNullOrWhiteSpace(field) || field.Contains(' '))
                {
                    continue;
                }

                DocxPatch parsed = editor.ParsePatch(new StringReader($"docxpatch 1\n\nop {operation.Name}\n{field} v\nend\n"));
                Assert.DoesNotContain(parsed.Diagnostics, diagnostic => diagnostic.Code == "E2010");
                Assert.DoesNotContain(parsed.Diagnostics, diagnostic => diagnostic.Code == "E2011");
            }
        }

        DocxPatch unknown = editor.ParsePatch(new StringReader("docxpatch 1\n\nop no-such-operation\nend\n"));
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

    private static void AddEntry(ZipArchive archive, string name, string text)
    {
        ZipArchiveEntry entry = archive.CreateEntry(name);
        using Stream stream = entry.Open();
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static string ReadEntry(MemoryStream docx, string entryName)
    {
        docx.Position = 0;
        using var archive = new ZipArchive(docx, ZipArchiveMode.Read, leaveOpen: true);
        ZipArchiveEntry entry = archive.GetEntry(entryName)
            ?? throw new InvalidDataException($"Missing {entryName}.");
        using Stream stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    }
}
