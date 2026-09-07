using System.IO.Compression;
using System.Text;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

public static class OoxmlUtilityTests
{
    [Theory]
    [InlineData("1in", 914400)]
    [InlineData("2.54cm", 914400)]
    [InlineData("72pt", 914400)]
    [InlineData("96px", 914400)]
    [InlineData("123emu", 123)]
    public static void ImageSizeDimensionsConvertToExpectedEmus(string dimension, long expectedEmus)
    {
        using MemoryStream stream = CreateDocxWithImage();
        using var output = new MemoryStream();
        using var patch = new StringReader($"""
            docxpatch 1

            op set-image-size
            target M.I0001
            width {dimension}
            height {dimension}
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(stream, patch, output);

        Assert.True(result.Success);
        output.Position = 0;
        string documentXml = ReadEntry(output, "word/document.xml");
        Assert.Contains($"cx=\"{expectedEmus}\"", documentXml, StringComparison.Ordinal);
        Assert.Contains($"cy=\"{expectedEmus}\"", documentXml, StringComparison.Ordinal);
    }

    [Fact]
    public static void InsertImageAllocatesFreshRelationshipIdAndMediaPartName()
    {
        using MemoryStream stream = CreateDocxWithMediaPair();
        using var output = new MemoryStream();
        using var patch = new StringReader("""
            docxpatch 1

            op insert-image-after
            target M.P0001
            asset new.png
            width 1in
            end
            """);

        DocxApplyResult result = new DocxEditor().Apply(stream, patch, output, new DocxEditOptions { AssetProvider = SingleAssetProvider.Png });

        Assert.True(result.Success);
        output.Position = 0;
        string relationshipsXml = ReadEntry(output, "word/_rels/document.xml.rels");
        Assert.Contains("rId3", relationshipsXml, StringComparison.Ordinal);
        Assert.True(HasEntry(output, "word/media/image3.png"));
    }

    [Fact]
    public static void StyleDefaultFlagsParseCommonWordprocessingValues()
    {
        using MemoryStream stream = CreateDocxWithStyleFlags();
        var editor = new DocxEditor();

        DocxStylesResult result = editor.Styles(stream);

        Assert.True(result.Success);
        IReadOnlyDictionary<string, bool> flags = result.Styles.ToDictionary(style => style.StyleId, style => style.IsDefault, StringComparer.Ordinal);
        Assert.Equal(new Dictionary<string, bool>(StringComparer.Ordinal)
        {
            ["S01"] = false,
            ["S02"] = true,
            ["S03"] = true,
            ["S04"] = true,
            ["S05"] = true,
            ["S06"] = true,
            ["S07"] = false,
            ["S08"] = false,
            ["S09"] = false,
            ["S10"] = false,
            ["S11"] = false,
        }, flags);
    }

    private static MemoryStream CreateDocxWithImage()
    {
        return CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", PackageRelationshipsXml());
            AddEntry(archive, "word/document.xml", DocumentXmlWithImage());
            AddEntry(archive, "word/_rels/document.xml.rels", RelationshipsXml("media/image1.png"));
            AddEntry(archive, "word/media/image1.png", "png");
        });
    }

    private static MemoryStream CreateDocxWithMediaPair()
    {
        return CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", PackageRelationshipsXml());
            AddEntry(archive, "word/document.xml", DocumentXml());
            AddEntry(archive, "word/_rels/document.xml.rels", """
                <?xml version="1.0" encoding="utf-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/>
                  <Relationship Id="custom" Type="http://example.test/custom" Target="other.bin"/>
                </Relationships>
                """);
            AddEntry(archive, "word/media/image1.png", "png1");
            AddEntry(archive, "word/media/image2.png", "png2");
            AddEntry(archive, "word/other.bin", "other");
        });
    }

    private static MemoryStream CreateDocxWithStyleFlags()
    {
        return CreatePackage(archive =>
        {
            AddEntry(archive, "[Content_Types].xml", ContentTypesXml());
            AddEntry(archive, "_rels/.rels", PackageRelationshipsXml());
            AddEntry(archive, "word/document.xml", DocumentXml());
            AddEntry(archive, "word/_rels/document.xml.rels", StylesRelationshipsXml());
            AddEntry(archive, "word/styles.xml", StylesXmlWithFlags());
        });
    }

    private sealed class SingleAssetProvider : IDocxAssetProvider
    {
        public static readonly SingleAssetProvider Png = new();

        public bool TryOpen(
            string reference,
            out Stream stream,
            out string? contentTypeHint,
            out string? fileNameHint)
        {
            // Minimal structurally valid PNG: signature, IHDR, empty IDAT, IEND.
            byte[] bytes =
            [
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
                0x00, 0x00, 0x00, 0x04, 0x00, 0x00, 0x00, 0x03,
                0x08, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00, 0x00, 0x49, 0x44, 0x41,
                0x54, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60,
                0x82
            ];
            stream = new MemoryStream(bytes, writable: false);
            contentTypeHint = "image/png";
            fileNameHint = "new.png";
            return true;
        }
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



    private static bool HasEntry(MemoryStream docx, string entryName)
    {
        docx.Position = 0;
        using var archive = new ZipArchive(docx, ZipArchiveMode.Read, leaveOpen: true);
        return archive.GetEntry(entryName) is not null;
    }

    private static string ContentTypesXml()
    {
        return """
            <?xml version="1.0" encoding="utf-8"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
              <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
            </Types>
            """;
    }

    private static string PackageRelationshipsXml()
    {
        return """
            <?xml version="1.0" encoding="utf-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """;
    }

    private static string RelationshipsXml(string imageTarget)
    {
        return $$"""
            <?xml version="1.0" encoding="utf-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{{imageTarget}}"/>
            </Relationships>
            """;
    }

    private static string StylesRelationshipsXml()
    {
        return """
            <?xml version="1.0" encoding="utf-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
            </Relationships>
            """;
    }

    private static string DocumentXml()
    {
        return """
            <?xml version="1.0" encoding="utf-8"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
              </w:body>
            </w:document>
            """;
    }

    private static string DocumentXmlWithImage()
    {
        return """
            <?xml version="1.0" encoding="utf-8"?>
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
                        <wp:extent cx="914400" cy="914400"/>
                        <wp:docPr id="1" name="Picture"/>
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
            """;
    }

    private static string StylesXmlWithFlags()
    {
        return """
            <?xml version="1.0" encoding="utf-8"?>
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:style w:type="paragraph" w:styleId="S01"><w:name w:val="S01"/></w:style>
              <w:style w:type="paragraph" w:styleId="S02" w:default="1"><w:name w:val="S02"/></w:style>
              <w:style w:type="paragraph" w:styleId="S03" w:default="true"><w:name w:val="S03"/></w:style>
              <w:style w:type="paragraph" w:styleId="S04" w:default="TRUE"><w:name w:val="S04"/></w:style>
              <w:style w:type="paragraph" w:styleId="S05" w:default="on"><w:name w:val="S05"/></w:style>
              <w:style w:type="paragraph" w:styleId="S06" w:default="ON"><w:name w:val="S06"/></w:style>
              <w:style w:type="paragraph" w:styleId="S07" w:default="0"><w:name w:val="S07"/></w:style>
              <w:style w:type="paragraph" w:styleId="S08" w:default="false"><w:name w:val="S08"/></w:style>
              <w:style w:type="paragraph" w:styleId="S09" w:default="off"><w:name w:val="S09"/></w:style>
              <w:style w:type="paragraph" w:styleId="S10" w:default="OFF"><w:name w:val="S10"/></w:style>
              <w:style w:type="paragraph" w:styleId="S11" w:default=""><w:name w:val="S11"/></w:style>
            </w:styles>
            """;
    }
}
