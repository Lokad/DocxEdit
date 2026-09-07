using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;

namespace Lokad.DocxEdit.OfficeTests;

public static class OfficeGateTests
{
    [Fact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationOpenSaveRoundTripIsOptIn()
    {
        if (!string.Equals(Environment.GetEnvironmentVariable("DOCXEDIT_ENABLE_OFFICE_TESTS"), "1", StringComparison.Ordinal))
        {
            return;
        }

        if (!OperatingSystem.IsWindows())
        {
            throw new InvalidOperationException("Office integration tests require Windows.");
        }

        Type? wordApplicationType = Type.GetTypeFromProgID("Word.Application");
        if (wordApplicationType is null)
        {
            throw new InvalidOperationException("Microsoft Word is not installed or is not available through COM.");
        }

        string directory = Path.Combine(Path.GetTempPath(), "docxedit-office-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);

        string inputPath = Path.Combine(directory, "input.docx");
        string outputPath = Path.Combine(directory, "output.docx");

        try
        {
            CreateDocx(inputPath, "Office smoke input");

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op replace-text
                target M.P0001
                find input
                with output
                end

                op insert-after
                target M.P0001
                text Office inserted tracked paragraph
                end

                op append-row
                target M.T0001
                cell East
                cell West
                end

                op insert-image-after
                target M.P0001
                asset chart.png
                alt Office chart
                end
                """))
            using (FileStream output = File.Create(outputPath))
            {
                DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
                {
                    TrackChanges = TrackChangesMode.Suggest,
                    Author = "Office Reviewer",
                    TimestampUtc = DateTimeOffset.Parse("2026-06-11T12:00:00Z").ToUniversalTime(),
                    AssetProvider = new MemoryAssetProvider("chart.png", "office-png", "chart.png"),
                    MarkFieldsDirtyWhenEditing = false
                });
                Assert.True(result.Success, string.Join(Environment.NewLine, result.Diagnostics.Select(FormatDiagnostic)));
                Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error);
            }

            OpenSaveWithWord(wordApplicationType, outputPath);

            using FileStream saved = File.OpenRead(outputPath);
            DocxReadResult read = new DocxEditor().Read(saved);
            Assert.True(read.Success, string.Join(Environment.NewLine, read.Diagnostics.Select(FormatDiagnostic)));
            Assert.Contains(read.Paragraphs, paragraph => paragraph.Text == "Office smoke output");
            Assert.Contains(read.Paragraphs, paragraph => paragraph.Text == "Office inserted tracked paragraph");
            Assert.Contains(read.Tables.SelectMany(table => table.Cells), cell => cell.Text == "East");
            Assert.Contains(read.Tables.SelectMany(table => table.Cells), cell => cell.Text == "West");
            Assert.Single(read.Images);

            using FileStream savedChanges = File.OpenRead(outputPath);
            DocxChangesResult changes = new DocxEditor().Changes(savedChanges);
            Assert.True(changes.Success, string.Join(Environment.NewLine, changes.Diagnostics.Select(FormatDiagnostic)));
            Assert.Contains(changes.Summary, summary => summary.Type == "deleted-run" && summary.Count == 1);
            Assert.Contains(changes.Summary, summary => summary.Type == "inserted-run" && summary.Count >= 2);
            Assert.Contains(changes.Summary, summary => summary.Type == "row-inserted" && summary.Count == 1);
        }
        finally
        {
            try
            {
                Directory.Delete(directory, recursive: true);
            }
            catch
            {
                // Keep the test failure focused on Office/docx behavior if cleanup is blocked.
            }
        }
    }

    private static void OpenSaveWithWord(Type wordApplicationType, string path)
    {
        object? application = null;
        object? document = null;

        try
        {
            application = Activator.CreateInstance(wordApplicationType)
                ?? throw new InvalidOperationException("Could not create Word.Application.");

            dynamic word = application;
            word.Visible = false;
            word.DisplayAlerts = 0;

            dynamic documents = word.Documents;
            document = documents.Open(path, ReadOnly: false, AddToRecentFiles: false, Visible: false);

            dynamic doc = document;
            doc.Save();
            doc.Close(SaveChanges: false);
            document = null;

            word.Quit(SaveChanges: false);
        }
        finally
        {
            if (document is not null)
            {
                try
                {
                    ((dynamic)document).Close(SaveChanges: false);
                }
                catch
                {
                }

                ReleaseComObject(document);
            }

            if (application is not null)
            {
                try
                {
                    ((dynamic)application).Quit(SaveChanges: false);
                }
                catch
                {
                }

                ReleaseComObject(application);
            }
        }
    }

    private static void ReleaseComObject(object value)
    {
        if (OperatingSystem.IsWindows() && Marshal.IsComObject(value))
        {
            Marshal.FinalReleaseComObject(value);
        }
    }

    private static void CreateDocx(string path, string text)
    {
        using FileStream file = File.Create(path);
        using var archive = new ZipArchive(file, ZipArchiveMode.Create);

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
                <w:p><w:r><w:t>{{text}}</w:t></w:r></w:p>
                <w:tbl>
                  <w:tr>
                    <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                    <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                  </w:tr>
                </w:tbl>
              </w:body>
            </w:document>
            """);
    }

    private static void AddEntry(ZipArchive archive, string name, string text)
    {
        ZipArchiveEntry entry = archive.CreateEntry(name);
        using Stream stream = entry.Open();
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        stream.Write(bytes, 0, bytes.Length);
    }

    private sealed class MemoryAssetProvider : IDocxAssetProvider
    {
        private readonly string reference;
        private readonly byte[] bytes;
        private readonly string fileNameHint;

        public MemoryAssetProvider(string reference, string text, string fileNameHint)
        {
            this.reference = reference;
            bytes = Encoding.UTF8.GetBytes(text);
            this.fileNameHint = fileNameHint;
        }

        public bool TryOpen(string requestedReference, out Stream stream, out string? contentTypeHint, out string? fileNameHint)
        {
            if (!string.Equals(reference, requestedReference, StringComparison.Ordinal))
            {
                stream = Stream.Null;
                contentTypeHint = null;
                fileNameHint = null;
                return false;
            }

            stream = new MemoryStream(bytes, writable: false);
            contentTypeHint = null;
            fileNameHint = this.fileNameHint;
            return true;
        }
    }

    private static string FormatDiagnostic(DocxDiagnostic diagnostic) => $"{diagnostic.Code}: {diagnostic.Message}";
}
