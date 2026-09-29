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
                    AssetProvider = new MemoryAssetProvider("chart.png", MinimalPng(), "chart.png"),
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
    [Fact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationIterativeTrackedRoundTripIsOptIn()
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
            CreateRichDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op replace-text
                target M.P0001
                find Alpha
                with Omega
                occurrence 1
                end

                op add-comment
                target M.P0001
                text Review note
                end
                """))
            using (FileStream output = File.Create(outputPath))
            {
                DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
                {
                    TrackChanges = TrackChangesMode.Require,
                    Author = "Office Reviewer",
                    TimestampUtc = DateTimeOffset.Parse("2026-06-11T12:00:00Z").ToUniversalTime(),
                    MarkFieldsDirtyWhenEditing = false
                });
                Assert.True(result.Success, string.Join(Environment.NewLine, result.Diagnostics.Select(FormatDiagnostic)));
                Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error);
            }

            OpenSaveWithWord(wordApplicationType, outputPath);

            using FileStream saved = File.OpenRead(outputPath);
            DocxReadResult read = new DocxEditor().Read(saved, new DocxReadOptions { IncludeHeadersFooters = true });
            Assert.True(read.Success, string.Join(Environment.NewLine, read.Diagnostics.Select(FormatDiagnostic)));
            Assert.Contains(read.Paragraphs, paragraph => paragraph.Text == "Omega Alpha");
            Assert.Contains(read.Paragraphs, paragraph => paragraph.Text == "Header line");
            Assert.NotEmpty(read.Fields);

            using FileStream savedChanges = File.OpenRead(outputPath);
            DocxChangesResult changes = new DocxEditor().Changes(savedChanges);
            Assert.True(changes.Success, string.Join(Environment.NewLine, changes.Diagnostics.Select(FormatDiagnostic)));
            Assert.Contains(changes.Summary, summary => summary.Type == "inserted-run" && summary.Count >= 2);
            Assert.NotEmpty(changes.CommentSummary);
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

    [Fact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationBreakStructureRoundTripIsOptIn()
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
            CreateDocx(inputPath, "Alpha");

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op replace-paragraph
                target M.P0001
                text "Line one\nLine two\tTab"
                end

                op insert-after
                target M.P0001
                text "Second\nPart"
                end
                """))
            using (FileStream output = File.Create(outputPath))
            {
                DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
                Assert.True(result.Success, string.Join(Environment.NewLine, result.Diagnostics.Select(FormatDiagnostic)));
                Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error);
            }

            OpenSaveWithWord(wordApplicationType, outputPath);

            using FileStream saved = File.OpenRead(outputPath);
            DocxReadResult read = new DocxEditor().Read(saved);
            Assert.True(read.Success, string.Join(Environment.NewLine, read.Diagnostics.Select(FormatDiagnostic)));
            Assert.Equal("Line one\nLine two\tTab", read.Paragraphs[0].Text);
            Assert.Equal("Second\nPart", read.Paragraphs[1].Text);

            using FileStream xmlStream = File.OpenRead(outputPath);
            using var archive = new ZipArchive(xmlStream, ZipArchiveMode.Read);
            ZipArchiveEntry entry = archive.GetEntry("word/document.xml") ?? throw new InvalidOperationException("Missing word/document.xml.");
            using Stream entryStream = entry.Open();
            using var reader = new StreamReader(entryStream, Encoding.UTF8);
            string xml = reader.ReadToEnd();
            Assert.Equal(2, CountOccurrences(xml, "<w:br"));
            Assert.Equal(1, CountOccurrences(xml, "<w:tab"));
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


    [Fact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationTrackedRowRoundTripIsOptIn()
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
            CreateDocx(inputPath, "Office rows");

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op append-row
                target M.T0001
                cell East
                cell West
                end
                """))
            using (FileStream output = File.Create(outputPath))
            {
                DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
                {
                    TrackChanges = TrackChangesMode.Require,
                    Author = "Office Reviewer",
                    TimestampUtc = DateTimeOffset.Parse("2026-06-11T12:00:00Z").ToUniversalTime()
                });
                Assert.True(result.Success, string.Join(Environment.NewLine, result.Diagnostics.Select(FormatDiagnostic)));
                Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error);
            }

            OpenSaveWithWord(wordApplicationType, outputPath);

            using FileStream saved = File.OpenRead(outputPath);
            DocxReadResult read = new DocxEditor().Read(saved);
            Assert.True(read.Success, string.Join(Environment.NewLine, read.Diagnostics.Select(FormatDiagnostic)));
            DocxTableInfo table = Assert.Single(read.Tables);
            Assert.Equal(2, table.RowCount);
            Assert.Contains(table.Cells, cell => cell.Text == "East");
            Assert.Contains(table.Cells, cell => cell.Text == "West");

            using FileStream savedChanges = File.OpenRead(outputPath);
            DocxChangesResult changes = new DocxEditor().Changes(savedChanges);
            Assert.True(changes.Success, string.Join(Environment.NewLine, changes.Diagnostics.Select(FormatDiagnostic)));
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
    [Fact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationTrackedPropertyRoundTripIsOptIn()
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
            CreatePropertyDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op set-cell-shading
                target M.T0001.R01.C01
                expect-fill none
                fill 4472C4
                end

                op set-section-columns
                target M.S0001
                count 2
                end
                """))
            using (FileStream output = File.Create(outputPath))
            {
                DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
                {
                    TrackChanges = TrackChangesMode.Require,
                    Author = "Office Reviewer",
                    TimestampUtc = DateTimeOffset.Parse("2026-06-11T12:00:00Z").ToUniversalTime()
                });
                Assert.True(result.Success, string.Join(Environment.NewLine, result.Diagnostics.Select(FormatDiagnostic)));
                Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error);
            }

            OpenSaveWithWord(wordApplicationType, outputPath);

            using FileStream saved = File.OpenRead(outputPath);
            DocxReadResult read = new DocxEditor().Read(saved);
            Assert.True(read.Success, string.Join(Environment.NewLine, read.Diagnostics.Select(FormatDiagnostic)));
            Assert.Equal(2, Assert.Single(read.Sections).Columns);

            using FileStream savedChanges = File.OpenRead(outputPath);
            DocxChangesResult changes = new DocxEditor().Changes(savedChanges);
            Assert.True(changes.Success, string.Join(Environment.NewLine, changes.Diagnostics.Select(FormatDiagnostic)));
            Assert.Contains(changes.Summary, summary => summary.Type == "cell-properties-change" && summary.Count == 1);
            Assert.Contains(changes.Summary, summary => summary.Type == "section-properties-change" && summary.Count == 1);
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
        object? documents = null;
        object? document = null;

        try
        {
            application = Activator.CreateInstance(wordApplicationType)
                ?? throw new InvalidOperationException("Could not create Word.Application.");

            dynamic word = application;
            word.Visible = false;
            word.DisplayAlerts = 0;

            documents = word.Documents;
            document = ((dynamic)documents).Open(path, ReadOnly: false, AddToRecentFiles: false, Visible: false);

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

            if (documents is not null)
            {
                ReleaseComObject(documents);
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

    private static void CreateRichDocx(string path)
    {
        using FileStream file = File.Create(path);
        using var archive = new ZipArchive(file, ZipArchiveMode.Create);

        AddEntry(archive, "[Content_Types].xml", """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
              <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
              <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
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
            </Relationships>
            """);
        AddEntry(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <w:body>
                <w:p><w:r><w:t>Alpha Alpha</w:t></w:r></w:p>
                <w:p><w:ins w:id="9" w:author="Main" w:date="2026-06-01T00:00:00Z"><w:r><w:t>Beta</w:t></w:r></w:ins></w:p>
                <w:p>
                  <w:bookmarkStart w:id="1" w:name="ClientName"/>
                  <w:r><w:t>Acme Corp</w:t></w:r>
                  <w:bookmarkEnd w:id="1"/>
                </w:p>
                <w:p>
                  <w:fldSimple w:instr=" REF ClientName \h ">
                    <w:r><w:t>Acme Corp</w:t></w:r>
                  </w:fldSimple>
                </w:p>
                <w:sectPr><w:headerReference w:type="default" r:id="rHeader"/></w:sectPr>
              </w:body>
            </w:document>
            """);
        AddEntry(archive, "word/header1.xml", """
            <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:p><w:r><w:t>Header line</w:t></w:r></w:p>
            </w:hdr>
            """);
        AddEntry(archive, "word/comments.xml", """
            <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
            """);
    }

    private static void CreatePropertyDocx(string path)
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
        AddEntry(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:tbl>
                  <w:tr>
                    <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                    <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                  </w:tr>
                </w:tbl>
                <w:sectPr>
                  <w:pgSz w:w="12240" w:h="15840"/>
                  <w:cols w:num="1"/>
                </w:sectPr>
              </w:body>
            </w:document>
            """);
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

    private static byte[] MinimalPng()
    {
        using var stream = new MemoryStream();
        stream.Write(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A });
        WritePngChunk(stream, "IHDR", [0, 0, 0, 2, 0, 0, 0, 2, 0x08, 0x02, 0x00, 0x00, 0x00]);
        WritePngChunk(stream, "IDAT", []);
        WritePngChunk(stream, "IEND", []);
        return stream.ToArray();
    }

    private static void WritePngChunk(MemoryStream stream, string type, byte[] data)
    {
        stream.Write([(byte)((data.Length >> 24) & 0xFF), (byte)((data.Length >> 16) & 0xFF), (byte)((data.Length >> 8) & 0xFF), (byte)(data.Length & 0xFF)]);
        stream.Write(Encoding.ASCII.GetBytes(type));
        stream.Write(data);
        stream.Write([0, 0, 0, 0]);
    }

    private sealed class MemoryAssetProvider : IDocxAssetProvider
    {
        private readonly string reference;
        private readonly byte[] bytes;
        private readonly string fileNameHint;

        public MemoryAssetProvider(string reference, byte[] imageBytes, string fileNameHint)
        {
            this.reference = reference;
            bytes = imageBytes;
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
