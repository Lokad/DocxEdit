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

    [Fact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationTrackedStyleRoundTripIsOptIn()
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
            CreateStyledDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op set-style
                target M.P0001
                style heading 1
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
            Assert.Equal("Heading1", read.Paragraphs[0].StyleId);

            using FileStream savedChanges = File.OpenRead(outputPath);
            DocxChangesResult changes = new DocxEditor().Changes(savedChanges);
            Assert.True(changes.Success, string.Join(Environment.NewLine, changes.Diagnostics.Select(FormatDiagnostic)));
            Assert.Contains(changes.Summary, summary => summary.Type == "paragraph-properties-change" && summary.Count == 1);
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
    public static void OfficeAutomationResolveCommentRoundTripIsOptIn()
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
            CreateResolvableCommentDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op resolve-comment
                target comment:3
                end
                """))
            using (FileStream output = File.Create(outputPath))
            {
                DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
                Assert.True(result.Success, string.Join(Environment.NewLine, result.Diagnostics.Select(FormatDiagnostic)));
                Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error);
            }

            OpenSaveWithWord(wordApplicationType, outputPath);

            using FileStream savedChanges = File.OpenRead(outputPath);
            DocxChangesResult changes = new DocxEditor().Changes(savedChanges);
            Assert.True(changes.Success, string.Join(Environment.NewLine, changes.Diagnostics.Select(FormatDiagnostic)));
            DocxCommentThreadSummary summary = Assert.Single(changes.CommentSummary);
            Assert.Equal("3", summary.CommentId);
            Assert.True(summary.Resolved);
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
    public static void OfficeAutomationCommentReplyRoundTripIsOptIn()
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
            CreateResolvableCommentDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op add-comment-reply
                target comment:3
                text Reply body
                author Second Reviewer
                initials SR
                date 2026-06-08T09:30:00Z
                end
                """))
            using (FileStream output = File.Create(outputPath))
            {
                DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
                Assert.True(result.Success, string.Join(Environment.NewLine, result.Diagnostics.Select(FormatDiagnostic)));
                Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error);
            }

            OpenSaveWithWord(wordApplicationType, outputPath);

            using FileStream savedChanges = File.OpenRead(outputPath);
            DocxChangesResult changes = new DocxEditor().Changes(savedChanges, new DocxChangesOptions { IncludeCommentText = true });
            Assert.True(changes.Success, string.Join(Environment.NewLine, changes.Diagnostics.Select(FormatDiagnostic)));
            Assert.Contains(changes.CommentSummary, thread => thread.CommentId == "3");
            DocxCommentThreadSummary reply = Assert.Single(changes.CommentSummary, thread => thread.IsReply == true);
            Assert.Contains("Reply body", reply.TextSnippet, StringComparison.Ordinal);
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
    public static void OfficeAutomationFooterEditRoundTripIsOptIn()
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
            CreateFooterDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op replace-text
                target F001.P0001
                find old
                with new
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
            DocxReadResult read = new DocxEditor().Read(saved, new DocxReadOptions { IncludeHeadersFooters = true });
            Assert.True(read.Success, string.Join(Environment.NewLine, read.Diagnostics.Select(FormatDiagnostic)));
            Assert.Contains(read.Paragraphs, paragraph => paragraph.Id.ToWireValue() == "F001.P0001" && paragraph.Text == "Footer new line");
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
    public static void OfficeAutomationBookmarkTextRoundTripIsOptIn()
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
            CreateBookmarkDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op replace-bookmark-text
                target M.B0001
                text New Client
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
            Assert.Equal("Before New Client After", Assert.Single(read.Paragraphs).Text);
            DocxBookmarkInfo bookmark = Assert.Single(read.Bookmarks);
            Assert.Equal("ClientName", bookmark.Name);
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
    public static void OfficeAutomationContentControlRoundTripIsOptIn()
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
            CreateControlDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op set-content-control-text
                target M.CC0001
                text New Name
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
            Assert.Equal("New Name", Assert.Single(read.Paragraphs).Text);
            DocxContentControlInfo control = Assert.Single(read.ContentControls);
            Assert.Equal("client_name", control.Tag);
            Assert.Equal("Client Name", control.Alias);
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
    public static void OfficeAutomationHyperlinkRoundTripIsOptIn()
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
            CreateHyperlinkDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op set-hyperlink-target
                target M.L0001
                uri https://example.test/new
                tooltip Updated link
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
            DocxHyperlinkInfo hyperlink = Assert.Single(read.Hyperlinks);
            Assert.Equal("https://example.test/new", hyperlink.Uri);
            Assert.Equal("Updated link", hyperlink.Tooltip);
            Assert.False(hyperlink.IsBroken);
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

    private static void CreateStyledDocx(string path)
    {
        using FileStream file = File.Create(path);
        using var archive = new ZipArchive(file, ZipArchiveMode.Create);

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
                <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
              </w:body>
            </w:document>
            """);
        AddEntry(archive, "word/styles.xml", """
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
                <w:name w:val="Normal"/>
              </w:style>
              <w:style w:type="paragraph" w:styleId="Heading1">
                <w:name w:val="heading 1"/>
              </w:style>
            </w:styles>
            """);
    }

    private static void CreateResolvableCommentDocx(string path)
    {
        using FileStream file = File.Create(path);
        using var archive = new ZipArchive(file, ZipArchiveMode.Create);

        AddEntry(archive, "[Content_Types].xml", """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
              <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
              <Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>
            </Types>
            """);
        AddEntry(archive, "_rels/.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """);
        AddEntry(archive, "word/_rels/document.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
              <Relationship Id="rCommentsExtended" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>
            </Relationships>
            """);
        AddEntry(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:commentRangeStart w:id="3"/>
                  <w:r><w:t>Commented</w:t></w:r>
                  <w:commentRangeEnd w:id="3"/>
                  <w:r><w:commentReference w:id="3"/></w:r>
                </w:p>
              </w:body>
            </w:document>
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
              <w15:commentEx w15:paraId="00ABCDEF" w15:done="0"/>
            </w15:commentsEx>
            """);
    }

    private static void CreateFooterDocx(string path)
    {
        using FileStream file = File.Create(path);
        using var archive = new ZipArchive(file, ZipArchiveMode.Create);

        AddEntry(archive, "[Content_Types].xml", """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
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
              <Relationship Id="rFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
            </Relationships>
            """);
        AddEntry(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <w:body>
                <w:p><w:r><w:t>Body</w:t></w:r></w:p>
                <w:sectPr><w:footerReference w:type="default" r:id="rFooter"/></w:sectPr>
              </w:body>
            </w:document>
            """);
        AddEntry(archive, "word/footer1.xml", """
            <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:p><w:r><w:t>Footer old line</w:t></w:r></w:p>
            </w:ftr>
            """);
    }

    private static void CreateBookmarkDocx(string path)
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
                <w:p>
                  <w:r><w:t>Before </w:t></w:r>
                  <w:bookmarkStart w:id="4" w:name="ClientName"/>
                  <w:r><w:t>Old Client</w:t></w:r>
                  <w:bookmarkEnd w:id="4"/>
                  <w:r><w:t> After</w:t></w:r>
                </w:p>
              </w:body>
            </w:document>
            """);
    }

    private static void CreateControlDocx(string path)
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
                <w:p>
                  <w:sdt>
                    <w:sdtPr>
                      <w:id w:val="77"/>
                      <w:alias w:val="Client Name"/>
                      <w:tag w:val="client_name"/>
                      <w:text/>
                    </w:sdtPr>
                    <w:sdtContent><w:r><w:t>Acme</w:t></w:r></w:sdtContent>
                  </w:sdt>
                </w:p>
              </w:body>
            </w:document>
            """);
    }

    private static void CreateHyperlinkDocx(string path)
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
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/old" TargetMode="External"/>
            </Relationships>
            """);
        AddEntry(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <w:body>
                <w:p>
                  <w:hyperlink r:id="rLink" w:tooltip="Old link">
                    <w:r><w:t>Linked text</w:t></w:r>
                  </w:hyperlink>
                </w:p>
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
