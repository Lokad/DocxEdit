using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;

namespace Lokad.DocxEdit.OfficeTests;

public static class OfficeGateTests
{
    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationOpenSaveRoundTripIsOptIn()
    {

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
    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationIterativeTrackedRoundTripIsOptIn()
    {

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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationBreakStructureRoundTripIsOptIn()
    {

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

            string xml = ReadEntryText(outputPath, "word/document.xml");
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


    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationTrackedRowRoundTripIsOptIn()
    {

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
    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationTrackedPropertyRoundTripIsOptIn()
    {

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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationTrackedStyleRoundTripIsOptIn()
    {

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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationResolveCommentRoundTripIsOptIn()
    {

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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationCommentReplyRoundTripIsOptIn()
    {

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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationFooterEditRoundTripIsOptIn()
    {

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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationBookmarkTextRoundTripIsOptIn()
    {

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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationContentControlRoundTripIsOptIn()
    {

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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationHyperlinkRoundTripIsOptIn()
    {

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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationFieldResultRoundTripIsOptIn()
    {

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
            CreateFieldDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op set-field-result
                target M.F0001
                expect-result Old cached result
                text New cached result
                end
                """))
            using (FileStream output = File.Create(outputPath))
            {
                DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { MarkFieldsDirtyWhenEditing = false });
                Assert.True(result.Success, string.Join(Environment.NewLine, result.Diagnostics.Select(FormatDiagnostic)));
                Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error);
            }

            OpenSaveWithWord(wordApplicationType, outputPath);

            using FileStream saved = File.OpenRead(outputPath);
            DocxReadResult read = new DocxEditor().Read(saved);
            Assert.True(read.Success, string.Join(Environment.NewLine, read.Diagnostics.Select(FormatDiagnostic)));
            DocxFieldInfo field = Assert.Single(read.Fields);
            Assert.Equal("New cached result", field.CachedResultText);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationFormattedTrackedRoundTripIsOptIn()
    {

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
            CreateBoldDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op replace-text
                target M.P0001
                find Alpha
                with Omega
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
            Assert.Equal("Omega Beta", Assert.Single(read.Paragraphs).Text);

            string xml = ReadEntryText(outputPath, "word/document.xml");
            Assert.Contains("<w:b", xml, StringComparison.Ordinal);

            using FileStream savedChanges = File.OpenRead(outputPath);
            DocxChangesResult changes = new DocxEditor().Changes(savedChanges);
            Assert.True(changes.Success, string.Join(Environment.NewLine, changes.Diagnostics.Select(FormatDiagnostic)));
            Assert.Contains(changes.Summary, summary => summary.Type == "inserted-run" && summary.Count >= 1);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationMergeCellRoundTripIsOptIn()
    {

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
            CreateMergeDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op set-cell
                target M.T0001.MG0001
                expect-text Old merged
                text New merged
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
            DocxTableInfo table = Assert.Single(read.Tables);
            DocxTableCellInfo cell = Assert.Single(table.Cells);
            Assert.Equal("M.T0001.MG0001", cell.MergeGroupId?.ToWireValue());
            Assert.Equal(2, cell.VisualColumnEndIndex);
            Assert.Equal("New merged", cell.Text);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationTrackedRowDeleteRoundTripIsOptIn()
    {

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
            CreateTwoRowDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op delete-row
                target M.T0001.R02
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
            Assert.Equal(1, table.RowCount);
            Assert.Contains(table.Cells, cell => cell.Text == "North");

            using FileStream savedChanges = File.OpenRead(outputPath);
            DocxChangesResult changes = new DocxEditor().Changes(savedChanges);
            Assert.True(changes.Success, string.Join(Environment.NewLine, changes.Diagnostics.Select(FormatDiagnostic)));
            Assert.Contains(changes.Summary, summary => summary.Type == "row-deleted" && summary.Count == 1);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationAddBookmarkRoundTripIsOptIn()
    {

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
            CreateDocx(inputPath, "Office bookmark");

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op add-bookmark
                target M.P0001
                name OfficeMark
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
            DocxBookmarkInfo bookmark = Assert.Single(read.Bookmarks);
            Assert.Equal("OfficeMark", bookmark.Name);
            Assert.True(bookmark.IsComplete);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationHyperlinkInsertRoundTripIsOptIn()
    {

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
            CreateDocx(inputPath, "Anchor");

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op insert-hyperlink-after
                target M.P0001
                text Docs
                uri https://example.test/docs
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
            Assert.Equal("https://example.test/docs", hyperlink.Uri);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationDeleteCommentRoundTripIsOptIn()
    {

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

                op delete-comment
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

            using FileStream saved = File.OpenRead(outputPath);
            DocxReadResult read = new DocxEditor().Read(saved);
            Assert.True(read.Success, string.Join(Environment.NewLine, read.Diagnostics.Select(FormatDiagnostic)));
            Assert.Equal("Commented", Assert.Single(read.Paragraphs).Text);

            using FileStream savedChanges = File.OpenRead(outputPath);
            DocxChangesResult changes = new DocxEditor().Changes(savedChanges);
            Assert.True(changes.Success, string.Join(Environment.NewLine, changes.Diagnostics.Select(FormatDiagnostic)));
            Assert.Empty(changes.CommentSummary);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationTableStyleRoundTripIsOptIn()
    {

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
            CreateTableStyleDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op set-table-style
                target M.T0001
                style TableGrid
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
            Assert.Equal("TableGrid", Assert.Single(read.Tables).StyleId);

            using FileStream savedChanges = File.OpenRead(outputPath);
            DocxChangesResult changes = new DocxEditor().Changes(savedChanges);
            Assert.True(changes.Success, string.Join(Environment.NewLine, changes.Diagnostics.Select(FormatDiagnostic)));
            Assert.Contains(changes.Summary, summary => summary.Type == "table-properties-change" && summary.Count == 1);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationDeleteBookmarkRoundTripIsOptIn()
    {

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

                op delete-bookmark
                target M.B0001
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
            Assert.Equal("Before Old Client After", Assert.Single(read.Paragraphs).Text);
            Assert.Empty(read.Bookmarks);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationVerticalMergeRoundTripIsOptIn()
    {

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
            CreateVerticalMergeDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op set-cell
                target M.T0001.MG0001
                expect-text North
                text Merged
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
            DocxTableInfo table = Assert.Single(read.Tables);
            Assert.Contains(table.Cells, cell => cell.Id.ToWireValue() == "M.T0001.R01.C01" && cell.Text == "Merged");
            Assert.Contains(table.Cells, cell => cell.Id.ToWireValue() == "M.T0001.R02.C01" && cell.Text == "South");
            Assert.Contains(table.Cells, cell => cell.MergeGroupId != null && cell.MergeGroupId.Value.ToWireValue() == "M.T0001.MG0001");
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationAnchoredImageRoundTripIsOptIn()
    {

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
            CreateAnchoredImageDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op replace-text
                target M.P0001
                find Body
                with Edited body
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
            Assert.Contains(read.Paragraphs, paragraph => paragraph.Id.ToWireValue() == "M.P0001" && paragraph.Text == "Edited body");

            byte[] expected = MinimalPng();
            Assert.Equal(expected, ReadEntryBytes(outputPath, "word/media/image1.png"));

            Assert.Contains("wp:anchor", ReadEntryText(outputPath, "word/document.xml"), StringComparison.Ordinal);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationReopenCommentRoundTripIsOptIn()
    {

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
            CreateResolvableCommentDocx(inputPath, "1");

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op reopen-comment
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
            Assert.False(summary.Resolved);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationImageAltRoundTripIsOptIn()
    {

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
            CreateDocx(inputPath, "Anchor");

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op insert-image-after
                target M.P0001
                as newImage
                asset chart.png
                alt Office chart
                end

                op set-image-alt
                target @newImage
                expect-alt Office chart
                alt Updated chart
                end
                """))
            using (FileStream output = File.Create(outputPath))
            {
                DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
                {
                    AssetProvider = new MemoryAssetProvider("chart.png", MinimalPng(), "chart.png")
                });
                Assert.True(result.Success, string.Join(Environment.NewLine, result.Diagnostics.Select(FormatDiagnostic)));
                Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error);
            }

            OpenSaveWithWord(wordApplicationType, outputPath);

            using FileStream saved = File.OpenRead(outputPath);
            DocxReadResult read = new DocxEditor().Read(saved);
            Assert.True(read.Success, string.Join(Environment.NewLine, read.Diagnostics.Select(FormatDiagnostic)));
            Assert.Equal("Updated chart", Assert.Single(read.Images).Description);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationImageFormatSwapRoundTripIsOptIn()
    {
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
        string insertedPath = Path.Combine(directory, "inserted.docx");
        string outputPath = Path.Combine(directory, "output.docx");

        try
        {
            CreateDocx(inputPath, "Anchor");

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op insert-image-after
                target M.P0001
                asset chart.png
                end
                """))
            using (FileStream inserted = File.Create(insertedPath))
            {
                DocxApplyResult insert = new DocxEditor().Apply(input, patch, inserted, new DocxEditOptions
                {
                    AssetProvider = new MemoryAssetProvider("chart.png", MinimalPng(), "chart.png")
                });
                Assert.True(insert.Success, string.Join(Environment.NewLine, insert.Diagnostics.Select(FormatDiagnostic)));
            }

            using (FileStream middle = File.OpenRead(insertedPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op replace-image
                target M.I0001
                asset photo.jpeg
                end
                """))
            using (FileStream output = File.Create(outputPath))
            {
                DocxApplyResult replace = new DocxEditor().Apply(middle, patch, output, new DocxEditOptions
                {
                    AssetProvider = new MemoryAssetProvider("photo.jpeg", MinimalJpeg(), "photo.jpeg")
                });
                Assert.True(replace.Success, string.Join(Environment.NewLine, replace.Diagnostics.Select(FormatDiagnostic)));
                Assert.DoesNotContain(replace.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error);
            }

            OpenSaveWithWord(wordApplicationType, outputPath);

            using FileStream saved = File.OpenRead(outputPath);
            DocxReadResult read = new DocxEditor().Read(saved);
            Assert.True(read.Success, string.Join(Environment.NewLine, read.Diagnostics.Select(FormatDiagnostic)));
            DocxImageInfo image = Assert.Single(read.Images);
            Assert.Equal("image/jpeg", image.ContentType);
        }
        finally
        {
            try
            {
                Directory.Delete(directory, recursive: true);
            }
            catch
            {
            }
        }
    }

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationSharedHeaderMediaPreservationRoundTripIsOptIn()
    {
        // L01: the header drawing owns the same media part as the selected main
        // drawing through a different relationship. Replacing the main drawing
        // must clone fresh media for it while the header bytes stay untouched,
        // before and after Word opens and saves the result. Byte and reference
        // checks are the primary oracle; the Word pass only proves the shared
        // media fixture is accepted by Word. Package root thumbnail fixtures
        // stay byte level only: a valid JPEG root thumbnail round-trips through
        // Word while the PNG control is rejected, so PNG rejection proves no
        // universal limitation. The shared VML test below separately covers
        // preservation with default edit options on a field-free document.
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
            CreateSharedHeaderImageDocx(inputPath);
            byte[] sharedBefore = ReadEntryBytes(inputPath, "word/media/shared.png");

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op replace-image
                target M.I0001
                asset photo.jpeg
                end
                """))
            using (FileStream output = File.Create(outputPath))
            {
                DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
                {
                    AssetProvider = new MemoryAssetProvider("photo.jpeg", MinimalJpeg(), "photo.jpeg")
                });
                Assert.True(result.Success, string.Join(Environment.NewLine, result.Diagnostics.Select(FormatDiagnostic)));
                Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error);
            }

            Assert.Equal(sharedBefore, ReadEntryBytes(outputPath, "word/media/shared.png"));
            string headerRels = Encoding.UTF8.GetString(ReadEntryBytes(outputPath, "word/_rels/header1.xml.rels"));
            Assert.Contains("media/shared.png", headerRels, StringComparison.Ordinal);

            using (FileStream edited = File.OpenRead(outputPath))
            {
                DocxReadResult read = new DocxEditor().Read(edited, new DocxReadOptions { IncludeHeadersFooters = true });
                Assert.True(read.Success, string.Join(Environment.NewLine, read.Diagnostics.Select(FormatDiagnostic)));
                Assert.Equal(2, read.Images.Count);
                DocxImageInfo mainImage = read.Images.Single(static image => image.Id.ToWireValue() == "M.I0001");
                Assert.Equal("image/jpeg", mainImage.ContentType);
                Assert.True(mainImage.RelationshipId is not null, "Selected placement must resolve its relationship.");
                DocxImageInfo headerImage = read.Images.Single(static image => image.Id.ToWireValue() == "H001.I0001");
                Assert.Equal("image/png", headerImage.ContentType);
                Assert.True(headerImage.RelationshipId is not null, "Unselected placement must resolve its relationship.");
            }

            using (FileStream extracted = File.OpenRead(outputPath))
            {
                DocxMediaExtractResult media = new DocxEditor().ExtractMedia(extracted, new DocxMediaOptions { IncludeHeadersFooters = true });
                Assert.True(media.Success, string.Join(Environment.NewLine, media.Diagnostics.Select(FormatDiagnostic)));
                Assert.Equal(2, media.Files.Count);
                Assert.Equal(MinimalJpeg(), media.Files.Single(static file => file.ImageId.ToWireValue() == "M.I0001").Content);
                Assert.Equal(sharedBefore, media.Files.Single(static file => file.ImageId.ToWireValue() == "H001.I0001").Content);
            }

            OpenSaveWithWord(wordApplicationType, outputPath);

            using (FileStream savedExtract = File.OpenRead(outputPath))
            {
                DocxMediaExtractResult savedMedia = new DocxEditor().ExtractMedia(savedExtract, new DocxMediaOptions { IncludeHeadersFooters = true });
                Assert.True(savedMedia.Success, string.Join(Environment.NewLine, savedMedia.Diagnostics.Select(FormatDiagnostic)));
                Assert.Equal(2, savedMedia.Files.Count);
                Assert.Equal(MinimalJpeg(), savedMedia.Files.Single(static file => file.ImageId.ToWireValue() == "M.I0001").Content);
                Assert.Equal(sharedBefore, savedMedia.Files.Single(static file => file.ImageId.ToWireValue() == "H001.I0001").Content);
            }

            using FileStream saved = File.OpenRead(outputPath);
            DocxReadResult reread = new DocxEditor().Read(saved, new DocxReadOptions { IncludeHeadersFooters = true });
            Assert.True(reread.Success, string.Join(Environment.NewLine, reread.Diagnostics.Select(FormatDiagnostic)));
            Assert.Equal(2, reread.Images.Count);
            DocxImageInfo savedMain = reread.Images.Single(static image => image.Id.ToWireValue() == "M.I0001");
            Assert.Equal("image/jpeg", savedMain.ContentType);
            Assert.True(savedMain.RelationshipId is not null, "Selected placement must resolve its relationship after Word saves.");
            DocxImageInfo savedHeader = reread.Images.Single(static image => image.Id.ToWireValue() == "H001.I0001");
            Assert.Equal("image/png", savedHeader.ContentType);
            Assert.True(savedHeader.RelationshipId is not null, "Unselected placement must resolve its relationship after Word saves.");
        }
        finally
        {
            try
            {
                Directory.Delete(directory, recursive: true);
            }
            catch
            {
            }
        }
    }

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationSharedVmlDefaultOptionsRoundTripIsOptIn()
    {
        // VML preservation boundary: the DrawingML placement and the legacy VML
        // reference share one relationship, and replacing the drawing must leave
        // the VML bytes and relationship untouched. Default options must not
        // introduce field refresh into this field-free document: enabling it
        // previously stalled Word Open even though the package validated.
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
            CreateSharedVmlDocx(inputPath);
            byte[] sharedBefore = ReadEntryBytes(inputPath, "word/media/shared.png");

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("docxpatch 1\n\nop replace-image\ntarget M.I0001\nasset photo.jpeg\nend\n"))
            using (FileStream output = File.Create(outputPath))
            {
                DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
                {
                    AssetProvider = new MemoryAssetProvider("photo.jpeg", MinimalJpeg(), "photo.jpeg")
                });
                Assert.True(result.Success, string.Join(Environment.NewLine, result.Diagnostics.Select(FormatDiagnostic)));
                Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error);
            }

            using (var archive = ZipFile.OpenRead(outputPath))
            {
                Assert.Null(archive.GetEntry("word/settings.xml"));
            }
            Assert.Equal(sharedBefore, ReadEntryBytes(outputPath, "word/media/shared.png"));
            string documentRels = Encoding.UTF8.GetString(ReadEntryBytes(outputPath, "word/_rels/document.xml.rels"));
            Assert.Contains("rShared", documentRels, StringComparison.Ordinal);
            string document = Encoding.UTF8.GetString(ReadEntryBytes(outputPath, "word/document.xml"));
            Assert.Contains("v:imagedata", document, StringComparison.Ordinal);
            Assert.Contains("r:id=\"rShared\"", document, StringComparison.Ordinal);

            using (FileStream edited = File.OpenRead(outputPath))
            {
                DocxReadResult read = new DocxEditor().Read(edited);
                Assert.True(read.Success, string.Join(Environment.NewLine, read.Diagnostics.Select(FormatDiagnostic)));
                DocxImageInfo image = Assert.Single(read.Images);
                Assert.Equal("image/jpeg", image.ContentType);
            }

            OpenSaveWithWord(wordApplicationType, outputPath);

            Assert.Equal(sharedBefore, ReadEntryBytes(outputPath, "word/media/shared.png"));
            string savedDocument = Encoding.UTF8.GetString(ReadEntryBytes(outputPath, "word/document.xml"));
            Assert.Contains("v:imagedata", savedDocument, StringComparison.Ordinal);
            Assert.Contains("r:id=\"rShared\"", savedDocument, StringComparison.Ordinal);

            using FileStream saved = File.OpenRead(outputPath);
            DocxReadResult reread = new DocxEditor().Read(saved);
            Assert.True(reread.Success, string.Join(Environment.NewLine, reread.Diagnostics.Select(FormatDiagnostic)));
            DocxImageInfo savedImage = Assert.Single(reread.Images);
            Assert.Equal("image/jpeg", savedImage.ContentType);
        }
        finally
        {
            try
            {
                Directory.Delete(directory, recursive: true);
            }
            catch
            {
            }
        }
    }

    private static void CreateSharedVmlDocx(string path)
    {
        using FileStream file = File.Create(path);
        using var archive = new ZipArchive(file, ZipArchiveMode.Create);

        AddEntry(archive, "[Content_Types].xml", """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Default Extension="png" ContentType="image/png"/>
              <Default Extension="jpeg" ContentType="image/jpeg"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>
            """);
        AddPackageRelsEntry(archive);
        AddEntry(archive, "word/_rels/document.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/shared.png"/>
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
                <w:p><w:r><w:t>Shared VML body</w:t></w:r></w:p>
                <w:p>
                  <w:r>
                    <w:drawing>
                      <wp:inline>
                        <wp:extent cx="914400" cy="457200"/>
                        <wp:docPr id="1" name="Picture 1"/>
                        <a:graphic>
                          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                            <pic:pic>
                              <pic:nvPicPr><pic:cNvPr id="1" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr>
                              <pic:blipFill><a:blip r:embed="rShared"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                              <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
                            </pic:pic>
                          </a:graphicData>
                        </a:graphic>
                      </wp:inline>
                    </w:drawing>
                  </w:r>
                </w:p>
                <w:p><w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="Legacy" style="width:72pt;height:72pt"><v:imagedata r:id="rShared"/></v:shape></w:pict></w:r></w:p>
              </w:body>
            </w:document>
            """);
        AddBinaryEntry(archive, "word/media/shared.png", MinimalPng());
    }

    private static void CreateSharedHeaderImageDocx(string path)
    {
        using FileStream file = File.Create(path);
        using var archive = new ZipArchive(file, ZipArchiveMode.Create);

        AddEntry(archive, "[Content_Types].xml", """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Default Extension="png" ContentType="image/png"/>
              <Default Extension="jpeg" ContentType="image/jpeg"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
              <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
            </Types>
            """);
        AddPackageRelsEntry(archive);
        AddEntry(archive, "word/_rels/document.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/shared.png"/>
              <Relationship Id="rHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
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
                <w:p><w:r><w:t>Shared header body</w:t></w:r></w:p>
                <w:p>
                  <w:r>
                    <w:drawing>
                      <wp:inline>
                        <wp:extent cx="914400" cy="457200"/>
                        <wp:docPr id="1" name="Picture 1"/>
                        <a:graphic>
                          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                            <pic:pic>
                              <pic:nvPicPr><pic:cNvPr id="1" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr>
                              <pic:blipFill><a:blip r:embed="rImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                              <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
                            </pic:pic>
                          </a:graphicData>
                        </a:graphic>
                      </wp:inline>
                    </w:drawing>
                  </w:r>
                </w:p>
                <w:sectPr><w:headerReference w:type="default" r:id="rHeader"/></w:sectPr>
              </w:body>
            </w:document>
            """);
        AddEntry(archive, "word/_rels/header1.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/shared.png"/>
            </Relationships>
            """);
        AddEntry(archive, "word/header1.xml", """
            <w:hdr
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <w:p><w:r><w:t>Shared header line</w:t></w:r></w:p>
              <w:p>
                <w:r>
                  <w:drawing>
                    <wp:inline>
                      <wp:extent cx="914400" cy="457200"/>
                      <wp:docPr id="2" name="Picture 2"/>
                      <a:graphic>
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                          <pic:pic>
                            <pic:nvPicPr><pic:cNvPr id="2" name="Picture 2"/><pic:cNvPicPr/></pic:nvPicPr>
                            <pic:blipFill><a:blip r:embed="rImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                            <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
                          </pic:pic>
                        </a:graphicData>
                      </a:graphic>
                    </wp:inline>
                  </w:drawing>
                </w:r>
              </w:p>
            </w:hdr>
            """);
        AddBinaryEntry(archive, "word/media/shared.png", MinimalPng());
    }

    // Small genuinely decodable assets shared by the format-swap and ownership
    // tests, validated by ImageFixtureBytesDecodeIndependently below with an
    // independent structural check instead of a decoder package.
    [Fact]
    public static void ImageFixtureBytesDecodeIndependently()
    {
        // The shared Office image fixtures must be genuinely decodable image
        // content, not header or magic stubs. These structural checks verify
        // chunk integrity independently of the library under test and need no
        // decoder package: every PNG chunk passes its CRC and the scan data
        // inflates to the expected pixel rows, while the JPEG frame header,
        // scan data, and end marker are all present with expected dimensions.
        AssertPngDecodes(MinimalPng(), 4, 3);
        AssertJpegStructure(MinimalJpeg(), 4, 3);
    }

    private static void AssertPngDecodes(byte[] bytes, int expectedWidth, int expectedHeight)
    {
        byte[] signature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
        Assert.True(bytes.Length > signature.Length + 25, "PNG too short for header and footer chunks.");
        for (int i = 0; i < signature.Length; i++)
        {
            Assert.True(bytes[i] == signature[i], "PNG signature mismatch.");
        }

        uint[] crc = BuildCrc32Table();
        int offset = 8;
        bool seenHeader = false;
        var scan = new MemoryStream();
        while (true)
        {
            Assert.True(offset + 12 <= bytes.Length, "PNG ends inside a chunk header.");
            int length = (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];
            Assert.True(offset + 12 + length <= bytes.Length, "PNG chunk overruns the buffer.");
            string type = Encoding.ASCII.GetString(bytes, offset + 4, 4);
            uint actual = Crc32(crc, bytes, offset + 4, 4 + length);
            uint stored = ((uint)bytes[offset + 8 + length] << 24) | ((uint)bytes[offset + 9 + length] << 16) | ((uint)bytes[offset + 10 + length] << 8) | bytes[offset + 11 + length];
            Assert.True(actual == stored, "PNG chunk fails its integrity check: " + type);
            if (type == "IHDR")
            {
                seenHeader = true;
                int width = (bytes[offset + 8] << 24) | (bytes[offset + 9] << 16) | (bytes[offset + 10] << 8) | bytes[offset + 11];
                int height = (bytes[offset + 12] << 24) | (bytes[offset + 13] << 16) | (bytes[offset + 14] << 8) | bytes[offset + 15];
                Assert.Equal(expectedWidth, width);
                Assert.Equal(expectedHeight, height);
                Assert.True(bytes[offset + 16] == 8, "PNG fixture must use 8-bit samples.");
                Assert.True(bytes[offset + 17] == 2, "PNG fixture must use truecolor.");
            }
            if (type == "IDAT" && length > 0)
            {
                scan.Write(bytes, offset + 8, length);
            }
            offset += 12 + length;
            if (type == "IEND")
            {
                break;
            }
        }
        Assert.True(seenHeader, "PNG has no header chunk.");
        Assert.True(scan.Length > 0, "PNG has no scan data.");
        byte[] payload = scan.ToArray();
        Assert.True(payload.Length > 6, "PNG scan stream too short for framing bytes.");
        using var zlib = new MemoryStream(payload, 2, payload.Length - 6, writable: false);
        using var inflate = new System.IO.Compression.DeflateStream(zlib, System.IO.Compression.CompressionMode.Decompress);
        using var pixels = new MemoryStream();
        inflate.CopyTo(pixels);
        Assert.Equal((1 + 3 * expectedWidth) * expectedHeight, (int)pixels.Length);
    }

    private static void AssertJpegStructure(byte[] bytes, int expectedWidth, int expectedHeight)
    {
        Assert.True(bytes.Length > 4, "JPEG too short for markers.");
        Assert.True(bytes[0] == 0xFF && bytes[1] == 0xD8, "JPEG must start with its image marker.");
        Assert.True(bytes[bytes.Length - 2] == 0xFF && bytes[bytes.Length - 1] == 0xD9, "JPEG must end with its end marker.");
        int offset = 2;
        bool seenFrame = false;
        bool seenScan = false;
        while (offset + 4 <= bytes.Length)
        {
            Assert.True(bytes[offset] == 0xFF, "JPEG segment must start with a marker prefix.");
            byte marker = bytes[offset + 1];
            if (marker == 0xD9)
            {
                break;
            }
            if (marker == 0xDA)
            {
                seenScan = true;
                break;
            }
            if (marker == 0xD8 || (marker >= 0xD0 && marker <= 0xD7))
            {
                offset += 2;
                continue;
            }
            int length = (bytes[offset + 2] << 8) | bytes[offset + 3];
            Assert.True(length >= 2, "JPEG segment too short.");
            Assert.True(offset + 2 + length <= bytes.Length, "JPEG segment overruns the buffer.");
            if (marker == 0xC0 || marker == 0xC2)
            {
                seenFrame = true;
                int height = (bytes[offset + 5] << 8) | bytes[offset + 6];
                int width = (bytes[offset + 7] << 8) | bytes[offset + 8];
                Assert.Equal(expectedHeight, height);
                Assert.Equal(expectedWidth, width);
            }
            offset += 2 + length;
        }
        Assert.True(seenFrame, "JPEG has no frame header.");
        Assert.True(seenScan, "JPEG has no scan data.");
    }

    private static uint[] BuildCrc32Table()
    {
        var table = new uint[256];
        for (uint i = 0; i < 256; i++)
        {
            uint value = i;
            for (int bit = 0; bit < 8; bit++)
            {
                value = ((value & 1u) == 1u) ? (0xEDB88320u ^ (value >> 1)) : (value >> 1);
            }
            table[i] = value;
        }
        return table;
    }

    private static uint Crc32(uint[] table, byte[] bytes, int offset, int length)
    {
        uint value = 0xFFFFFFFFu;
        for (int i = 0; i < length; i++)
        {
            value = table[(int)((value ^ bytes[offset + i]) & 0xFF)] ^ (value >> 8);
        }
        return value ^ 0xFFFFFFFFu;
    }

    private static byte[] MinimalJpeg()
    {
        return Convert.FromBase64String("/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAADAAQDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDl6KKKZ+jn/9k=");
    }


    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationTableMetadataRoundTripIsOptIn()
    {

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
            CreateCaptionedTableDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op set-table-metadata
                target M.T0001
                expect-caption Old caption
                caption Revenue table
                description Quarterly figures
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
            DocxTableInfo table = Assert.Single(read.Tables);
            Assert.Equal("Revenue table", table.Caption);
            Assert.Equal("Quarterly figures", table.Description);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationHyperlinkTextRoundTripIsOptIn()
    {

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

                op set-hyperlink-text
                target M.L0001
                text New label
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
            Assert.Equal("New label", Assert.Single(read.Paragraphs).Text);
            DocxHyperlinkInfo hyperlink = Assert.Single(read.Hyperlinks);
            Assert.Equal("https://example.test/old", hyperlink.Uri);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationCheckboxRoundTripIsOptIn()
    {

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
            CreateCheckboxDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op set-content-control-checkbox
                target M.CC0001
                checked true
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
            DocxContentControlInfo control = Assert.Single(read.ContentControls);
            Assert.True(control.Checked);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationDeleteBlockRoundTripIsOptIn()
    {

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
            CreateDocx(inputPath, "Office delete");

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op delete-block
                target M.P0001
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
            Assert.Empty(read.Paragraphs);
            Assert.Single(read.Tables);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationRemoveHyperlinkRoundTripIsOptIn()
    {

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

                op remove-hyperlink
                target M.L0001
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
            Assert.Equal("Linked text", Assert.Single(read.Paragraphs).Text);
            Assert.Empty(read.Hyperlinks);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationImageSizeRoundTripIsOptIn()
    {

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
            CreateDocx(inputPath, "Anchor");

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op insert-image-after
                target M.P0001
                as newImage
                asset chart.png
                alt Office chart
                end

                op set-image-size
                target @newImage
                width 2in
                end
                """))
            using (FileStream output = File.Create(outputPath))
            {
                DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions
                {
                    AssetProvider = new MemoryAssetProvider("chart.png", MinimalPng(), "chart.png")
                });
                Assert.True(result.Success, string.Join(Environment.NewLine, result.Diagnostics.Select(FormatDiagnostic)));
                Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error);
            }

            long? libraryWidth;
            long? libraryHeight;
            using (FileStream preWord = File.OpenRead(outputPath))
            {
                DocxImageInfo preImage = Assert.Single(new DocxEditor().Read(preWord).Images);
                libraryWidth = preImage.WidthEmu;
                libraryHeight = preImage.HeightEmu;
            }

            OpenSaveWithWord(wordApplicationType, outputPath);

            using FileStream saved = File.OpenRead(outputPath);
            DocxReadResult read = new DocxEditor().Read(saved);
            Assert.True(read.Success, string.Join(Environment.NewLine, read.Diagnostics.Select(FormatDiagnostic)));
            DocxImageInfo image = Assert.Single(read.Images);
            Assert.Equal(libraryWidth, image.WidthEmu);
            Assert.Equal(libraryHeight, image.HeightEmu);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationDropdownRoundTripIsOptIn()
    {

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
            CreateDropdownDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op set-content-control-choice
                target M.CC0001
                value south
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
            DocxContentControlInfo control = Assert.Single(read.ContentControls);
            Assert.Equal("dropdown-list", control.Kind);
            Assert.Equal("South", Assert.Single(read.Paragraphs).Text);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationDateControlRoundTripIsOptIn()
    {

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
            CreateDateControlDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op set-content-control-date
                target M.CC0001
                value 2026-07-01T00:00:00Z
                display-text 2026-07-01
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
            DocxContentControlInfo control = Assert.Single(read.ContentControls);
            Assert.Equal("date", control.Kind);
            Assert.Equal("2026-07-01T00:00:00Z", control.DateValue);
            Assert.Equal("2026-07-01", Assert.Single(read.Paragraphs).Text);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationRenameBookmarkRoundTripIsOptIn()
    {

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

                op rename-bookmark
                target M.B0001
                name RenamedMark
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
            DocxBookmarkInfo bookmark = Assert.Single(read.Bookmarks);
            Assert.Equal("RenamedMark", bookmark.Name);
            Assert.True(bookmark.IsComplete);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationRowHeaderRoundTripIsOptIn()
    {

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
            CreateDocx(inputPath, "Office header");

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op set-row-header
                target M.T0001.R01
                expect-header false
                header true
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
            DocxTableInfo table = Assert.Single(read.Tables);
            Assert.True(Assert.Single(table.Rows).IsHeader);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationSectionOrientationRoundTripIsOptIn()
    {

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
            CreateSectionDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op set-section-orientation
                target M.S0001
                orientation landscape
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
            Assert.Equal(DocxOrientation.Landscape, Assert.Single(read.Sections).Orientation);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationFieldCodeRoundTripIsOptIn()
    {

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
            CreateFieldDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op set-field-code
                target M.F0001
                expect-code REF ClientName \h
                code REF ClientName
                end
                """))
            using (FileStream output = File.Create(outputPath))
            {
                DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { MarkFieldsDirtyWhenEditing = false });
                Assert.True(result.Success, string.Join(Environment.NewLine, result.Diagnostics.Select(FormatDiagnostic)));
                Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error);
            }

            OpenSaveWithWord(wordApplicationType, outputPath);

            using FileStream saved = File.OpenRead(outputPath);
            DocxReadResult read = new DocxEditor().Read(saved);
            Assert.True(read.Success, string.Join(Environment.NewLine, read.Diagnostics.Select(FormatDiagnostic)));
            Assert.Equal("REF ClientName", Assert.Single(read.Fields).Code);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationCommentTextRoundTripIsOptIn()
    {

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

                op set-comment-text
                target comment:3
                expect-text Comment body
                text Updated body
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
            DocxCommentThreadSummary summary = Assert.Single(changes.CommentSummary);
            Assert.Equal("3", summary.CommentId);
            Assert.Contains("Updated body", summary.TextSnippet, StringComparison.Ordinal);
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

    [OfficeFact]
    [Trait("Category", "RequiresWord")]
    public static void OfficeAutomationFieldFlagsRoundTripIsOptIn()
    {

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
            CreateFieldDocx(inputPath);

            using (FileStream input = File.OpenRead(inputPath))
            using (var patch = new StringReader("""
                docxpatch 1

                op set-field-dirty
                target M.F0001
                dirty true
                end

                op set-field-lock
                target M.F0001
                locked true
                end
                """))
            using (FileStream output = File.Create(outputPath))
            {
                DocxApplyResult result = new DocxEditor().Apply(input, patch, output, new DocxEditOptions { MarkFieldsDirtyWhenEditing = false });
                Assert.True(result.Success, string.Join(Environment.NewLine, result.Diagnostics.Select(FormatDiagnostic)));
                Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == DocxSeverity.Error);
            }

            OpenSaveWithWord(wordApplicationType, outputPath);

            using FileStream saved = File.OpenRead(outputPath);
            DocxReadResult read = new DocxEditor().Read(saved);
            Assert.True(read.Success, string.Join(Environment.NewLine, read.Diagnostics.Select(FormatDiagnostic)));
            DocxFieldInfo field = Assert.Single(read.Fields);
            Assert.True(field.IsDirty);
            Assert.True(field.IsLocked);
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
        AddPackageRelsEntry(archive);
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
        AddPackageRelsEntry(archive);
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
        AddPackageRelsEntry(archive);
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

    private static void CreateResolvableCommentDocx(string path, string done = "0")
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
        AddPackageRelsEntry(archive);
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
        AddEntry(archive, "word/commentsExtended.xml", $$"""
            <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
              <w15:commentEx w15:paraId="00ABCDEF" w15:done="{{done}}"/>
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
        AddPackageRelsEntry(archive);
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
        AddPackageRelsEntry(archive);
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
        AddPackageRelsEntry(archive);
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
        AddPackageRelsEntry(archive);
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

    private static void CreateFieldDocx(string path)
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
        AddPackageRelsEntry(archive);
        AddEntry(archive, "word/_rels/document.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
            """);
        AddEntry(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:bookmarkStart w:id="1" w:name="ClientName"/>
                  <w:r><w:t>Acme Corp</w:t></w:r>
                  <w:bookmarkEnd w:id="1"/>
                </w:p>
                <w:p>
                  <w:fldSimple w:instr=" REF ClientName \h " w:dirty="false" w:fldLock="0">
                    <w:r><w:t>Old cached result</w:t></w:r>
                  </w:fldSimple>
                </w:p>
              </w:body>
            </w:document>
            """);
    }

    private static void CreateBoldDocx(string path)
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
        AddPackageRelsEntry(archive);
        AddEntry(archive, "word/_rels/document.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
            """);
        AddEntry(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Alpha Beta</w:t></w:r></w:p>
              </w:body>
            </w:document>
            """);
    }

    private static void CreateMergeDocx(string path)
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
        AddPackageRelsEntry(archive);
        AddEntry(archive, "word/_rels/document.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
            """);
        AddEntry(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:tbl>
                  <w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>
                  <w:tr>
                    <w:tc>
                      <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
                      <w:p><w:r><w:t>Old merged</w:t></w:r></w:p>
                    </w:tc>
                  </w:tr>
                </w:tbl>
              </w:body>
            </w:document>
            """);
    }

    private static void CreateTwoRowDocx(string path)
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
        AddPackageRelsEntry(archive);
        AddEntry(archive, "word/_rels/document.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
            """);
        AddEntry(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
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
              </w:body>
            </w:document>
            """);
    }

    private static void CreateTableStyleDocx(string path)
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
        AddPackageRelsEntry(archive);
        AddEntry(archive, "word/_rels/document.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
            </Relationships>
            """);
        AddEntry(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p><w:r><w:t>Office style</w:t></w:r></w:p>
                <w:tbl>
                  <w:tr>
                    <w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>
                    <w:tc><w:p><w:r><w:t>South</w:t></w:r></w:p></w:tc>
                  </w:tr>
                </w:tbl>
              </w:body>
            </w:document>
            """);
        AddEntry(archive, "word/styles.xml", """
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
                <w:name w:val="Normal"/>
              </w:style>
              <w:style w:type="table" w:styleId="TableGrid">
                <w:name w:val="Table Grid"/>
              </w:style>
            </w:styles>
            """);
    }

    private static void CreateVerticalMergeDocx(string path)
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
        AddPackageRelsEntry(archive);
        AddEntry(archive, "word/_rels/document.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
            """);
        AddEntry(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
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
              </w:body>
            </w:document>
            """);
    }

    private static void CreateAnchoredImageDocx(string path)
    {
        using FileStream file = File.Create(path);
        using var archive = new ZipArchive(file, ZipArchiveMode.Create);

        AddEntry(archive, "[Content_Types].xml", """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Default Extension="png" ContentType="image/png"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>
            """);
        AddPackageRelsEntry(archive);
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
                <w:p><w:r><w:t>Body</w:t></w:r></w:p>
                <w:p>
                  <w:r>
                    <w:drawing>
                      <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
                        <wp:simplePos x="0" y="0"/>
                        <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
                        <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
                        <wp:extent cx="914400" cy="457200"/>
                        <wp:effectExtent l="0" t="0" r="0" b="0"/>
                        <wp:wrapNone/>
                        <wp:docPr id="1" name="Picture 1"/>
                        <wp:cNvGraphicFramePr/>
                        <a:graphic>
                          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                            <pic:pic>
                              <pic:nvPicPr><pic:cNvPr id="1" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr>
                              <pic:blipFill><a:blip r:embed="rImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                              <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
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
        AddBinaryEntry(archive, "word/media/image1.png", MinimalPng());
    }

    private static void AddBinaryEntry(ZipArchive archive, string name, byte[] bytes)
    {
        ZipArchiveEntry entry = archive.CreateEntry(name);
        using Stream stream = entry.Open();
        stream.Write(bytes, 0, bytes.Length);
    }

    private static byte[] ReadEntryBytes(string docxPath, string entryName)
    {
        using FileStream file = File.OpenRead(docxPath);
        using var archive = new ZipArchive(file, ZipArchiveMode.Read);
        ZipArchiveEntry entry = archive.GetEntry(entryName) ?? throw new InvalidOperationException("Missing " + entryName + ".");
        using Stream stream = entry.Open();
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }

    private static string ReadEntryText(string docxPath, string entryName)
    {
        using FileStream file = File.OpenRead(docxPath);
        using var archive = new ZipArchive(file, ZipArchiveMode.Read);
        ZipArchiveEntry entry = archive.GetEntry(entryName) ?? throw new InvalidOperationException("Missing " + entryName + ".");
        using Stream stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    }

    private static void CreateCaptionedTableDocx(string path)
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
        AddPackageRelsEntry(archive);
        AddEntry(archive, "word/_rels/document.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
            """);
        AddEntry(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:tbl>
                  <w:tblPr>
                    <w:tblCaption w:val="Old caption"/>
                    <w:tblDescription w:val="Old description"/>
                  </w:tblPr>
                  <w:tr>
                    <w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc>
                  </w:tr>
                </w:tbl>
              </w:body>
            </w:document>
            """);
    }

    private static void CreateCheckboxDocx(string path)
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
        AddPackageRelsEntry(archive);
        AddEntry(archive, "word/_rels/document.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
            """);
        AddEntry(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
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
              </w:body>
            </w:document>
            """);
    }

    private static void CreateDropdownDocx(string path)
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
        AddPackageRelsEntry(archive);
        AddEntry(archive, "word/_rels/document.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
            """);
        AddEntry(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
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
              </w:body>
            </w:document>
            """);
    }

    private static void CreateDateControlDocx(string path)
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
        AddPackageRelsEntry(archive);
        AddEntry(archive, "word/_rels/document.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
            """);
        AddEntry(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
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
              </w:body>
            </w:document>
            """);
    }

    private static void CreateSectionDocx(string path)
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
        AddPackageRelsEntry(archive);
        AddEntry(archive, "word/_rels/document.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />
            """);
        AddEntry(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p><w:r><w:t>Main text</w:t></w:r></w:p>
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
        AddPackageRelsEntry(archive);
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

    private static void AddPackageRelsEntry(ZipArchive archive)
    {
        AddEntry(archive, "_rels/.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """);
    }

    private static byte[] MinimalPng()
    {
        return Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAQAAAADCAIAAAA7ljmRAAAAGklEQVR4nGNkYGCwZdCHIBaGYH0GBihC4QAARUkDbIjjm+8AAAAASUVORK5CYII=");
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
