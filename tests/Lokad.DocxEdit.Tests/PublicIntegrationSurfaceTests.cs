using System.IO;

namespace Lokad.DocxEdit.Tests;

public static class PublicIntegrationSurfaceTests
{
    [Fact]
    public static void CommandCatalogExposesReusableAgentGuidance()
    {
        DocxCommandCatalog catalog = DocxHelp.Catalog;

        Assert.Equal("docxedit", catalog.ToolName);
        DocxCommandInfo changes = catalog.Commands.Single(command => command.Name == "changes");
        Assert.Contains(changes.OutputFields, field => field.Name == "TargetSummary");
        Assert.Contains(changes.OutputFields, field => field.Name == "CommentSummary");
        Assert.Contains(changes.PrivacyNotes, note => note.Contains("private-text-free", StringComparison.Ordinal));

        Assert.True(DocxHelp.TryRenderTopic("changes", out string help));
        Assert.Contains("GroupSummary", help, StringComparison.Ordinal);
        Assert.Contains("paired-change-id", help, StringComparison.Ordinal);
        Assert.DoesNotContain("PowerShell", help, StringComparison.Ordinal);
        Assert.DoesNotContain("ConvertFrom-Json", help, StringComparison.Ordinal);

        Assert.True(DocxHelp.TryGetPatchOperation("replace-text", out DocxPatchOperationInfo replaceText));
        Assert.Contains("target", replaceText.RequiredFields);
        Assert.Contains("find", replaceText.RequiredFields);
        Assert.Contains("with", replaceText.RequiredFields);
        Assert.Contains("preserve-runs", replaceText.OptionalFields);
        Assert.Equal("text-run", replaceText.TrackChangesSupportClass);
        Assert.Equal("tracked-simple", replaceText.TrackChangesSupport);
        Assert.True(replaceText.GeneratesTrackedChanges);
        Assert.Contains("E6002", replaceText.TrackChangesNote, StringComparison.Ordinal);
        Assert.True(DocxHelp.TryGetPatchOperation("set-content-control-text", out DocxPatchOperationInfo setContentControlText));
        Assert.Contains("expect-text", setContentControlText.OptionalFields);
        Assert.Equal("text-run", setContentControlText.TrackChangesSupportClass);
        Assert.Equal("tracked-content-control-text", setContentControlText.TrackChangesSupport);
        Assert.True(setContentControlText.GeneratesTrackedChanges);
        Assert.True(DocxHelp.TryGetPatchOperation("set-field-result", out DocxPatchOperationInfo setFieldResult));
        Assert.Contains("expect-result", setFieldResult.OptionalFields);
        Assert.Equal("text-run", setFieldResult.TrackChangesSupportClass);
        Assert.Equal("tracked-field-result", setFieldResult.TrackChangesSupport);
        Assert.True(setFieldResult.GeneratesTrackedChanges);
        Assert.True(DocxHelp.TryGetPatchOperation("refresh-field-result", out DocxPatchOperationInfo refreshFieldResult));
        Assert.Contains("expect-code", refreshFieldResult.OptionalFields);
        Assert.Contains("expect-result", refreshFieldResult.OptionalFields);
        Assert.Equal("preserve-only", refreshFieldResult.TrackChangesSupportClass);
        Assert.Equal("preserve-only", refreshFieldResult.TrackChangesSupport);
        Assert.True(DocxHelp.TryGetPatchOperation("set-cell", out DocxPatchOperationInfo setCell));
        Assert.Equal("text-run", setCell.TrackChangesSupportClass);
        Assert.Equal("tracked-cell-simple", setCell.TrackChangesSupport);
        Assert.True(setCell.GeneratesTrackedChanges);
        Assert.Contains("E6002", setCell.TrackChangesNote, StringComparison.Ordinal);
        Assert.True(DocxHelp.TryGetPatchOperation("set-table-style", out DocxPatchOperationInfo setTableStyle));
        Assert.Contains("expect-style", setTableStyle.OptionalFields);
        Assert.Equal("table-property", setTableStyle.TrackChangesSupportClass);
        Assert.Equal("tracked-table-style", setTableStyle.TrackChangesSupport);
        Assert.True(setTableStyle.GeneratesTrackedChanges);
        Assert.True(DocxHelp.TryGetPatchOperation("set-table-metadata", out DocxPatchOperationInfo setTableMetadata));
        Assert.Contains("expect-caption", setTableMetadata.OptionalFields);
        Assert.Contains("expect-description", setTableMetadata.OptionalFields);
        Assert.Equal("preserve-only", setTableMetadata.TrackChangesSupportClass);
        Assert.Equal("preserve-only", setTableMetadata.TrackChangesSupport);
        Assert.True(DocxHelp.TryGetPatchOperation("set-row-header", out DocxPatchOperationInfo setRowHeader));
        Assert.Contains("expect-header", setRowHeader.OptionalFields);
        Assert.Equal("row-property", setRowHeader.TrackChangesSupportClass);
        Assert.Equal("tracked-row-header", setRowHeader.TrackChangesSupport);
        Assert.True(setRowHeader.GeneratesTrackedChanges);
        Assert.True(DocxHelp.TryGetPatchOperation("remove-hyperlink", out DocxPatchOperationInfo removeHyperlink));
        Assert.Equal("preserve-only", removeHyperlink.TrackChangesSupportClass);
        Assert.Equal("preserve-only", removeHyperlink.TrackChangesSupport);
        Assert.Contains("does not create new revision markup", removeHyperlink.TrackChangesNote, StringComparison.Ordinal);
        Assert.Contains("E6001", removeHyperlink.TrackChangesNote, StringComparison.Ordinal);
        Assert.True(DocxHelp.TryGetPatchOperation("insert-hyperlink-after", out DocxPatchOperationInfo insertHyperlinkAfter));
        Assert.Equal("text-run", insertHyperlinkAfter.TrackChangesSupportClass);
        Assert.Equal("tracked-hyperlink-insert", insertHyperlinkAfter.TrackChangesSupport);
        Assert.True(insertHyperlinkAfter.GeneratesTrackedChanges);
        Assert.True(DocxHelp.TryGetPatchOperation("add-comment-reply", out DocxPatchOperationInfo addCommentReply));
        Assert.Equal("unsupported", addCommentReply.TrackChangesSupportClass);
        Assert.Equal("unsupported", addCommentReply.TrackChangesSupport);
        Assert.False(addCommentReply.GeneratesTrackedChanges);
        Assert.Contains("E4314", addCommentReply.TrackChangesNote, StringComparison.Ordinal);
        Assert.True(DocxHelp.TryGetPatchOperation("add-repeating-section-item", out DocxPatchOperationInfo addRepeatingSectionItem));
        Assert.Equal("unsupported", addRepeatingSectionItem.TrackChangesSupportClass);
        Assert.Equal("unsupported", addRepeatingSectionItem.TrackChangesSupport);
        Assert.Contains("E4315", addRepeatingSectionItem.TrackChangesNote, StringComparison.Ordinal);
        Assert.True(DocxHelp.TryGetPatchOperation("insert-column-before", out DocxPatchOperationInfo insertColumnBefore));
        Assert.Equal("unsupported", insertColumnBefore.TrackChangesSupportClass);
        Assert.Equal("unsupported", insertColumnBefore.TrackChangesSupport);
        Assert.Contains(insertColumnBefore.RequiredFields, field => field.Contains("column", StringComparison.Ordinal));
        Assert.Contains("E4316", insertColumnBefore.TrackChangesNote, StringComparison.Ordinal);
        Assert.True(DocxHelp.TryGetPatchOperation("add-bookmark", out DocxPatchOperationInfo addBookmark));
        Assert.Contains("expect-text", addBookmark.OptionalFields);
        Assert.Equal("preserve-only", addBookmark.TrackChangesSupportClass);
        Assert.Equal("preserve-only", addBookmark.TrackChangesSupport);
        Assert.True(DocxHelp.TryGetPatchOperation("replace-bookmark-text", out DocxPatchOperationInfo replaceBookmarkText));
        Assert.Equal("text-run", replaceBookmarkText.TrackChangesSupportClass);
        Assert.Equal("tracked-bookmark-text", replaceBookmarkText.TrackChangesSupport);
        Assert.True(replaceBookmarkText.GeneratesTrackedChanges);
        Assert.True(DocxHelp.TryGetPatchOperation("set-comment-text", out DocxPatchOperationInfo setCommentText));
        Assert.Equal("text-run", setCommentText.TrackChangesSupportClass);
        Assert.Equal("tracked-comment-text", setCommentText.TrackChangesSupport);
        Assert.True(setCommentText.GeneratesTrackedChanges);
        Assert.True(DocxHelp.TryRenderTopic("patch", out string patchHelp));
        Assert.Contains("replace-text | text-run | tracked-simple", patchHelp, StringComparison.Ordinal);
        Assert.Contains("preserve-only", patchHelp, StringComparison.Ordinal);
    }

    [Fact]
    public static void PatchTrackChangeSupportTableStaysInSyncWithHelpAndDocs()
    {
        string table = DocxHelp.RenderPatchTrackChangesSupportTable();

        Assert.StartsWith("operation | support class | support value | behavior", table, StringComparison.Ordinal);
        Assert.Contains("replace-text | text-run | tracked-simple", table, StringComparison.Ordinal);
        Assert.Contains("replace-bookmark-text | text-run | tracked-bookmark-text", table, StringComparison.Ordinal);
        Assert.Contains("set-comment-text | text-run | tracked-comment-text", table, StringComparison.Ordinal);
        Assert.Contains("insert-hyperlink-after | text-run | tracked-hyperlink-insert", table, StringComparison.Ordinal);
        Assert.Contains("set-cell | text-run | tracked-cell-simple", table, StringComparison.Ordinal);
        Assert.Contains("set-table-style | table-property | tracked-table-style", table, StringComparison.Ordinal);
        Assert.Contains("set-row-header | row-property | tracked-row-header", table, StringComparison.Ordinal);
        Assert.Contains(NormalizeLineEndings(table), NormalizeLineEndings(DocxHelp.RenderTopic("patch")), StringComparison.Ordinal);

        string repoRoot = FindRepoRoot();
        string patchFormat = File.ReadAllText(Path.Combine(repoRoot, "docs", "patch-format.md"));
        Assert.Contains(NormalizeLineEndings(table).TrimEnd(), NormalizeLineEndings(patchFormat), StringComparison.Ordinal);

        string readme = File.ReadAllText(Path.Combine(repoRoot, "README.md"));
        Assert.Contains("DocxHelp.RenderPatchTrackChangesSupportTable()", readme, StringComparison.Ordinal);
        Assert.Contains("docs/patch-format.md", readme, StringComparison.Ordinal);
    }

    [Fact]
    public static void PatchCatalogUsesExplicitTrackChangeSupportClasses()
    {
        string[] validClasses =
        [
            "text-run",
            "paragraph-block",
            "paragraph-property",
            "table-property",
            "row-property",
            "cell-property",
            "section-property",
            "relationship-metadata",
            "preserve-only",
            "unsupported"
        ];

        foreach (DocxPatchOperationInfo operation in DocxHelp.Catalog.PatchOperations)
        {
            Assert.Contains(operation.TrackChangesSupportClass, validClasses);
            Assert.Equal(
                operation.TrackChangesSupportClass is not "preserve-only" and not "unsupported",
                operation.GeneratesTrackedChanges);
            if (operation.TrackChangesSupportClass == "preserve-only")
            {
                Assert.False(
                    operation.TrackChangesNote.StartsWith("Existing tracked-change markup is preserved", StringComparison.Ordinal),
                    operation.Name);
                Assert.Contains("Existing tracked-change markup is preserved", operation.TrackChangesNote, StringComparison.Ordinal);
                Assert.Contains("Require fails with E6001", operation.TrackChangesNote, StringComparison.Ordinal);
            }
        }
    }

    [Fact]
    public static void PublicTextRendererProvidesPrivacySafeSummaries()
    {
        var result = new DocxReadResult
        {
            Success = true,
            Diagnostics = [],
            PartNames = ["/word/document.xml"],
            MainDocumentPartName = "/word/document.xml",
            Paragraphs =
            [
                new DocxParagraphInfo("M.P0001", "main", "Sensitive customer text", null, null, [])
            ],
            Tables = [],
            Images = [],
            Sections = []
        };

        string summary = DocxTextRenderer.RenderReadSummary(result);

        Assert.Contains("parts count=1", summary, StringComparison.Ordinal);
        Assert.Contains("paragraphs count=1", summary, StringComparison.Ordinal);
        Assert.Contains("story=\"main\" paragraphs=1", summary, StringComparison.Ordinal);
        Assert.DoesNotContain("Sensitive customer text", summary, StringComparison.Ordinal);
        Assert.Equal(summary, DocxPrivacyPresets.RenderReadSummary(result));
    }

    [Fact]
    public static void PublicTextRendererProvidesChangeAndPatchSummaries()
    {
        var changes = new DocxChangesResult
        {
            Success = true,
            Diagnostics = [],
            Summary = [new DocxChangeSummary("inserted-run", 1)],
            GroupSummary = [new DocxChangeGroupSummary("target", "M.P0001", "inserted-run", 1)],
            TargetSummary =
            [
                new DocxChangeTargetSummary
                {
                    TargetId = "M.P0001",
                    Count = 1,
                    Summary = [new DocxChangeSummary("inserted-run", 1)]
                }
            ],
            Changes =
            [
                new DocxChangeInfo
                {
                    Id = "M.CH0001",
                    Type = "inserted-run",
                    Story = "main",
                    PartName = "/word/document.xml",
                    TargetId = "M.P0001",
                    TargetStatus = "targeted",
                    TargetSource = "ancestor",
                    TextLength = 14,
                    ChildElementCount = 1,
                    Author = "Reviewer"
                }
            ]
        };
        var operations = new[]
        {
            new DocxPatchOperationReport(1, "replace-text", "M.P0001", true, [])
        };

        string changeText = DocxTextRenderer.RenderChanges(changes);
        string operationText = DocxTextRenderer.RenderOperationSummary(operations);
        string validateText = DocxTextRenderer.RenderValidate(new DocxValidateResult
        {
            Success = false,
            Diagnostics =
            [
                new DocxDiagnostic(
                    DocxSeverity.Error,
                    "E9105",
                    "Drawing references missing relationship 'rMissing'.",
                    PartName: "/word/document.xml")
            ],
            PartNames = ["/word/document.xml"],
            MainDocumentPartName = "/word/document.xml"
        });

        Assert.Contains("inserted-run count=1", changeText, StringComparison.Ordinal);
        Assert.Contains("target-summary target=M.P0001 count=1", changeText, StringComparison.Ordinal);
        Assert.Contains("target-source=ancestor", changeText, StringComparison.Ordinal);
        Assert.Contains("text-length=14", changeText, StringComparison.Ordinal);
        Assert.Contains("operation index=1 name=replace-text target=M.P0001 success=True", operationText, StringComparison.Ordinal);
        Assert.Contains("docxedit validate: FAILED", validateText, StringComparison.Ordinal);
        Assert.Contains("profile=structural", validateText, StringComparison.Ordinal);
        Assert.Contains("Error E9105 part=/word/document.xml", validateText, StringComparison.Ordinal);
    }

    [Fact]
    public static void PrivacyPresetsExposeMetadataOnlyDefaults()
    {
        Assert.Equal(0, DocxPrivacyPresets.ReadSummary.MaxText);
        Assert.Equal(0, DocxPrivacyPresets.ContextMetadataOnly.MaxText);
        Assert.Equal(1, DocxPrivacyPresets.ContextMetadataOnly.Radius);
        Assert.True(DocxPrivacyPresets.ChangesMarkupOnly.LeaveInputOpen);
    }

    private static string FindRepoRoot()
    {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Lokad.DocxEdit.slnx")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate repository root.");
    }

    private static string NormalizeLineEndings(string value)
    {
        return value.Replace("\r\n", "\n", StringComparison.Ordinal);
    }
}
