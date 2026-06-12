namespace DocxEdit.Tests;

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
        Assert.Equal("tracked-simple", replaceText.TrackChangesSupport);
        Assert.Contains("E6002", replaceText.TrackChangesNote, StringComparison.Ordinal);
        Assert.True(DocxHelp.TryGetPatchOperation("set-cell", out DocxPatchOperationInfo setCell));
        Assert.Equal("tracked-cell-simple", setCell.TrackChangesSupport);
        Assert.Contains("E6002", setCell.TrackChangesNote, StringComparison.Ordinal);
        Assert.True(DocxHelp.TryGetPatchOperation("remove-hyperlink", out DocxPatchOperationInfo removeHyperlink));
        Assert.Equal("preserve-only", removeHyperlink.TrackChangesSupport);
        Assert.Contains("does not create new revision markup", removeHyperlink.TrackChangesNote, StringComparison.Ordinal);
        Assert.Contains("E6001", removeHyperlink.TrackChangesNote, StringComparison.Ordinal);
        Assert.True(DocxHelp.TryRenderTopic("patch", out string patchHelp));
        Assert.Contains("preserve-only", patchHelp, StringComparison.Ordinal);
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
}
