using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.Json;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// Serializes with CliTests: both redirect the process-wide Console.
[Collection("ConsoleCli")]
public static class EditCaseTests
{
    public static IEnumerable<object[]> EditCases()
    {
        string root = FindRepoRoot();
        foreach (string path in Directory.EnumerateFiles(Path.Combine(root, "edit-cases", "cases"), "*.json"))
        {
            yield return [Path.GetFileNameWithoutExtension(path)];
        }
    }

    public static IEnumerable<object[]> EditMatrixCases()
    {
        string root = FindRepoRoot();

        using JsonDocument family = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "edit-cases", "families", "tracked-change-matrix.json")));
        foreach (JsonElement mode in family.RootElement.GetProperty("trackChangesModes").EnumerateArray())
        {
            foreach (JsonElement caseId in family.RootElement.GetProperty("cases").EnumerateArray())
            {
                string modeName = mode.GetString() ?? string.Empty;
                string id = caseId.GetString() ?? string.Empty;
                yield return [id, modeName, modeName is "off" or "preserve"];
            }
        }
    }


    private static string GetManifestText(JsonElement value)
    {
        if (value.ValueKind == JsonValueKind.Array)
        {
            return string.Join(Environment.NewLine, value.EnumerateArray().Select(item => item.GetString() ?? string.Empty));
        }

        return value.ValueKind == JsonValueKind.String ? value.GetString() ?? string.Empty : string.Empty;
    }

    private static string ReadMediaProperty(JsonElement input, string propertyName)
    {
        if (input.ValueKind == JsonValueKind.Object
            && input.TryGetProperty("media", out JsonElement mediaValue)
            && mediaValue.ValueKind == JsonValueKind.Object
            && mediaValue.TryGetProperty(propertyName, out JsonElement propertyValue)
            && propertyValue.ValueKind == JsonValueKind.String)
        {
            return propertyValue.GetString() ?? string.Empty;
        }

        return string.Empty;
    }

    private static bool IsFilenameSafeId(string value)
    {
        if (value.Length == 0 || !IsAsciiLetterOrDigit(value[0]))
        {
            return false;
        }

        return value.All(character => IsAsciiLetterOrDigit(character) || character == '.' || character == '_' || character == '-');
    }

    private static bool IsAsciiLetterOrDigit(char value)
    {
        return value is >= 'a' and <= 'z' or >= 'A' and <= 'Z' or >= '0' and <= '9';
    }
    private static void BuildCaseDocx(string path, string bodyXml, string headerXml, string footerXml, string stylesXml, string commentsXml, string mediaFileName, string mediaContentType, string mediaText, string hyperlinkTarget, string commentsExtendedXml, string commentsIdsXml)
    {
        bool header = !string.IsNullOrWhiteSpace(headerXml);
        bool footer = !string.IsNullOrWhiteSpace(footerXml);
        bool styles = !string.IsNullOrWhiteSpace(stylesXml);
        bool comments = !string.IsNullOrWhiteSpace(commentsXml);
        bool commentsExtended = !string.IsNullOrWhiteSpace(commentsExtendedXml);
        bool commentsIds = !string.IsNullOrWhiteSpace(commentsIdsXml);
        bool media = !string.IsNullOrWhiteSpace(mediaFileName);
        bool hyperlink = !string.IsNullOrWhiteSpace(hyperlinkTarget);
        string headerOverride = header ? """<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>""" : string.Empty;
        string footerOverride = footer ? """<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>""" : string.Empty;
        string stylesOverride = styles ? """<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>""" : string.Empty;
        string commentsOverride = comments ? """<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>""" : string.Empty;
        string headerRelationship = header ? """<Relationship Id="rHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>""" : string.Empty;
        string footerRelationship = footer ? """<Relationship Id="rFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>""" : string.Empty;
        string stylesRelationship = styles ? """<Relationship Id="rStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>""" : string.Empty;
        string commentsRelationship = comments ? """<Relationship Id="rComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>""" : string.Empty;
        string commentsExtendedOverride = commentsExtended ? """<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.ms-word.commentsExtended+xml"/>""" : string.Empty;
        string commentsIdsOverride = commentsIds ? """<Override PartName="/word/commentsIds.xml" ContentType="application/vnd.ms-word.commentsIds+xml"/>""" : string.Empty;
        string commentsExtendedRelationship = commentsExtended ? """<Relationship Id="rCommentsExtended" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>""" : string.Empty;
        string commentsIdsRelationship = commentsIds ? """<Relationship Id="rCommentsIds" Type="http://schemas.microsoft.com/office/2016/09/relationships/commentsIds" Target="commentsIds.xml"/>""" : string.Empty;
        string mediaExtension = media ? mediaFileName.Substring(mediaFileName.LastIndexOf(".") + 1) : string.Empty;
        string mediaDefault = media ? "<Default Extension=\"" + mediaExtension + "\" ContentType=\"" + mediaContentType + "\"/>" : string.Empty;
        string mediaRelationship = media ? "<Relationship Id=\"rImage\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/" + mediaFileName + "\"/>" : string.Empty;
        string hyperlinkRelationship = hyperlink ? "<Relationship Id=\"rLink\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"" + hyperlinkTarget + "\" TargetMode=\"External\"/>" : string.Empty;
        using var archive = new ZipArchive(File.Create(path), ZipArchiveMode.Create);
        AddEntry(archive, "[Content_Types].xml", """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            """ + headerOverride + footerOverride + stylesOverride + commentsOverride + commentsExtendedOverride + commentsIdsOverride + mediaDefault + """
            </Types>
            """);
        AddEntry(archive, "_rels/.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """);
        AddEntry(archive, "word/_rels/document.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            """ + headerRelationship + footerRelationship + stylesRelationship + commentsRelationship + commentsExtendedRelationship + commentsIdsRelationship + mediaRelationship + hyperlinkRelationship + """
            </Relationships>
            """);
        AddEntry(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
            """ + bodyXml + """
              </w:body>
            </w:document>
            """);
        if (header)
        {
            AddEntry(archive, "word/header1.xml", """
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                """ + headerXml + """
                </w:hdr>
                """);
        }

        if (footer)
        {
            AddEntry(archive, "word/footer1.xml", """
                <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                """ + footerXml + """
                </w:ftr>
                """);
        }

        if (styles)
        {
            AddEntry(archive, "word/styles.xml", """
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                """ + stylesXml + """
                </w:styles>
                """);
        }

        if (comments)
        {
            AddEntry(archive, "word/comments.xml", commentsXml);
        }

        if (commentsExtended)
        {
            AddEntry(archive, "word/commentsExtended.xml", commentsExtendedXml);
        }

        if (commentsIds)
        {
            AddEntry(archive, "word/commentsIds.xml", commentsIdsXml);
        }

        if (media)
        {
            AddEntry(archive, "word/media/" + mediaFileName, mediaText);
        }
    }

    [Theory]
    [MemberData(nameof(EditCases))]
    public static void EditCaseApplies(string caseId)
    {
        RunEditCase(caseId, null, false);
    }

    [Theory]
    [MemberData(nameof(EditMatrixCases))]
    public static void EditCaseMatrixApplies(string caseId, string trackChangesOverride, bool expectNoChangeSummary)
    {
        RunEditCase(caseId, trackChangesOverride, expectNoChangeSummary);
    }

    private static void RunEditCase(string caseId, string? trackChangesOverride, bool expectNoChangeSummary)
    {
        string repoRoot = FindRepoRoot();
        using JsonDocument manifestDocument = JsonDocument.Parse(File.ReadAllText(Path.Combine(repoRoot, "edit-cases", "cases", caseId + ".json")));
        JsonElement manifest = manifestDocument.RootElement;
        string manifestId = manifest.GetProperty("id").GetString() ?? string.Empty;
        Assert.True(IsFilenameSafeId(manifestId), "Case manifest must define a filename-safe id.");

        Assert.True(manifest.TryGetProperty("patch", out JsonElement patchValue), "Case " + manifestId + " must define patch.");
        string patchText = GetManifestText(patchValue);
        Assert.False(string.IsNullOrWhiteSpace(patchText), "Case " + manifestId + " must define patch.");

        manifest.TryGetProperty("input", out JsonElement input);
        string fixture = input.ValueKind == JsonValueKind.Object && input.TryGetProperty("fixture", out JsonElement fixtureValue) && fixtureValue.ValueKind == JsonValueKind.String ? fixtureValue.GetString() ?? string.Empty : string.Empty;
        string bodyXml = input.ValueKind == JsonValueKind.Object && input.TryGetProperty("bodyXml", out JsonElement bodyValue) ? GetManifestText(bodyValue) : string.Empty;
        string headerXml = input.ValueKind == JsonValueKind.Object && input.TryGetProperty("headerXml", out JsonElement headerValue) ? GetManifestText(headerValue) : string.Empty;
        string footerXml = input.ValueKind == JsonValueKind.Object && input.TryGetProperty("footerXml", out JsonElement footerValue) ? GetManifestText(footerValue) : string.Empty;
        string stylesXml = input.ValueKind == JsonValueKind.Object && input.TryGetProperty("stylesXml", out JsonElement stylesValue) ? GetManifestText(stylesValue) : string.Empty;
        string commentsXml = input.ValueKind == JsonValueKind.Object && input.TryGetProperty("commentsXml", out JsonElement commentsValue) ? GetManifestText(commentsValue) : string.Empty;
        string commentsExtendedXml = input.ValueKind == JsonValueKind.Object && input.TryGetProperty("commentsExtendedXml", out JsonElement commentsExtendedValue) ? GetManifestText(commentsExtendedValue) : string.Empty;
        string commentsIdsXml = input.ValueKind == JsonValueKind.Object && input.TryGetProperty("commentsIdsXml", out JsonElement commentsIdsValue) ? GetManifestText(commentsIdsValue) : string.Empty;
        string mediaFileName = ReadMediaProperty(input, "fileName");
        string mediaContentType = ReadMediaProperty(input, "contentType");
        string mediaText = ReadMediaProperty(input, "text");
        string hyperlinkTarget = input.ValueKind == JsonValueKind.Object && input.TryGetProperty("hyperlink", out JsonElement hyperlinkValue) ? GetManifestText(hyperlinkValue) : string.Empty;
        Assert.True(!string.IsNullOrWhiteSpace(fixture) || !string.IsNullOrWhiteSpace(bodyXml), "Case " + manifestId + " must define input.bodyXml or input.fixture.");
        using TempDirectory temp = TempDirectory.Create();
        string inputPath = Path.Combine(temp.Path, "input.docx");
        string patchPath = Path.Combine(temp.Path, "edit.docxpatch");
        string outputPath = Path.Combine(temp.Path, "output.docx");

        if (manifest.TryGetProperty("assets", out JsonElement assets) && assets.ValueKind == JsonValueKind.Object)
        {
            string assetDir = Path.Combine(temp.Path, "assets");
            Directory.CreateDirectory(assetDir);
            foreach (JsonProperty asset in assets.EnumerateObject())
            {
                string assetPath = Path.Combine(assetDir, asset.Name);
                string assetText = asset.Value.GetString() ?? string.Empty;
                if (assetText.StartsWith("base64:", StringComparison.Ordinal))
                {
                    File.WriteAllBytes(assetPath, Convert.FromBase64String(assetText["base64:".Length..]));
                }
                else
                {
                    File.WriteAllText(assetPath, assetText, Encoding.UTF8);
                }

                patchText = patchText.Replace("{{asset:" + asset.Name + "}}", assetPath);
            }
        }

        if (!string.IsNullOrWhiteSpace(fixture))
        {
            string fixtureRoot = Path.GetFullPath(Path.Combine(repoRoot, "edit-cases", "fixtures")) + Path.DirectorySeparatorChar;
            string resolved = Path.GetFullPath(Path.Combine(repoRoot, fixture));
            Assert.StartsWith(fixtureRoot, resolved, StringComparison.OrdinalIgnoreCase);
            File.Copy(resolved, inputPath);
        }
        else
        {
            BuildCaseDocx(inputPath, bodyXml, headerXml, footerXml, stylesXml, commentsXml, mediaFileName, mediaContentType, mediaText, hyperlinkTarget, commentsExtendedXml, commentsIdsXml);
        }

        File.WriteAllText(patchPath, patchText + Environment.NewLine, Encoding.UTF8);
        manifest.TryGetProperty("applyOptions", out JsonElement applyOptions);
        var applyArgs = new List<string> { "apply", inputPath, patchPath, "-o", outputPath, "--json" };
        if (applyOptions.ValueKind == JsonValueKind.Object)
        {
            if (applyOptions.TryGetProperty("author", out JsonElement author) && author.ValueKind == JsonValueKind.String)
            {
                applyArgs.Add("--author");
                applyArgs.Add(author.GetString() ?? string.Empty);
            }

            if (applyOptions.TryGetProperty("timestampUtc", out JsonElement timestamp) && timestamp.ValueKind == JsonValueKind.String)
            {
                applyArgs.Add("--timestamp-utc");
                applyArgs.Add(timestamp.GetString() ?? string.Empty);
            }
        }

        string? mode = trackChangesOverride;
        if (mode is null && applyOptions.ValueKind == JsonValueKind.Object && applyOptions.TryGetProperty("trackChanges", out JsonElement trackChanges) && trackChanges.ValueKind == JsonValueKind.String)
        {
            mode = trackChanges.GetString();
        }

        if (mode is not null)
        {
            applyArgs.Add("--track-changes");
            applyArgs.Add(mode);
        }

        CliTests.CliResult apply = CliTests.RunCli(applyArgs.ToArray());
        manifest.TryGetProperty("expect", out JsonElement expect);
        bool expectedApplySuccess = true;
        if (expect.ValueKind == JsonValueKind.Object && expect.TryGetProperty("applySuccess", out JsonElement applySuccessValue) && applySuccessValue.ValueKind == JsonValueKind.False)
        {
            expectedApplySuccess = false;
        }
        Assert.True(expectedApplySuccess == (apply.ExitCode == 0), "Case " + manifestId + " apply success mismatch.");
        Assert.False(string.IsNullOrWhiteSpace(apply.Output), "Case " + manifestId + " apply did not emit JSON output.");
        using JsonDocument applyJson = JsonDocument.Parse(apply.Output);
        List<JsonElement> diagnostics = applyJson.RootElement.GetProperty("Diagnostics").EnumerateArray().ToList();
        int expectedDiagnosticCount = expect.ValueKind == JsonValueKind.Object && expect.TryGetProperty("diagnosticCount", out JsonElement diagnosticCountValue) ? diagnosticCountValue.GetInt32() : 0;
        Assert.True(diagnostics.Count == expectedDiagnosticCount, "Case " + manifestId + " diagnostic count mismatch.");

        if (expect.ValueKind == JsonValueKind.Object && expect.TryGetProperty("diagnosticCodes", out JsonElement diagnosticCodes))
        {
            List<string> actualCodes = diagnostics.Select(diagnostic => diagnostic.GetProperty("Code").GetString() ?? string.Empty).ToList();
            foreach (JsonElement code in diagnosticCodes.EnumerateArray())
            {
                Assert.Contains(code.GetString() ?? string.Empty, actualCodes);
            }
        }

        if (expect.ValueKind == JsonValueKind.Object && expect.TryGetProperty("diagnosticMessagesContain", out JsonElement diagnosticMessages))
        {
            List<string> actualMessages = diagnostics.Select(diagnostic => diagnostic.GetProperty("Message").GetString() ?? string.Empty).ToList();
            foreach (JsonElement expectedMessage in diagnosticMessages.EnumerateArray())
            {
                string snippet = expectedMessage.GetString() ?? string.Empty;
                Assert.Contains(actualMessages, message => message.Contains(snippet, StringComparison.Ordinal));
            }
        }

        if (!expectedApplySuccess)
        {
            return;
        }
        CliTests.CliResult read = CliTests.RunCli("read", outputPath, "--json");
        Assert.True(read.ExitCode == 0, "Case " + manifestId + " readback failed.");
        using JsonDocument readJson = JsonDocument.Parse(read.Output);
        if (expect.ValueKind == JsonValueKind.Object && expect.TryGetProperty("paragraphs", out JsonElement paragraphs))
        {
            AssertStringArraysEqual(manifestId, "Paragraphs", paragraphs.EnumerateArray().Select(item => item.GetString() ?? string.Empty).ToList(), readJson.RootElement.GetProperty("Paragraphs").EnumerateArray().Select(item => item.GetProperty("Text").GetString() ?? string.Empty).ToList());
        }

        if (expect.ValueKind == JsonValueKind.Object && expect.TryGetProperty("tableCells", out JsonElement tableCells))
        {
            List<string> actualCells = new();
            foreach (JsonElement table in readJson.RootElement.GetProperty("Tables").EnumerateArray())
            {
                foreach (JsonElement cell in table.GetProperty("Cells").EnumerateArray())
                {
                    actualCells.Add(cell.GetProperty("Text").GetString() ?? string.Empty);
                }
            }

            AssertStringArraysEqual(manifestId, "Table cells", tableCells.EnumerateArray().Select(item => item.GetString() ?? string.Empty).ToList(), actualCells);
        }
        if (expect.ValueKind == JsonValueKind.Object && expect.TryGetProperty("sections", out JsonElement sections))
        {
            List<string> actualSections = new();
            foreach (JsonElement section in readJson.RootElement.GetProperty("Sections").EnumerateArray())
            {
                actualSections.Add(section.GetProperty("Columns").GetInt32().ToString() + "|" + section.GetProperty("Orientation").GetString());
            }

            AssertStringArraysEqual(manifestId, "Sections", sections.EnumerateArray().Select(item => item.GetString() ?? string.Empty).ToList(), actualSections);
        }

        if (expect.ValueKind == JsonValueKind.Object && expect.TryGetProperty("images", out JsonElement images))
        {
            AssertStringArraysEqual(manifestId, "Images", images.EnumerateArray().Select(item => item.GetString() ?? string.Empty).ToList(), readJson.RootElement.GetProperty("Images").EnumerateArray().Select(item => item.GetProperty("PartName").GetString() ?? string.Empty).ToList());
        }

        if (expect.ValueKind == JsonValueKind.Object && expect.TryGetProperty("imageDescriptions", out JsonElement imageDescriptions))
        {
            AssertStringArraysEqual(manifestId, "Image descriptions", imageDescriptions.EnumerateArray().Select(item => item.GetString() ?? string.Empty).ToList(), readJson.RootElement.GetProperty("Images").EnumerateArray().Select(item => item.GetProperty("Description").GetString() ?? string.Empty).ToList());
        }

        if (expect.ValueKind == JsonValueKind.Object && expect.TryGetProperty("imageExtents", out JsonElement imageExtents))
        {
            AssertStringArraysEqual(manifestId, "Image extents", imageExtents.EnumerateArray().Select(item => item.GetString() ?? string.Empty).ToList(), readJson.RootElement.GetProperty("Images").EnumerateArray().Select(item => item.GetProperty("WidthEmu").GetInt64().ToString() + "x" + item.GetProperty("HeightEmu").GetInt64().ToString()).ToList());
        }

        if (expect.ValueKind == JsonValueKind.Object && expect.TryGetProperty("controlChecked", out JsonElement controlChecked))
        {
            AssertStringArraysEqual(manifestId, "Control checked states", controlChecked.EnumerateArray().Select(item => item.GetString() ?? string.Empty).ToList(), readJson.RootElement.GetProperty("ContentControls").EnumerateArray().Select(item => item.GetProperty("Checked").GetBoolean().ToString().ToLowerInvariant()).ToList());
        }

        if (expect.ValueKind == JsonValueKind.Object && expect.TryGetProperty("hyperlinkUris", out JsonElement hyperlinkUris))
        {
            AssertStringArraysEqual(manifestId, "Hyperlink URIs", hyperlinkUris.EnumerateArray().Select(item => item.GetString() ?? string.Empty).ToList(), readJson.RootElement.GetProperty("Hyperlinks").EnumerateArray().Select(item => item.GetProperty("Uri").GetString() ?? string.Empty).ToList());
        }

        if (expect.ValueKind == JsonValueKind.Object && expect.TryGetProperty("allStoryParagraphs", out JsonElement allStoryParagraphs))
        {
            CliTests.CliResult allStoryRead = CliTests.RunCli("read", outputPath, "--headers-footers", "--json");
            Assert.True(allStoryRead.ExitCode == 0, "Case " + manifestId + " all-story readback failed.");
            using JsonDocument allStoryJson = JsonDocument.Parse(allStoryRead.Output);
            AssertStringArraysEqual(manifestId, "All-story paragraphs", allStoryParagraphs.EnumerateArray().Select(item => item.GetString() ?? string.Empty).ToList(), allStoryJson.RootElement.GetProperty("Paragraphs").EnumerateArray().Select(item => item.GetProperty("Text").GetString() ?? string.Empty).ToList());
        }
        if (expectNoChangeSummary)
        {
            CliTests.CliResult noChange = CliTests.RunCli("changes", outputPath, "--json");
            Assert.True(noChange.ExitCode == 0, "Case " + manifestId + " changes readback failed.");
            using JsonDocument changesJson = JsonDocument.Parse(noChange.Output);
            int total = 0;
            foreach (JsonElement summary in changesJson.RootElement.GetProperty("Summary").EnumerateArray())
            {
                total += summary.GetProperty("Count").GetInt32();
            }

            Assert.True(total == 0, "Case " + manifestId + " expected no changes.");
            return;
        }

        if (expect.ValueKind == JsonValueKind.Object && expect.TryGetProperty("changeSummary", out JsonElement changeSummary))
        {
            CliTests.CliResult changes = CliTests.RunCli("changes", outputPath, "--json");
            Assert.True(changes.ExitCode == 0, "Case " + manifestId + " changes readback failed.");
            using JsonDocument changesJson = JsonDocument.Parse(changes.Output);
            var actualSummary = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (JsonElement summary in changesJson.RootElement.GetProperty("Summary").EnumerateArray())
            {
                actualSummary[summary.GetProperty("Type").GetString() ?? string.Empty] = summary.GetProperty("Count").GetInt32();
            }

            foreach (JsonProperty expected in changeSummary.EnumerateObject())
            {
                int actual = actualSummary.TryGetValue(expected.Name, out int count) ? count : 0;
                Assert.True(actual == expected.Value.GetInt32(), "Case " + manifestId + " change summary mismatch.");
            }
        }

        if (expect.ValueKind == JsonValueKind.Object && expect.TryGetProperty("commentResolved", out JsonElement commentResolved))
        {
            CliTests.CliResult commentState = CliTests.RunCli("changes", outputPath, "--json");
            Assert.True(commentState.ExitCode == 0, "Case " + manifestId + " comment state readback failed.");
            using JsonDocument commentStateJson = JsonDocument.Parse(commentState.Output);
            var actualResolved = new Dictionary<string, bool>(StringComparer.Ordinal);
            foreach (JsonElement summary in commentStateJson.RootElement.GetProperty("CommentSummary").EnumerateArray())
            {
                string commentKey = summary.GetProperty("CommentId").GetString() ?? string.Empty;
                bool commentResolvedValue = summary.TryGetProperty("Resolved", out JsonElement resolvedValue) && resolvedValue.ValueKind == JsonValueKind.True;
                actualResolved[commentKey] = commentResolvedValue;
            }

            foreach (JsonProperty expected in commentResolved.EnumerateObject())
            {
                Assert.True(actualResolved.TryGetValue(expected.Name, out bool resolved) && resolved == expected.Value.GetBoolean(), "Case " + manifestId + " comment resolved mismatch for " + expected.Name + " (actual: " + string.Join(",", actualResolved.Select(pair => pair.Key + "=" + pair.Value)) + ").");
                Assert.True(actualResolved.Count == commentResolved.EnumerateObject().Count(), "Case " + manifestId + " comment resolved count mismatch (actual count " + actualResolved.Count + ").");
            }
        }

        if (expect.ValueKind == JsonValueKind.Object && expect.TryGetProperty("bookmarkNames", out JsonElement bookmarkNames))
        {
            CliTests.CliResult bookmarkState = CliTests.RunCli("read", outputPath, "--json");
            Assert.True(bookmarkState.ExitCode == 0, "Case " + manifestId + " bookmark readback failed.");
            using JsonDocument bookmarkStateJson = JsonDocument.Parse(bookmarkState.Output);
            AssertStringArraysEqual(manifestId, "Bookmark names", bookmarkNames.EnumerateArray().Select(item => item.GetString() ?? string.Empty).ToList(), bookmarkStateJson.RootElement.GetProperty("Bookmarks").EnumerateArray().Select(item => item.GetProperty("Name").GetString() ?? string.Empty).ToList());
        }
    }

    private static void AssertStringArraysEqual(string caseId, string label, IReadOnlyList<string> expected, IReadOnlyList<string> actual)
    {
        Assert.True(expected.Count == actual.Count, "Case " + caseId + " " + label + " count mismatch.");
        for (int i = 0; i < expected.Count; i++)
        {
            Assert.True(string.Equals(expected[i], actual[i], StringComparison.Ordinal), "Case " + caseId + " " + label + " mismatch.");
        }
    }
}
