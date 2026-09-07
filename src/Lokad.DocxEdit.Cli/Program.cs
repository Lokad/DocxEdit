using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Runtime.InteropServices;
using Lokad.DocxEdit;

return ProgramMain.Run(args);

/// <summary>Command-line harness entry point. Public for in-process testing; the published surface is the <c>docxedit</c> command line.</summary>
public static class ProgramMain
{
    private static readonly JsonSerializerOptions IndentedJsonOptions = CreateJsonOptions(writeIndented: true);
    private static readonly JsonSerializerOptions CompactJsonOptions = CreateJsonOptions(writeIndented: false);

    private static JsonSerializerOptions CreateJsonOptions(bool writeIndented)
    {
        return new JsonSerializerOptions
        {
            WriteIndented = writeIndented,
            Converters = { new DocxOrientationJsonConverter(), new DocxTargetStatusJsonConverter(), new DocxTargetSourceJsonConverter(), new DocxTargetReasonJsonConverter(), new DocxRefreshPolicyJsonConverter(), new DocxLabelStatusJsonConverter(), new DocxLabelSourceJsonConverter(), new DocxVerticalMergeJsonConverter() }
        };
    }

    private static JsonSerializerOptions JsonOptionsFor(ParsedOptions options)
    {
        return options.Flags.Contains("--compact") ? CompactJsonOptions : IndentedJsonOptions;
    }

    public static int Run(string[] args)
    {
        try
        {
            if (args.Length == 2 && args[0] == "help")
            {
                if (DocxHelp.TryRenderTopic(args[1], out string helpText))
                {
                    Console.Write(helpText);
                    return 0;
                }

                return InvalidUsage($"Unknown help topic '{args[1]}'.");
            }

            if (args.Length == 0 || IsHelp(args[0]))
            {
                Console.Write(DocxHelp.RenderOverview());
                return 0;
            }

            ParsedOptions options = ParsedOptions.Parse(args);
            if (options.Error is not null)
            {
                Console.Error.WriteLine(options.Error);
                return 2;
            }

            return options.Command switch
            {
                "read" => RunRead(options),
                "outline" => RunOutline(options),
                "find" => RunFind(options),
                "dump" => RunDump(options),
                "context" => RunContext(options),
                "styles" => RunStyles(options),
                "media" => RunMedia(options),
                "validate" => RunValidate(options),
                "changes" => RunChanges(options),
                "check" => RunCheck(options),
                "apply" => RunApply(options),
                "catalog" => RunCatalog(options),
                "version" => RunVersion(options),
                _ => InvalidUsage($"Unknown command '{options.Command}'.\nValid commands: {ValidCommands()}\nSee 'docxedit --help'.")
            };
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.Message);
            return 4;
        }
    }

    private static Stream OpenInputFile(string path)
    {
        return path == "-" ? Console.OpenStandardInput() : File.OpenRead(path);
    }

    private static TextReader OpenPatchFile(string path)
    {
        return path == "-" ? Console.In : File.OpenText(path);
    }

    private static Stream CreateOutputFile(string path)
    {
        return path == "-" ? Console.OpenStandardOutput() : File.Create(path);
    }

    private static StringComparison PathComparison =>
        OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;

    private static bool IsStandardStreamPath(string? path) => string.IsNullOrEmpty(path) || path == "-";

    private static bool SameCliPath(string first, string second) =>
        string.Equals(Path.GetFullPath(first), Path.GetFullPath(second), PathComparison);

    // Validates writer/reader path collisions before any writer is opened.
    // Returns an InvalidUsage message, or null when the patch file set is safe.
    // Writers are --output/--report/--diagnostics; readers are the input
    // document, the patch file, and (for changes) the operation report.
    private static string? ValidatePatchCollisions(
        ParsedOptions options,
        string inputPath,
        string patchPath,
        string? operationReportPath)
    {
        var readers = new List<(string Label, string Path)>();
        if (!IsStandardStreamPath(inputPath))
        {
            readers.Add(("input document", inputPath));
        }

        if (!IsStandardStreamPath(patchPath))
        {
            readers.Add(("patch file", patchPath));
        }

        if (!IsStandardStreamPath(operationReportPath))
        {
            readers.Add(("operation report", operationReportPath!));
        }

        var writers = new List<(string Label, string? Path)>
        {
            ("--output", options.OutputPath),
            ("--report", options.ReportPath),
            ("--diagnostics", options.DiagnosticsPath)
        };
        var seenWriters = new List<(string Label, string Path)>();
        foreach ((string label, string? path) in writers)
        {
            if (IsStandardStreamPath(path))
            {
                continue;
            }

            foreach ((string readerLabel, string readerPath) in readers)
            {
                if (SameCliPath(path!, readerPath))
                {
                    return $"Invalid {label} '{path}': it must differ from the {readerLabel} '{readerPath}'.";
                }
            }

            foreach ((string seenLabel, string seenPath) in seenWriters)
            {
                if (SameCliPath(path!, seenPath))
                {
                    return $"Invalid {label} '{path}': it must differ from the {seenLabel} '{seenPath}'.";
                }
            }

            seenWriters.Add((label, path!));
        }

        return null;
    }

    private static string? ValidateInputCollisions(ParsedOptions options, string inputPath)
    {
        if (!IsStandardStreamPath(options.DiagnosticsPath) &&
            !IsStandardStreamPath(inputPath) &&
            SameCliPath(options.DiagnosticsPath!, inputPath))
        {
            return $"Invalid --diagnostics '{options.DiagnosticsPath}': it must differ from the input document '{inputPath}'.";
        }

        return null;
    }

    private static int RunInputCommand<T>(
        ParsedOptions options,
        string usage,
        Func<Stream, T> execute,
        Func<T, IReadOnlyList<DocxDiagnostic>> getDiagnostics,
        Func<T, bool> getSuccess,
        Action<T> writeText)
        where T : notnull
    {
        if (options.Positionals.Count != 1)
        {
            return InvalidUsage(usage);
        }

        if (ValidateInputCollisions(options, options.Positionals[0]) is { } inputCollision)
        {
            return InvalidUsage(inputCollision);
        }

        using Stream input = OpenInputFile(options.Positionals[0]);
        T result = execute(input);
        return FinishCommand(options, result, getDiagnostics, getSuccess, writeText);
    }

    private static int RunPatchCommand<T>(
        ParsedOptions options,
        string usage,
        Func<Stream, TextReader, T> execute,
        Func<T, IReadOnlyList<DocxDiagnostic>> getDiagnostics,
        Func<T, bool> getSuccess,
        Action<T> writeText,
        Action<T> writeReport)
        where T : notnull
    {
        if (options.Positionals.Count != 2)
        {
            return InvalidUsage(usage);
        }

        using Stream input = OpenInputFile(options.Positionals[0]);
        using TextReader patch = OpenPatchFile(options.Positionals[1]);
        T result = execute(input, patch);
        writeReport(result);
        return FinishCommand(options, result, getDiagnostics, getSuccess, writeText);
    }

    private static int FinishCommand<T>(
        ParsedOptions options,
        T result,
        Func<T, IReadOnlyList<DocxDiagnostic>> getDiagnostics,
        Func<T, bool> getSuccess,
        Action<T> writeText)
        where T : notnull
    {
        IReadOnlyList<DocxDiagnostic> diagnostics = getDiagnostics(result);
        WriteDiagnostics(options.DiagnosticsPath, diagnostics, JsonOptionsFor(options));
        if (!options.Json)
        {
            WriteErrorDiagnostics(diagnostics, options.Strict);
        }

        if (options.Json)
        {
            WriteJson(result, JsonOptionsFor(options));
        }
        else
        {
            writeText(result);
        }

        return ExitCode(getSuccess(result), diagnostics, options.Strict);
    }

    private static void WriteErrorDiagnostics(IReadOnlyList<DocxDiagnostic> diagnostics, bool strict)
    {
        foreach (DocxDiagnostic diagnostic in diagnostics)
        {
            bool show = diagnostic.Severity == DocxSeverity.Error ||
                (strict && diagnostic.Severity == DocxSeverity.Warning);
            if (!show)
            {
                continue;
            }

            var message = new StringBuilder();
            message.Append(diagnostic.Severity).Append(' ').Append(diagnostic.Code).Append(": ").Append(diagnostic.Message);
            if (diagnostic.TargetId is not null)
            {
                message.Append(" target=").Append(diagnostic.TargetId);
            }

            if (diagnostic.PartName is not null)
            {
                message.Append(" part=").Append(diagnostic.PartName);
            }

            Console.Error.WriteLine(message.ToString());
        }
    }

    private static int RunRead(ParsedOptions options)
    {
        if (ValidateInputCollisions(options, options.Positionals[0]) is { } readCollision)
        {
            return InvalidUsage(readCollision);
        }

        using Stream input = OpenInputFile(options.Positionals[0]);
        DocxReadResult result = new DocxEditor().Read(input, new DocxReadOptions
        {
            IncludeHeadersFooters = options.Flags.Contains("--headers-footers"),
            TextView = options.TextView,
            MaxText = options.MaxText ?? 4_000
        });
        bool summary = options.Flags.Contains("--summary");
        WriteDiagnostics(options.DiagnosticsPath, result.Diagnostics, JsonOptionsFor(options));
        if (!options.Json)
        {
            WriteErrorDiagnostics(result.Diagnostics, options.Strict);
        }

        if (options.Json)
        {
            WriteJson(summary ? BuildReadSummaryJson(result) : result, JsonOptionsFor(options));
        }
        else if (summary)
        {
            Console.Write(DocxTextRenderer.RenderReadSummary(result));
        }
        else
        {
            Console.Write(DocxTextRenderer.RenderRead(result));
        }

        return ExitCode(result.Success, result.Diagnostics, options.Strict);
    }

    private static object BuildReadSummaryJson(DocxReadResult result)
    {
        return new
        {
            result.Success,
            result.Diagnostics,
            result.PartNames,
            result.MainDocumentPartName,
            Counts = new
            {
                Paragraphs = result.Paragraphs.Count,
                Tables = result.Tables.Count,
                Images = result.Images.Count,
                Sections = result.Sections.Count,
                Bookmarks = result.Bookmarks.Count,
                ContentControls = result.ContentControls.Count,
                Fields = result.Fields.Count,
                Hyperlinks = result.Hyperlinks.Count
            },
            Stories = result.Paragraphs
                .GroupBy(static paragraph => paragraph.Story)
                .OrderBy(static group => group.Key, StringComparer.Ordinal)
                .Select(static group => new { Story = group.Key, Paragraphs = group.Count() })
                .ToArray()
        };
    }

    private static int RunOutline(ParsedOptions options)
    {
        return RunInputCommand(
            options,
            CommandUsageError("outline"),
            input => new DocxEditor().Outline(input, new DocxOutlineOptions
            {
                IncludeHeadersFooters = options.Flags.Contains("--headers-footers"),
                TextView = options.TextView,
                MaxText = options.MaxText ?? 4_000
            }),
            static result => result.Diagnostics,
            static result => result.Success,
            result => Console.Write(DocxTextRenderer.RenderOutline(result)));
    }

    private static int RunFind(ParsedOptions options)
    {
        // Find takes document plus query, unlike the single-input commands, so it opens
        // and finishes explicitly instead of routing through the one-positional helper. Positional
        // arity itself is validated from the command catalog during parsing.
        string query = options.Positionals[1];
        if (ValidateInputCollisions(options, options.Positionals[0]) is { } findCollision)
        {
            return InvalidUsage(findCollision);
        }

        using Stream input = OpenInputFile(options.Positionals[0]);
        DocxFindResult result = new DocxEditor().Find(input, query, new DocxFindOptions
        {
            IncludeHeadersFooters = options.Flags.Contains("--headers-footers"),
            TextView = options.TextView,
            MaxText = options.MaxText ?? 4_000
        });
        return FinishCommand(
            options,
            result,
            static findResult => findResult.Diagnostics,
            static findResult => findResult.Success,
            findResult => Console.Write(DocxTextRenderer.RenderFind(findResult)));
    }

    private static int RunDump(ParsedOptions options)
    {
        if (options.Id is null)
        {
            return InvalidUsage(CommandUsageError("dump"));
        }

        string id = options.Id;
        return RunInputCommand(
            options,
            CommandUsageError("dump"),
            input => new DocxEditor().Dump(input, id, new DocxDumpOptions
            {
                IncludeRuns = options.Flags.Contains("--runs"),
                TextView = options.TextView,
                MaxText = options.MaxText ?? 4_000
            }),
            static result => result.Diagnostics,
            static result => result.Success,
            result => Console.Write(DocxTextRenderer.RenderDump(result)));
    }

    private static int RunContext(ParsedOptions options)
    {
        if (options.Id is null)
        {
            return InvalidUsage(CommandUsageError("context"));
        }

        string id = options.Id;
        return RunInputCommand(
            options,
            CommandUsageError("context"),
            input => new DocxEditor().Context(input, id, new DocxContextOptions
            {
                IncludeHeadersFooters = options.Flags.Contains("--headers-footers"),
                TextView = options.TextView,
                Radius = options.Radius ?? 1,
                MaxText = options.MaxText ?? 0
            }),
            static result => result.Diagnostics,
            static result => result.Success,
            result => Console.Write(DocxTextRenderer.RenderContext(result)));
    }

    private static int RunStyles(ParsedOptions options)
    {
        return RunInputCommand(
            options,
            CommandUsageError("styles"),
            static input => new DocxEditor().Styles(input),
            static result => result.Diagnostics,
            static result => result.Success,
            result => Console.Write(DocxTextRenderer.RenderStyles(result)));
    }

    private static int RunMedia(ParsedOptions options)
    {
        string inputPath = options.Positionals[0];
        if (ValidateInputCollisions(options, inputPath) is { } mediaCollision)
        {
            return InvalidUsage(mediaCollision);
        }

        using Stream input = OpenInputFile(inputPath);
        if (options.ExtractPath is string extractDirectory)
        {
            return RunExtractMedia(options, input, extractDirectory);
        }

        DocxMediaResult result = new DocxEditor().Media(input);
        return FinishCommand(options, result, static r => r.Diagnostics, static r => r.Success, static r => Console.Write(DocxTextRenderer.RenderMedia(r)));
    }

    private static int RunExtractMedia(ParsedOptions options, Stream input, string directory)
    {
        DocxMediaExtractResult result = new DocxEditor().ExtractMedia(input);
        if (result.Success)
        {
            Directory.CreateDirectory(directory);
            foreach (DocxMediaFile file in result.Files)
            {
                File.WriteAllBytes(Path.Combine(directory, file.FileName), file.Content);
            }
        }

        WriteDiagnostics(options.DiagnosticsPath, result.Diagnostics, JsonOptionsFor(options));
        if (options.Json)
        {
            WriteJson(
                new
                {
                    result.Success,
                    result.Diagnostics,
                    Files = result.Files.Select(static file => new
                    {
                        file.ImageId,
                        file.PartName,
                        file.FileName,
                        ByteLength = file.Content.Length
                    }).ToArray()
                },
                JsonOptionsFor(options));
        }
        else
        {
            WriteErrorDiagnostics(result.Diagnostics, options.Strict);
            Console.Write(DocxTextRenderer.RenderMediaExtract(result));
        }

        return ExitCode(result.Success, result.Diagnostics, options.Strict);
    }

    private static int RunValidate(ParsedOptions options)
    {
        return RunInputCommand(
            options,
            CommandUsageError("validate"),
            input => new DocxEditor().Validate(input, new DocxValidateOptions
            {
                Profile = options.ValidationProfile,
                MaxDiagnostics = options.MaxDiagnostics ?? 500
            }),
            static result => result.Diagnostics,
            static result => result.Success,
            result => Console.Write(DocxTextRenderer.RenderValidate(result)));
    }

    private static int RunChanges(ParsedOptions options)
    {
        if (options.Positionals.Count == 1 &&
            ValidatePatchCollisions(options, options.Positionals[0], string.Empty, options.OperationReportPath) is { } changesCollision)
        {
            return InvalidUsage(changesCollision);
        }

        return RunInputCommand(
            options,
            CommandUsageError("changes"),
            input => new DocxEditor().Changes(input, new DocxChangesOptions
            {
                IncludeCommentText = options.Flags.Contains("--include-comment-text"),
                MaxCommentText = options.MaxCommentText ?? 240,
                OperationReports = ReadOperationReports(options.OperationReportPath)
            }),
            static result => result.Diagnostics,
            static result => result.Success,
            result => Console.Write(DocxTextRenderer.RenderChanges(result)));
    }

    private static int RunCheck(ParsedOptions options)
    {
        if (options.Positionals.Count == 2)
        {
            if (ValidatePatchCollisions(options, options.Positionals[0], options.Positionals[1], null) is { } checkCollision)
            {
                return InvalidUsage(checkCollision);
            }

            if (IsStandardStreamPath(options.Positionals[0]) && IsStandardStreamPath(options.Positionals[1]))
            {
                return InvalidUsage("Invalid input: the input document and the patch file cannot both use standard input ('-').");
            }
        }

        return RunPatchCommand(
            options,
            CommandUsageError("check"),
            (input, patch) => new DocxEditor().Check(input, patch, ToEditOptions(options)),
            static result => result.Diagnostics,
            static result => result.Success,
            result =>
            {
                Console.WriteLine($"{(result.Success ? "docxedit check: OK" : "docxedit check: FAILED")} (author={result.Author} timestamp={result.TimestampUtc:O})");
                Console.Write(DocxTextRenderer.RenderOperationSummary(result.Operations));
            },
            result => WriteReport(options.ReportPath, result, JsonOptionsFor(options)));
    }

    private static int RunApply(ParsedOptions options)
    {
        if (options.OutputPath is null)
        {
            return InvalidUsage(CommandUsageError("apply"));
        }

        string outputPath = options.OutputPath;
        string inputPath = options.Positionals[0];
        string patchPath = options.Positionals[1];
        if (ValidatePatchCollisions(options, inputPath, patchPath, null) is { } applyCollision)
        {
            return InvalidUsage(applyCollision);
        }

        if (IsStandardStreamPath(inputPath) && IsStandardStreamPath(patchPath))
        {
            return InvalidUsage("Invalid input: the input document and the patch file cannot both use standard input ('-').");
        }

        if (outputPath == "-")
        {
            if (options.Json)
            {
                return InvalidUsage("Invalid --json with '--output -': JSON status cannot share standard output with document bytes. Use --report <path> to capture the operation report.");
            }

            // Binary stdout owns stdout exclusively: document bytes go to
            // stdout while status and diagnostics go to stderr (or files), so
            // the captured stream is exactly the package.
            using Stream binaryInput = OpenInputFile(inputPath);
            using TextReader binaryPatch = OpenPatchFile(patchPath);
            DocxApplyResult binaryResult;
            using (Stream binaryOutput = Console.OpenStandardOutput())
            {
                binaryResult = new DocxEditor().Apply(binaryInput, binaryPatch, binaryOutput, ToEditOptions(options));
            }

            WriteReport(options.ReportPath, binaryResult, JsonOptionsFor(options));
            WriteDiagnostics(options.DiagnosticsPath, binaryResult.Diagnostics, JsonOptionsFor(options));
            WriteErrorDiagnostics(binaryResult.Diagnostics, options.Strict);
            Console.Error.WriteLine($"{(binaryResult.Success ? "docxedit apply: OK" : "docxedit apply: FAILED")} (author={binaryResult.Author} timestamp={binaryResult.TimestampUtc:O})");
            Console.Error.Write(DocxTextRenderer.RenderOperationSummary(binaryResult.Operations));
            return ExitCode(binaryResult.Success, binaryResult.Diagnostics, options.Strict);
        }

        // File outputs publish atomically: the edit lands in a temporary
        // sibling and replaces the destination only after successful editing,
        // so failed operations preserve any pre-existing destination file.
        string fullDestination = Path.GetFullPath(outputPath);
        string? directory = Path.GetDirectoryName(fullDestination);
        string tempPath = Path.Combine(
            directory ?? Directory.GetCurrentDirectory(),
            Path.GetFileName(fullDestination) + ".tmp-" + Guid.NewGuid().ToString("N") + ".docxedit-tmp");
        try
        {
            DocxApplyResult result;
            using (Stream input = OpenInputFile(inputPath))
            using (TextReader patch = OpenPatchFile(patchPath))
            using (Stream tempOutput = File.Create(tempPath))
            {
                result = new DocxEditor().Apply(input, patch, tempOutput, ToEditOptions(options));
            }

            WriteReport(options.ReportPath, result, JsonOptionsFor(options));
            if (result.Success)
            {
                File.Move(tempPath, fullDestination, overwrite: true);
            }

            return FinishCommand(
                options,
                result,
                static applyResult => applyResult.Diagnostics,
                static applyResult => applyResult.Success,
                applyResult =>
                {
                    Console.WriteLine($"{(applyResult.Success ? "docxedit apply: OK" : "docxedit apply: FAILED")} (author={applyResult.Author} timestamp={applyResult.TimestampUtc:O})");
                    Console.Write(DocxTextRenderer.RenderOperationSummary(applyResult.Operations));
                });
        }
        finally
        {
            try
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
            catch
            {
            }
        }
    }

    private static int RunCatalog(ParsedOptions options)
    {
        if (options.Json)
        {
            WriteJson(DocxHelp.Catalog, JsonOptionsFor(options));
        }
        else
        {
            Console.Write(DocxHelp.RenderOverview());
            Console.WriteLine();
            Console.Write(DocxHelp.RenderPatchTrackChangesSupportTable());
        }

        return 0;
    }

    private static int RunVersion(ParsedOptions options)
    {
        string version = typeof(DocxEditor).Assembly.GetName().Version?.ToString() ?? "unknown";
        string framework = RuntimeInformation.FrameworkDescription;
        int operations = DocxHelp.Catalog.PatchOperations.Count;
        if (options.Json)
        {
            WriteJson(new
            {
                Tool = DocxHelp.Catalog.ToolName,
                Version = version,
                Framework = framework,
                PatchOperations = operations,
            }, JsonOptionsFor(options));
        }
        else
        {
            Console.WriteLine($"docxedit {version} ({framework}), {operations} patch operations");
        }

        return 0;
    }

    private static void WriteDiagnostics(string? path, IReadOnlyList<DocxDiagnostic> diagnostics, JsonSerializerOptions jsonOptions)
    {
        if (path is null)
        {
            return;
        }

        EnsureParentDirectory(path);
        File.WriteAllText(path, JsonSerializer.Serialize(diagnostics, jsonOptions));
    }

    private static IReadOnlyList<DocxPatchOperationReport> ReadOperationReports(string? path)
    {
        if (path is null)
        {
            return [];
        }

        using Stream input = File.OpenRead(path);
        DocxApplyResult? report = JsonSerializer.Deserialize<DocxApplyResult>(input, IndentedJsonOptions);
        if (report?.Operations is null)
        {
            throw new InvalidDataException($"Operation report '{path}' is not a docxedit check/apply report JSON object.");
        }

        return report.Operations;
    }

    private static void WriteReport(string? path, object report, JsonSerializerOptions jsonOptions)
    {
        if (path is null)
        {
            return;
        }

        EnsureParentDirectory(path);
        File.WriteAllText(path, JsonSerializer.Serialize(report, jsonOptions));
    }

    private static void WriteJson(object value, JsonSerializerOptions jsonOptions)
    {
        Console.WriteLine(JsonSerializer.Serialize(value, jsonOptions));
    }

    private static DocxEditOptions ToEditOptions(ParsedOptions options)
    {
        string? patchDirectory = Path.GetDirectoryName(Path.GetFullPath(options.Positionals[1]));

        return new DocxEditOptions
        {
            TrackChanges = options.TrackChanges,
            Author = options.Author ?? "docxedit",
            TimestampUtc = options.TimestampUtc ?? DateTimeOffset.UtcNow,
            AssetProvider = new FileSystemAssetProvider(patchDirectory)
        };
    }

    private sealed class DocxOrientationJsonConverter : JsonConverter<DocxOrientation>
    {
        public override DocxOrientation Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            string? value = reader.GetString();
            if (DocxOrientationExtensions.TryParseWireValue(value, out DocxOrientation orientation))
            {
                return orientation;
            }

            throw new JsonException($"Unsupported section orientation '{value}'. Expected portrait or landscape.");
        }

        public override void Write(Utf8JsonWriter writer, DocxOrientation value, JsonSerializerOptions options)
        {
            writer.WriteStringValue(value.ToWireValue());
        }
    }

    private sealed class DocxTargetStatusJsonConverter : JsonConverter<DocxTargetStatus>
    {
        public override DocxTargetStatus Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            string? value = reader.GetString();
            if (DocxTargetStatusExtensions.TryParseWireValue(value, out DocxTargetStatus parsed))
            {
                return parsed;
            }

            throw new JsonException($"Unsupported target status '{value}'.");
        }

        public override void Write(Utf8JsonWriter writer, DocxTargetStatus value, JsonSerializerOptions options)
        {
            writer.WriteStringValue(value.ToWireValue());
        }
    }

    private sealed class DocxTargetSourceJsonConverter : JsonConverter<DocxTargetSource>
    {
        public override DocxTargetSource Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            string? value = reader.GetString();
            if (DocxTargetSourceExtensions.TryParseWireValue(value, out DocxTargetSource parsed))
            {
                return parsed;
            }

            throw new JsonException($"Unsupported target source '{value}'.");
        }

        public override void Write(Utf8JsonWriter writer, DocxTargetSource value, JsonSerializerOptions options)
        {
            writer.WriteStringValue(value.ToWireValue());
        }
    }

    private sealed class DocxTargetReasonJsonConverter : JsonConverter<DocxTargetReason>
    {
        public override DocxTargetReason Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            string? value = reader.GetString();
            if (DocxTargetReasonExtensions.TryParseWireValue(value, out DocxTargetReason parsed))
            {
                return parsed;
            }

            throw new JsonException($"Unsupported target reason '{value}'.");
        }

        public override void Write(Utf8JsonWriter writer, DocxTargetReason value, JsonSerializerOptions options)
        {
            writer.WriteStringValue(value.ToWireValue());
        }
    }

    private sealed class DocxRefreshPolicyJsonConverter : JsonConverter<DocxRefreshPolicy>
    {
        public override DocxRefreshPolicy Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            string? value = reader.GetString();
            if (DocxRefreshPolicyExtensions.TryParseWireValue(value, out DocxRefreshPolicy parsed))
            {
                return parsed;
            }

            throw new JsonException($"Unsupported refresh policy '{value}'.");
        }

        public override void Write(Utf8JsonWriter writer, DocxRefreshPolicy value, JsonSerializerOptions options)
        {
            writer.WriteStringValue(value.ToWireValue());
        }
    }

    private sealed class DocxLabelStatusJsonConverter : JsonConverter<DocxLabelStatus>
    {
        public override DocxLabelStatus Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            string? value = reader.GetString();
            if (DocxLabelStatusExtensions.TryParseWireValue(value, out DocxLabelStatus parsed))
            {
                return parsed;
            }

            throw new JsonException($"Unsupported label status '{value}'.");
        }

        public override void Write(Utf8JsonWriter writer, DocxLabelStatus value, JsonSerializerOptions options)
        {
            writer.WriteStringValue(value.ToWireValue());
        }
    }

    private sealed class DocxLabelSourceJsonConverter : JsonConverter<DocxLabelSource>
    {
        public override DocxLabelSource Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            string? value = reader.GetString();
            if (DocxLabelSourceExtensions.TryParseWireValue(value, out DocxLabelSource parsed))
            {
                return parsed;
            }

            throw new JsonException($"Unsupported label source '{value}'.");
        }

        public override void Write(Utf8JsonWriter writer, DocxLabelSource value, JsonSerializerOptions options)
        {
            writer.WriteStringValue(value.ToWireValue());
        }
    }

    private sealed class DocxVerticalMergeJsonConverter : JsonConverter<DocxVerticalMerge>
    {
        public override DocxVerticalMerge Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            string? value = reader.GetString();
            if (DocxVerticalMergeExtensions.TryParseWireValue(value, out DocxVerticalMerge parsed))
            {
                return parsed;
            }

            throw new JsonException("Unsupported vertical merge value.");
        }

        public override void Write(Utf8JsonWriter writer, DocxVerticalMerge value, JsonSerializerOptions options)
        {
            writer.WriteStringValue(value.ToWireValue());
        }
    }

    private sealed class FileSystemAssetProvider : IDocxAssetProvider
    {
        private readonly string? _baseDirectory;

        public FileSystemAssetProvider(string? baseDirectory)
        {
            _baseDirectory = baseDirectory;
        }

        public bool TryOpen(
            string reference,
            out Stream stream,
            out string? contentTypeHint,
            out string? fileNameHint)
        {
            stream = Stream.Null;
            contentTypeHint = null;
            fileNameHint = null;

            try
            {
                if (_baseDirectory is not null && !Path.IsPathRooted(reference))
                {
                    string sibling = Path.GetFullPath(Path.Combine(_baseDirectory, reference));
                    if (File.Exists(sibling))
                    {
                        stream = File.OpenRead(sibling);
                        fileNameHint = Path.GetFileName(sibling);
                        return true;
                    }
                }

                string path = Path.GetFullPath(reference);
                if (!File.Exists(path))
                {
                    return false;
                }

                stream = File.OpenRead(path);
                fileNameHint = Path.GetFileName(path);
                return true;
            }
            catch (Exception ex) when (ex is ArgumentException or IOException or NotSupportedException or UnauthorizedAccessException)
            {
                return false;
            }
        }
    }


    private static void EnsureParentDirectory(string path)
    {
        string? directory = Path.GetDirectoryName(Path.GetFullPath(path));
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }

    private static int ExitCode(bool success, IReadOnlyList<DocxDiagnostic> diagnostics, bool strict)
    {
        if (!success)
        {
            return 1;
        }

        return strict && diagnostics.Any(diagnostic => diagnostic.Severity is DocxSeverity.Warning or DocxSeverity.Error)
            ? 3
            : 0;
    }

    private static string CommandUsage(string command)
    {
        if (DocxHelp.TryGetCommand(command, out DocxCommandInfo info) && !string.IsNullOrWhiteSpace(info.Usage))
        {
            return "Usage: " + info.Usage;
        }

        return $"Usage: docxedit {command} [options]";
    }

    private static string CommandUsageError(string command)
    {
        return $"{CommandUsage(command)}\nSee 'docxedit help {command}'.";
    }

    private static string ValidCommands()
    {
        return string.Join(", ", DocxHelp.Catalog.Commands.Select(static command => command.Name));
    }

    private static int InvalidUsage(string message)
    {
        Console.Error.WriteLine(message);
        return 2;
    }

    private static bool IsHelp(string arg)
    {
        return arg is "-h" or "--help" or "help";
    }

    private sealed class ParsedOptions
    {
        public required string Command { get; init; }
        public required IReadOnlyList<string> Positionals { get; init; }
        public required HashSet<string> Flags { get; init; }
        public required bool Json { get; init; }
        public required bool Strict { get; init; }
        public required string? DiagnosticsPath { get; init; }
        public required string? ReportPath { get; init; }
        public required string? OperationReportPath { get; init; }
        public required string? OutputPath { get; init; }
        public required string? Id { get; init; }
        public required string? ExtractPath { get; init; }
        public required int? MaxText { get; init; }
        public required int? MaxCommentText { get; init; }
        public required int? MaxDiagnostics { get; init; }
        public required int? Radius { get; init; }
        public required TrackChangesMode TrackChanges { get; init; }
        public required string? Author { get; init; }
        public required DateTimeOffset? TimestampUtc { get; init; }
        public required DocxTextView TextView { get; init; }
        public required DocxValidationProfile ValidationProfile { get; init; }
        public required string? Error { get; init; }

        public static ParsedOptions Parse(string[] args)
        {
            string command = args[0];
            var positionals = new List<string>();
            var flags = new HashSet<string>(StringComparer.Ordinal);
            bool json = false;
            bool strict = false;
            string? diagnosticsPath = null;
            string? reportPath = null;
            string? operationReportPath = null;
            string? outputPath = null;
            string? id = null;
            string? extractPath = null;
            int? maxText = null;
            int? maxCommentText = null;
            int? maxDiagnostics = null;
            int? radius = null;
            TrackChangesMode trackChanges = TrackChangesMode.Off;
            string? author = null;
            DateTimeOffset? timestampUtc = null;
            DocxTextView textView = DocxTextView.Final;
            DocxValidationProfile validationProfile = DocxValidationProfile.Structural;
            var seenFlags = new HashSet<string>(StringComparer.Ordinal);

            static bool TryParseTimestampUtc(string value, out DateTimeOffset timestamp)
            {
                string[] formats = ["O", "yyyy-MM-ddTHH:mm:ss.FFFFFFFK", "yyyy-MM-ddTHH:mm:ssK", "yyyy-MM-dd"];
                if (DateTimeOffset.TryParseExact(value, formats, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out timestamp))
                {
                    timestamp = timestamp.ToUniversalTime();
                    return true;
                }

                timestamp = default;
                return false;
            }

            for (int i = 1; i < args.Length; i++)
            {
                string arg = args[i];
                switch (arg)
                {
                    case "--json":
                        seenFlags.Add("--json");
                        json = true;
                        break;
                    case "--strict":
                        seenFlags.Add("--strict");
                        strict = true;
                        break;
                    case "--compact":
                        seenFlags.Add(arg);
                        flags.Add(arg);
                        break;
                    case "--runs":
                    case "--headers-footers":
                    case "--summary":
                    case "--include-comment-text":
                        seenFlags.Add(arg);
                        flags.Add(arg);
                        break;
                    case "--diagnostics":
                        seenFlags.Add("--diagnostics");
                        if (!TryReadValue(args, ref i, out diagnosticsPath))
                        {
                            return WithError(command, "Missing value for --diagnostics.");
                        }

                        break;
                    case "--report":
                        seenFlags.Add("--report");
                        if (!TryReadValue(args, ref i, out reportPath))
                        {
                            return WithError(command, "Missing value for --report.");
                        }

                        break;
                    case "--operation-report":
                        seenFlags.Add("--operation-report");
                        if (!TryReadValue(args, ref i, out operationReportPath))
                        {
                            return WithError(command, "Missing value for --operation-report.");
                        }

                        break;
                    case "--extract":
                        seenFlags.Add("--extract");
                        if (!TryReadValue(args, ref i, out extractPath))
                        {
                            return WithError(command, "Missing value for --extract.");
                        }

                        break;
                    case "--max-text":
                        seenFlags.Add("--max-text");
                        if (!TryReadValue(args, ref i, out string? maxTextValue))
                        {
                            return WithError(command, "Missing value for --max-text.");
                        }

                        if (!int.TryParse(maxTextValue, out int parsedMaxText) || parsedMaxText < 0)
                        {
                            return WithError(command, "Invalid value for --max-text.");
                        }

                        maxText = parsedMaxText;
                        break;
                    case "--max-comment-text":
                        seenFlags.Add("--max-comment-text");
                        if (!TryReadValue(args, ref i, out string? maxCommentTextValue))
                        {
                            return WithError(command, "Missing value for --max-comment-text.");
                        }

                        if (!int.TryParse(maxCommentTextValue, out int parsedMaxCommentText) || parsedMaxCommentText < 0)
                        {
                            return WithError(command, "Invalid value for --max-comment-text.");
                        }

                        maxCommentText = parsedMaxCommentText;
                        break;
                    case "--max-diagnostics":
                        seenFlags.Add("--max-diagnostics");
                        if (!TryReadValue(args, ref i, out string? maxDiagnosticsValue))
                        {
                            return WithError(command, "Missing value for --max-diagnostics.");
                        }

                        if (!int.TryParse(maxDiagnosticsValue, out int parsedMaxDiagnostics) || parsedMaxDiagnostics < 1)
                        {
                            return WithError(command, "Invalid value for --max-diagnostics.");
                        }

                        maxDiagnostics = parsedMaxDiagnostics;
                        break;
                    case "--radius":
                        seenFlags.Add("--radius");
                        if (!TryReadValue(args, ref i, out string? radiusValue))
                        {
                            return WithError(command, "Missing value for --radius.");
                        }

                        if (!int.TryParse(radiusValue, out int parsedRadius) || parsedRadius < 0)
                        {
                            return WithError(command, "Invalid value for --radius.");
                        }

                        radius = parsedRadius;
                        break;
                    case "--track-changes":
                        seenFlags.Add("--track-changes");
                        if (!TryReadValue(args, ref i, out string? trackChangesValue))
                        {
                            return WithError(command, "Missing value for --track-changes.");
                        }

                        if (!TryParseTrackChangesMode(trackChangesValue, out trackChanges))
                        {
                            return WithError(command, "Invalid value for --track-changes. Expected off, preserve, suggest, or require.");
                        }

                        break;
                    case "--view":
                        seenFlags.Add("--view");
                        if (!TryReadValue(args, ref i, out string? textViewValue))
                        {
                            return WithError(command, "Missing value for --view.");
                        }

                        if (!TryParseTextView(textViewValue, out textView))
                        {
                            return WithError(command, "Invalid value for --view. Expected final, original, or markup.");
                        }

                        break;
                    case "--profile":
                        seenFlags.Add("--profile");
                        if (!TryReadValue(args, ref i, out string? profileValue))
                        {
                            return WithError(command, "Missing value for --profile.");
                        }

                        if (!TryParseValidationProfile(profileValue, out validationProfile))
                        {
                            return WithError(command, "Invalid value for --profile. Expected structural or package.");
                        }

                        break;
                    case "--author":
                        seenFlags.Add("--author");
                        if (!TryReadValue(args, ref i, out author))
                        {
                            return WithError(command, "Missing value for --author.");
                        }

                        break;
                    case "--timestamp-utc":
                        seenFlags.Add("--timestamp-utc");
                        if (!TryReadValue(args, ref i, out string? timestampValue))
                        {
                            return WithError(command, "Missing value for --timestamp-utc.");
                        }

                        if (!TryParseTimestampUtc(timestampValue, out DateTimeOffset parsedTimestamp))
                        {
                            return WithError(command, $"Invalid value for --timestamp-utc '{timestampValue}'. Expected ISO-8601 UTC, e.g. 2026-01-01T00:00:00Z.");
                        }

                        timestampUtc = parsedTimestamp;
                        break;
                    case "-o":
                    case "--output":
                        seenFlags.Add("--output");
                        if (!TryReadValue(args, ref i, out outputPath))
                        {
                            return WithError(command, $"Missing value for {arg}.");
                        }

                        break;
                    case "--id":
                        seenFlags.Add("--id");
                        if (!TryReadValue(args, ref i, out id))
                        {
                            return WithError(command, "Missing value for --id.");
                        }

                        break;
                    default:
                        if (arg != "-" && arg.StartsWith('-'))
                        {
                            return WithError(command, $"Unknown option '{arg}'.");
                        }

                        positionals.Add(arg);
                        break;
                }
            }

            if (DocxHelp.TryGetCommand(command, out DocxCommandInfo commandInfo))
            {
                foreach (string seen in seenFlags.OrderBy(static flag => flag, StringComparer.Ordinal))
                {
                    if (!commandInfo.Options.Any(option => option.Flags.Contains(seen, StringComparer.Ordinal)))
                    {
                        return WithError(command, $"{seen} is not an option of the {command} command.");
                    }
                }

                if (positionals.Count < commandInfo.MinPositionals || positionals.Count > commandInfo.MaxPositionals)
                {
                    return WithError(command, ArityError(command, commandInfo, positionals.Count));
                }
            }

            static string ArityError(string commandName, DocxCommandInfo info, int actual)
            {
                if (info.MinPositionals == info.MaxPositionals && info.MinPositionals == 1)
                {
                    return $"The {commandName} command expects exactly 1 positional argument; received {actual}.";
                }

                string expected = info.MinPositionals == info.MaxPositionals
                    ? $"exactly {info.MinPositionals} positional arguments"
                    : $"between {info.MinPositionals} and {info.MaxPositionals} positional arguments";
                return $"The {commandName} command expects {expected}; received {actual}.";
            }

            if (diagnosticsPath == "-")
            {
                return WithError(command, "--diagnostics does not accept '-'; diagnostics output is file-only.");
            }

            if (reportPath == "-")
            {
                return WithError(command, "--report does not accept '-'; reports are file-only.");
            }

            if (operationReportPath == "-")
            {
                return WithError(command, "--operation-report does not accept '-'; operation reports are file-only.");
            }

            return new ParsedOptions
            {
                Command = command,
                Positionals = positionals,
                Flags = flags,
                Json = json,
                Strict = strict,
                DiagnosticsPath = diagnosticsPath,
                ReportPath = reportPath,
                OperationReportPath = operationReportPath,
                OutputPath = outputPath,
                Id = id,
                ExtractPath = extractPath,
                MaxText = maxText,
                MaxCommentText = maxCommentText,
                MaxDiagnostics = maxDiagnostics,
                Radius = radius,
                TrackChanges = trackChanges,
                Author = author,
                TimestampUtc = timestampUtc,
                TextView = textView,
                ValidationProfile = validationProfile,
                Error = null
            };
        }

        private static bool TryReadValue(string[] args, ref int index, [NotNullWhen(true)] out string? value)
        {
            if (index + 1 >= args.Length)
            {
                value = null;
                return false;
            }

            value = args[++index];
            return true;
        }

        private static ParsedOptions WithError(string command, string message)
        {
            string detail = DocxHelp.TryGetCommand(command, out DocxCommandInfo _) ? CommandUsageError(command) : $"Valid commands: {ValidCommands()}\nSee 'docxedit --help'.";
            return new ParsedOptions
            {
                Command = command,
                Positionals = [],
                Flags = new HashSet<string>(StringComparer.Ordinal),
                Json = false,
                Strict = false,
                DiagnosticsPath = null,
                ReportPath = null,
                OperationReportPath = null,
                OutputPath = null,
                Id = null,
                ExtractPath = null,
                MaxText = null,
                MaxCommentText = null,
                MaxDiagnostics = null,
                Radius = null,
                TrackChanges = TrackChangesMode.Off,
                Author = null,
                TimestampUtc = null,
                TextView = DocxTextView.Final,
                ValidationProfile = DocxValidationProfile.Structural,
                Error = $"{message}\n{detail}"
            };
        }

        private static bool TryParseTrackChangesMode(string value, out TrackChangesMode mode)
        {
            string normalized = value.ToLowerInvariant();
            mode = normalized switch
            {
                "off" => TrackChangesMode.Off,
                "preserve" => TrackChangesMode.Preserve,
                "suggest" => TrackChangesMode.Suggest,
                "require" => TrackChangesMode.Require,
                _ => TrackChangesMode.Off
            };
            return normalized is "off" or "preserve" or "suggest" or "require";
        }

        private static bool TryParseValidationProfile(string value, out DocxValidationProfile profile)
        {
            switch (value.ToLowerInvariant())
            {
                case "structural":
                    profile = DocxValidationProfile.Structural;
                    return true;
                case "package":
                    profile = DocxValidationProfile.Package;
                    return true;
                default:
                    profile = DocxValidationProfile.Structural;
                    return false;
            }
        }

        private static bool TryParseTextView(string value, out DocxTextView textView)
        {
            switch (value.ToLowerInvariant())
            {
                case "final":
                    textView = DocxTextView.Final;
                    return true;
                case "original":
                    textView = DocxTextView.Original;
                    return true;
                case "markup":
                    textView = DocxTextView.Markup;
                    return true;
                default:
                    textView = DocxTextView.Final;
                    return false;
            }
        }
    }
}
