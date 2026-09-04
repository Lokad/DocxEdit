using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Runtime.InteropServices;
using Lokad.DocxEdit;

return ProgramMain.Run(args);

internal static class ProgramMain
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Converters = { new DocxOrientationJsonConverter(), new DocxTargetStatusJsonConverter(), new DocxTargetSourceJsonConverter(), new DocxTargetReasonJsonConverter(), new DocxRefreshPolicyJsonConverter(), new DocxLabelStatusJsonConverter(), new DocxLabelSourceJsonConverter(), new DocxVerticalMergeJsonConverter() }
    };

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

        using Stream input = File.OpenRead(options.Positionals[0]);
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

        using Stream input = File.OpenRead(options.Positionals[0]);
        using TextReader patch = File.OpenText(options.Positionals[1]);
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
        WriteDiagnostics(options.DiagnosticsPath, diagnostics);
        if (!options.Json)
        {
            WriteErrorDiagnostics(diagnostics, options.Strict);
        }

        if (options.Json)
        {
            WriteJson(result);
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
        return RunInputCommand(
            options,
            CommandUsageError("read"),
            input => new DocxEditor().Read(input, new DocxReadOptions
            {
                IncludeHeadersFooters = options.Flags.Contains("--headers-footers"),
                IncludeAllStories = options.Flags.Contains("--all-stories"),
                TextView = options.TextView,
                MaxText = options.MaxText ?? 4_000
            }),
            static result => result.Diagnostics,
            static result => result.Success,
            result =>
            {
                if (options.Flags.Contains("--summary"))
                {
                    Console.Write(DocxTextRenderer.RenderReadSummary(result));
                }
                else
                {
                    Console.Write(DocxTextRenderer.RenderRead(result));
                }
            });
    }

    private static int RunOutline(ParsedOptions options)
    {
        return RunInputCommand(
            options,
            CommandUsageError("outline"),
            input => new DocxEditor().Outline(input, new DocxOutlineOptions
            {
                IncludeHeadersFooters = options.Flags.Contains("--headers-footers"),
                TextView = options.TextView
            }),
            static result => result.Diagnostics,
            static result => result.Success,
            result => Console.Write(DocxTextRenderer.RenderOutline(result)));
    }

    private static int RunFind(ParsedOptions options)
    {
        if (options.Positionals.Count != 2)
        {
            return InvalidUsage(CommandUsageError("find"));
        }

        string query = options.Positionals[1];
        return RunInputCommand(
            options,
            CommandUsageError("find"),
            input => new DocxEditor().Find(input, query, new DocxFindOptions
            {
                IncludeHeadersFooters = options.Flags.Contains("--headers-footers"),
                TextView = options.TextView,
                MaxText = options.MaxText ?? 4_000
            }),
            static result => result.Diagnostics,
            static result => result.Success,
            result => Console.Write(DocxTextRenderer.RenderFind(result)));
    }

    private static int RunDump(ParsedOptions options)
    {
        if (options.Positionals.Count != 1 || options.Id is null)
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
        if (options.Positionals.Count != 1 || options.Id is null)
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
        if (options.Positionals.Count != 1)
        {
            return InvalidUsage(CommandUsageError("media"));
        }

        string inputPath = options.Positionals[0];
        using Stream input = File.OpenRead(inputPath);
        DocxMediaResult result = new DocxEditor().Media(input);
        if (result.Success && options.ExtractPath is not null)
        {
            ExtractMedia(inputPath, result.Images, options.ExtractPath);
        }

        return FinishCommand(options, result, static r => r.Diagnostics, static r => r.Success, static r => Console.Write(DocxTextRenderer.RenderMedia(r)));
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
        return RunPatchCommand(
            options,
            CommandUsageError("check"),
            (input, patch) => new DocxEditor().Check(input, patch, ToEditOptions(options)),
            static result => result.Diagnostics,
            static result => result.Success,
            result =>
            {
                Console.WriteLine(result.Success ? "docxedit check: OK" : "docxedit check: FAILED");
                Console.Write(DocxTextRenderer.RenderOperationSummary(result.Operations));
            },
            result => WriteReport(options.ReportPath, result));
    }

    private static int RunApply(ParsedOptions options)
    {
        if (options.OutputPath is null)
        {
            return InvalidUsage(CommandUsageError("apply"));
        }

        string outputPath = options.OutputPath;
        return RunPatchCommand(
            options,
            CommandUsageError("apply"),
            (input, patch) =>
            {
                using Stream output = File.Create(outputPath);
                return new DocxEditor().Apply(input, patch, output, ToEditOptions(options));
            },
            static result => result.Diagnostics,
            static result => result.Success,
            result =>
            {
                Console.WriteLine(result.Success ? "docxedit apply: OK" : "docxedit apply: FAILED");
                Console.Write(DocxTextRenderer.RenderOperationSummary(result.Operations));
            },
            result => WriteReport(options.ReportPath, result));
    }

    private static int RunCatalog(ParsedOptions options)
    {
        if (options.Positionals.Count != 0)
        {
            return InvalidUsage(CommandUsageError("catalog"));
        }

        if (options.Json)
        {
            WriteJson(DocxHelp.Catalog);
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
        if (options.Positionals.Count != 0)
        {
            return InvalidUsage(CommandUsageError("version"));
        }

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
            });
        }
        else
        {
            Console.WriteLine($"docxedit {version} ({framework}), {operations} patch operations");
        }

        return 0;
    }

    private static void WriteDiagnostics(string? path, IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        if (path is null)
        {
            return;
        }

        EnsureParentDirectory(path);
        File.WriteAllText(path, JsonSerializer.Serialize(diagnostics, JsonOptions));
    }

    private static IReadOnlyList<DocxPatchOperationReport> ReadOperationReports(string? path)
    {
        if (path is null)
        {
            return [];
        }

        using Stream input = File.OpenRead(path);
        DocxApplyResult? report = JsonSerializer.Deserialize<DocxApplyResult>(input, JsonOptions);
        return report?.Operations ?? [];
    }

    private static void WriteReport(string? path, object report)
    {
        if (path is null)
        {
            return;
        }

        EnsureParentDirectory(path);
        File.WriteAllText(path, JsonSerializer.Serialize(report, JsonOptions));
    }

    private static void WriteJson(object value)
    {
        Console.WriteLine(JsonSerializer.Serialize(value, JsonOptions));
    }

    private static DocxEditOptions ToEditOptions(ParsedOptions options)
    {
        return new DocxEditOptions
        {
            TrackChanges = options.TrackChanges,
            Author = options.Author ?? "docxedit",
            TimestampUtc = options.TimestampUtc ?? DateTimeOffset.UtcNow,
            AssetProvider = FileSystemAssetProvider.Instance
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
        public static readonly FileSystemAssetProvider Instance = new();

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

    private static void ExtractMedia(string inputPath, IReadOnlyList<DocxImageInfo> images, string outputDirectory)
    {
        Directory.CreateDirectory(outputDirectory);
        using FileStream input = File.OpenRead(inputPath);
        using var archive = new System.IO.Compression.ZipArchive(input, System.IO.Compression.ZipArchiveMode.Read);
        foreach (DocxImageInfo image in images)
        {
            string entryName = image.PartName.TrimStart('/');
            System.IO.Compression.ZipArchiveEntry? entry = archive.GetEntry(entryName);
            if (entry is null)
            {
                continue;
            }

            string fileName = $"{image.Id}-{Path.GetFileName(entryName)}";
            string outputPath = Path.Combine(outputDirectory, fileName);
            using Stream source = entry.Open();
            using Stream destination = File.Create(outputPath);
            source.CopyTo(destination);
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

            for (int i = 1; i < args.Length; i++)
            {
                string arg = args[i];
                switch (arg)
                {
                    case "--json":
                        json = true;
                        break;
                    case "--strict":
                        strict = true;
                        break;
                    case "--runs":
                    case "--headers-footers":
                    case "--all-stories":
                    case "--summary":
                    case "--include-comment-text":
                        flags.Add(arg);
                        break;
                    case "--diagnostics":
                        if (!TryReadValue(args, ref i, out diagnosticsPath))
                        {
                            return WithError(command, "Missing value for --diagnostics.");
                        }

                        break;
                    case "--report":
                        if (!TryReadValue(args, ref i, out reportPath))
                        {
                            return WithError(command, "Missing value for --report.");
                        }

                        break;
                    case "--operation-report":
                        if (!TryReadValue(args, ref i, out operationReportPath))
                        {
                            return WithError(command, "Missing value for --operation-report.");
                        }

                        break;
                    case "--extract":
                        if (!TryReadValue(args, ref i, out extractPath))
                        {
                            return WithError(command, "Missing value for --extract.");
                        }

                        break;
                    case "--max-text":
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
                        if (!TryReadValue(args, ref i, out author))
                        {
                            return WithError(command, "Missing value for --author.");
                        }

                        break;
                    case "--timestamp-utc":
                        if (!TryReadValue(args, ref i, out string? timestampValue))
                        {
                            return WithError(command, "Missing value for --timestamp-utc.");
                        }

                        if (!DateTimeOffset.TryParse(timestampValue, out DateTimeOffset parsedTimestamp))
                        {
                            return WithError(command, "Invalid value for --timestamp-utc.");
                        }

                        timestampUtc = parsedTimestamp.ToUniversalTime();
                        break;
                    case "-o":
                    case "--output":
                        if (!TryReadValue(args, ref i, out outputPath))
                        {
                            return WithError(command, $"Missing value for {arg}.");
                        }

                        break;
                    case "--id":
                        if (!TryReadValue(args, ref i, out id))
                        {
                            return WithError(command, "Missing value for --id.");
                        }

                        break;
                    default:
                        if (arg.StartsWith('-'))
                        {
                            return WithError(command, $"Unknown option '{arg}'.");
                        }

                        positionals.Add(arg);
                        break;
                }
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
