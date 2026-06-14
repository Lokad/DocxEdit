using System.Text.Json;
using Lokad.DocxEdit;

return ProgramMain.Run(args);

internal static class ProgramMain
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
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
                _ => InvalidUsage($"Unknown command '{options.Command}'.")
            };
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.Message);
            return 4;
        }
    }

    private static int RunRead(ParsedOptions options)
    {
        if (options.Positionals.Count != 1)
        {
            return InvalidUsage("Usage: docxedit read input.docx [--summary] [--view final|original|markup] [--max-text <chars>] [--json] [--diagnostics <path>] [--strict]");
        }

        using Stream input = File.OpenRead(options.Positionals[0]);
        DocxReadResult result = new DocxEditor().Read(input, new DocxReadOptions
        {
            IncludeHeadersFooters = options.Flags.Contains("--headers-footers"),
            IncludeAllStories = options.Flags.Contains("--all-stories"),
            TextView = options.TextView,
            MaxText = options.MaxText ?? 4_000
        });
        WriteDiagnostics(options.DiagnosticsPath, result.Diagnostics);
        if (options.Json)
        {
            WriteJson(result);
        }
        else if (options.Flags.Contains("--summary"))
        {
            Console.Write(DocxTextRenderer.RenderReadSummary(result));
        }
        else
        {
            Console.Write(DocxTextRenderer.RenderRead(result));
        }

        return ExitCode(result.Success, result.Diagnostics, options.Strict);
    }

    private static int RunOutline(ParsedOptions options)
    {
        if (options.Positionals.Count != 1)
        {
            return InvalidUsage("Usage: docxedit outline input.docx [--view final|original|markup] [--json] [--diagnostics <path>] [--strict]");
        }

        using Stream input = File.OpenRead(options.Positionals[0]);
        DocxOutlineResult result = new DocxEditor().Outline(input, new DocxOutlineOptions
        {
            IncludeHeadersFooters = options.Flags.Contains("--headers-footers"),
            TextView = options.TextView
        });
        WriteDiagnostics(options.DiagnosticsPath, result.Diagnostics);
        if (options.Json)
        {
            WriteJson(result);
        }
        else
        {
            Console.Write(DocxTextRenderer.RenderOutline(result));
        }

        return ExitCode(result.Success, result.Diagnostics, options.Strict);
    }

    private static int RunFind(ParsedOptions options)
    {
        if (options.Positionals.Count != 2)
        {
            return InvalidUsage("Usage: docxedit find input.docx \"text\" [--view final|original|markup] [--max-text <chars>] [--json] [--diagnostics <path>] [--strict]");
        }

        using Stream input = File.OpenRead(options.Positionals[0]);
        DocxFindResult result = new DocxEditor().Find(input, options.Positionals[1], new DocxFindOptions
        {
            IncludeHeadersFooters = options.Flags.Contains("--headers-footers"),
            TextView = options.TextView,
            MaxText = options.MaxText ?? 4_000
        });
        WriteDiagnostics(options.DiagnosticsPath, result.Diagnostics);
        if (options.Json)
        {
            WriteJson(result);
        }
        else
        {
            Console.Write(DocxTextRenderer.RenderFind(result));
        }

        return ExitCode(result.Success, result.Diagnostics, options.Strict);
    }

    private static int RunDump(ParsedOptions options)
    {
        if (options.Positionals.Count != 1 || options.Id is null)
        {
            return InvalidUsage("Usage: docxedit dump input.docx --id M.P0001 [--runs] [--view final|original|markup] [--max-text <chars>] [--json] [--diagnostics <path>] [--strict]");
        }

        using Stream input = File.OpenRead(options.Positionals[0]);
        DocxDumpResult result = new DocxEditor().Dump(input, options.Id, new DocxDumpOptions
        {
            IncludeRuns = options.Flags.Contains("--runs"),
            TextView = options.TextView,
            MaxText = options.MaxText ?? 4_000
        });
        WriteDiagnostics(options.DiagnosticsPath, result.Diagnostics);
        if (options.Json)
        {
            WriteJson(result);
        }
        else
        {
            Console.Write(DocxTextRenderer.RenderDump(result));
        }

        return ExitCode(result.Success, result.Diagnostics, options.Strict);
    }

    private static int RunContext(ParsedOptions options)
    {
        if (options.Positionals.Count != 1 || options.Id is null)
        {
            return InvalidUsage("Usage: docxedit context input.docx --id M.P0001 [--radius <count>] [--view final|original|markup] [--max-text <chars>] [--json] [--diagnostics <path>] [--strict]");
        }

        using Stream input = File.OpenRead(options.Positionals[0]);
        DocxContextResult result = new DocxEditor().Context(input, options.Id, new DocxContextOptions
        {
            IncludeHeadersFooters = options.Flags.Contains("--headers-footers"),
            TextView = options.TextView,
            Radius = options.Radius ?? 1,
            MaxText = options.MaxText ?? 0
        });
        WriteDiagnostics(options.DiagnosticsPath, result.Diagnostics);
        if (options.Json)
        {
            WriteJson(result);
        }
        else
        {
            Console.Write(DocxTextRenderer.RenderContext(result));
        }

        return ExitCode(result.Success, result.Diagnostics, options.Strict);
    }

    private static int RunStyles(ParsedOptions options)
    {
        if (options.Positionals.Count != 1)
        {
            return InvalidUsage("Usage: docxedit styles input.docx [--json] [--diagnostics <path>] [--strict]");
        }

        using Stream input = File.OpenRead(options.Positionals[0]);
        DocxStylesResult result = new DocxEditor().Styles(input);
        WriteDiagnostics(options.DiagnosticsPath, result.Diagnostics);
        if (options.Json)
        {
            WriteJson(result);
        }
        else
        {
            Console.Write(DocxTextRenderer.RenderStyles(result));
        }

        return ExitCode(result.Success, result.Diagnostics, options.Strict);
    }

    private static int RunMedia(ParsedOptions options)
    {
        if (options.Positionals.Count != 1)
        {
            return InvalidUsage("Usage: docxedit media input.docx [--extract <dir>] [--json] [--diagnostics <path>] [--strict]");
        }

        string inputPath = options.Positionals[0];
        using Stream input = File.OpenRead(inputPath);
        DocxMediaResult result = new DocxEditor().Media(input);
        WriteDiagnostics(options.DiagnosticsPath, result.Diagnostics);
        if (result.Success && options.ExtractPath is not null)
        {
            ExtractMedia(inputPath, result.Images, options.ExtractPath);
        }

        if (options.Json)
        {
            WriteJson(result);
        }
        else
        {
            Console.Write(DocxTextRenderer.RenderMedia(result));
        }

        return ExitCode(result.Success, result.Diagnostics, options.Strict);
    }

    private static int RunValidate(ParsedOptions options)
    {
        if (options.Positionals.Count != 1)
        {
            return InvalidUsage("Usage: docxedit validate input.docx [--profile structural|package] [--max-diagnostics <count>] [--json] [--diagnostics <path>] [--strict]");
        }

        using Stream input = File.OpenRead(options.Positionals[0]);
        DocxValidateResult result = new DocxEditor().Validate(input, new DocxValidateOptions
        {
            Profile = options.ValidationProfile,
            MaxDiagnostics = options.MaxDiagnostics ?? 500
        });
        WriteDiagnostics(options.DiagnosticsPath, result.Diagnostics);
        if (options.Json)
        {
            WriteJson(result);
        }
        else
        {
            Console.Write(DocxTextRenderer.RenderValidate(result));
        }

        return ExitCode(result.Success, result.Diagnostics, options.Strict);
    }

    private static int RunChanges(ParsedOptions options)
    {
        if (options.Positionals.Count != 1)
        {
            return InvalidUsage("Usage: docxedit changes input.docx [--operation-report <path>] [--include-comment-text] [--max-comment-text <chars>] [--json] [--diagnostics <path>] [--strict]");
        }

        using Stream input = File.OpenRead(options.Positionals[0]);
        DocxChangesResult result = new DocxEditor().Changes(input, new DocxChangesOptions
        {
            IncludeCommentText = options.Flags.Contains("--include-comment-text"),
            MaxCommentText = options.MaxCommentText ?? 240,
            OperationReports = ReadOperationReports(options.OperationReportPath)
        });
        WriteDiagnostics(options.DiagnosticsPath, result.Diagnostics);
        if (options.Json)
        {
            WriteJson(result);
        }
        else
        {
            Console.Write(DocxTextRenderer.RenderChanges(result));
        }

        return ExitCode(result.Success, result.Diagnostics, options.Strict);
    }

    private static int RunCheck(ParsedOptions options)
    {
        if (options.Positionals.Count != 2)
        {
            return InvalidUsage("Usage: docxedit check input.docx edits.docxpatch [--track-changes <mode>] [--author <name>] [--timestamp-utc <instant>] [--json] [--report <path>] [--diagnostics <path>] [--strict]");
        }

        using Stream input = File.OpenRead(options.Positionals[0]);
        using TextReader patch = File.OpenText(options.Positionals[1]);
        DocxCheckResult result = new DocxEditor().Check(input, patch, ToEditOptions(options));
        WriteDiagnostics(options.DiagnosticsPath, result.Diagnostics);
        WriteReport(options.ReportPath, result);
        if (options.Json)
        {
            WriteJson(result);
        }
        else
        {
            Console.WriteLine(result.Success ? "docxedit check: OK" : "docxedit check: FAILED");
            Console.Write(DocxTextRenderer.RenderOperationSummary(result.Operations));
        }

        return ExitCode(result.Success, result.Diagnostics, options.Strict);
    }

    private static int RunApply(ParsedOptions options)
    {
        if (options.Positionals.Count != 2 || options.OutputPath is null)
        {
            return InvalidUsage("Usage: docxedit apply input.docx edits.docxpatch --output output.docx [--track-changes <mode>] [--author <name>] [--timestamp-utc <instant>] [--json] [--report <path>] [--diagnostics <path>] [--strict]");
        }

        using Stream input = File.OpenRead(options.Positionals[0]);
        using TextReader patch = File.OpenText(options.Positionals[1]);
        using Stream output = File.Create(options.OutputPath);
        DocxApplyResult result = new DocxEditor().Apply(input, patch, output, ToEditOptions(options));
        WriteDiagnostics(options.DiagnosticsPath, result.Diagnostics);
        WriteReport(options.ReportPath, result);
        if (options.Json)
        {
            WriteJson(result);
        }
        else
        {
            Console.WriteLine(result.Success ? "docxedit apply: OK" : "docxedit apply: FAILED");
            Console.Write(DocxTextRenderer.RenderOperationSummary(result.Operations));
        }

        return ExitCode(result.Success, result.Diagnostics, options.Strict);
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

    private static int InvalidUsage(string message)
    {
        Console.Error.WriteLine(message);
        return 2;
    }

    private static bool IsHelp(string arg)
    {
        return arg is "-h" or "--help" or "help";
    }

    private sealed record ParsedOptions(
        string Command,
        IReadOnlyList<string> Positionals,
        HashSet<string> Flags,
        bool Json,
        bool Strict,
        bool Verbose,
        string? DiagnosticsPath,
        string? ReportPath,
        string? OperationReportPath,
        string? OutputPath,
        string? Id,
        string? ExtractPath,
        int? MaxText,
        int? MaxCommentText,
        int? MaxDiagnostics,
        int? Radius,
        TrackChangesMode TrackChanges,
        string? Author,
        DateTimeOffset? TimestampUtc,
        DocxTextView TextView,
        DocxValidationProfile ValidationProfile,
        string? Error)
    {
        public static ParsedOptions Parse(string[] args)
        {
            string command = args[0];
            var positionals = new List<string>();
            var flags = new HashSet<string>(StringComparer.Ordinal);
            bool json = false;
            bool strict = false;
            bool verbose = false;
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
                    case "--verbose":
                        verbose = true;
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

                        if (!TryParseTrackChangesMode(trackChangesValue!, out trackChanges))
                        {
                            return WithError(command, "Invalid value for --track-changes. Expected off, preserve, suggest, or require.");
                        }

                        break;
                    case "--view":
                        if (!TryReadValue(args, ref i, out string? textViewValue))
                        {
                            return WithError(command, "Missing value for --view.");
                        }

                        if (!TryParseTextView(textViewValue!, out textView))
                        {
                            return WithError(command, "Invalid value for --view. Expected final, original, or markup.");
                        }

                        break;
                    case "--profile":
                        if (!TryReadValue(args, ref i, out string? profileValue))
                        {
                            return WithError(command, "Missing value for --profile.");
                        }

                        if (!TryParseValidationProfile(profileValue!, out validationProfile))
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

            return new ParsedOptions(command, positionals, flags, json, strict, verbose, diagnosticsPath, reportPath, operationReportPath, outputPath, id, extractPath, maxText, maxCommentText, maxDiagnostics, radius, trackChanges, author, timestampUtc, textView, validationProfile, null);
        }

        private static bool TryReadValue(string[] args, ref int index, out string? value)
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
            return new ParsedOptions(command, [], new HashSet<string>(StringComparer.Ordinal), false, false, false, null, null, null, null, null, null, null, null, null, null, TrackChangesMode.Off, null, null, DocxTextView.Final, DocxValidationProfile.Structural, message);
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
