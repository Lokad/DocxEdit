using System.Text.Json;
using DocxEdit;

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
            if (args.Length == 2 && args[0] == "help" && args[1] == "patch")
            {
                WritePatchHelp();
                return 0;
            }

            if (args.Length == 0 || IsHelp(args[0]))
            {
                WriteHelp();
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
                "styles" => RunStyles(options),
                "media" => RunMedia(options),
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
            return InvalidUsage("Usage: docxedit read input.docx [--view final|original|markup] [--max-text <chars>] [--json] [--diagnostics <path>] [--strict]");
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
        else
        {
            Console.Write(result.Text);
        }

        return ExitCode(result.Success, result.Diagnostics, options.Strict);
    }

    private static int RunOutline(ParsedOptions options)
    {
        if (options.Positionals.Count != 1)
        {
            return InvalidUsage("Usage: docxedit outline input.docx [--json] [--diagnostics <path>] [--strict]");
        }

        using Stream input = File.OpenRead(options.Positionals[0]);
        DocxOutlineResult result = new DocxEditor().Outline(input, new DocxOutlineOptions
        {
            IncludeHeadersFooters = options.Flags.Contains("--headers-footers")
        });
        WriteDiagnostics(options.DiagnosticsPath, result.Diagnostics);
        if (options.Json)
        {
            WriteJson(result);
        }
        else
        {
            foreach (string line in result.Lines)
            {
                Console.WriteLine(line);
            }
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
            foreach (string match in result.Matches)
            {
                Console.WriteLine(match);
            }
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
        else if (result.Text is not null)
        {
            Console.WriteLine(result.Text);
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
            foreach (DocxStyleInfo style in result.Styles.OrderBy(style => style.Type, StringComparer.Ordinal).ThenBy(style => style.StyleId, StringComparer.Ordinal))
            {
                string defaultText = style.IsDefault ? " default=true" : string.Empty;
                Console.WriteLine($"{style.Type} styleId={style.StyleId} name=\"{EscapeText(style.Name)}\"{defaultText}");
            }
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
            foreach (DocxImageInfo image in result.Images)
            {
                Console.WriteLine($"{image.Id} {image.PartName} {image.ContentType ?? "unknown"} {image.ByteLength} bytes");
            }
        }

        return ExitCode(result.Success, result.Diagnostics, options.Strict);
    }

    private static int RunChanges(ParsedOptions options)
    {
        if (options.Positionals.Count != 1)
        {
            return InvalidUsage("Usage: docxedit changes input.docx [--json] [--diagnostics <path>] [--strict]");
        }

        using Stream input = File.OpenRead(options.Positionals[0]);
        DocxChangesResult result = new DocxEditor().Changes(input);
        WriteDiagnostics(options.DiagnosticsPath, result.Diagnostics);
        if (options.Json)
        {
            WriteJson(result);
        }
        else
        {
            foreach (DocxChangeSummary summary in result.Summary)
            {
                Console.WriteLine($"{summary.Type} count={summary.Count}");
            }

            foreach (DocxChangeInfo change in result.Changes)
            {
                string target = change.TargetId is null ? "target=unknown" : $"target={change.TargetId}";
                string revision = change.RevisionId is null ? string.Empty : $" revision-id={EscapeText(change.RevisionId)}";
                string author = change.Author is null ? string.Empty : $" author=\"{EscapeText(change.Author)}\"";
                string timestamp = change.TimestampUtc is null ? string.Empty : $" timestamp-utc={change.TimestampUtc:O}";
                string commentId = change.CommentId is null ? string.Empty : $" comment-id={EscapeText(change.CommentId)}";
                string commentAuthor = change.CommentAuthor is null ? string.Empty : $" comment-author=\"{EscapeText(change.CommentAuthor)}\"";
                string commentTimestamp = change.CommentTimestampUtc is null ? string.Empty : $" comment-timestamp-utc={change.CommentTimestampUtc:O}";
                string commentInitials = change.CommentInitials is null ? string.Empty : $" comment-initials=\"{EscapeText(change.CommentInitials)}\"";
                Console.WriteLine($"{change.Id} {change.Type} story=\"{EscapeText(change.Story)}\" part={change.PartName} {target} text-length={change.TextLength} children={change.ChildElementCount}{revision}{author}{timestamp}{commentId}{commentAuthor}{commentTimestamp}{commentInitials}");
            }
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
        }

        return ExitCode(result.Success, result.Diagnostics, options.Strict);
    }

    private static int RunApply(ParsedOptions options)
    {
        if (options.Positionals.Count != 2 || options.OutputPath is null)
        {
            return InvalidUsage("Usage: docxedit apply input.docx edits.docxpatch -o output.docx [--track-changes <mode>] [--author <name>] [--timestamp-utc <instant>] [--json] [--report <path>] [--diagnostics <path>] [--strict]");
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
            TimestampUtc = options.TimestampUtc ?? DateTimeOffset.UtcNow
        };
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

    private static string EscapeText(string text)
    {
        return text.Replace("\\", "\\\\", StringComparison.Ordinal).Replace("\"", "\\\"", StringComparison.Ordinal);
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

    private static void WriteHelp()
    {
        Console.WriteLine("""
            docxedit - read and patch .docx documents

            Usage:
              docxedit <command> [options]

            Read / explore:
              read       Produce an agent-friendly structural view of a .docx
              outline    Show headings, tables, images, sections, headers, footers
              find       Find text and print stable edit targets
              dump       Dump one target in detail
              styles     List paragraph, character, and table styles
              media      List embedded images
              changes    List tracked-change and comment markup without printing private text

            Patch:
              check      Validate a .docxpatch file without writing output
              apply      Apply a .docxpatch file and write a new .docx

            Help:
              help patch Show the .docxpatch syntax with examples

            Examples:
              docxedit read report.docx [--view final|original|markup]
              docxedit dump report.docx --id M.P0004 --runs
              docxedit media report.docx --extract media
              docxedit changes report.docx
              docxedit check report.docx edits.docxpatch
              docxedit apply report.docx edits.docxpatch -o report.edited.docx
            """);
    }

    private static void WritePatchHelp()
    {
        Console.WriteLine("""
            docxpatch 1

            op replace-text
            target M.P0004
            expect-text <<<
            current text
            >>>
            find <<<
            current
            >>>
            with <<<
            new text
            >>>
            end

            op set-cell
            target M.T0001.R02.C03
            text <<<
            updated cell text
            >>>
            end

            op replace-image
            target M.I0001
            asset chart.png
            alt Chart after update
            end

            op delete-row
            target M.T0001.R02
            expect-contains row text
            end

            Selectors may use explicit IDs, heading:"Text", heading:2:"Text", text:"contained text", bookmark:"Name", or content-control:"TagOrAlias" for paragraph targets.
            delete-row expect-contains checks that the target row's final visible text contains the supplied value.
            replace-image alt updates the image DrawingML description while replacing the media bytes.
            Unsupported fields are rejected. expect-hash, preserve-size, and caption are not supported.
            """);
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
        string? OutputPath,
        string? Id,
        string? ExtractPath,
        int? MaxText,
        TrackChangesMode TrackChanges,
        string? Author,
        DateTimeOffset? TimestampUtc,
        DocxTextView TextView,
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
            string? outputPath = null;
            string? id = null;
            string? extractPath = null;
            int? maxText = null;
            TrackChangesMode trackChanges = TrackChangesMode.Off;
            string? author = null;
            DateTimeOffset? timestampUtc = null;
            DocxTextView textView = DocxTextView.Final;

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

            return new ParsedOptions(command, positionals, flags, json, strict, verbose, diagnosticsPath, reportPath, outputPath, id, extractPath, maxText, trackChanges, author, timestampUtc, textView, null);
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
            return new ParsedOptions(command, [], new HashSet<string>(StringComparer.Ordinal), false, false, false, null, null, null, null, null, null, TrackChangesMode.Off, null, null, DocxTextView.Final, message);
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
