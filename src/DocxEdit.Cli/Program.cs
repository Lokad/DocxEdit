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
            if (args.Length == 0 || IsHelp(args[0]))
            {
                WriteHelp();
                return 0;
            }

            if (args.Length == 2 && args[0] == "help" && args[1] == "patch")
            {
                WritePatchHelp();
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
            return InvalidUsage("Usage: docxedit read input.docx [--json] [--diagnostics <path>] [--strict]");
        }

        using Stream input = File.OpenRead(options.Positionals[0]);
        DocxReadResult result = new DocxEditor().Read(input, new DocxReadOptions
        {
            IncludeHeadersFooters = options.Flags.Contains("--headers-footers"),
            IncludeAllStories = options.Flags.Contains("--all-stories")
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
            return InvalidUsage("Usage: docxedit find input.docx \"text\" [--json] [--diagnostics <path>] [--strict]");
        }

        using Stream input = File.OpenRead(options.Positionals[0]);
        DocxFindResult result = new DocxEditor().Find(input, options.Positionals[1], new DocxFindOptions
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
            return InvalidUsage("Usage: docxedit dump input.docx --id M.P0001 [--runs] [--json] [--diagnostics <path>] [--strict]");
        }

        using Stream input = File.OpenRead(options.Positionals[0]);
        DocxDumpResult result = new DocxEditor().Dump(input, options.Id, new DocxDumpOptions
        {
            IncludeRuns = options.Flags.Contains("--runs")
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
            return InvalidUsage("Usage: docxedit media input.docx [--json] [--diagnostics <path>] [--strict]");
        }

        using Stream input = File.OpenRead(options.Positionals[0]);
        DocxMediaResult result = new DocxEditor().Media(input);
        WriteDiagnostics(options.DiagnosticsPath, result.Diagnostics);
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

    private static int RunCheck(ParsedOptions options)
    {
        if (options.Positionals.Count != 2)
        {
            return InvalidUsage("Usage: docxedit check input.docx edits.docxpatch [--json] [--report <path>] [--diagnostics <path>] [--strict]");
        }

        using Stream input = File.OpenRead(options.Positionals[0]);
        using TextReader patch = File.OpenText(options.Positionals[1]);
        DocxCheckResult result = new DocxEditor().Check(input, patch);
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
            return InvalidUsage("Usage: docxedit apply input.docx edits.docxpatch -o output.docx [--json] [--report <path>] [--diagnostics <path>] [--strict]");
        }

        using Stream input = File.OpenRead(options.Positionals[0]);
        using TextReader patch = File.OpenText(options.Positionals[1]);
        using Stream output = File.Create(options.OutputPath);
        DocxApplyResult result = new DocxEditor().Apply(input, patch, output);
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

            Patch:
              check      Validate a .docxpatch file without writing output
              apply      Apply a .docxpatch file and write a new .docx

            Help:
              help patch Show the .docxpatch syntax with examples
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
            with <<<
            new text
            >>>
            end

            The expect-hash feature is unsupported and is rejected.
            """);
    }

    private sealed record ParsedOptions(
        string Command,
        IReadOnlyList<string> Positionals,
        HashSet<string> Flags,
        bool Json,
        bool Strict,
        string? DiagnosticsPath,
        string? ReportPath,
        string? OutputPath,
        string? Id,
        string? Error)
    {
        public static ParsedOptions Parse(string[] args)
        {
            string command = args[0];
            var positionals = new List<string>();
            var flags = new HashSet<string>(StringComparer.Ordinal);
            bool json = false;
            bool strict = false;
            string? diagnosticsPath = null;
            string? reportPath = null;
            string? outputPath = null;
            string? id = null;

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

            return new ParsedOptions(command, positionals, flags, json, strict, diagnosticsPath, reportPath, outputPath, id, null);
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
            return new ParsedOptions(command, [], new HashSet<string>(StringComparer.Ordinal), false, false, null, null, null, null, message);
        }
    }
}
