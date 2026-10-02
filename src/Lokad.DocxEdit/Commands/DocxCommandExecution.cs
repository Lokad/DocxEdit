using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

namespace Lokad.DocxEdit;

internal sealed partial class DocxCommandExecution(
    IDocxCommandHost host,
    DocxCommandOptions options,
    CancellationToken cancellationToken)
{
    private readonly IDocxCommandHost _host = host;
    private readonly DocxCommandOptions _options = options;
    private readonly CancellationToken _cancellationToken = cancellationToken;
    private readonly StringWriter _stdout = new();
    private readonly StringWriter _stderr = new();

    public async Task<int> ExecuteAsync(string[] args)
    {
        int exitCode = await RunAsync(args).ConfigureAwait(false);
        _cancellationToken.ThrowIfCancellationRequested();
        await _host.StandardOutput.WriteAsync(_stdout.ToString().AsMemory(), _cancellationToken).ConfigureAwait(false);
        await _host.StandardError.WriteAsync(_stderr.ToString().AsMemory(), _cancellationToken).ConfigureAwait(false);
        await _host.StandardOutput.FlushAsync(_cancellationToken).ConfigureAwait(false);
        await _host.StandardError.FlushAsync(_cancellationToken).ConfigureAwait(false);
        return exitCode;
    }

    private static readonly JsonSerializerOptions IndentedJsonOptions = DocxJson.CreateOptions(writeIndented: true);
    private static readonly JsonSerializerOptions CompactJsonOptions = DocxJson.CreateOptions(writeIndented: false);

    private JsonSerializerOptions JsonOptionsFor(ParsedOptions options)
    {
        return options.Flags.Contains("--compact") ? CompactJsonOptions : IndentedJsonOptions;
    }

    public async Task<int> RunAsync(string[] args)
    {
        try
        {
            _cancellationToken.ThrowIfCancellationRequested();
            if (args.Length == 2 && args[0] == "help")
            {
                if (DocxHelp.TryRenderTopic(args[1], out string helpText))
                {
                    _stdout.Write(helpText);
                    return 0;
                }

                return InvalidUsage($"Unknown help topic '{args[1]}'.");
            }

            if (args.Length == 0 || IsHelp(args[0]))
            {
                _stdout.Write(DocxHelp.RenderOverview());
                return 0;
            }

            ParsedOptions options = ParsedOptions.Parse(args);
            if (options.Error is not null)
            {
                _stderr.WriteLine(options.Error);
                return 2;
            }

            return options.Command switch
            {
                "read" => await RunRead(options).ConfigureAwait(false),
                "outline" => await RunOutline(options).ConfigureAwait(false),
                "find" => await RunFind(options).ConfigureAwait(false),
                "dump" => await RunDump(options).ConfigureAwait(false),
                "context" => await RunContext(options).ConfigureAwait(false),
                "capabilities" => await RunCapabilities(options).ConfigureAwait(false),
                "template" => await RunTemplate(options).ConfigureAwait(false),
                "lint" => await RunLint(options).ConfigureAwait(false),
                "styles" => await RunStyles(options).ConfigureAwait(false),
                "media" => await RunMedia(options).ConfigureAwait(false),
                "validate" => await RunValidate(options).ConfigureAwait(false),
                "changes" => await RunChanges(options).ConfigureAwait(false),
                "check" => await RunCheck(options).ConfigureAwait(false),
                "apply" => await RunApply(options).ConfigureAwait(false),
                "catalog" => RunCatalog(options),
                "version" => RunVersion(options),
                _ => InvalidUsage($"Unknown command '{options.Command}'.\nValid commands: {ValidCommands()}\nSee 'docxedit --help'.")
            };
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _stderr.WriteLine(ex.Message);
            return 4;
        }
    }

    private bool IsStandardStreamPath(string? path) => string.IsNullOrEmpty(path) || path == "-";

    private bool SameCliPath(string first, string second) => _host.PathsEqual(first, second);

    // Validates writer/reader path collisions before any writer is opened.
    // Returns an InvalidUsage message, or null when the patch file set is safe.
    // Writers are --output/--report/--diagnostics; readers are the input
    // document, the patch file, and (for changes) the operation report.
    private string? ValidatePatchCollisions(
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

    private string? ValidateInputCollisions(ParsedOptions options, string inputPath)
    {
        if (!IsStandardStreamPath(options.DiagnosticsPath) &&
            !IsStandardStreamPath(inputPath) &&
            SameCliPath(options.DiagnosticsPath!, inputPath))
        {
            return $"Invalid --diagnostics '{options.DiagnosticsPath}': it must differ from the input document '{inputPath}'.";
        }

        return null;
    }

    private async Task<int> RunInputCommand<T>(
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

        using Stream input = await OpenInputFileAsync(options.Positionals[0]).ConfigureAwait(false);
        T result = execute(input);
        return await FinishCommand(options, result, getDiagnostics, getSuccess, writeText).ConfigureAwait(false);
    }

    private async Task<int> RunPatchCommand<T>(
        ParsedOptions options,
        string usage,
        Func<Stream, DocxPatch, DocxEditOptions, T> execute,
        Func<T, IReadOnlyList<DocxDiagnostic>> getDiagnostics,
        Func<T, bool> getSuccess,
        Action<T> writeText,
        Func<T, Task> writeReport)
        where T : notnull
    {
        if (options.Positionals.Count != 2)
        {
            return InvalidUsage(usage);
        }

        using Stream input = await OpenInputFileAsync(options.Positionals[0]).ConfigureAwait(false);
        using TextReader patch = await OpenPatchFileAsync(options.Positionals[1]).ConfigureAwait(false);
        DocxPatch parsed = new DocxEditor().ParsePatch(patch, _cancellationToken);
        DocxEditOptions editOptions = await ToEditOptionsAsync(options, parsed).ConfigureAwait(false);
        T result = execute(input, parsed, editOptions);
        await writeReport(result).ConfigureAwait(false);
        return await FinishCommand(options, result, getDiagnostics, getSuccess, writeText).ConfigureAwait(false);
    }

    private async Task<int> FinishCommand<T>(
        ParsedOptions options,
        T result,
        Func<T, IReadOnlyList<DocxDiagnostic>> getDiagnostics,
        Func<T, bool> getSuccess,
        Action<T> writeText)
        where T : notnull
    {
        IReadOnlyList<DocxDiagnostic> diagnostics = getDiagnostics(result);
        await WriteDiagnosticsAsync(options.DiagnosticsPath, diagnostics, JsonOptionsFor(options)).ConfigureAwait(false);
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

    private void WriteErrorDiagnostics(IReadOnlyList<DocxDiagnostic> diagnostics, bool strict)
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

            if (diagnostic.OperationIndex is not null)
            {
                message.Append(" op=").Append(diagnostic.OperationIndex.Value);
            }

            if (diagnostic.Line is not null)
            {
                message.Append(" line=").Append(diagnostic.Line.Value);
                if (diagnostic.Column is not null)
                {
                    message.Append(" col=").Append(diagnostic.Column.Value);
                }
            }

            if (diagnostic.HelpTopic is not null)
            {
                message.Append(" help=").Append(diagnostic.HelpTopic);
            }
            if (diagnostic.MatchCount is not null)
            {
                message.Append(" matches=").Append(diagnostic.MatchCount.Value);
            }
            _stderr.WriteLine(message.ToString());
        }
    }

    private async Task<int> RunRead(ParsedOptions options)
    {
        if (ValidateInputCollisions(options, options.Positionals[0]) is { } readCollision)
        {
            return InvalidUsage(readCollision);
        }

        using Stream input = await OpenInputFileAsync(options.Positionals[0]).ConfigureAwait(false);
        DocxReadResult result = new DocxEditor().Read(input, new DocxReadOptions
        {
            Quotas = _options.Quotas,
            IncludeHeadersFooters = options.Flags.Contains("--headers-footers"),
            TextView = options.TextView,
            MaxText = options.MaxText ?? 4_000
        }, _cancellationToken);
        bool summary = options.Flags.Contains("--summary");
        await WriteDiagnosticsAsync(options.DiagnosticsPath, result.Diagnostics, JsonOptionsFor(options)).ConfigureAwait(false);
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
            _stdout.Write(DocxTextRenderer.RenderReadSummary(result));
        }
        else
        {
            _stdout.Write(DocxTextRenderer.RenderRead(result));
        }

        return ExitCode(result.Success, result.Diagnostics, options.Strict);
    }

    private object BuildReadSummaryJson(DocxReadResult result)
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

    private async Task<int> RunOutline(ParsedOptions options)
    {
        return await RunInputCommand(
            options,
            CommandUsageError("outline"),
            input => new DocxEditor().Outline(input, new DocxOutlineOptions
            {
                Quotas = _options.Quotas,
                IncludeHeadersFooters = options.Flags.Contains("--headers-footers"),
                TextView = options.TextView,
                MaxText = options.MaxText ?? 4_000
            }, _cancellationToken),
            static result => result.Diagnostics,
            static result => result.Success,
            result => _stdout.Write(DocxTextRenderer.RenderOutline(result))).ConfigureAwait(false);
    }

    private async Task<int> RunFind(ParsedOptions options)
    {
        // Find takes document plus query, unlike the single-input commands, so it opens
        // and finishes explicitly instead of routing through the one-positional helper. Positional
        // arity itself is validated from the command catalog during parsing.
        string query = options.Positionals[1];
        if (ValidateInputCollisions(options, options.Positionals[0]) is { } findCollision)
        {
            return InvalidUsage(findCollision);
        }

        using Stream input = await OpenInputFileAsync(options.Positionals[0]).ConfigureAwait(false);
        DocxFindResult result = new DocxEditor().Find(input, query, new DocxFindOptions
        {
            Quotas = _options.Quotas,
            IncludeHeadersFooters = options.Flags.Contains("--headers-footers"),
            TextView = options.TextView,
            MaxText = options.MaxText ?? 4_000
        }, _cancellationToken);
        return await FinishCommand(
            options,
            result,
            static findResult => findResult.Diagnostics,
            static findResult => findResult.Success,
            findResult => _stdout.Write(DocxTextRenderer.RenderFind(findResult))).ConfigureAwait(false);
    }

    private async Task<int> RunDump(ParsedOptions options)
    {
        if (options.Id is null)
        {
            return InvalidUsage(CommandUsageError("dump"));
        }

        string id = options.Id;
        return await RunInputCommand(
            options,
            CommandUsageError("dump"),
            input => new DocxEditor().Dump(input, id, new DocxDumpOptions
            {
                Quotas = _options.Quotas,
                IncludeRuns = options.Flags.Contains("--runs"),
                TextView = options.TextView,
                MaxText = options.MaxText ?? 4_000
            }, _cancellationToken),
            static result => result.Diagnostics,
            static result => result.Success,
            result => _stdout.Write(DocxTextRenderer.RenderDump(result))).ConfigureAwait(false);
    }

    private async Task<int> RunContext(ParsedOptions options)
    {
        if (options.Id is null)
        {
            return InvalidUsage(CommandUsageError("context"));
        }

        string id = options.Id;
        return await RunInputCommand(
            options,
            CommandUsageError("context"),
            input => new DocxEditor().Context(input, id, new DocxContextOptions
            {
                Quotas = _options.Quotas,
                IncludeHeadersFooters = options.Flags.Contains("--headers-footers"),
                TextView = options.TextView,
                Radius = options.Radius ?? 1,
                MaxText = options.MaxText ?? 0
            }, _cancellationToken),
            static result => result.Diagnostics,
            static result => result.Success,
            result => _stdout.Write(DocxTextRenderer.RenderContext(result))).ConfigureAwait(false);
    }

    private async Task<int> RunCapabilities(ParsedOptions options)
    {
        if (options.Id is null)
        {
            return InvalidUsage(CommandUsageError("capabilities"));
        }

        string id = options.Id;
        return await RunInputCommand(
            options,
            CommandUsageError("capabilities"),
            input => new DocxEditor().GetCapabilities(input, id, new DocxCapabilitiesOptions
            {
                Quotas = _options.Quotas,
                TrackChanges = options.TrackChanges
            }, _cancellationToken),
            static result => result.Diagnostics,
            static result => result.Success,
            result => _stdout.Write(DocxTextRenderer.RenderCapabilities(result))).ConfigureAwait(false);
    }

    private async Task<int> RunLint(ParsedOptions options)
    {
        if (options.Positionals.Count != 1)
        {
            return InvalidUsage(CommandUsageError("lint"));
        }

        string patchPath = options.Positionals[0];
        if (ValidateInputCollisions(options, patchPath) is { } lintCollision)
        {
            return InvalidUsage(lintCollision);
        }

        using TextReader patch = await OpenPatchFileAsync(patchPath).ConfigureAwait(false);
        DocxLintResult result = new DocxEditor().Lint(patch, _cancellationToken);
        return await FinishCommand(
            options,
            result,
            static lintResult => lintResult.Diagnostics,
            static lintResult => lintResult.Success,
            lintResult => _stdout.Write(DocxTextRenderer.RenderLint(lintResult))).ConfigureAwait(false);
    }

    private async Task<int> RunTemplate(ParsedOptions options)
    {
        if (options.Id is null)
        {
            return InvalidUsage(CommandUsageError("template"));
        }

        string id = options.Id;
        return await RunInputCommand(
            options,
            CommandUsageError("template"),
            input => new DocxEditor().GetTemplate(input, id, new DocxTemplateOptions
            {
                Quotas = _options.Quotas,
                TrackChanges = options.TrackChanges
            }, _cancellationToken),
            static result => result.Diagnostics,
            static result => result.Success,
            result => _stdout.Write(DocxTextRenderer.RenderTemplate(result))).ConfigureAwait(false);
    }

    private async Task<int> RunStyles(ParsedOptions options)
    {
        return await RunInputCommand(
            options,
            CommandUsageError("styles"),
            input => new DocxEditor().Styles(input, new DocxStylesOptions { Quotas = _options.Quotas }, _cancellationToken),
            static result => result.Diagnostics,
            static result => result.Success,
            result => _stdout.Write(DocxTextRenderer.RenderStyles(result))).ConfigureAwait(false);
    }

    private async Task<int> RunMedia(ParsedOptions options)
    {
        string inputPath = options.Positionals[0];
        if (ValidateInputCollisions(options, inputPath) is { } mediaCollision)
        {
            return InvalidUsage(mediaCollision);
        }

        using Stream input = await OpenInputFileAsync(inputPath).ConfigureAwait(false);
        if (options.ExtractPath is string extractDirectory)
        {
            return await RunExtractMedia(options, input, extractDirectory).ConfigureAwait(false);
        }

        DocxMediaResult result = new DocxEditor().Media(input, new DocxMediaOptions { Quotas = _options.Quotas, IncludeHeadersFooters = options.Flags.Contains("--headers-footers"), ImageId = options.Id }, _cancellationToken);
        return await FinishCommand(options, result, static r => r.Diagnostics, static r => r.Success, r => _stdout.Write(DocxTextRenderer.RenderMedia(r))).ConfigureAwait(false);
    }

    private async Task<int> RunExtractMedia(ParsedOptions options, Stream input, string directory)
    {
        DocxMediaExtractResult result = new DocxEditor().ExtractMedia(input, new DocxMediaOptions { Quotas = _options.Quotas, IncludeHeadersFooters = options.Flags.Contains("--headers-footers"), ImageId = options.Id }, _cancellationToken);
        if (result.Success)
        {
            foreach (DocxMediaFile file in result.Files)
            {
                using var bytes = new MemoryStream(file.Content, writable: false);
                await _host.PublishFileAsync(_host.CombinePath(directory, file.FileName), bytes, _cancellationToken).ConfigureAwait(false);
            }
        }

        await WriteDiagnosticsAsync(options.DiagnosticsPath, result.Diagnostics, JsonOptionsFor(options)).ConfigureAwait(false);
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
                        file.ContentType,
                        ByteLength = file.Content.Length
                    }).ToArray()
                },
                JsonOptionsFor(options));
        }
        else
        {
            WriteErrorDiagnostics(result.Diagnostics, options.Strict);
            _stdout.Write(DocxTextRenderer.RenderMediaExtract(result));
        }

        return ExitCode(result.Success, result.Diagnostics, options.Strict);
    }

    private async Task<int> RunValidate(ParsedOptions options)
    {
        return await RunInputCommand(
            options,
            CommandUsageError("validate"),
            input => new DocxEditor().Validate(input, new DocxValidateOptions
            {
                Quotas = _options.Quotas,
                Profile = options.ValidationProfile,
                MaxDiagnostics = options.MaxDiagnostics ?? 500
            }, _cancellationToken),
            static result => result.Diagnostics,
            static result => result.Success,
            result => _stdout.Write(DocxTextRenderer.RenderValidate(result))).ConfigureAwait(false);
    }

    private async Task<int> RunChanges(ParsedOptions options)
    {
        if (options.Positionals.Count == 1 &&
            ValidatePatchCollisions(options, options.Positionals[0], string.Empty, options.OperationReportPath) is { } changesCollision)
        {
            return InvalidUsage(changesCollision);
        }

        IReadOnlyList<DocxPatchOperationReport> reports = await ReadOperationReportsAsync(options.OperationReportPath).ConfigureAwait(false);
        return await RunInputCommand(
            options,
            CommandUsageError("changes"),
            input => new DocxEditor().Changes(input, new DocxChangesOptions
            {
                Quotas = _options.Quotas,
                IncludeCommentText = options.Flags.Contains("--include-comment-text"),
                MaxCommentText = options.MaxCommentText ?? 240,
                OperationReports = reports
            }, _cancellationToken),
            static result => result.Diagnostics,
            static result => result.Success,
            result => _stdout.Write(DocxTextRenderer.RenderChanges(result))).ConfigureAwait(false);
    }

    private async Task<int> RunCheck(ParsedOptions options)
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

        return await RunPatchCommand(
            options,
            CommandUsageError("check"),
            (input, patch, editOptions) => new DocxEditor().CheckParsed(input, patch, editOptions, _cancellationToken),
            static result => result.Diagnostics,
            static result => result.Success,
            result =>
            {
                _stdout.WriteLine($"{(result.Success ? "docxedit check: OK" : "docxedit check: FAILED")} (author={result.Author} timestamp={result.TimestampUtc:O} track-changes={result.TrackChanges.ToWireValue()})");
                _stdout.Write(DocxTextRenderer.RenderOperationSummary(result.Operations));
            },
            result => WriteReportAsync(options.ReportPath, result, JsonOptionsFor(options))).ConfigureAwait(false);
    }

    private async Task<int> RunApply(ParsedOptions options)
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

        if (outputPath == "-" && options.Json)
        {
            return InvalidUsage("Invalid --json with '--output -': JSON status cannot share standard output with document bytes. Use --report <path> to capture the operation report.");
        }

        using Stream input = await OpenInputFileAsync(inputPath).ConfigureAwait(false);
        using TextReader patch = await OpenPatchFileAsync(patchPath).ConfigureAwait(false);
        DocxPatch parsed = new DocxEditor().ParsePatch(patch, _cancellationToken);
        DocxEditOptions editOptions = await ToEditOptionsAsync(options, parsed).ConfigureAwait(false);
        using var staged = new MemoryStream();
        DocxApplyResult result = new DocxEditor().ApplyParsed(input, parsed, staged, editOptions, _cancellationToken);
        await WriteReportAsync(options.ReportPath, result, JsonOptionsFor(options)).ConfigureAwait(false);
        if (result.Success)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            staged.Position = 0;
            if (outputPath == "-")
            {
                await _host.WriteStandardOutputAsync(staged, _cancellationToken).ConfigureAwait(false);
            }
            else
            {
                await _host.PublishFileAsync(outputPath, staged, _cancellationToken).ConfigureAwait(false);
            }
        }

        if (outputPath == "-")
        {
            await WriteDiagnosticsAsync(options.DiagnosticsPath, result.Diagnostics, JsonOptionsFor(options)).ConfigureAwait(false);
            WriteErrorDiagnostics(result.Diagnostics, options.Strict);
            _stderr.WriteLine($"{(result.Success ? "docxedit apply: OK" : "docxedit apply: FAILED")} (author={result.Author} timestamp={result.TimestampUtc:O} track-changes={result.TrackChanges.ToWireValue()})");
            _stderr.Write(DocxTextRenderer.RenderOperationSummary(result.Operations));
            return ExitCode(result.Success, result.Diagnostics, options.Strict);
        }

        return await FinishCommand(
            options, result, static r => r.Diagnostics, static r => r.Success,
            r =>
            {
                _stdout.WriteLine($"{(r.Success ? "docxedit apply: OK" : "docxedit apply: FAILED")} (author={r.Author} timestamp={r.TimestampUtc:O} track-changes={r.TrackChanges.ToWireValue()})");
                _stdout.Write(DocxTextRenderer.RenderOperationSummary(r.Operations));
            }).ConfigureAwait(false);
    }

    private int RunCatalog(ParsedOptions options)
    {
        if (options.Json)
        {
            WriteJson(DocxHelp.Catalog, JsonOptionsFor(options));
        }
        else
        {
            _stdout.Write(DocxHelp.RenderOverview());
            _stdout.WriteLine();
            _stdout.Write(DocxHelp.RenderPatchTrackChangesSupportTable());
        }

        return 0;
    }

    private int RunVersion(ParsedOptions options)
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
            _stdout.WriteLine($"docxedit {version} ({framework}), {operations} patch operations");
        }

        return 0;
    }

    private async Task WriteDiagnosticsAsync(string? path, IReadOnlyList<DocxDiagnostic> diagnostics, JsonSerializerOptions jsonOptions)
    {
        if (path is null)
        {
            return;
        }

        await PublishTextAsync(path, JsonSerializer.Serialize(diagnostics, jsonOptions)).ConfigureAwait(false);
    }

    private async Task<IReadOnlyList<DocxPatchOperationReport>> ReadOperationReportsAsync(string? path)
    {
        if (path is null)
        {
            return [];
        }

        using Stream input = await OpenInputFileAsync(path).ConfigureAwait(false);
        DocxApplyResult? report = JsonSerializer.Deserialize<DocxApplyResult>(input, IndentedJsonOptions);
        if (report?.Operations is null)
        {
            throw new InvalidDataException($"Operation report '{path}' is not a docxedit check/apply report JSON object.");
        }

        return report.Operations;
    }

    private async Task WriteReportAsync(string? path, object report, JsonSerializerOptions jsonOptions)
    {
        if (path is null)
        {
            return;
        }

        await PublishTextAsync(path, JsonSerializer.Serialize(report, jsonOptions)).ConfigureAwait(false);
    }

    private void WriteJson(object value, JsonSerializerOptions jsonOptions)
    {
        _stdout.WriteLine(JsonSerializer.Serialize(value, jsonOptions));
    }

    private async Task<DocxEditOptions> ToEditOptionsAsync(ParsedOptions options, DocxPatch patch)
    {
        IDocxAssetProvider? assets = null;
        if (patch.Success && patch.Operations.Any(operation => operation.Fields.ContainsKey("asset")) &&
            _host.GetAssetProvider(options.Positionals[1]) is { } provider)
        {
            assets = await BufferedDocxAssets.LoadAsync(patch, provider, _options.Quotas.MaxSinglePartBytes, _cancellationToken).ConfigureAwait(false);
        }
        return new DocxEditOptions
        {
            TrackChanges = options.TrackChanges,
            Author = options.Author ?? "docxedit",
            TimestampUtc = options.TimestampUtc ?? DateTimeOffset.UtcNow,
            AssetProvider = assets,
            Quotas = _options.Quotas,
            MaxPreviewChars = options.MaxPreviewChars ?? 0
        };
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

    private int InvalidUsage(string message)
    {
        _stderr.WriteLine(message);
        return 2;
    }

    private static bool IsHelp(string arg)
    {
        return arg is "-h" or "--help" or "help";
    }

}
