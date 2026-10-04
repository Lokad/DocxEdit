namespace Lokad.DocxEdit;

internal sealed partial class DocxCommandExecution
{
    private async Task<int> RunCreate(ParsedOptions options)
    {
        if (string.IsNullOrWhiteSpace(options.OutputPath))
        {
            return InvalidUsage(CommandUsageError("create"));
        }
        if (ValidateOutputCollisions(options) is { } collision)
        {
            return InvalidUsage(collision);
        }
        if (options.OutputPath == "-" && options.Json)
        {
            return InvalidUsage("Invalid --json with '--output -': JSON status cannot share standard output with document bytes. Use --report <path> to capture the creation report.");
        }

        using var staged = new MemoryStream();
        DocxCreateResult result = new DocxEditor().Create(staged, new DocxCreateOptions
        {
            PaperSize = options.PaperSize,
            Orientation = options.Orientation,
            Quotas = _options.Quotas
        }, _cancellationToken);
        await WriteReportAsync(options.ReportPath, result, JsonOptionsFor(options)).ConfigureAwait(false);
        if (result.Success)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            staged.Position = 0;
            if (options.OutputPath == "-")
            {
                await _host.WriteStandardOutputAsync(staged, _cancellationToken).ConfigureAwait(false);
            }
            else
            {
                await _host.PublishFileAsync(options.OutputPath, staged, _cancellationToken).ConfigureAwait(false);
            }
        }

        if (options.OutputPath == "-")
        {
            await WriteDiagnosticsAsync(options.DiagnosticsPath, result.Diagnostics, JsonOptionsFor(options)).ConfigureAwait(false);
            WriteErrorDiagnostics(result.Diagnostics, options.Strict);
            _stderr.Write(DocxTextRenderer.RenderCreate(result));
            return ExitCode(result.Success, result.Diagnostics, options.Strict);
        }

        return await FinishCommand(options, result, static r => r.Diagnostics, static r => r.Success,
            r => _stdout.Write(DocxTextRenderer.RenderCreate(r))).ConfigureAwait(false);
    }
}
