# Hosting docxedit commands

`Lokad.DocxEdit` includes the same command parser, execution, text output, JSON
wire format, and exit-code policy used by the development CLI. Hosts can expose
these commands without launching a process or making machine files available:

```csharp
int exitCode = await DocxCommand.RunAsync(
    new[] { "read", "/documents/report.docx", "--summary", "--json" },
    host,
    new DocxCommandOptions
    {
        Quotas = new DocxPackageLimits(2_000, 64L * 1024 * 1024, 16L * 1024 * 1024)
    },
    cancellationToken);
```

Implement `IDocxCommandHost` for your filesystem or document store. Arguments
exclude the executable name. Path comparison and path joining use the host's
namespace, working directory, and case rules. All command file operations go
through this host, including reports, diagnostics, and extracted images.

`OpenReadAsync` and `OpenTextAsync` return owned handles; the command disposes
them after reading, including cancellation and failure. A `-` input denotes
standard input, so return a borrowed-stream wrapper when the underlying input
must remain open. The stdout/stderr text writers are borrowed and never closed.
Binary stdout uses `WriteStandardOutputAsync` and keeps status text on stderr.

Input is read asynchronously into bounded memory before synchronous document
processing. Cancellation is passed to host I/O and document operations, and
`OperationCanceledException` propagates rather than becoming a command error.
The command writes a complete edited document to staging memory before calling
`PublishFileAsync`. The host must preserve the old destination on failed or
canceled publication. Staged streams are borrowed until the callback completes;
consume or copy them before returning. Publication of a document and its reports
is not a transaction across multiple files. Binary stdout cannot be rolled back
after an output failure.

Exit codes match the CLI: 0 success, 1 failed document/patch operation, 2 invalid
usage, 3 warning-bearing success under `--strict`, and 4 command I/O/runtime
failure. Strict mode can return 3 after publishing a successful edit. Errors
writing the result to a host output writer propagate to the caller.

Image assets are supplied through the host's patch-relative asset provider;
the editor owns and disposes each returned asset stream. A host should expose
only authorized assets. The command never resolves host paths with the machine's
filesystem APIs.

For direct library integrations, `DocxEditor`, `DocxHelp.Catalog`,
`DocxTextRenderer`, and `DocxJson.CreateOptions(bool)` remain independently usable.
The JSON options include the CLI's enum wire-value converters and are newly
created for each caller.

Package quotas bound package loading and individual image assets; they are not
an aggregate memory or output budget. `MaxText` and preview lengths are per item.
Hosts should bound responses, concurrent work, and publication size separately.
