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

`create --output /documents/new.docx [--paper a4|letter] [--orientation portrait|landscape]`
uses the same host without opening an input. It honors host quotas and stages the
validated document before publication. `--output -` uses binary stdout; status stays
on stderr. File output replaces an existing destination on successful publication,
as with `apply`. Direct integrations can use `DocxEditor.Create` with `DocxCreateOptions`.

`OpenReadAsync` and `OpenTextAsync` return owned handles; the command disposes
them after reading, including cancellation and failure. A `-` input denotes
standard input, so return a borrowed-stream wrapper when the underlying input
must remain open. The stdout/stderr text writers are borrowed, asynchronously
flushed before return, and never closed.
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

`GetAssetProvider` returns an `IDocxAsyncAssetProvider` bound to the patch's host
location, or null when assets are unsupported. Its `OpenAsync` receives a decoded
asset reference and cancellation token, and returns a `DocxAsset` or null for a
missing asset. The runner asynchronously reads and disposes each returned content
stream. Each distinct reference is opened once per invocation and its bytes are
reused by the editor. Quoted and heredoc references use the normal patch parser.
Image magic bytes and structure are still validated by the existing engine.

Assets are prefetched before document editing; providers must permit read-only
opens even when a later document guard will reject the patch. Missing assets and
I/O errors are reported only if execution reaches the referencing operation.
Cancellation always propagates immediately. A host should expose only authorized
assets. The command never resolves host paths with the machine's filesystem APIs.
Separate check and apply invocations obtain separate snapshots; hosts that need
version consistency across requests must bind them to immutable document and
asset versions. The direct synchronous `IDocxAssetProvider` API remains available
for callers of `DocxEditor` working with already-acquired bytes.

For direct library integrations, `DocxEditor`, `DocxHelp.Catalog`,
`DocxTextRenderer`, and `DocxJson.CreateOptions(bool)` remain independently usable.
The JSON options include the CLI's enum wire-value converters and are newly
created for each caller.

Package quotas bound package loading and individual image assets. The engine
also checks the final uncompressed package total after editing; these checks are
not an aggregate memory or output budget. `MaxText` and preview lengths are per item.
Hosts should bound responses, concurrent work, and publication size separately.
