using DocxEdit.Ooxml;
using DocxEdit.Model;
using DocxEdit.Rendering;
using System.Xml;

namespace DocxEdit;

public sealed class DocxEditor
{
    public DocxReadResult Read(
        Stream input,
        DocxReadOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(input);
        options ??= new DocxReadOptions();
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxReadResult
            {
                Success = false,
                Diagnostics = diagnostics
            };
        }

        if (!TryDocumentOperation(() => DocxDocumentScanner.Scan(package, options.IncludeHeadersFooters || options.IncludeAllStories, options.TextView, cancellationToken), out DocxDocumentModel? model, out IReadOnlyList<DocxDiagnostic> scanDiagnostics))
        {
            return new DocxReadResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(scanDiagnostics).ToArray(),
                PartNames = package.Parts.Keys.Order(StringComparer.Ordinal).ToArray(),
                MainDocumentPartName = package.MainDocumentPartName
            };
        }

        return new DocxReadResult
        {
            Success = true,
            Diagnostics = diagnostics
                .Concat(DocxUnsupportedFeatureScanner.Scan(package, options.IncludeHeadersFooters || options.IncludeAllStories, cancellationToken))
                .ToArray(),
            PartNames = package.Parts.Keys.Order(StringComparer.Ordinal).ToArray(),
            MainDocumentPartName = package.MainDocumentPartName,
            Text = TextRenderers.RenderRead(model!, options.MaxText),
            Paragraphs = model!.Paragraphs,
            Tables = model.Tables,
            Images = model.Images,
            Sections = model.Sections,
            Bookmarks = model.Bookmarks,
            ContentControls = model.ContentControls,
            Fields = model.Fields,
            Hyperlinks = model.Hyperlinks
        };
    }

    public DocxOutlineResult Outline(
        Stream input,
        DocxOutlineOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(input);
        options ??= new DocxOutlineOptions();
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxOutlineResult { Success = false, Diagnostics = diagnostics };
        }

        if (!TryDocumentOperation(() => DocxDocumentScanner.Scan(package, options.IncludeHeadersFooters, cancellationToken: cancellationToken), out DocxDocumentModel? model, out IReadOnlyList<DocxDiagnostic> scanDiagnostics))
        {
            return new DocxOutlineResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(scanDiagnostics).ToArray(),
                PartNames = package.Parts.Keys.Order(StringComparer.Ordinal).ToArray(),
                MainDocumentPartName = package.MainDocumentPartName
            };
        }

        return new DocxOutlineResult
        {
            Success = true,
            Diagnostics = diagnostics
                .Concat(DocxUnsupportedFeatureScanner.Scan(package, options.IncludeHeadersFooters, cancellationToken))
                .ToArray(),
            PartNames = package.Parts.Keys.Order(StringComparer.Ordinal).ToArray(),
            MainDocumentPartName = package.MainDocumentPartName,
            Lines = TextRenderers.RenderOutline(model!)
        };
    }

    public DocxFindResult Find(
        Stream input,
        string query,
        DocxFindOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentException.ThrowIfNullOrWhiteSpace(query);
        options ??= new DocxFindOptions();
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxFindResult { Success = false, Diagnostics = diagnostics, Query = query };
        }

        if (!TryDocumentOperation(() => DocxDocumentScanner.Scan(package, options.IncludeHeadersFooters, options.TextView, cancellationToken), out DocxDocumentModel? model, out IReadOnlyList<DocxDiagnostic> scanDiagnostics))
        {
            return new DocxFindResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(scanDiagnostics).ToArray(),
                Query = query
            };
        }

        return new DocxFindResult
        {
            Success = true,
            Diagnostics = diagnostics
                .Concat(DocxUnsupportedFeatureScanner.Scan(package, options.IncludeHeadersFooters, cancellationToken))
                .ToArray(),
            Query = query,
            Matches = TextRenderers.Find(model!, query, options.MaxText)
        };
    }

    public DocxDumpResult Dump(
        Stream input,
        string targetId,
        DocxDumpOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentException.ThrowIfNullOrWhiteSpace(targetId);
        options ??= new DocxDumpOptions();
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxDumpResult { Success = false, Diagnostics = diagnostics, TargetId = targetId };
        }

        if (!TryDocumentOperation(() => DocxDocumentScanner.Scan(package, textView: options.TextView, cancellationToken: cancellationToken), out DocxDocumentModel? model, out IReadOnlyList<DocxDiagnostic> scanDiagnostics))
        {
            return new DocxDumpResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(scanDiagnostics).ToArray(),
                TargetId = targetId
            };
        }

        return new DocxDumpResult
        {
            Success = true,
            Diagnostics = diagnostics
                .Concat(DocxUnsupportedFeatureScanner.Scan(package, includeHeadersFooters: false, cancellationToken))
                .ToArray(),
            TargetId = targetId,
            Text = TextRenderers.Dump(model!, targetId, options.IncludeRuns, options.MaxText),
            Runs = options.IncludeRuns ? TextRenderers.DumpRuns(model!, targetId, options.MaxText) : []
        };
    }

    public DocxContextResult Context(
        Stream input,
        string targetId,
        DocxContextOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentException.ThrowIfNullOrWhiteSpace(targetId);
        options ??= new DocxContextOptions();
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxContextResult { Success = false, Diagnostics = diagnostics, TargetId = targetId };
        }

        if (!TryDocumentOperation(() => DocxDocumentScanner.Scan(package, options.IncludeHeadersFooters, options.TextView, cancellationToken), out DocxDocumentModel? model, out IReadOnlyList<DocxDiagnostic> scanDiagnostics))
        {
            return new DocxContextResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(scanDiagnostics).ToArray(),
                TargetId = targetId
            };
        }

        IReadOnlyList<DocxContextItem> items = TextRenderers.Context(model!, targetId, options.Radius, options.MaxText);
        IReadOnlyList<DocxDiagnostic> contextDiagnostics = items.Count == 0
            ? [new DocxDiagnostic(DocxSeverity.Error, "E2001", $"Target '{targetId}' was not found.", TargetId: targetId)]
            : [];
        return new DocxContextResult
        {
            Success = items.Count > 0,
            Diagnostics = diagnostics
                .Concat(DocxUnsupportedFeatureScanner.Scan(package, options.IncludeHeadersFooters, cancellationToken))
                .Concat(contextDiagnostics)
                .ToArray(),
            TargetId = targetId,
            Items = items,
            Text = TextRenderers.RenderContext(items)
        };
    }

    public DocxStylesResult Styles(
        Stream input,
        DocxStylesOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(input);
        options ??= new DocxStylesOptions();
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxStylesResult { Success = false, Diagnostics = diagnostics };
        }

        if (!TryDocumentOperation(() => DocxStyleScanner.Scan(package, cancellationToken), out IReadOnlyList<DocxStyleInfo>? styles, out IReadOnlyList<DocxDiagnostic> scanDiagnostics))
        {
            return new DocxStylesResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(scanDiagnostics).ToArray()
            };
        }

        return new DocxStylesResult
        {
            Success = true,
            Diagnostics = diagnostics,
            Styles = styles!
        };
    }

    public DocxMediaResult Media(
        Stream input,
        DocxMediaOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(input);
        options ??= new DocxMediaOptions();
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxMediaResult { Success = false, Diagnostics = diagnostics };
        }

        if (!TryDocumentOperation(() => DocxDocumentScanner.Scan(package, cancellationToken: cancellationToken), out DocxDocumentModel? model, out IReadOnlyList<DocxDiagnostic> scanDiagnostics))
        {
            return new DocxMediaResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(scanDiagnostics).ToArray()
            };
        }

        return new DocxMediaResult
        {
            Success = true,
            Diagnostics = diagnostics
                .Concat(DocxUnsupportedFeatureScanner.Scan(package, includeHeadersFooters: false, cancellationToken))
                .ToArray(),
            Images = model!.Images
        };
    }

    public DocxChangesResult Changes(
        Stream input,
        DocxChangesOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(input);
        options ??= new DocxChangesOptions();
        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxChangesResult { Success = false, Diagnostics = diagnostics };
        }

        if (!TryDocumentOperation(() => DocxChangeScanner.Scan(package, options.IncludeCommentText, options.MaxCommentText, cancellationToken), out IReadOnlyList<DocxChangeInfo>? changes, out IReadOnlyList<DocxDiagnostic> scanDiagnostics))
        {
            return new DocxChangesResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(scanDiagnostics).ToArray(),
                PartNames = package.Parts.Keys.Order(StringComparer.Ordinal).ToArray(),
                MainDocumentPartName = package.MainDocumentPartName
            };
        }

        return new DocxChangesResult
        {
            Success = true,
            Diagnostics = diagnostics,
            PartNames = package.Parts.Keys.Order(StringComparer.Ordinal).ToArray(),
            MainDocumentPartName = package.MainDocumentPartName,
            Changes = changes!,
            Summary = DocxChangeScanner.Summarize(changes!),
            GroupSummary = DocxChangeScanner.SummarizeGroups(changes!),
            TargetSummary = DocxChangeScanner.SummarizeTargets(changes!),
            CommentSummary = DocxChangeScanner.SummarizeComments(changes!)
        };
    }

    public DocxPatch ParsePatch(
        TextReader patchReader,
        DocxPatchParseOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(patchReader);
        _ = options;
        cancellationToken.ThrowIfCancellationRequested();
        string text = patchReader.ReadToEnd();
        cancellationToken.ThrowIfCancellationRequested();
        return DocxPatchParser.Parse(text);
    }

    public DocxCheckResult Check(
        Stream input,
        TextReader patchReader,
        DocxEditOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentNullException.ThrowIfNull(patchReader);
        options ??= new DocxEditOptions();

        DocxPatch patch = ParsePatch(patchReader, null, cancellationToken);
        if (!patch.Success)
        {
            return new DocxCheckResult { Success = false, Diagnostics = patch.Diagnostics };
        }

        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxCheckResult
            {
                Success = false,
                Diagnostics = diagnostics
            };
        }

        if (!TryDocumentOperation(() => DocxPatchEngine.Check(package, patch, options, cancellationToken), out PatchExecutionResult? execution, out IReadOnlyList<DocxDiagnostic> executionDiagnostics))
        {
            return new DocxCheckResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(executionDiagnostics).ToArray()
            };
        }

        return new DocxCheckResult
        {
            Success = execution!.Success,
            Diagnostics = diagnostics.Concat(execution.Diagnostics).ToArray(),
            Operations = execution.Reports
        };
    }

    public DocxApplyResult Apply(
        Stream input,
        TextReader patchReader,
        Stream output,
        DocxEditOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentNullException.ThrowIfNull(patchReader);
        ArgumentNullException.ThrowIfNull(output);
        if (!output.CanWrite)
        {
            throw new ArgumentException("Output stream must be writable.", nameof(output));
        }

        options ??= new DocxEditOptions();
        DocxPatch patch = ParsePatch(patchReader, null, cancellationToken);
        if (!patch.Success)
        {
            return new DocxApplyResult { Success = false, Diagnostics = patch.Diagnostics };
        }

        OoxmlPackage? package = TryLoad(input, ToPackageOptions(options), cancellationToken, out IReadOnlyList<DocxDiagnostic> diagnostics);
        if (package is null)
        {
            return new DocxApplyResult { Success = false, Diagnostics = diagnostics };
        }

        if (!TryDocumentOperation(() => DocxPatchEngine.Apply(package, patch, options, cancellationToken), out PatchExecutionResult? execution, out IReadOnlyList<DocxDiagnostic> executionDiagnostics))
        {
            return new DocxApplyResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(executionDiagnostics).ToArray()
            };
        }

        if (!execution!.Success)
        {
            return new DocxApplyResult
            {
                Success = false,
                Diagnostics = diagnostics.Concat(execution.Diagnostics).ToArray(),
                Operations = execution.Reports
            };
        }

        package.Save(output, cancellationToken);
        if (!options.LeaveOutputOpen)
        {
            output.Dispose();
        }

        return new DocxApplyResult
        {
            Success = diagnostics.All(d => d.Severity != DocxSeverity.Error),
            Diagnostics = diagnostics.Concat(execution.Diagnostics).ToArray(),
            Operations = execution.Reports
        };
    }

    private static OoxmlPackage? TryLoad(
        Stream input,
        OoxmlPackageOptions options,
        CancellationToken cancellationToken,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
    {
        if (!input.CanRead)
        {
            throw new ArgumentException("Input stream must be readable.", nameof(input));
        }

        try
        {
            OoxmlPackage package = OoxmlPackage.Load(input, options, cancellationToken);
            diagnostics = [];
            return package;
        }
        catch (Exception ex) when (IsExpectedDocumentException(ex))
        {
            diagnostics =
            [
                ToDocumentDiagnostic(ex)
            ];
            return null;
        }
    }

    private static bool TryDocumentOperation<T>(
        Func<T> operation,
        out T? result,
        out IReadOnlyList<DocxDiagnostic> diagnostics)
        where T : class
    {
        try
        {
            result = operation();
            diagnostics = [];
            return true;
        }
        catch (Exception ex) when (IsExpectedDocumentException(ex))
        {
            result = null;
            diagnostics = [ToDocumentDiagnostic(ex)];
            return false;
        }
    }

    private static bool IsExpectedDocumentException(Exception exception)
    {
        return exception is InvalidDataException or XmlException or IOException;
    }

    private static DocxDiagnostic ToDocumentDiagnostic(Exception exception)
    {
        if (exception is XmlException xmlException)
        {
            return new DocxDiagnostic(
                DocxSeverity.Error,
                "E0001",
                xmlException.Message,
                Line: xmlException.LineNumber > 0 ? xmlException.LineNumber : null,
                Column: xmlException.LinePosition > 0 ? xmlException.LinePosition : null);
        }

        return new DocxDiagnostic(DocxSeverity.Error, "E0001", exception.Message);
    }

    private static OoxmlPackageOptions ToPackageOptions(DocxReadOptions options)
    {
        return new OoxmlPackageOptions(options.LeaveInputOpen, options.MaxZipEntries, options.MaxUncompressedBytes, options.MaxSinglePartBytes);
    }

    private static OoxmlPackageOptions ToPackageOptions(DocxOutlineOptions options)
    {
        return new OoxmlPackageOptions(options.LeaveInputOpen, options.MaxZipEntries, options.MaxUncompressedBytes, options.MaxSinglePartBytes);
    }

    private static OoxmlPackageOptions ToPackageOptions(DocxFindOptions options)
    {
        return new OoxmlPackageOptions(options.LeaveInputOpen, options.MaxZipEntries, options.MaxUncompressedBytes, options.MaxSinglePartBytes);
    }

    private static OoxmlPackageOptions ToPackageOptions(DocxDumpOptions options)
    {
        return new OoxmlPackageOptions(options.LeaveInputOpen, options.MaxZipEntries, options.MaxUncompressedBytes, options.MaxSinglePartBytes);
    }

    private static OoxmlPackageOptions ToPackageOptions(DocxContextOptions options)
    {
        return new OoxmlPackageOptions(options.LeaveInputOpen, options.MaxZipEntries, options.MaxUncompressedBytes, options.MaxSinglePartBytes);
    }

    private static OoxmlPackageOptions ToPackageOptions(DocxStylesOptions options)
    {
        return new OoxmlPackageOptions(options.LeaveInputOpen, options.MaxZipEntries, options.MaxUncompressedBytes, options.MaxSinglePartBytes);
    }

    private static OoxmlPackageOptions ToPackageOptions(DocxMediaOptions options)
    {
        return new OoxmlPackageOptions(options.LeaveInputOpen, options.MaxZipEntries, options.MaxUncompressedBytes, options.MaxSinglePartBytes);
    }

    private static OoxmlPackageOptions ToPackageOptions(DocxChangesOptions options)
    {
        return new OoxmlPackageOptions(options.LeaveInputOpen, options.MaxZipEntries, options.MaxUncompressedBytes, options.MaxSinglePartBytes);
    }

    private static OoxmlPackageOptions ToPackageOptions(DocxEditOptions options)
    {
        return new OoxmlPackageOptions(
            options.LeaveInputOpen,
            options.MaxZipEntries,
            options.MaxUncompressedBytes,
            options.MaxSinglePartBytes,
            options.AllowMacroEnabledDocuments);
    }
}
