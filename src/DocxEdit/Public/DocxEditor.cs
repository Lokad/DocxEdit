using DocxEdit.Ooxml;
using DocxEdit.Model;
using DocxEdit.Rendering;

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
        DocxDocumentModel model = package is null ? DocxDocumentModel.Empty : DocxDocumentScanner.Scan(package, cancellationToken);
        return new DocxReadResult
        {
            Success = package is not null,
            Diagnostics = diagnostics,
            PartNames = package?.Parts.Keys.Order(StringComparer.Ordinal).ToArray() ?? [],
            MainDocumentPartName = package?.MainDocumentPartName,
            Text = TextRenderers.RenderRead(model),
            Paragraphs = model.Paragraphs,
            Tables = model.Tables,
            Images = model.Images
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
        DocxDocumentModel model = package is null ? DocxDocumentModel.Empty : DocxDocumentScanner.Scan(package, cancellationToken);
        return new DocxOutlineResult
        {
            Success = package is not null,
            Diagnostics = diagnostics,
            PartNames = package?.Parts.Keys.Order(StringComparer.Ordinal).ToArray() ?? [],
            MainDocumentPartName = package?.MainDocumentPartName,
            Lines = TextRenderers.RenderOutline(model)
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
        DocxDocumentModel model = package is null ? DocxDocumentModel.Empty : DocxDocumentScanner.Scan(package, cancellationToken);
        return new DocxFindResult
        {
            Success = package is not null,
            Diagnostics = diagnostics,
            Query = query,
            Matches = TextRenderers.Find(model, query)
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
        DocxDocumentModel model = package is null ? DocxDocumentModel.Empty : DocxDocumentScanner.Scan(package, cancellationToken);
        return new DocxDumpResult
        {
            Success = package is not null,
            Diagnostics = diagnostics,
            TargetId = targetId,
            Text = TextRenderers.Dump(model, targetId)
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
        IReadOnlyList<DocxStyleInfo> styles = package is null ? [] : DocxStyleScanner.Scan(package, cancellationToken);
        return new DocxStylesResult
        {
            Success = package is not null,
            Diagnostics = diagnostics,
            Styles = styles
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
        DocxDocumentModel model = package is null ? DocxDocumentModel.Empty : DocxDocumentScanner.Scan(package, cancellationToken);
        return new DocxMediaResult
        {
            Success = package is not null,
            Diagnostics = diagnostics,
            Images = model.Images
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

        PatchExecutionResult execution = DocxPatchEngine.Check(package, patch, cancellationToken);
        return new DocxCheckResult
        {
            Success = execution.Success,
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

        PatchExecutionResult execution = DocxPatchEngine.Apply(package, patch, cancellationToken);
        if (!execution.Success)
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
        catch (InvalidDataException ex)
        {
            diagnostics =
            [
                new DocxDiagnostic(DocxSeverity.Error, "E0001", ex.Message)
            ];
            return null;
        }
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

    private static OoxmlPackageOptions ToPackageOptions(DocxStylesOptions options)
    {
        return new OoxmlPackageOptions(options.LeaveInputOpen, options.MaxZipEntries, options.MaxUncompressedBytes, options.MaxSinglePartBytes);
    }

    private static OoxmlPackageOptions ToPackageOptions(DocxMediaOptions options)
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
