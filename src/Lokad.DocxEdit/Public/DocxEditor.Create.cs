using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

/// <summary>Options for creating an empty document with one section and one empty paragraph.</summary>
public sealed class DocxCreateOptions
{
    /// <summary>Paper preset; defaults to A4 independently of machine locale.</summary>
    public DocxPaperSize PaperSize { get; init; } = DocxPaperSize.A4;
    /// <summary>Page orientation; defaults to portrait.</summary>
    public DocxOrientation Orientation { get; init; } = DocxOrientation.Portrait;
    /// <summary>Limits generated ZIP entry count, total uncompressed bytes, and single-part bytes.</summary>
    public DocxPackageLimits Quotas { get; init; } = DocxPackageLimits.Default;
    /// <summary>Leave output open by default. When false, dispose it on success, failure, or cancellation after accepting a writable stream.</summary>
    public bool LeaveOutputOpen { get; init; } = true;
}

/// <summary>Creation status. A failed result writes no output; output I/O errors and cancellation propagate as exceptions.</summary>
public sealed record DocxCreateResult : DocxOperationResult
{
    /// <summary>Requested paper preset.</summary>
    public DocxPaperSize PaperSize { get; init; }
    /// <summary>Requested orientation.</summary>
    public DocxOrientation Orientation { get; init; }
}

public sealed partial class DocxEditor
{
    /// <summary>Creates an empty A4 portrait document with one-inch margins. Leaves output open.</summary>
    public DocxCreateResult Create(Stream output) => Create(output, new DocxCreateOptions(), CancellationToken.None);

    /// <summary>Creates an empty document with explicit options and no cancellation.</summary>
    public DocxCreateResult Create(Stream output, DocxCreateOptions options) => Create(output, options, CancellationToken.None);

    /// <summary>Creates an empty document with explicit options and cancellation. Includes Normal and Heading1–Heading9 styles.</summary>
    /// <remarks>Writes at the current output position without seeking or truncating. Invalid enum values throw. Stage output before publishing when I/O failure or cancellation must preserve a destination.</remarks>
    public DocxCreateResult Create(Stream output, DocxCreateOptions options, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(output);
        ArgumentNullException.ThrowIfNull(options);
        if (!output.CanWrite)
        {
            throw new ArgumentException("Output stream must be writable.", nameof(output));
        }

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(options.Quotas);
            _ = options.PaperSize.ToWireValue();
            _ = options.Orientation.ToWireValue();
            if (!TryDocumentOperation(
                () => OoxmlPackage.CreateBlank(options.PaperSize, options.Orientation, options.Quotas, cancellationToken),
                out OoxmlPackage? package, out IReadOnlyList<DocxDiagnostic> diagnostics))
            {
                return Result(false, diagnostics);
            }

            if (!TryDocumentOperation(
                () => DocxPackageValidator.Validate(package, DocxValidationProfile.Structural, cancellationToken),
                out IReadOnlyList<DocxDiagnostic>? validation, out diagnostics))
            {
                return Result(false, diagnostics);
            }

            if (validation.Any(d => d.Severity == DocxSeverity.Error))
            {
                return Result(false, validation);
            }

            cancellationToken.ThrowIfCancellationRequested();
            package.Save(output, cancellationToken);
            return Result(true, validation);
        }
        finally
        {
            if (!options.LeaveOutputOpen) output.Dispose();
        }

        DocxCreateResult Result(bool success, IReadOnlyList<DocxDiagnostic> diagnostics) => new()
        {
            Success = success,
            Diagnostics = diagnostics,
            PaperSize = options.PaperSize,
            Orientation = options.Orientation
        };
    }
}
