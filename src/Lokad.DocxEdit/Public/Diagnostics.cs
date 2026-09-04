namespace Lokad.DocxEdit;

/// <summary>
/// Diagnostic severity. Errors fail the operation; warnings ride along successes
/// (and fail strict mode); info is informational.
/// </summary>
public enum DocxSeverity
{
    /// <summary>Informational diagnostic; never fails anything.</summary>
    Info,

    /// <summary>Warning: the effect happened, but review the message.</summary>
    Warning,

    /// <summary>Error: the requested effect did not happen.</summary>
    Error
}

/// <summary>
/// One machine-readable finding with a stable code, a human message, and optional
/// location and classification context.
/// </summary>
/// <remarks>
/// <para>
/// <c>Code</c> meanings are cataloged per code in <c>docs/diagnostics.md</c>;
/// the message carries the instance-specific details.
/// </para>
/// </remarks>
/// <param name="Severity">Whether this finding fails the operation.</param>
/// <param name="Code">Stable diagnostic code (for example <c>E1201</c>).</param>
/// <param name="Message">Human-readable detail for this instance.</param>
/// <param name="TargetId">Stable target ID when the finding is about one target.</param>
/// <param name="PartName">Package part path when the finding is about one part.</param>
/// <param name="Story">Story label when the finding is about one story.</param>
/// <param name="Feature">Machine-readable feature tag for fallback classification.</param>
/// <param name="Fallback">Machine-readable fallback tag describing what was done instead.</param>
/// <param name="OperationIndex">Zero-based patch operation index, when from a patch.</param>
/// <param name="Line">1-based patch source line, when from a patch.</param>
/// <param name="Column">1-based patch source column, when from a patch.</param>
public sealed record DocxDiagnostic(
    DocxSeverity Severity,
    string Code,
    string Message,
    string? TargetId = null,
    string? PartName = null,
    string? Story = null,
    string? Feature = null,
    string? Fallback = null,
    int? OperationIndex = null,
    int? Line = null,
    int? Column = null);
