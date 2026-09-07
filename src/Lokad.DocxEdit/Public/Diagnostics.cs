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
public sealed record DocxDiagnostic(
    DocxSeverity Severity,
    string Code,
    string Message)
{
    /// <summary>Target ID when the finding is about one target: a wire ID when resolved, or raw patch-target text (selectors included) when echoing input; not guaranteed parseable.</summary>
    public string? TargetId { get; init; }

    /// <summary>Package part path when the finding is about one part.</summary>
    public string? PartName { get; init; }

    /// <summary>Story label when the finding is about one story.</summary>
    public string? Story { get; init; }

    /// <summary>Machine-readable feature tag for fallback classification.</summary>
    public string? Feature { get; init; }

    /// <summary>Machine-readable fallback tag describing what was done instead.</summary>
    public string? Fallback { get; init; }

    /// <summary>1-based patch operation index, when from a patch.</summary>
    public int? OperationIndex { get; init; }

    /// <summary>1-based patch source line, when from a patch.</summary>
    public int? Line { get; init; }

    /// <summary>1-based patch source column, when from a patch.</summary>
    public int? Column { get; init; }
}
