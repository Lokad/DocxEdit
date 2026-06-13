namespace Lokad.DocxEdit;

public enum DocxSeverity
{
    Info,
    Warning,
    Error
}

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

