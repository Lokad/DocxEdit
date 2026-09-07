namespace Lokad.DocxEdit;

/// <summary>
/// Parsed <c>docxpatch</c> document: the operations to check or apply plus any syntax diagnostics.
/// </summary>
/// <param name="Success">Whether the patch text parsed without errors.</param>
/// <param name="MajorVersion">Patch format major version from the header line.</param>
/// <param name="Operations">Operations in file order. Empty when parsing failed.</param>
/// <param name="Diagnostics">Syntax diagnostics. Empty when parsing succeeded.</param>
public sealed record DocxPatch(
    bool Success,
    int MajorVersion,
    IReadOnlyList<DocxPatchOperation> Operations,
    IReadOnlyList<DocxDiagnostic> Diagnostics);

/// <summary>
/// One parsed patch operation: its index, name, and field map.
/// </summary>
/// <param name="Index">1-based operation index: the first operation in the patch is 1, matching the parser; echoed in diagnostics and reports.</param>
/// <param name="OperationName">Operation name as written (for example <c>replace-text</c>).</param>
/// <param name="Fields">Field map (last value wins on duplicates).</param>
public sealed record DocxPatchOperation(
    int Index,
    string OperationName,
    IReadOnlyDictionary<string, string> Fields)
{
    /// <summary>Field values with source positions, in file order: one entry per occurrence, including repeats of repeatable fields (see catalog repeatable fields).</summary>
    public IReadOnlyList<DocxPatchField> FieldValues { get; init; } = [];
}

/// <summary>
/// One raw patch field with its source position.
/// </summary>
/// <param name="Name">Field name as written.</param>
/// <param name="Value">Field value as written (heredocs already unwrapped).</param>
/// <param name="Line">1-based source line where the field starts (for heredoc fields, the key line rather than the closing marker).</param>
/// <param name="Column">1-based source column of the field name.</param>
public sealed record DocxPatchField(string Name, string Value, int Line, int Column);
