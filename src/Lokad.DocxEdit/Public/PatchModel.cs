namespace Lokad.DocxEdit;

public sealed record DocxPatch(
    bool Success,
    int MajorVersion,
    IReadOnlyList<DocxPatchOperation> Operations,
    IReadOnlyList<DocxDiagnostic> Diagnostics);

public sealed record DocxPatchOperation(
    int Index,
    string OperationName,
    IReadOnlyDictionary<string, string> Fields)
{
    public IReadOnlyList<DocxPatchField> FieldValues { get; init; } = [];
}

public sealed record DocxPatchField(string Name, string Value, int Line, int Column);
