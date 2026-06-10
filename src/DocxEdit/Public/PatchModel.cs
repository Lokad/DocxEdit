namespace DocxEdit;

public sealed record DocxPatch(
    bool Success,
    int MajorVersion,
    IReadOnlyList<DocxPatchOperation> Operations,
    IReadOnlyList<DocxDiagnostic> Diagnostics);

public sealed record DocxPatchOperation(
    int Index,
    string OperationName,
    IReadOnlyDictionary<string, string> Fields);

