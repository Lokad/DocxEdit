namespace DocxEdit;

public abstract record DocxOperationResult
{
    public required bool Success { get; init; }
    public required IReadOnlyList<DocxDiagnostic> Diagnostics { get; init; }
}

public sealed record DocxReadResult : DocxOperationResult
{
    public IReadOnlyList<string> PartNames { get; init; } = [];
    public string? MainDocumentPartName { get; init; }
}

public sealed record DocxOutlineResult : DocxOperationResult
{
    public IReadOnlyList<string> PartNames { get; init; } = [];
    public string? MainDocumentPartName { get; init; }
}

public sealed record DocxFindResult : DocxOperationResult
{
    public string Query { get; init; } = string.Empty;
    public IReadOnlyList<string> Matches { get; init; } = [];
}

public sealed record DocxDumpResult : DocxOperationResult
{
    public string TargetId { get; init; } = string.Empty;
    public string? Text { get; init; }
}

public sealed record DocxStylesResult : DocxOperationResult
{
    public IReadOnlyList<DocxStyleInfo> Styles { get; init; } = [];
}

public sealed record DocxMediaResult : DocxOperationResult
{
    public IReadOnlyList<DocxImageInfo> Images { get; init; } = [];
}

public sealed record DocxCheckResult : DocxOperationResult
{
    public IReadOnlyList<DocxPatchOperationReport> Operations { get; init; } = [];
}

public sealed record DocxApplyResult : DocxOperationResult
{
    public IReadOnlyList<DocxPatchOperationReport> Operations { get; init; } = [];
}

public sealed record DocxPatchOperationReport(
    int Index,
    string OperationName,
    string? Target,
    bool Success,
    IReadOnlyList<DocxDiagnostic> Diagnostics);

public sealed record DocxStyleInfo(string StyleId, string Name, string Type, bool IsDefault);

public sealed record DocxImageInfo(string Id, string PartName, string? ContentType, long ByteLength);

