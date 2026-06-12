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
    public string Text { get; init; } = string.Empty;
    public IReadOnlyList<DocxParagraphInfo> Paragraphs { get; init; } = [];
    public IReadOnlyList<DocxTableInfo> Tables { get; init; } = [];
    public IReadOnlyList<DocxImageInfo> Images { get; init; } = [];
    public IReadOnlyList<DocxSectionInfo> Sections { get; init; } = [];
}

public sealed record DocxOutlineResult : DocxOperationResult
{
    public IReadOnlyList<string> PartNames { get; init; } = [];
    public string? MainDocumentPartName { get; init; }
    public IReadOnlyList<string> Lines { get; init; } = [];
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
    public IReadOnlyList<DocxDumpRunInfo> Runs { get; init; } = [];
}

public sealed record DocxContextResult : DocxOperationResult
{
    public string TargetId { get; init; } = string.Empty;
    public string Text { get; init; } = string.Empty;
    public IReadOnlyList<DocxContextItem> Items { get; init; } = [];
}

public sealed record DocxStylesResult : DocxOperationResult
{
    public IReadOnlyList<DocxStyleInfo> Styles { get; init; } = [];
}

public sealed record DocxMediaResult : DocxOperationResult
{
    public IReadOnlyList<DocxImageInfo> Images { get; init; } = [];
}

public sealed record DocxChangesResult : DocxOperationResult
{
    public IReadOnlyList<string> PartNames { get; init; } = [];
    public string? MainDocumentPartName { get; init; }
    public IReadOnlyList<DocxChangeInfo> Changes { get; init; } = [];
    public IReadOnlyList<DocxChangeSummary> Summary { get; init; } = [];
    public IReadOnlyList<DocxChangeGroupSummary> GroupSummary { get; init; } = [];
    public IReadOnlyList<DocxChangeTargetSummary> TargetSummary { get; init; } = [];
    public IReadOnlyList<DocxCommentThreadSummary> CommentSummary { get; init; } = [];
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

public sealed record DocxStyleInfo(string StyleId, string Name, string Type, bool IsDefault)
{
    public string? BasedOnStyleId { get; init; }
    public string? NextStyleId { get; init; }
    public string? LinkedStyleId { get; init; }
    public string? NumberingId { get; init; }
    public int? NumberingLevel { get; init; }
}

public sealed record DocxImageInfo(string Id, string PartName, string? ContentType, long ByteLength);

public sealed record DocxChangeInfo
{
    public string Id { get; init; } = string.Empty;
    public string Type { get; init; } = string.Empty;
    public string Story { get; init; } = string.Empty;
    public string PartName { get; init; } = string.Empty;
    public string? TargetId { get; init; }
    public string? Author { get; init; }
    public DateTimeOffset? TimestampUtc { get; init; }
    public string? RevisionId { get; init; }
    public int TextLength { get; init; }
    public int ChildElementCount { get; init; }
    public string? CommentId { get; init; }
    public string? CommentAuthor { get; init; }
    public DateTimeOffset? CommentTimestampUtc { get; init; }
    public string? CommentInitials { get; init; }
    public string? CommentAnchorTargetId { get; init; }
    public string? CommentReferenceTargetId { get; init; }
    public string? CommentAnchorStory { get; init; }
    public string? CommentAnchorPartName { get; init; }
    public string TargetStatus { get; init; } = "targetless";
    public string TargetSource { get; init; } = "none";
    public string? TargetReason { get; init; }
    public string? NearestTargetId { get; init; }
    public string? TargetNote { get; init; }
    public string? PairedChangeId { get; init; }
}

public sealed record DocxChangeSummary(string Type, int Count);

public sealed record DocxChangeGroupSummary(string Group, string Key, string Type, int Count);

public sealed record DocxChangeTargetSummary
{
    public string TargetId { get; init; } = string.Empty;
    public int Count { get; init; }
    public IReadOnlyList<DocxChangeSummary> Summary { get; init; } = [];
}

public sealed record DocxCommentThreadSummary
{
    public string CommentId { get; init; } = string.Empty;
    public string? AnchorTargetId { get; init; }
    public string? ReferenceTargetId { get; init; }
    public string? AnchorStory { get; init; }
    public string? AnchorPartName { get; init; }
    public string? Author { get; init; }
    public DateTimeOffset? TimestampUtc { get; init; }
    public int Count { get; init; }
    public IReadOnlyList<DocxChangeSummary> Summary { get; init; } = [];
}

public sealed record DocxDumpRunInfo
{
    public string Id { get; init; } = string.Empty;
    public string Text { get; init; } = string.Empty;
    public string? MarkupType { get; init; }
    public string? RevisionId { get; init; }
    public string? Author { get; init; }
    public DateTimeOffset? TimestampUtc { get; init; }
    public string? CommentId { get; init; }
}

public sealed record DocxContextItem
{
    public string Id { get; init; } = string.Empty;
    public string Kind { get; init; } = string.Empty;
    public string Relation { get; init; } = string.Empty;
    public string Story { get; init; } = string.Empty;
    public string? ParentId { get; init; }
    public string Text { get; init; } = string.Empty;
    public int? HeadingLevel { get; init; }
    public string? StyleId { get; init; }
    public string? StyleName { get; init; }
    public DocxListInfo? List { get; init; }
    public int? RowCount { get; init; }
    public int? ColumnCount { get; init; }
    public int? RowIndex { get; init; }
    public int? ColumnIndex { get; init; }
    public int? ColumnSpan { get; init; }
    public string? VerticalMerge { get; init; }
    public bool HasNestedTable { get; init; }
}

public sealed record DocxParagraphInfo(
    string Id,
    string Story,
    string Text,
    int? HeadingLevel,
    DocxListInfo? List,
    IReadOnlyList<DocxRunInfo> Runs)
{
    public string? StyleId { get; init; }
    public string? StyleName { get; init; }
}

public sealed record DocxListInfo(string NumberingId, int Level)
{
    public string? AbstractNumberingId { get; init; }
    public string? Format { get; init; }
    public string? LevelText { get; init; }
    public string? ParagraphStyleId { get; init; }
    public string Source { get; init; } = "direct";
}

public sealed record DocxRunInfo(string Text)
{
    public string? MarkupType { get; init; }
    public string? RevisionId { get; init; }
    public string? Author { get; init; }
    public DateTimeOffset? TimestampUtc { get; init; }
    public string? CommentId { get; init; }
}

public sealed record DocxTableInfo(
    string Id,
    string Story,
    int RowCount,
    int ColumnCount,
    IReadOnlyList<DocxTableCellInfo> Cells);

public sealed record DocxTableCellInfo(
    string Id,
    int RowIndex,
    int ColumnIndex,
    string Text,
    int ColumnSpan,
    string? VerticalMerge,
    bool HasNestedTable);

public sealed record DocxSectionInfo(
    string Id,
    string Story,
    int Columns,
    string Orientation);
