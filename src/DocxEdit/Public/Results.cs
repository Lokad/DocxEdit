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
    public IReadOnlyList<DocxBookmarkInfo> Bookmarks { get; init; } = [];
    public IReadOnlyList<DocxContentControlInfo> ContentControls { get; init; } = [];
    public IReadOnlyList<DocxFieldInfo> Fields { get; init; } = [];
    public IReadOnlyList<DocxHyperlinkInfo> Hyperlinks { get; init; } = [];
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

public sealed record DocxValidateResult : DocxOperationResult
{
    public IReadOnlyList<string> PartNames { get; init; } = [];
    public string? MainDocumentPartName { get; init; }
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
    IReadOnlyList<DocxDiagnostic> Diagnostics)
{
    public IReadOnlyList<DocxPatchAffectedTarget> AffectedTargets { get; init; } = [];
}

public sealed record DocxPatchAffectedTarget(string Id, string Kind, string Action)
{
    public string? ParentId { get; init; }
    public int? RowIndex { get; init; }
    public int? ColumnIndex { get; init; }
    public int? RowCountBefore { get; init; }
    public int? RowCountAfter { get; init; }
    public int? ColumnCount { get; init; }
    public int? CellCount { get; init; }
}

public sealed record DocxStyleInfo(string StyleId, string Name, string Type, bool IsDefault)
{
    public string? BasedOnStyleId { get; init; }
    public string? NextStyleId { get; init; }
    public string? LinkedStyleId { get; init; }
    public string? NumberingId { get; init; }
    public int? NumberingLevel { get; init; }
}

public sealed record DocxImageInfo(string Id, string PartName, string? ContentType, long ByteLength)
{
    public string LayoutKind { get; init; } = "unknown";
    public string? RelationshipId { get; init; }
    public string? ContainingTargetId { get; init; }
    public long? WidthEmu { get; init; }
    public long? HeightEmu { get; init; }
    public string? Name { get; init; }
    public string? Description { get; init; }
    public string? Title { get; init; }
    public string? WrapMode { get; init; }
    public bool BehindDoc { get; init; }
}

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
    public int? CommentTextLength { get; init; }
    public string? CommentTextSnippet { get; init; }
    public bool CommentTextTruncated { get; init; }
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
    public string? Initials { get; init; }
    public int? TextLength { get; init; }
    public string? TextSnippet { get; init; }
    public bool TextTruncated { get; init; }
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
    public string? HyperlinkRelationshipId { get; init; }
    public string? HyperlinkAnchor { get; init; }
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
    public IReadOnlyList<string> BookmarkNames { get; init; } = [];
    public IReadOnlyList<string> ContentControlIds { get; init; } = [];
    public IReadOnlyList<string> ContentControlTags { get; init; } = [];
    public IReadOnlyList<string> ContentControlAliases { get; init; } = [];
    public IReadOnlyList<string> FieldIds { get; init; } = [];
    public IReadOnlyList<string> FieldCodes { get; init; } = [];
    public IReadOnlyList<string> FieldKinds { get; init; } = [];
    public IReadOnlyList<string> HyperlinkIds { get; init; } = [];
    public IReadOnlyList<string> HyperlinkTargets { get; init; } = [];
    public IReadOnlyList<string> CommentIds { get; init; } = [];
    public IReadOnlyList<string> CommentBodyIds { get; init; } = [];
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
    public string? LabelText { get; init; }
    public string LabelStatus { get; init; } = "not-resolved";
    public IReadOnlyList<string> LabelWarnings { get; init; } = [];
    public int? StartValue { get; init; }
    public string? Suffix { get; init; }
    public bool IsLegal { get; init; }
    public int? RestartAfterLevel { get; init; }
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
    public string? HyperlinkRelationshipId { get; init; }
    public string? HyperlinkAnchor { get; init; }
}

public sealed record DocxBookmarkInfo
{
    public string Id { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;
    public string? OoxmlId { get; init; }
    public string Story { get; init; } = string.Empty;
    public string PartName { get; init; } = string.Empty;
    public string? StartTargetId { get; init; }
    public string? EndTargetId { get; init; }
    public bool IsComplete { get; init; }
}

public sealed record DocxContentControlInfo
{
    public string Id { get; init; } = string.Empty;
    public string Story { get; init; } = string.Empty;
    public string PartName { get; init; } = string.Empty;
    public string? TargetId { get; init; }
    public string Kind { get; init; } = "unknown";
    public string? OoxmlId { get; init; }
    public string? Tag { get; init; }
    public string? Alias { get; init; }
    public string? Lock { get; init; }
    public int TextLength { get; init; }
}

public sealed record DocxFieldInfo
{
    public string Id { get; init; } = string.Empty;
    public string Story { get; init; } = string.Empty;
    public string PartName { get; init; } = string.Empty;
    public string? TargetId { get; init; }
    public string Kind { get; init; } = "unknown";
    public string Code { get; init; } = string.Empty;
    public int ResultTextLength { get; init; }
    public bool? IsDirty { get; init; }
    public bool? IsLocked { get; init; }
    public bool IsComplete { get; init; }
}

public sealed record DocxHyperlinkInfo
{
    public string Id { get; init; } = string.Empty;
    public string Story { get; init; } = string.Empty;
    public string PartName { get; init; } = string.Empty;
    public string? TargetId { get; init; }
    public string? RelationshipId { get; init; }
    public string? Uri { get; init; }
    public string? Anchor { get; init; }
    public string? Tooltip { get; init; }
    public string? TargetPartName { get; init; }
    public bool IsExternal { get; init; }
    public bool IsBroken { get; init; }
    public int DisplayTextLength { get; init; }
}

public sealed record DocxTableInfo(
    string Id,
    string Story,
    int RowCount,
    int ColumnCount,
    IReadOnlyList<DocxTableCellInfo> Cells)
{
    public string? StyleId { get; init; }
    public int? GridColumnCount { get; init; }
    public bool HasHeaderRow { get; init; }
    public bool HasMergedCells { get; init; }
    public bool HasNestedTables { get; init; }
    public IReadOnlyList<DocxTableRowInfo> Rows { get; init; } = [];
}

public sealed record DocxTableRowInfo
{
    public string Id { get; init; } = string.Empty;
    public int RowIndex { get; init; }
    public int CellCount { get; init; }
    public int GridBefore { get; init; }
    public int GridAfter { get; init; }
    public bool IsHeader { get; init; }
    public bool CantSplit { get; init; }
}

public sealed record DocxTableCellInfo(
    string Id,
    int RowIndex,
    int ColumnIndex,
    string Text,
    int ColumnSpan,
    string? VerticalMerge,
    bool HasNestedTable)
{
    public int PhysicalColumnIndex { get; init; }
}

public sealed record DocxSectionInfo(
    string Id,
    string Story,
    int Columns,
    string Orientation);
