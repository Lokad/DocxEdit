namespace Lokad.DocxEdit;

/// <summary>
/// Base contract for every <see cref="DocxEditor"/> operation result.
/// </summary>
/// <remarks>
/// <para>
/// <see cref="Success"/> reports whether the requested effect happened, not
/// whether <see cref="Diagnostics"/> is empty. Warnings (for example
/// unsupported-structure reports) are returned alongside <see cref="Success"/>
/// set to <c>true</c>; consumers must inspect <see cref="Diagnostics"/> even on
/// success. The CLI maps warning-bearing successes to its strict-mode exit code
/// instead of failing the operation.
/// </para>
/// <para>
/// <see cref="Success"/> set to <c>false</c> means the requested effect did not
/// happen (<c>Apply</c> writes no output in that case). <see cref="Diagnostics"/>
/// always explains why and contains at least one <see cref="DocxSeverity.Error"/>
/// entry on current failure paths; keep it that way when adding new ones.
/// </para>
/// <para>
/// <see cref="Diagnostics"/> is ordered causally: package-load diagnostics first,
/// then operation/scan diagnostics, then unsupported-feature reports. On failure the
/// remaining payload properties keep their defaults, and <c>PartNames</c> /
/// <c>MainDocumentPartName</c> are only populated when the package itself loaded —
/// do not rely on them when <see cref="Success"/> is <c>false</c>.
/// </para>
/// </remarks>
public abstract record DocxOperationResult
{
/// <summary>Whether the requested effect happened. See remarks.</summary>
    public required bool Success { get; init; }
/// <summary>All diagnostics produced along the way, in causal order. May contain warnings even when <see cref="Success"/> is <c>true</c>.</summary>
    public required IReadOnlyList<DocxDiagnostic> Diagnostics { get; init; }
}

/// <summary>Structural view of the document. <see cref="DocxOperationResult.Success"/> means the package loaded and scanned; payload properties are populated.</summary>
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

/// <summary>Heading/paragraph outline lines. <see cref="DocxOperationResult.Success"/> means the package loaded and scanned.</summary>
public sealed record DocxOutlineResult : DocxOperationResult
{
    public IReadOnlyList<string> PartNames { get; init; } = [];
    public string? MainDocumentPartName { get; init; }
    public IReadOnlyList<string> Lines { get; init; } = [];
}

/// <summary>Text search matches. An empty <c>Matches</c> list is a successful search with no hits, not a failure.</summary>
public sealed record DocxFindResult : DocxOperationResult
{
    public string Query { get; init; } = string.Empty;
    public IReadOnlyList<string> Matches { get; init; } = [];
}

/// <summary>Target-scoped dump. <see cref="DocxOperationResult.Success"/> reports only that the document scanned: an unknown target yields success with null <c>Text</c> and empty <c>Runs</c>. Check the payload, not just success.</summary>
public sealed record DocxDumpResult : DocxOperationResult
{
    public string TargetId { get; init; } = string.Empty;
    public string? Text { get; init; }
    public IReadOnlyList<DocxDumpRunInfo> Runs { get; init; } = [];
}

/// <summary>Target neighborhood. Unlike dump, an unknown target fails with <c>E2001</c>; <see cref="DocxOperationResult.Success"/> is equivalent to "target found".</summary>
public sealed record DocxContextResult : DocxOperationResult
{
    public string TargetId { get; init; } = string.Empty;
    public string Text { get; init; } = string.Empty;
    public IReadOnlyList<DocxContextItem> Items { get; init; } = [];
}

/// <summary>Style catalog. <see cref="DocxOperationResult.Success"/> means the package loaded and the styles part (when present) parsed.</summary>
public sealed record DocxStylesResult : DocxOperationResult
{
    public IReadOnlyList<DocxStyleInfo> Styles { get; init; } = [];
}

/// <summary>Image inventory. <see cref="DocxOperationResult.Success"/> means the package loaded and scanned.</summary>
public sealed record DocxMediaResult : DocxOperationResult
{
    public IReadOnlyList<DocxImageInfo> Images { get; init; } = [];
}

/// <summary>Validation outcome. <see cref="DocxOperationResult.Success"/> means no <see cref="DocxSeverity.Error"/> diagnostic was produced; warnings (including the <c>W9199</c>/<c>E9199</c> cap marker) do not fail validation by themselves.</summary>
public sealed record DocxValidateResult : DocxOperationResult
{
    public DocxValidationProfile Profile { get; init; } = DocxValidationProfile.Structural;
    public IReadOnlyList<string> PartNames { get; init; } = [];
    public string? MainDocumentPartName { get; init; }
}

/// <summary>Tracked-change and comment-markup inventory. <see cref="DocxOperationResult.Success"/> means the package loaded and scanned.</summary>
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

/// <summary>Patch dry-run. <see cref="DocxOperationResult.Success"/> means the patch parsed, the package loaded, and every operation reported success; nothing was written.</summary>
public sealed record DocxCheckResult : DocxOperationResult
{
    public IReadOnlyList<DocxPatchOperationReport> Operations { get; init; } = [];
}

/// <summary>Patch application. <see cref="DocxOperationResult.Success"/> means the patch parsed, every operation reported success, and the output was written. No output is written on failure.</summary>
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
    public IReadOnlyList<string> GeneratedRevisionIds { get; init; } = [];
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
    public int? VisualColumnEndIndex { get; init; }
    public int? GridBefore { get; init; }
    public int? GridAfter { get; init; }
    public string? MergeGroupId { get; init; }
    public string? NestedTablePath { get; init; }
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
    public long? WrapDistanceTopEmu { get; init; }
    public long? WrapDistanceBottomEmu { get; init; }
    public long? WrapDistanceLeftEmu { get; init; }
    public long? WrapDistanceRightEmu { get; init; }
    public long? RelativeHeight { get; init; }
    public bool? AllowOverlap { get; init; }
    public bool? LockAspectRatio { get; init; }
    public string? HorizontalPositionRelativeFrom { get; init; }
    public long? HorizontalPositionOffsetEmu { get; init; }
    public string? HorizontalPositionAlign { get; init; }
    public string? VerticalPositionRelativeFrom { get; init; }
    public long? VerticalPositionOffsetEmu { get; init; }
    public string? VerticalPositionAlign { get; init; }
    public decimal? CropLeftPercent { get; init; }
    public decimal? CropTopPercent { get; init; }
    public decimal? CropRightPercent { get; init; }
    public decimal? CropBottomPercent { get; init; }
}

public sealed record DocxChangeInfo
{
    public string Id { get; init; } = string.Empty;
    public string Type { get; init; } = string.Empty;
    public string Story { get; init; } = string.Empty;
    public string PartName { get; init; } = string.Empty;
    public string? ParentType { get; init; }
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
    public string? CommentParaId { get; init; }
    public string? CommentParentParaId { get; init; }
    public string? CommentRootParaId { get; init; }
    public string? CommentDurableId { get; init; }
    public bool? CommentIsReply { get; init; }
    public bool? CommentResolved { get; init; }
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
    public int? OperationIndex { get; init; }
    public string? OperationName { get; init; }
    public string? OperationTarget { get; init; }
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
    public string? ParaId { get; init; }
    public string? ParentParaId { get; init; }
    public string? RootParaId { get; init; }
    public string? DurableId { get; init; }
    public bool? IsReply { get; init; }
    public bool? Resolved { get; init; }
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
    public IReadOnlyList<string> FieldTypes { get; init; } = [];
    public IReadOnlyList<string> HyperlinkIds { get; init; } = [];
    public IReadOnlyList<string> HyperlinkTargets { get; init; } = [];
    public IReadOnlyList<string> CommentIds { get; init; } = [];
    public IReadOnlyList<string> CommentBodyIds { get; init; } = [];
    public IReadOnlyList<string> CommentParaIds { get; init; } = [];
    public IReadOnlyList<string> CommentParentParaIds { get; init; } = [];
    public IReadOnlyList<string> CommentRootParaIds { get; init; } = [];
    public IReadOnlyList<string> CommentDurableIds { get; init; } = [];
    public IReadOnlyList<string> CommentReplyIds { get; init; } = [];
    public IReadOnlyList<string> CommentResolvedIds { get; init; } = [];
    public string? Caption { get; init; }
    public string? Description { get; init; }
    public int? RowCount { get; init; }
    public int? ColumnCount { get; init; }
    public int? RowIndex { get; init; }
    public int? ColumnIndex { get; init; }
    public int? ColumnSpan { get; init; }
    public int? VisualColumnEndIndex { get; init; }
    public string? MergeGroupId { get; init; }
    public string? VerticalMerge { get; init; }
    public string? VerticalMergeRootCellId { get; init; }
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
    public IReadOnlyList<DocxListLabelComponent> LabelComponents { get; init; } = [];
    public string LabelStatus { get; init; } = "not-resolved";
    public IReadOnlyList<string> LabelWarnings { get; init; } = [];
    public int? StartValue { get; init; }
    public string? Suffix { get; init; }
    public bool IsLegal { get; init; }
    public int? RestartAfterLevel { get; init; }
    public string? ParagraphStyleId { get; init; }
    public string Source { get; init; } = "direct";
}

public sealed record DocxListLabelComponent(int Level, int Value, string Text, string Format);

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
    public bool IsNameDuplicate { get; init; }
    public IReadOnlyList<string> DuplicateNameBookmarkIds { get; init; } = [];
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
    public string? PlaceholderDocPart { get; init; }
    public bool IsShowingPlaceholderText { get; init; }
    public string? DataBindingXPath { get; init; }
    public string? DataBindingStoreItemId { get; init; }
    public string? DataBindingPrefixMappings { get; init; }
    public string? RepeatingSectionTitle { get; init; }
    public int? RepeatingSectionItemCount { get; init; }
    public string? ParentContentControlId { get; init; }
    public IReadOnlyList<string> ChildContentControlIds { get; init; } = [];
    public string SafeEditStatus { get; init; } = "unknown";
    public string? SafeEditReason { get; init; }
    public bool IsTagDuplicate { get; init; }
    public IReadOnlyList<string> DuplicateTagControlIds { get; init; } = [];
    public bool IsAliasDuplicate { get; init; }
    public IReadOnlyList<string> DuplicateAliasControlIds { get; init; } = [];
    public string? Lock { get; init; }
    public bool? Checked { get; init; }
    public string? CheckedSymbol { get; init; }
    public string? UncheckedSymbol { get; init; }
    public IReadOnlyList<DocxContentControlListItemInfo> ListItems { get; init; } = [];
    public string? DateFormat { get; init; }
    public string? DateLanguage { get; init; }
    public string? DateCalendar { get; init; }
    public string? DateValue { get; init; }
    public int TextLength { get; init; }
}

public sealed record DocxContentControlListItemInfo(string? DisplayText, string? Value);

public sealed record DocxFieldInfo
{
    public string Id { get; init; } = string.Empty;
    public string Story { get; init; } = string.Empty;
    public string PartName { get; init; } = string.Empty;
    public string? TargetId { get; init; }
    public string Kind { get; init; } = "unknown";
    public string? FieldType { get; init; }
    public string Code { get; init; } = string.Empty;
    public IReadOnlyList<string> Arguments { get; init; } = [];
    public IReadOnlyList<string> Switches { get; init; } = [];
    public string CachedResultText { get; init; } = string.Empty;
    public int ResultTextLength { get; init; }
    public int NestingDepth { get; init; }
    public IReadOnlyList<string> BookmarkDependencies { get; init; } = [];
    public IReadOnlyList<string> HyperlinkDependencies { get; init; } = [];
    public string RefreshPolicy { get; init; } = "unsupported";
    public string? RefreshReason { get; init; }
    public bool CanRefreshDeterministically { get; init; }
    public string SafeEditStatus { get; init; } = "unknown";
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
    public string? RelationshipPartName { get; init; }
    public string? RelationshipTargetMode { get; init; }
    public string? Uri { get; init; }
    public string? UriScheme { get; init; }
    public bool? IsUriValid { get; init; }
    public string? UriValidationReason { get; init; }
    public string? Anchor { get; init; }
    public bool? IsAnchorMissing { get; init; }
    public bool? IsAnchorDuplicate { get; init; }
    public string? Tooltip { get; init; }
    public string? TargetFrame { get; init; }
    public bool? History { get; init; }
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
    public string? Caption { get; init; }
    public string? Description { get; init; }
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
    public int VisualColumnEndIndex { get; init; }
    public string? MergeGroupId { get; init; }
    public string? VerticalMergeRootCellId { get; init; }
}

public sealed record DocxSectionInfo(
    string Id,
    string Story,
    int Columns,
    DocxOrientation Orientation);
