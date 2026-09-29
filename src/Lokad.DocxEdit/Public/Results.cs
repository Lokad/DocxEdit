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
    /// <summary>Scanned package part paths.</summary>
    public IReadOnlyList<string> PartNames { get; init; } = [];
    /// <summary>Main document part path, when found.</summary>
    public string? MainDocumentPartName { get; init; }
    /// <summary>Full read text.</summary>
    public string Text { get; init; } = string.Empty;
    /// <summary>Paragraphs in document order.</summary>
    public IReadOnlyList<DocxParagraphInfo> Paragraphs { get; init; } = [];
    /// <summary>Tables in document order.</summary>
    public IReadOnlyList<DocxTableInfo> Tables { get; init; } = [];
    /// <summary>Embedded images.</summary>
    public IReadOnlyList<DocxImageInfo> Images { get; init; } = [];
    /// <summary>Sections in document order.</summary>
    public IReadOnlyList<DocxSectionInfo> Sections { get; init; } = [];
    /// <summary>Bookmarks.</summary>
    public IReadOnlyList<DocxBookmarkInfo> Bookmarks { get; init; } = [];
    /// <summary>Content controls.</summary>
    public IReadOnlyList<DocxContentControlInfo> ContentControls { get; init; } = [];
    /// <summary>Fields.</summary>
    public IReadOnlyList<DocxFieldInfo> Fields { get; init; } = [];
    /// <summary>Hyperlinks.</summary>
    public IReadOnlyList<DocxHyperlinkInfo> Hyperlinks { get; init; } = [];
}

/// <summary>
/// One outline entry: an ID, a kind, and the kind-specific structural data.
/// Text bounds are applied when the entry is built; renderers format entries without re-truncating.
/// </summary>
public sealed record DocxOutlineItem
{
    /// <summary>Entry target wire ID.</summary>
    public string TargetId { get; init; } = string.Empty;
    /// <summary>Entry kind: heading, table, section, image, bookmark, content-control, field, or hyperlink.</summary>
    public string Kind { get; init; } = string.Empty;
    /// <summary>Entry text, bounded by the request maximum text; headings carry their text, other kinds leave this empty.</summary>
    public string Text { get; init; } = string.Empty;
    /// <summary>Heading level 1-9, when a heading.</summary>
    public int? HeadingLevel { get; init; }
    /// <summary>Numbering label info, when numbered.</summary>
    public DocxListInfo? List { get; init; }
    /// <summary>Table row count, when a table.</summary>
    public int RowCount { get; init; }
    /// <summary>Table or section column count, when a table or section.</summary>
    public int Columns { get; init; }
    /// <summary>Table style ID, when styled.</summary>
    public string? StyleId { get; init; }
    /// <summary>Table caption, bounded by the request maximum text.</summary>
    public string? Caption { get; init; }
    /// <summary>Table description, bounded by the request maximum text.</summary>
    public string? Description { get; init; }
    /// <summary>Table grid column count, when known.</summary>
    public int? GridColumnCount { get; init; }
    /// <summary>Whether the table has a header row.</summary>
    public bool HasHeaderRow { get; init; }
    /// <summary>Whether the table has merged cells.</summary>
    public bool HasMergedCells { get; init; }
    /// <summary>Whether the table has nested tables.</summary>
    public bool HasNestedTables { get; init; }
    /// <summary>Section orientation wire value, when a section.</summary>
    public string Orientation { get; init; } = string.Empty;
    /// <summary>Image layout kind, when an image.</summary>
    public string LayoutKind { get; init; } = string.Empty;
    /// <summary>Containing target wire ID, when an image inside a paragraph or cell.</summary>
    public string? ContainingTargetId { get; init; }
    /// <summary>Image part path, when an image.</summary>
    public string PartName { get; init; } = string.Empty;
    /// <summary>Bookmark name, when a bookmark.</summary>
    public string Name { get; init; } = string.Empty;
    /// <summary>Bookmark range start target wire ID, when known.</summary>
    public string? StartTargetId { get; init; }
    /// <summary>Bookmark range end target wire ID, when known.</summary>
    public string? EndTargetId { get; init; }
    /// <summary>Whether the bookmark name is duplicated.</summary>
    public bool IsNameDuplicate { get; init; }
    /// <summary>Bookmark IDs sharing the duplicated name.</summary>
    public IReadOnlyList<string> DuplicateNameBookmarkIds { get; init; } = [];
    /// <summary>Content-control kind, when a content control.</summary>
    public string ControlKind { get; init; } = string.Empty;
    /// <summary>Content-control target wire ID, when known.</summary>
    public string? ControlTargetId { get; init; }
    /// <summary>Content-control tag, when set.</summary>
    public string? Tag { get; init; }
    /// <summary>Content-control alias, when set.</summary>
    public string? Alias { get; init; }
    /// <summary>Parent content-control wire ID, when nested.</summary>
    public string? ParentControlId { get; init; }
    /// <summary>Child content-control wire IDs.</summary>
    public IReadOnlyList<string> ChildControlIds { get; init; } = [];
    /// <summary>Safe-edit status, when a content control or field.</summary>
    public string SafeEditStatus { get; init; } = string.Empty;
    /// <summary>Safe-edit reason, when present.</summary>
    public string? SafeEditReason { get; init; }
    /// <summary>Whether the content-control tag is duplicated.</summary>
    public bool IsTagDuplicate { get; init; }
    /// <summary>Control IDs sharing the duplicated tag.</summary>
    public IReadOnlyList<string> DuplicateTagControlIds { get; init; } = [];
    /// <summary>Whether the content-control alias is duplicated.</summary>
    public bool IsAliasDuplicate { get; init; }
    /// <summary>Control IDs sharing the duplicated alias.</summary>
    public IReadOnlyList<string> DuplicateAliasControlIds { get; init; } = [];
    /// <summary>Checkbox state, when a checkbox.</summary>
    public bool? Checked { get; init; }
    /// <summary>Dropdown item count, when a list control.</summary>
    public int? ListItemCount { get; init; }
    /// <summary>Field kind, when a field.</summary>
    public string FieldKind { get; init; } = string.Empty;
    /// <summary>Field type, when known.</summary>
    public string? FieldType { get; init; }
    /// <summary>Field target wire ID, when known.</summary>
    public string? FieldTargetId { get; init; }
    /// <summary>Normalized field code.</summary>
    public string Code { get; init; } = string.Empty;
    /// <summary>Field nesting depth.</summary>
    public int NestingDepth { get; init; }
    /// <summary>Field refresh-policy wire value.</summary>
    public string RefreshPolicy { get; init; } = string.Empty;
    /// <summary>Whether the field refreshes deterministically.</summary>
    public bool DeterministicRefresh { get; init; }
    /// <summary>Hyperlink target wire ID, when known.</summary>
    public string? HyperlinkTarget { get; init; }
    /// <summary>Hyperlink destination: URI, anchor, part path, or unknown.</summary>
    public string? Destination { get; init; }
    /// <summary>Whether the hyperlink relationship is broken.</summary>
    public bool IsBroken { get; init; }
}

/// <summary>Structured document outline. <see cref="DocxOperationResult.Success"/> means the package loaded and scanned.</summary>
public sealed record DocxOutlineResult : DocxOperationResult
{
    /// <summary>Scanned package part paths.</summary>
    public IReadOnlyList<string> PartNames { get; init; } = [];
    /// <summary>Main document part path, when found.</summary>
    public string? MainDocumentPartName { get; init; }
    /// <summary>Outline entries in display order.</summary>
    public IReadOnlyList<DocxOutlineItem> Items { get; init; } = [];
}

/// <summary>
/// One text-search hit: where it matched and the bounded match text.
/// Text bounds are applied when the hit is built; renderers format hits without re-truncating.
/// </summary>
public sealed record DocxFindMatch
{
    /// <summary>Matched target wire ID.</summary>
    public string TargetId { get; init; } = string.Empty;
    /// <summary>Match kind: paragraph or cell.</summary>
    public string Kind { get; init; } = string.Empty;
    /// <summary>Containing table wire ID for cell matches; null for paragraphs.</summary>
    public string? ParentId { get; init; }
    /// <summary>Match text, bounded by the request maximum text.</summary>
    public string Text { get; init; } = string.Empty;
    /// <summary>Numbering label info, when the matched paragraph is numbered.</summary>
    public DocxListInfo? List { get; init; }
}

/// <summary>Text search matches. An empty <c>Matches</c> list is a successful search with no hits, not a failure.</summary>
public sealed record DocxFindResult : DocxOperationResult
{
    /// <summary>Searched text.</summary>
    public string Query { get; init; } = string.Empty;
    /// <summary>Matches in display order.</summary>
    public IReadOnlyList<DocxFindMatch> Matches { get; init; } = [];
}

/// <summary>Target-scoped dump. <see cref="DocxOperationResult.Success"/> is equivalent to "target found": an unknown target or a target with the wrong shape for dumping fails with <c>E1201</c> and leaves <c>Text</c> null.</summary>
public sealed record DocxDumpResult : DocxOperationResult
{
    /// <summary>Requested target, echoed byte-identical (explicit ID, comment-body ID, or comment reference); never parsed here.</summary>
    public string TargetId { get; init; } = string.Empty;
    /// <summary>Dumped target text; null when the target was not found (the result then carries the failure).</summary>
    public string? Text { get; init; }
    /// <summary>Run detail, when requested.</summary>
    public IReadOnlyList<DocxDumpRunInfo> Runs { get; init; } = [];
}

/// <summary>Editing capabilities for one target under an effective track-change policy. Guidance only: capabilities can change after an edit, and check remains authoritative for a patch.</summary>
public sealed record DocxCapabilitiesResult : DocxOperationResult
{
    /// <summary>Requested target, echoed byte-identical.</summary>
    public string TargetId { get; init; } = string.Empty;
    /// <summary>Capabilities for the resolved target; null when the target was not found or has no capability model.</summary>
    public DocxTargetCapabilities? Capabilities { get; init; }
}

/// <summary>Guarded patch template for one target. The active block passes check under the requested policy and changes nothing visible; commented blocks cover the remaining supported operations. Check remains authoritative before apply.</summary>
public sealed record DocxTemplateResult : DocxOperationResult
{
    /// <summary>Requested target, echoed byte-identical.</summary>
    public string TargetId { get; init; } = string.Empty;
    /// <summary>Patch template text; empty when the target was not found or has no capability model.</summary>
    public string Template { get; init; } = string.Empty;
    /// <summary>Capabilities the template was generated from; null on failure.</summary>
    public DocxTargetCapabilities? Capabilities { get; init; }
}

/// <summary>Document-independent patch shape validation. Reports missing fields, unsatisfied alternative groups, conflicting exclusive fields, and patch syntax errors without loading a document. Success leaves document-dependent targets, guards, assets, and shapes to check.</summary>
public sealed record DocxLintResult : DocxOperationResult
{
    /// <summary>Number of parsed operations; zero when parsing failed.</summary>
    public int OperationCount { get; init; }
    /// <summary>Parsed operations in file order; empty when parsing failed.</summary>
    public IReadOnlyList<DocxLintOperation> Operations { get; init; } = [];
}

/// <summary>One linted patch operation: its index and name.</summary>
/// <param name="Index">1-based operation index matching the parser.</param>
/// <param name="Name">Operation name as written.</param>
public sealed record DocxLintOperation(int Index, string Name);

/// <summary>Supported, conditional, and unsupported operations for one discovered target.</summary>
/// <param name="TargetId">Canonical wire ID of the resolved target.</param>
/// <param name="Kind">Target kind word (for example <c>paragraph</c>).</param>
/// <param name="Story">Story label (for example <c>main</c>).</param>
/// <param name="Operations">Per-operation guidance in canonical operation order.</param>
public sealed record DocxTargetCapabilities(
    string TargetId,
    string Kind,
    string Story,
    IReadOnlyList<DocxOperationCapability> Operations)
{
}

/// <summary>What one operation can do to one target under the effective policy.</summary>
/// <param name="Operation">Operation name; also the help topic name.</param>
/// <param name="Support">One of <c>supported</c>, <c>conditional</c>, or <c>unsupported</c>.</param>
/// <param name="Reason">Actionable reason, including preservation and fallback notes.</param>
/// <param name="HelpTopic">Help topic with operation reference.</param>
/// <param name="Alternative">Safe alternative operation or target form, when one exists.</param>
public sealed record DocxOperationCapability(
    string Operation,
    string Support,
    string Reason,
    string HelpTopic,
    string? Alternative)
{
}

/// <summary>Target neighborhood. An unknown target fails with <c>E1201</c>; <see cref="DocxOperationResult.Success"/> is equivalent to "target found".</summary>
public sealed record DocxContextResult : DocxOperationResult
{
    /// <summary>Requested target, echoed byte-identical (explicit ID, comment-body ID, or comment reference); never parsed here.</summary>
    public string TargetId { get; init; } = string.Empty;
    /// <summary>Target text.</summary>
    public string Text { get; init; } = string.Empty;
    /// <summary>Neighborhood items.</summary>
    public IReadOnlyList<DocxContextItem> Items { get; init; } = [];
}

/// <summary>Style catalog. <see cref="DocxOperationResult.Success"/> means the package loaded and the styles part (when present) parsed.</summary>
public sealed record DocxStylesResult : DocxOperationResult
{
    /// <summary>Style definitions.</summary>
    public IReadOnlyList<DocxStyleInfo> Styles { get; init; } = [];
}

/// <summary>Image inventory. <see cref="DocxOperationResult.Success"/> means the package loaded and scanned.</summary>
public sealed record DocxMediaResult : DocxOperationResult
{
    /// <summary>Embedded images.</summary>
    public IReadOnlyList<DocxImageInfo> Images { get; init; } = [];
}

/// <summary>Embedded image bytes. <see cref="DocxOperationResult.Success"/> means the package loaded, scanned, and every inventoried part resolved.</summary>
public sealed record DocxMediaExtractResult : DocxOperationResult
{
    /// <summary>Extracted image files.</summary>
    public IReadOnlyList<DocxMediaFile> Files { get; init; } = [];
}

/// <summary>One extracted image file.</summary>
/// <param name="ImageId">Stable image ID.</param>
/// <param name="PartName">Media part path.</param>
/// <param name="FileName">Suggested file name (image ID plus the part file name).</param>
/// <param name="Content">Media bytes.</param>
public sealed record DocxMediaFile(DocxTargetId ImageId, string PartName, string FileName, byte[] Content);

/// <summary>Validation outcome. <see cref="DocxOperationResult.Success"/> means no <see cref="DocxSeverity.Error"/> diagnostic was produced; warnings (including the <c>W9199</c>/<c>E9199</c> cap marker) do not fail validation by themselves.</summary>
public sealed record DocxValidateResult : DocxOperationResult
{
    /// <summary>Validation profile used.</summary>
    public DocxValidationProfile Profile { get; init; } = DocxValidationProfile.Structural;
    /// <summary>Scanned package part paths.</summary>
    public IReadOnlyList<string> PartNames { get; init; } = [];
    /// <summary>Main document part path, when found.</summary>
    public string? MainDocumentPartName { get; init; }
}

/// <summary>Tracked-change and comment-markup inventory. <see cref="DocxOperationResult.Success"/> means the package loaded and scanned.</summary>
public sealed record DocxChangesResult : DocxOperationResult
{
    /// <summary>Scanned package part paths.</summary>
    public IReadOnlyList<string> PartNames { get; init; } = [];
    /// <summary>Main document part path, when found.</summary>
    public string? MainDocumentPartName { get; init; }
    /// <summary>Change records.</summary>
    public IReadOnlyList<DocxChangeInfo> Changes { get; init; } = [];
    /// <summary>Change counts by type.</summary>
    public IReadOnlyList<DocxChangeSummary> Summary { get; init; } = [];
    /// <summary>Change counts by group.</summary>
    public IReadOnlyList<DocxChangeGroupSummary> GroupSummary { get; init; } = [];
    /// <summary>Change counts by target.</summary>
    public IReadOnlyList<DocxChangeTargetSummary> TargetSummary { get; init; } = [];
    /// <summary>Change counts by comment thread.</summary>
    public IReadOnlyList<DocxCommentThreadSummary> CommentSummary { get; init; } = [];
}

/// <summary>Patch dry-run. <see cref="DocxOperationResult.Success"/> means the patch parsed, the package loaded, and every operation reported success; nothing was written.</summary>
public sealed record DocxCheckResult : DocxOperationResult
{
    /// <summary>Per-operation reports.</summary>
    public IReadOnlyList<DocxPatchOperationReport> Operations { get; init; } = [];

    /// <summary>Effective author recorded on generated revisions: the requested <c>Author</c> trimmed. Present on failed results too.</summary>
    public string Author { get; init; } = "docxedit";

    /// <summary>Effective track-change policy the patch ran under. Present on failed results too.</summary>
    public TrackChangesMode TrackChanges { get; init; } = TrackChangesMode.Off;

    /// <summary>Effective UTC timestamp used for generated revisions: the requested <c>TimestampUtc</c> converted to UTC. Present on failed results too.</summary>
    public DateTimeOffset TimestampUtc { get; init; } = DateTimeOffset.UtcNow;

    /// <summary>Version of the producing Lokad.DocxEdit assembly.</summary>
    public string ToolVersion { get; init; } = typeof(DocxOperationResult).Assembly.GetName().Version?.ToString() ?? "unknown";
}

/// <summary>Patch application. <see cref="DocxOperationResult.Success"/> means the patch parsed, every operation reported success, and the output was written. No output is written on failure.</summary>
public sealed record DocxApplyResult : DocxOperationResult
{
    /// <summary>Per-operation reports.</summary>
    public IReadOnlyList<DocxPatchOperationReport> Operations { get; init; } = [];

    /// <summary>Effective author recorded on generated revisions: the requested <c>Author</c> trimmed. Present on failed results too.</summary>
    public string Author { get; init; } = "docxedit";

    /// <summary>Effective track-change policy the patch ran under. Present on failed results too.</summary>
    public TrackChangesMode TrackChanges { get; init; } = TrackChangesMode.Off;

    /// <summary>Effective UTC timestamp used for generated revisions: the requested <c>TimestampUtc</c> converted to UTC. Present on failed results too.</summary>
    public DateTimeOffset TimestampUtc { get; init; } = DateTimeOffset.UtcNow;

    /// <summary>Version of the producing Lokad.DocxEdit assembly.</summary>
    public string ToolVersion { get; init; } = typeof(DocxOperationResult).Assembly.GetName().Version?.ToString() ?? "unknown";
}

/// <summary>
/// One operation outcome: target, success, and diagnostics.
/// </summary>
/// <param name="Index">1-based operation index, matching the parsed patch operation.</param>
/// <param name="OperationName">Operation name.</param>
/// <param name="Target">Target selector text, when the operation has one.</param>
/// <param name="Success">Whether the operation reported success.</param>
/// <param name="Diagnostics">Operation diagnostics.</param>
public sealed record DocxPatchOperationReport(
    int Index,
    string OperationName,
    string? Target,
    bool Success,
    IReadOnlyList<DocxDiagnostic> Diagnostics)
{
    /// <summary>Affected targets.</summary>
    public IReadOnlyList<DocxPatchAffectedTarget> AffectedTargets { get; init; } = [];
    /// <summary>Generated revision IDs.</summary>
    public IReadOnlyList<string> GeneratedRevisionIds { get; init; } = [];
    /// <summary>Wire IDs created by the operation, such as an inserted paragraph or a new comment; empty when the operation creates no addressable targets.</summary>
    public IReadOnlyList<string> CreatedTargetIds { get; init; } = [];
    /// <summary>Bounded text before the operation, when previews are enabled.</summary>
    public string? PreviewBefore { get; init; }
    /// <summary>Bounded text after the operation, when previews are enabled.</summary>
    public string? PreviewAfter { get; init; }
    /// <summary>Whether either preview side was truncated to the preview budget.</summary>
    public bool PreviewTruncated { get; init; }
}

/// <summary>
/// One target touched by an operation, with shape metadata.
/// </summary>
/// <param name="Id">Affected target ID.</param>
/// <param name="Kind">Affected target kind.</param>
/// <param name="Action">What the operation did to the target.</param>
public sealed record DocxPatchAffectedTarget(DocxTargetId Id, string Kind, string Action)
{
    /// <summary>Containing target ID, when known.</summary>
    public DocxTargetId? ParentId { get; init; }
    /// <summary>1-based row ordinal.</summary>
    public int? RowIndex { get; init; }
    /// <summary>1-based visual column ordinal.</summary>
    public int? ColumnIndex { get; init; }
    /// <summary>Row count before the edit.</summary>
    public int? RowCountBefore { get; init; }
    /// <summary>Row count after the edit.</summary>
    public int? RowCountAfter { get; init; }
    /// <summary>Column count.</summary>
    public int? ColumnCount { get; init; }
    /// <summary>Cell count.</summary>
    public int? CellCount { get; init; }
    /// <summary>1-based visual end column ordinal.</summary>
    public int? VisualColumnEndIndex { get; init; }
    /// <summary>Leading grid columns.</summary>
    public int? GridBefore { get; init; }
    /// <summary>Trailing grid columns.</summary>
    public int? GridAfter { get; init; }
    /// <summary>Merge-group ID, when merged.</summary>
    public DocxTargetId? MergeGroupId { get; init; }
    /// <summary>Nested-table path, when nested.</summary>
    public string? NestedTablePath { get; init; }
}

/// <summary>
/// One style definition with inheritance and numbering links.
/// </summary>
/// <param name="StyleId">Style ID.</param>
/// <param name="Name">Style name.</param>
/// <param name="Type">Style type (paragraph, character, or table).</param>
/// <param name="IsDefault">Whether this is a default style.</param>
public sealed record DocxStyleInfo(string StyleId, string Name, string Type, bool IsDefault)
{
    /// <summary>Base style ID, when inherited.</summary>
    public string? BasedOnStyleId { get; init; }
    /// <summary>Next style ID, when set.</summary>
    public string? NextStyleId { get; init; }
    /// <summary>Linked style ID, when set.</summary>
    public string? LinkedStyleId { get; init; }
    /// <summary>Numbering instance ID, when numbered.</summary>
    public string? NumberingId { get; init; }
    /// <summary>Numbering level, when numbered.</summary>
    public int? NumberingLevel { get; init; }
}

/// <summary>
/// One embedded image with DrawingML layout metadata.
/// </summary>
/// <param name="Id">Stable image ID.</param>
/// <param name="PartName">Media part path.</param>
/// <param name="ContentType">Media content type, when known.</param>
/// <param name="ByteLength">Media byte length.</param>
public sealed record DocxImageInfo(DocxTargetId Id, string PartName, string? ContentType, long ByteLength)
{
    /// <summary>DrawingML layout kind.</summary>
    public string LayoutKind { get; init; } = "unknown";
    /// <summary>Package relationship ID, when known.</summary>
    public string? RelationshipId { get; init; }
    /// <summary>Containing target ID, when known.</summary>
    public DocxTargetId? ContainingTargetId { get; init; }
    /// <summary>Extent width in EMU, when known.</summary>
    public long? WidthEmu { get; init; }
    /// <summary>Extent height in EMU, when known.</summary>
    public long? HeightEmu { get; init; }
    /// <summary>Image name, when present.</summary>
    public string? Name { get; init; }
    /// <summary>Description, when present.</summary>
    public string? Description { get; init; }
    /// <summary>Title, when present.</summary>
    public string? Title { get; init; }
    /// <summary>Text-wrapping mode, when anchored.</summary>
    public string? WrapMode { get; init; }
    /// <summary>Whether the drawing sits behind text.</summary>
    public bool BehindDoc { get; init; }
    /// <summary>Top wrap distance in EMU, when set.</summary>
    public long? WrapDistanceTopEmu { get; init; }
    /// <summary>Bottom wrap distance in EMU, when set.</summary>
    public long? WrapDistanceBottomEmu { get; init; }
    /// <summary>Left wrap distance in EMU, when set.</summary>
    public long? WrapDistanceLeftEmu { get; init; }
    /// <summary>Right wrap distance in EMU, when set.</summary>
    public long? WrapDistanceRightEmu { get; init; }
    /// <summary>Relative height, when set.</summary>
    public long? RelativeHeight { get; init; }
    /// <summary>Whether overlap is allowed, when known.</summary>
    public bool? AllowOverlap { get; init; }
    /// <summary>Whether aspect ratio is locked, when known.</summary>
    public bool? LockAspectRatio { get; init; }
    /// <summary>Horizontal positioning datum, when anchored.</summary>
    public string? HorizontalPositionRelativeFrom { get; init; }
    /// <summary>Horizontal offset in EMU, when anchored.</summary>
    public long? HorizontalPositionOffsetEmu { get; init; }
    /// <summary>Horizontal alignment, when anchored.</summary>
    public string? HorizontalPositionAlign { get; init; }
    /// <summary>Vertical positioning datum, when anchored.</summary>
    public string? VerticalPositionRelativeFrom { get; init; }
    /// <summary>Vertical offset in EMU, when anchored.</summary>
    public long? VerticalPositionOffsetEmu { get; init; }
    /// <summary>Vertical alignment, when anchored.</summary>
    public string? VerticalPositionAlign { get; init; }
    /// <summary>Left crop percent.</summary>
    public decimal? CropLeftPercent { get; init; }
    /// <summary>Top crop percent.</summary>
    public decimal? CropTopPercent { get; init; }
    /// <summary>Right crop percent.</summary>
    public decimal? CropRightPercent { get; init; }
    /// <summary>Bottom crop percent.</summary>
    public decimal? CropBottomPercent { get; init; }
}

/// <summary>
/// One tracked change or markup item with target annotation.
/// </summary>
public sealed record DocxChangeInfo
{
    /// <summary>Stable change ID.</summary>
    public required DocxChangeId Id { get; init; }
    /// <summary>Change type.</summary>
    public string Type { get; init; } = string.Empty;
    /// <summary>Story label.</summary>
    public string Story { get; init; } = string.Empty;
    /// <summary>Package part path.</summary>
    public string PartName { get; init; } = string.Empty;
    /// <summary>Parent markup type, when nested.</summary>
    public string? ParentType { get; init; }
    /// <summary>Associated target ID, when known: an explicit target ID (<c>M.P0001</c>) or a comment-body ID (<c>C001.C0001</c>); null when targetless. Never a patch selector.</summary>
    public string? TargetId { get; init; }
    /// <summary>Revision author, when tracked.</summary>
    public string? Author { get; init; }
    /// <summary>Revision timestamp, when tracked.</summary>
    public DateTimeOffset? TimestampUtc { get; init; }
    /// <summary>Revision ID, when tracked.</summary>
    public string? RevisionId { get; init; }
    /// <summary>Text character length.</summary>
    public int TextLength { get; init; }
    /// <summary>Child element count.</summary>
    public int ChildElementCount { get; init; }
    /// <summary>Comment ID.</summary>
    public string? CommentId { get; init; }
    /// <summary>Comment author.</summary>
    public string? CommentAuthor { get; init; }
    /// <summary>Comment timestamp.</summary>
    public DateTimeOffset? CommentTimestampUtc { get; init; }
    /// <summary>Comment author initials.</summary>
    public string? CommentInitials { get; init; }
    /// <summary>Comment paragraph ID.</summary>
    public string? CommentParaId { get; init; }
    /// <summary>Comment parent paragraph ID.</summary>
    public string? CommentParentParaId { get; init; }
    /// <summary>Comment root paragraph ID.</summary>
    public string? CommentRootParaId { get; init; }
    /// <summary>Comment durable ID.</summary>
    public string? CommentDurableId { get; init; }
    /// <summary>Whether the comment is a reply.</summary>
    public bool? CommentIsReply { get; init; }
    /// <summary>Whether the comment is resolved.</summary>
    public bool? CommentResolved { get; init; }
    /// <summary>Comment anchor target ID, when resolved (same ID families as target IDs).</summary>
    public string? CommentAnchorTargetId { get; init; }
    /// <summary>Comment reference target ID, when resolved (same ID families as target IDs).</summary>
    public string? CommentReferenceTargetId { get; init; }
    /// <summary>Comment anchor story, when resolved.</summary>
    public string? CommentAnchorStory { get; init; }
    /// <summary>Comment anchor part, when resolved.</summary>
    public string? CommentAnchorPartName { get; init; }
    /// <summary>Comment text character length, when text is included.</summary>
    public int? CommentTextLength { get; init; }
    /// <summary>Comment text excerpt, when text is included.</summary>
    public string? CommentTextSnippet { get; init; }
    /// <summary>Whether the comment text excerpt was truncated.</summary>
    public bool CommentTextTruncated { get; init; }
    /// <summary>Target resolution status.</summary>
    public DocxTargetStatus TargetStatus { get; init; } = DocxTargetStatus.Targetless;
    /// <summary>How the target was resolved.</summary>
    public DocxTargetSource TargetSource { get; init; } = DocxTargetSource.None;
    /// <summary>Why there is no modeled target, when targetless.</summary>
    public DocxTargetReason? TargetReason { get; init; }
    /// <summary>Nearest modeled target ID, when targetless (same ID families as target IDs).</summary>
    public string? NearestTargetId { get; init; }
    /// <summary>Human-readable target note, when targetless.</summary>
    public string? TargetNote { get; init; }
    /// <summary>Paired change ID, when ranges pair.</summary>
    public DocxChangeId? PairedChangeId { get; init; }
    /// <summary>1-based patch operation index, when from a patch.</summary>
    public int? OperationIndex { get; init; }
    /// <summary>Patch operation name, when from a patch.</summary>
    public string? OperationName { get; init; }
    /// <summary>Raw patch operation target text, when from a patch: an explicit ID or any selector, echoed for attribution.</summary>
    public string? OperationTarget { get; init; }
}

/// <summary>
/// Change count by type.
/// </summary>
/// <param name="Type">Change type.</param>
/// <param name="Count">Number of changes.</param>
public sealed record DocxChangeSummary(string Type, int Count);

/// <summary>
/// Change count by group.
/// </summary>
/// <param name="Group">Group key kind.</param>
/// <param name="Key">Group key.</param>
/// <param name="Type">Change type.</param>
/// <param name="Count">Number of changes.</param>
public sealed record DocxChangeGroupSummary(string Group, string Key, string Type, int Count);

/// <summary>
/// Change counts for one target.
/// </summary>
public sealed record DocxChangeTargetSummary
{
    /// <summary>Grouped target key (same ID families as change target IDs).</summary>
    public string TargetId { get; init; } = string.Empty;
    /// <summary>Number of changes.</summary>
    public int Count { get; init; }
    /// <summary>Change counts by type.</summary>
    public IReadOnlyList<DocxChangeSummary> Summary { get; init; } = [];
}

/// <summary>
/// One comment thread with anchor metadata.
/// </summary>
public sealed record DocxCommentThreadSummary
{
    /// <summary>Comment ID.</summary>
    public string CommentId { get; init; } = string.Empty;
    /// <summary>Anchor target ID, when resolved (same ID families as target IDs).</summary>
    public string? AnchorTargetId { get; init; }
    /// <summary>Reference target ID, when resolved (same ID families as target IDs).</summary>
    public string? ReferenceTargetId { get; init; }
    /// <summary>Anchor story, when resolved.</summary>
    public string? AnchorStory { get; init; }
    /// <summary>Anchor part path, when resolved.</summary>
    public string? AnchorPartName { get; init; }
    /// <summary>Comment author.</summary>
    public string? Author { get; init; }
    /// <summary>Comment timestamp.</summary>
    public DateTimeOffset? TimestampUtc { get; init; }
    /// <summary>Comment author initials.</summary>
    public string? Initials { get; init; }
    /// <summary>Comment paragraph ID.</summary>
    public string? ParaId { get; init; }
    /// <summary>Comment parent paragraph ID.</summary>
    public string? ParentParaId { get; init; }
    /// <summary>Comment root paragraph ID.</summary>
    public string? RootParaId { get; init; }
    /// <summary>Comment durable ID.</summary>
    public string? DurableId { get; init; }
    /// <summary>Whether the comment is a reply.</summary>
    public bool? IsReply { get; init; }
    /// <summary>Whether the thread is resolved.</summary>
    public bool? Resolved { get; init; }
    /// <summary>Comment text character length, when text is included.</summary>
    public int? TextLength { get; init; }
    /// <summary>Comment text excerpt, when text is included.</summary>
    public string? TextSnippet { get; init; }
    /// <summary>Whether the comment text excerpt was truncated.</summary>
    public bool TextTruncated { get; init; }
    /// <summary>Number of changes.</summary>
    public int Count { get; init; }
    /// <summary>Change counts by type.</summary>
    public IReadOnlyList<DocxChangeSummary> Summary { get; init; } = [];
}

/// <summary>
/// One run in a dumped target.
/// </summary>
public sealed record DocxDumpRunInfo
{
    /// <summary>Run ID: the owning paragraph wire ID plus a run ordinal (<c>M.P0001.R0001</c>); a composite key, not an addressable target.</summary>
    public string Id { get; init; } = string.Empty;
    /// <summary>Run text.</summary>
    public string Text { get; init; } = string.Empty;
    /// <summary>Run markup type, when marked.</summary>
    public string? MarkupType { get; init; }
    /// <summary>Revision ID, when tracked.</summary>
    public string? RevisionId { get; init; }
    /// <summary>Revision author, when tracked.</summary>
    public string? Author { get; init; }
    /// <summary>Revision timestamp, when tracked.</summary>
    public DateTimeOffset? TimestampUtc { get; init; }
    /// <summary>Comment ID.</summary>
    public string? CommentId { get; init; }
    /// <summary>Hyperlink relationship ID, when linked.</summary>
    public string? HyperlinkRelationshipId { get; init; }
    /// <summary>Hyperlink anchor, when linked.</summary>
    public string? HyperlinkAnchor { get; init; }
}

/// <summary>
/// One neighborhood item with relation and attached markup.
/// </summary>
public sealed record DocxContextItem
{
    /// <summary>Item ID: a model or change wire ID, or a <c>comment:</c> composite for comment bodies; a composite key, not always addressable.</summary>
    public string Id { get; init; } = string.Empty;
    /// <summary>Item kind.</summary>
    public string Kind { get; init; } = string.Empty;
    /// <summary>Position relative to the target (before, target, or after).</summary>
    public string Relation { get; init; } = string.Empty;
    /// <summary>Story label.</summary>
    public string Story { get; init; } = string.Empty;
    /// <summary>Containing ID, when known: a table wire ID, a media part path, or a change-union member, depending on the item kind.</summary>
    public string? ParentId { get; init; }
    /// <summary>Item text.</summary>
    public string Text { get; init; } = string.Empty;
    /// <summary>Heading level 1–9, when a heading.</summary>
    public int? HeadingLevel { get; init; }
    /// <summary>Style ID, when styled.</summary>
    public string? StyleId { get; init; }
    /// <summary>Style display name, when styled.</summary>
    public string? StyleName { get; init; }
    /// <summary>Numbering label info, when numbered.</summary>
    public DocxListInfo? List { get; init; }
    /// <summary>Attached bookmark names.</summary>
    public IReadOnlyList<string> BookmarkNames { get; init; } = [];
    /// <summary>Attached content-control IDs.</summary>
    public IReadOnlyList<string> ContentControlIds { get; init; } = [];
    /// <summary>Attached content-control tags.</summary>
    public IReadOnlyList<string> ContentControlTags { get; init; } = [];
    /// <summary>Attached content-control aliases.</summary>
    public IReadOnlyList<string> ContentControlAliases { get; init; } = [];
    /// <summary>Attached field IDs.</summary>
    public IReadOnlyList<string> FieldIds { get; init; } = [];
    /// <summary>Attached field codes.</summary>
    public IReadOnlyList<string> FieldCodes { get; init; } = [];
    /// <summary>Attached field kinds.</summary>
    public IReadOnlyList<string> FieldKinds { get; init; } = [];
    /// <summary>Attached field types.</summary>
    public IReadOnlyList<string> FieldTypes { get; init; } = [];
    /// <summary>Attached hyperlink IDs.</summary>
    public IReadOnlyList<string> HyperlinkIds { get; init; } = [];
    /// <summary>Attached hyperlink targets.</summary>
    public IReadOnlyList<string> HyperlinkTargets { get; init; } = [];
    /// <summary>Attached comment IDs.</summary>
    public IReadOnlyList<string> CommentIds { get; init; } = [];
    /// <summary>Attached comment-body IDs.</summary>
    public IReadOnlyList<string> CommentBodyIds { get; init; } = [];
    /// <summary>Attached comment paragraph IDs.</summary>
    public IReadOnlyList<string> CommentParaIds { get; init; } = [];
    /// <summary>Attached comment parent paragraph IDs.</summary>
    public IReadOnlyList<string> CommentParentParaIds { get; init; } = [];
    /// <summary>Attached comment root paragraph IDs.</summary>
    public IReadOnlyList<string> CommentRootParaIds { get; init; } = [];
    /// <summary>Attached comment durable IDs.</summary>
    public IReadOnlyList<string> CommentDurableIds { get; init; } = [];
    /// <summary>Attached comment reply IDs.</summary>
    public IReadOnlyList<string> CommentReplyIds { get; init; } = [];
    /// <summary>Attached resolved comment IDs.</summary>
    public IReadOnlyList<string> CommentResolvedIds { get; init; } = [];
    /// <summary>Caption, when present.</summary>
    public string? Caption { get; init; }
    /// <summary>Description, when present.</summary>
    public string? Description { get; init; }
    /// <summary>Row count.</summary>
    public int? RowCount { get; init; }
    /// <summary>Column count.</summary>
    public int? ColumnCount { get; init; }
    /// <summary>1-based row ordinal.</summary>
    public int? RowIndex { get; init; }
    /// <summary>1-based visual column ordinal.</summary>
    public int? ColumnIndex { get; init; }
    /// <summary>Spanned columns.</summary>
    public int? ColumnSpan { get; init; }
    /// <summary>1-based visual end column ordinal.</summary>
    public int? VisualColumnEndIndex { get; init; }
    /// <summary>Merge-group ID, when merged.</summary>
    public string? MergeGroupId { get; init; }
    /// <summary>Vertical-merge marker, when merged.</summary>
    public DocxVerticalMerge? VerticalMerge { get; init; }
    /// <summary>Merge-root cell ID, when merged.</summary>
    public string? VerticalMergeRootCellId { get; init; }
    /// <summary>Whether a nested table is present.</summary>
    public bool HasNestedTable { get; init; }
}

/// <summary>
/// One paragraph with text, style, and numbering.
/// </summary>
/// <param name="Id">Stable paragraph ID.</param>
/// <param name="Story">Story label.</param>
/// <param name="Text">Visible paragraph text.</param>
/// <param name="HeadingLevel">Heading level 1–9, when a heading.</param>
/// <param name="List">Numbering label info, when numbered.</param>
/// <param name="Runs">Runs in document order.</param>
public sealed record DocxParagraphInfo(
    DocxTargetId Id,
    string Story,
    string Text,
    int? HeadingLevel,
    DocxListInfo? List,
    IReadOnlyList<DocxRunInfo> Runs)
{
    /// <summary>Style ID, when styled.</summary>
    public string? StyleId { get; init; }
    /// <summary>Style display name, when styled.</summary>
    public string? StyleName { get; init; }
}

/// <summary>
/// Numbering label detail for a paragraph.
/// </summary>
/// <param name="NumberingId">Numbering instance ID.</param>
/// <param name="Level">Zero-based numbering level.</param>
public sealed record DocxListInfo(string NumberingId, int Level)
{
    /// <summary>Abstract numbering ID, when known.</summary>
    public string? AbstractNumberingId { get; init; }
    /// <summary>Numbering format, when known.</summary>
    public string? Format { get; init; }
    /// <summary>Level text, when known.</summary>
    public string? LevelText { get; init; }
    /// <summary>Resolved label text, when known.</summary>
    public string? LabelText { get; init; }
    /// <summary>Label components.</summary>
    public IReadOnlyList<DocxListLabelComponent> LabelComponents { get; init; } = [];
    /// <summary>Label resolution status.</summary>
    public DocxLabelStatus LabelStatus { get; init; } = DocxLabelStatus.NotResolved;
    /// <summary>Label warnings.</summary>
    public IReadOnlyList<string> LabelWarnings { get; init; } = [];
    /// <summary>Start value override, when set.</summary>
    public int? StartValue { get; init; }
    /// <summary>Level suffix, when known.</summary>
    public string? Suffix { get; init; }
    /// <summary>Whether legal numbering applies.</summary>
    public bool IsLegal { get; init; }
    /// <summary>Restart-after level, when set.</summary>
    public int? RestartAfterLevel { get; init; }
    /// <summary>Paragraph style ID, when set.</summary>
    public string? ParagraphStyleId { get; init; }
    /// <summary>Label source.</summary>
    public DocxLabelSource Source { get; init; } = DocxLabelSource.Direct;
}

/// <summary>
/// One level-text component.
/// </summary>
/// <param name="Level">Numbering level.</param>
/// <param name="Value">Component value.</param>
/// <param name="Text">Component text.</param>
/// <param name="Format">Component format.</param>
public sealed record DocxListLabelComponent(int Level, int Value, string Text, string Format);

/// <summary>
/// One run with markup annotations.
/// </summary>
/// <param name="Text">Run text.</param>
public sealed record DocxRunInfo(string Text)
{
    /// <summary>Run markup type, when marked.</summary>
    public string? MarkupType { get; init; }
    /// <summary>Revision ID, when tracked.</summary>
    public string? RevisionId { get; init; }
    /// <summary>Revision author, when tracked.</summary>
    public string? Author { get; init; }
    /// <summary>Revision timestamp, when tracked.</summary>
    public DateTimeOffset? TimestampUtc { get; init; }
    /// <summary>Comment ID.</summary>
    public string? CommentId { get; init; }
    /// <summary>Hyperlink relationship ID, when linked.</summary>
    public string? HyperlinkRelationshipId { get; init; }
    /// <summary>Hyperlink anchor, when linked.</summary>
    public string? HyperlinkAnchor { get; init; }
}

/// <summary>
/// One bookmark with range resolution.
/// </summary>
public sealed record DocxBookmarkInfo
{
    /// <summary>Stable bookmark ID.</summary>
    public required DocxTargetId Id { get; init; }
    /// <summary>Bookmark name.</summary>
    public string Name { get; init; } = string.Empty;
    /// <summary>OOXML bookmark ID, when known.</summary>
    public string? OoxmlId { get; init; }
    /// <summary>Story label.</summary>
    public string Story { get; init; } = string.Empty;
    /// <summary>Package part path.</summary>
    public string PartName { get; init; } = string.Empty;
    /// <summary>Range-start target ID, when resolved.</summary>
    public DocxTargetId? StartTargetId { get; init; }
    /// <summary>Range-end target ID, when resolved.</summary>
    public DocxTargetId? EndTargetId { get; init; }
    /// <summary>Whether the markup is complete.</summary>
    public bool IsComplete { get; init; }
    /// <summary>Whether the name is duplicated.</summary>
    public bool IsNameDuplicate { get; init; }
    /// <summary>Bookmark IDs sharing the name.</summary>
    public IReadOnlyList<string> DuplicateNameBookmarkIds { get; init; } = [];
}

/// <summary>
/// One content control with kind-specific metadata.
/// </summary>
public sealed record DocxContentControlInfo
{
    /// <summary>Stable content-control ID.</summary>
    public required DocxTargetId Id { get; init; }
    /// <summary>Story label.</summary>
    public string Story { get; init; } = string.Empty;
    /// <summary>Package part path.</summary>
    public string PartName { get; init; } = string.Empty;
    /// <summary>Associated target ID, when known.</summary>
    public DocxTargetId? TargetId { get; init; }
    /// <summary>Content-control kind.</summary>
    public string Kind { get; init; } = "unknown";
    /// <summary>OOXML bookmark ID, when known.</summary>
    public string? OoxmlId { get; init; }
    /// <summary>Content-control tag, when set.</summary>
    public string? Tag { get; init; }
    /// <summary>Content-control alias, when set.</summary>
    public string? Alias { get; init; }
    /// <summary>Placeholder building-block name, when set.</summary>
    public string? PlaceholderDocPart { get; init; }
    /// <summary>Whether placeholder text is shown.</summary>
    public bool IsShowingPlaceholderText { get; init; }
    /// <summary>Data-binding XPath, when bound.</summary>
    public string? DataBindingXPath { get; init; }
    /// <summary>Data-binding store ID, when bound.</summary>
    public string? DataBindingStoreItemId { get; init; }
    /// <summary>Data-binding prefix mappings, when bound.</summary>
    public string? DataBindingPrefixMappings { get; init; }
    /// <summary>Repeating-section title, when set.</summary>
    public string? RepeatingSectionTitle { get; init; }
    /// <summary>Repeating-section item count, when set.</summary>
    public int? RepeatingSectionItemCount { get; init; }
    /// <summary>Parent content-control ID, when nested.</summary>
    public DocxTargetId? ParentContentControlId { get; init; }
    /// <summary>Child content-control IDs.</summary>
    public IReadOnlyList<string> ChildContentControlIds { get; init; } = [];
    /// <summary>Direct-edit safety status.</summary>
    public string SafeEditStatus { get; init; } = "unknown";
    /// <summary>Why the safety status applies, when known.</summary>
    public string? SafeEditReason { get; init; }
    /// <summary>Whether the tag is duplicated.</summary>
    public bool IsTagDuplicate { get; init; }
    /// <summary>Control IDs sharing the tag.</summary>
    public IReadOnlyList<string> DuplicateTagControlIds { get; init; } = [];
    /// <summary>Whether the alias is duplicated.</summary>
    public bool IsAliasDuplicate { get; init; }
    /// <summary>Control IDs sharing the alias.</summary>
    public IReadOnlyList<string> DuplicateAliasControlIds { get; init; } = [];
    /// <summary>Content-control lock, when set.</summary>
    public string? Lock { get; init; }
    /// <summary>Checked state, when a checkbox.</summary>
    public bool? Checked { get; init; }
    /// <summary>Checked symbol, when set.</summary>
    public string? CheckedSymbol { get; init; }
    /// <summary>Unchecked symbol, when set.</summary>
    public string? UncheckedSymbol { get; init; }
    /// <summary>Drop-down items.</summary>
    public IReadOnlyList<DocxContentControlListItemInfo> ListItems { get; init; } = [];
    /// <summary>Date format, when set.</summary>
    public string? DateFormat { get; init; }
    /// <summary>Date language, when set.</summary>
    public string? DateLanguage { get; init; }
    /// <summary>Date calendar, when set.</summary>
    public string? DateCalendar { get; init; }
    /// <summary>Date value, when set.</summary>
    public string? DateValue { get; init; }
    /// <summary>Text character length.</summary>
    public int TextLength { get; init; }
}

/// <summary>
/// One drop-down/combo-box item.
/// </summary>
/// <param name="DisplayText">Display text.</param>
/// <param name="Value">Item value.</param>
public sealed record DocxContentControlListItemInfo(string? DisplayText, string? Value);

/// <summary>
/// One field with code, result, and refresh metadata.
/// </summary>
public sealed record DocxFieldInfo
{
    /// <summary>Stable field ID.</summary>
    public required DocxTargetId Id { get; init; }
    /// <summary>Story label.</summary>
    public string Story { get; init; } = string.Empty;
    /// <summary>Package part path.</summary>
    public string PartName { get; init; } = string.Empty;
    /// <summary>Associated target ID, when known.</summary>
    public DocxTargetId? TargetId { get; init; }
    /// <summary>Field kind (simple or complex).</summary>
    public string Kind { get; init; } = "unknown";
    /// <summary>Field type code, when known.</summary>
    public string? FieldType { get; init; }
    /// <summary>Field code.</summary>
    public string Code { get; init; } = string.Empty;
    /// <summary>Field arguments.</summary>
    public IReadOnlyList<string> Arguments { get; init; } = [];
    /// <summary>Field switches.</summary>
    public IReadOnlyList<string> Switches { get; init; } = [];
    /// <summary>Cached field result text.</summary>
    public string CachedResultText { get; init; } = string.Empty;
    /// <summary>Result text character length.</summary>
    public int ResultTextLength { get; init; }
    /// <summary>Field nesting depth.</summary>
    public int NestingDepth { get; init; }
    /// <summary>Bookmark dependencies.</summary>
    public IReadOnlyList<string> BookmarkDependencies { get; init; } = [];
    /// <summary>Hyperlink dependencies.</summary>
    public IReadOnlyList<string> HyperlinkDependencies { get; init; } = [];
    /// <summary>Field refresh policy.</summary>
    public DocxRefreshPolicy RefreshPolicy { get; init; } = DocxRefreshPolicy.Unsupported;
    /// <summary>Why the field refreshes that way.</summary>
    public string? RefreshReason { get; init; }
    /// <summary>Whether the result can be recomputed without Word.</summary>
    public bool CanRefreshDeterministically { get; init; }
    /// <summary>Direct-edit safety status.</summary>
    public string SafeEditStatus { get; init; } = "unknown";
    /// <summary>Whether the field is dirty, when known.</summary>
    public bool? IsDirty { get; init; }
    /// <summary>Whether the field is locked, when known.</summary>
    public bool? IsLocked { get; init; }
    /// <summary>Whether the markup is complete.</summary>
    public bool IsComplete { get; init; }
}

/// <summary>
/// One hyperlink with destination validation.
/// </summary>
public sealed record DocxHyperlinkInfo
{
    /// <summary>Stable hyperlink ID.</summary>
    public required DocxTargetId Id { get; init; }
    /// <summary>Story label.</summary>
    public string Story { get; init; } = string.Empty;
    /// <summary>Package part path.</summary>
    public string PartName { get; init; } = string.Empty;
    /// <summary>Associated target ID, when known.</summary>
    public DocxTargetId? TargetId { get; init; }
    /// <summary>Package relationship ID, when known.</summary>
    public string? RelationshipId { get; init; }
    /// <summary>Relationship part path, when known.</summary>
    public string? RelationshipPartName { get; init; }
    /// <summary>Relationship target mode, when known.</summary>
    public string? RelationshipTargetMode { get; init; }
    /// <summary>Hyperlink URI, when present.</summary>
    public string? Uri { get; init; }
    /// <summary>Hyperlink URI scheme, when known.</summary>
    public string? UriScheme { get; init; }
    /// <summary>Whether the URI is valid, when known.</summary>
    public bool? IsUriValid { get; init; }
    /// <summary>URI validation reason, when known.</summary>
    public string? UriValidationReason { get; init; }
    /// <summary>Internal anchor, when present.</summary>
    public string? Anchor { get; init; }
    /// <summary>Whether the anchor target is missing.</summary>
    public bool? IsAnchorMissing { get; init; }
    /// <summary>Whether the anchor is duplicated.</summary>
    public bool? IsAnchorDuplicate { get; init; }
    /// <summary>Tooltip, when present.</summary>
    public string? Tooltip { get; init; }
    /// <summary>Target frame, when present.</summary>
    public string? TargetFrame { get; init; }
    /// <summary>History flag, when known.</summary>
    public bool? History { get; init; }
    /// <summary>Target part path, when known.</summary>
    public string? TargetPartName { get; init; }
    /// <summary>Whether the target is external.</summary>
    public bool IsExternal { get; init; }
    /// <summary>Whether the link is broken.</summary>
    public bool IsBroken { get; init; }
    /// <summary>Display text character length.</summary>
    public int DisplayTextLength { get; init; }
}

/// <summary>
/// One table with grid metadata.
/// </summary>
/// <param name="Id">Stable table ID.</param>
/// <param name="Story">Story label.</param>
/// <param name="RowCount">Row count.</param>
/// <param name="ColumnCount">Column count.</param>
/// <param name="Cells">Cells in row-major order.</param>
public sealed record DocxTableInfo(
    DocxTargetId Id,
    string Story,
    int RowCount,
    int ColumnCount,
    IReadOnlyList<DocxTableCellInfo> Cells)
{
    /// <summary>Style ID, when styled.</summary>
    public string? StyleId { get; init; }
    /// <summary>Caption, when present.</summary>
    public string? Caption { get; init; }
    /// <summary>Description, when present.</summary>
    public string? Description { get; init; }
    /// <summary>Grid columns.</summary>
    public int? GridColumnCount { get; init; }
    /// <summary>Whether a header row is present.</summary>
    public bool HasHeaderRow { get; init; }
    /// <summary>Whether merged cells are present.</summary>
    public bool HasMergedCells { get; init; }
    /// <summary>Whether nested tables are present.</summary>
    public bool HasNestedTables { get; init; }
    /// <summary>Rows in order.</summary>
    public IReadOnlyList<DocxTableRowInfo> Rows { get; init; } = [];
}

/// <summary>
/// One table row.
/// </summary>
public sealed record DocxTableRowInfo
{
    /// <summary>Stable row ID.</summary>
    public required DocxTargetId Id { get; init; }
    /// <summary>1-based row ordinal.</summary>
    public int RowIndex { get; init; }
    /// <summary>Cell count.</summary>
    public int CellCount { get; init; }
    /// <summary>Leading grid columns.</summary>
    public int GridBefore { get; init; }
    /// <summary>Trailing grid columns.</summary>
    public int GridAfter { get; init; }
    /// <summary>Whether the row is a header row.</summary>
    public bool IsHeader { get; init; }
    /// <summary>Whether the row cannot split across pages.</summary>
    public bool CantSplit { get; init; }
}

/// <summary>
/// One table cell.
/// </summary>
/// <param name="Id">Stable cell ID.</param>
/// <param name="RowIndex">1-based row ordinal.</param>
/// <param name="ColumnIndex">1-based visual column ordinal.</param>
/// <param name="Text">Cell text.</param>
/// <param name="ColumnSpan">Spanned columns.</param>
/// <param name="VerticalMerge">Vertical-merge marker, when merged.</param>
/// <param name="HasNestedTable">Whether a nested table is present.</param>
public sealed record DocxTableCellInfo(
    DocxTargetId Id,
    int RowIndex,
    int ColumnIndex,
    string Text,
    int ColumnSpan,
    DocxVerticalMerge? VerticalMerge,
    bool HasNestedTable)
{
    /// <summary>1-based physical column ordinal.</summary>
    public int PhysicalColumnIndex { get; init; }
    /// <summary>1-based visual end column ordinal.</summary>
    public int VisualColumnEndIndex { get; init; }
    /// <summary>Merge-group ID, when merged.</summary>
    public DocxTargetId? MergeGroupId { get; init; }
    /// <summary>Merge-root cell ID, when merged.</summary>
    public DocxTargetId? VerticalMergeRootCellId { get; init; }
}

/// <summary>
/// One section with columns and orientation.
/// </summary>
/// <param name="Id">Stable section ID.</param>
/// <param name="Story">Story label.</param>
/// <param name="Columns">Column count.</param>
/// <param name="Orientation">Page orientation.</param>
public sealed record DocxSectionInfo(
    DocxTargetId Id,
    string Story,
    int Columns,
    DocxOrientation Orientation);
