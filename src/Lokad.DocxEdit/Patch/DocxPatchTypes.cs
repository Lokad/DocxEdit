using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Lokad.DocxEdit.Model;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

internal sealed record PatchExecutionResult(
    bool Success,
    IReadOnlyList<DocxDiagnostic> Diagnostics,
    IReadOnlyList<DocxPatchOperationReport> Reports);

internal sealed record ContentControlTarget(string PartName, XDocument Document, XElement ContentControl);

internal sealed record ContentControlChoice(string DisplayText);

internal sealed record BookmarkTarget(string PartName, XDocument Document, XElement Start, XElement? End);

internal sealed record TrackedBookmarkReplacement(string DeletedText, XElement? RunProperties);

internal sealed record BookmarkTextSlot(XElement Container, XNode[] Nodes, XElement? AddAfter, XElement? AddBefore);

internal sealed record CommentsPartTarget(string PartName, XDocument Document, XElement Root);

internal sealed record CommentsExtendedPartTarget(string PartName, XDocument Document, XElement Root);

internal sealed record CommentsIdsPartTarget(string PartName, XDocument Document, XElement Root);

internal sealed record CommentTarget(string PartName, XDocument Document, XElement Comment);

internal sealed record CommentReplyTarget(string PartName, XDocument Document, XElement Comment, string ParaId, string ParentParaId);

internal sealed record CommentExtensionTarget(string PartName, XDocument Document, XElement CommentExtension);

internal sealed record HyperlinkTarget(string PartName, XDocument Document, XElement Hyperlink);

internal sealed record FieldTarget(string PartName, XDocument Document, XElement Element);

internal sealed record ImageBlipTarget(string PartName, XDocument Document, XElement Blip, string RelationshipId, OoxmlPart Part);

internal readonly record struct ImageCrop(int Left, int Top, int Right, int Bottom);

internal readonly record struct ImagePositionAxis(bool HasAny, string? RelativeFrom, long? OffsetEmus, string? Align);

internal sealed record ParagraphTarget(string PartName, XDocument Document, XElement Paragraph);

internal sealed record BlockTarget(string PartName, XDocument Document, XElement Block);

internal sealed record TableTarget(string PartName, XDocument Document, XElement Table);

internal sealed record RowTarget(string PartName, XDocument Document, XElement Table, XElement Row);

internal sealed record CellTarget(string PartName, XDocument Document, XElement Table, XElement Row, XElement Cell, int VisualColumnIndex);

internal sealed record TableOperationSnapshot(
    DocxTargetId ResolvedTarget,
    int? RowIndex,
    int? ColumnIndex,
    int RowCountBefore,
    int ColumnCount,
    int? CellCount,
    int? GridBefore,
    int? GridAfter,
    IReadOnlyList<TableCellSnapshot> Cells);

internal sealed record TableCellSnapshot(
    int ColumnIndex,
    int VisualColumnEndIndex,
    string? MergeGroupId,
    string? NestedTablePath);

internal sealed record TableCellGridSlot(XElement Cell, int ColumnIndex, int ColumnSpan);

internal sealed record SectionTarget(XDocument Document, XElement SectionProperties);

internal sealed record MergeGroupRootState(XElement Row, XElement Cell, int VisualColumnIndex);

internal readonly record struct TextRange(int Start, int Length);

internal sealed record TextPosition(XElement TextElement, int Offset);

/// <summary>
/// Parsed form of a patch <c>target</c> field. Implementors describe <i>which</i>
/// paragraphs a target can resolve to; resolution itself lives in
/// <c>ResolveMainParagraphElementBySelector</c> and the explicit-ID parsers.
/// </summary>
/// <remarks>
/// <para>
/// <c>Raw</c> is always the original target text and is used verbatim in
/// diagnostics (<c>E1201</c> no match, <c>E1202</c> ambiguous match,
/// <c>E1203</c> malformed selector), so it must survive parsing unchanged.
/// </para>
/// <para>
/// Predicate selectors (<c>Heading</c>, <c>ParagraphText</c>, <c>Bookmark</c>,
/// <c>ContentControl</c>) only ever match direct <c>w:p</c> children of the
/// <i>main</i> document story and require exactly one match: zero matches fail
/// with <c>E1201</c> (with up to 3 suggestions), two or more fail with
/// <c>E1202</c> (listing the candidate IDs). All text/name comparisons are
/// ordinal and case-sensitive. Add a new subtype only together with its branch
/// in <c>TryParseTargetSelector</c> and in
/// <c>ResolveMainParagraphElementBySelector</c>; the compiler does not enforce
/// the pairing, so cover it with a parser/resolver round-trip test.
/// </para>
/// </remarks>
/// <param name="Raw">Original target text, preserved for diagnostics.</param>
internal abstract record TargetSelector(string Raw);

/// <summary>
/// Catch-all for targets without a <c>prefix:</c> selector: a stable explicit ID
/// (<c>M.P0001</c>, <c>H001.P0002</c>, table/cell/image/section/comment IDs, …).
/// Never handled by the predicate path — resolution matches on <c>TargetId</c> instead.
/// </summary>
/// <c>TargetId</c> carries the parsed structure when <c>Raw</c> matches the explicit-ID
/// grammar, null otherwise; resolution choke points match on it instead of re-parsing <c>Raw</c>.
internal sealed record ExplicitIdTargetSelector(string Raw, DocxTargetId? TargetId) : TargetSelector(Raw);

/// <summary>
/// <c>heading:"Text"</c> or <c>heading:L:"Text"</c>: matches a main-story paragraph
/// whose heading level equals <c>Level</c> (any heading level when null) and whose
/// visible text equals <c>Text</c> ordinally. <c>Level</c> is validated to 1–9 at
/// parse time (<c>E1203</c> otherwise).
/// </summary>
internal sealed record HeadingTargetSelector(string Raw, int? Level, string Text) : TargetSelector(Raw);

/// <summary>
/// <c>text:"..."</c>: matches a main-story paragraph whose visible text <i>contains</i>
/// <c>Text</c> as an ordinal substring. Prefer an explicit ID when the text is not unique.
/// </summary>
internal sealed record ParagraphTextTargetSelector(string Raw, string Text) : TargetSelector(Raw);

/// <summary>
/// <c>bookmark:"Name"</c>: matches a main-story paragraph containing a descendant
/// <c>bookmarkStart</c> whose name equals <c>Name</c> ordinally.
/// </summary>
internal sealed record BookmarkTargetSelector(string Raw, string Name) : TargetSelector(Raw);

/// <summary>
/// <c>content-control:"TagOrAlias"</c>: matches a main-story paragraph containing a
/// descendant <c>sdt</c> whose tag or alias equals <c>Name</c> ordinally.
/// </summary>
internal sealed record ContentControlTargetSelector(string Raw, string Name) : TargetSelector(Raw);

internal sealed record ParagraphSelectorMatch(string Id, XElement Paragraph);
