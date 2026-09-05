// D11: single canonical patch-operation table.
// Parser field schemas (DocxPatchParser), help/catalog entries (DocxHelp), and
// engine dispatch (Execute) are all projections of OperationRegistrations below,
// so adding an operation means adding exactly one row. The compiler plus
// OperationRegistryTests enforce that the three views stay in agreement.
// This file is a second part of DocxPatchEngine (declared partial there) so the
// dispatch table can reference the private Execute* methods without widening
// their visibility; parser and catalog see only the internal projections.
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

/// <summary>Canonical registry row for one patch operation (see file remarks).</summary>
/// <param name="Name">Operation name as written in the patch file.</param>
/// <param name="AllowedFields">Every field the parser accepts for this operation.</param>
/// <param name="BooleanFields">Subset of <paramref name="AllowedFields"/> that must be <c>true</c>/<c>false</c>.</param>
/// <param name="IntegerFields">Subset of <paramref name="AllowedFields"/> that must be integers.</param>
/// <param name="Catalog">Help/catalog entry. <c>Name</c> must equal the row name (locked by test).</param>
/// <param name="Handler">Engine implementation. Narrower Execute* signatures are adapted with lambdas; shared implementations bind their extra arguments here.</param>
internal sealed record OperationRegistration(
    string Name,
    string[] AllowedFields,
    string[] BooleanFields,
    string[] IntegerFields,
    DocxPatchOperationInfo Catalog,
    PatchOperationHandler Handler);

/// <summary>Engine implementation of one patch operation.</summary>
internal delegate IReadOnlyList<DocxDiagnostic> PatchOperationHandler(
    OoxmlPackage package,
    DocxPatchOperation operation,
    DocxEditOptions options,
    bool apply,
    List<string> generatedRevisionIds,
    CancellationToken cancellationToken);

internal static partial class DocxPatchEngine
{
    private const string PreserveOnlyTrackChangesNote =
        "Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.";

    private const string TrackClassTextRun = "text-run";
    private const string TrackClassParagraphBlock = "paragraph-block";
    private const string TrackClassParagraphProperty = "paragraph-property";
    private const string TrackClassTableProperty = "table-property";
    private const string TrackClassRowProperty = "row-property";
    private const string TrackClassCellProperty = "cell-property";
    private const string TrackClassRowStructure = "row-structure";
    private const string TrackClassSectionProperty = "section-property";
    private const string TrackClassPreserveOnly = "preserve-only";
    private const string TrackClassUnsupported = "unsupported";

    private static readonly IReadOnlyList<OperationRegistration> OperationRegistrations =
    [
        new OperationRegistration(
            "replace-text",
            ["target", "expect-text", "find", "with", "preserve-runs", "occurrence"],
            ["preserve-runs"],
            ["occurrence"],
                    Tracked(
                        "replace-text",
                        ["target", "find", "with"],
                        TrackClassTextRun,
                        "tracked-simple",
                        "Suggest/Require emit tracked w:del/w:ins for supported simple text-only matches; unsupported shapes warn with W4002 or fail with E6002.",
                        ["expect-text", "preserve-runs", "occurrence"],
                        "Paragraphs And Blocks",
                        "Replaces matching text inside one target"),
            ExecuteReplaceText),
        new OperationRegistration(
            "replace-paragraph",
            ["target", "expect-text", "style", "text"],
            [],
            [],
                    Tracked(
                        "replace-paragraph",
                        ["target", "text"],
                        TrackClassParagraphBlock,
                        "tracked-paragraph",
                        "Suggest/Require emit whole-paragraph w:del/w:ins for simple text replacements and add w:pPrChange when a compatible style change is included; complex shapes warn with W4002 or fail with E6002.",
                        ["expect-text", "style"],
                        "Paragraphs And Blocks",
                        "Replaces the paragraph text, optionally setting style"),
            ExecuteReplaceParagraph),
        new OperationRegistration(
            "insert-before",
            ["target", "style", "copy-paragraph-properties", "text"],
            ["copy-paragraph-properties"],
            [],
                    Tracked(
                        "insert-before",
                        ["target", "text"],
                        TrackClassParagraphBlock,
                        "tracked-paragraph-insert",
                        "Suggest/Require emit inserted paragraph text as w:ins when the inserted text has no tabs or line breaks; unsupported shapes warn with W4002 or fail with E6002.",
                        ["style", "copy-paragraph-properties"],
                        "Paragraphs And Blocks",
                        "Inserts a paragraph/block before the target"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteInsertBlock(package, operation, options, insertAfter: false, apply, revisions, cancellationToken)),
        new OperationRegistration(
            "insert-after",
            ["target", "style", "copy-paragraph-properties", "text"],
            ["copy-paragraph-properties"],
            [],
                    Tracked(
                        "insert-after",
                        ["target", "text"],
                        TrackClassParagraphBlock,
                        "tracked-paragraph-insert",
                        "Suggest/Require emit inserted paragraph text as w:ins when the inserted text has no tabs or line breaks; unsupported shapes warn with W4002 or fail with E6002.",
                        ["style", "copy-paragraph-properties"],
                        "Paragraphs And Blocks",
                        "Inserts a paragraph/block after the target"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteInsertBlock(package, operation, options, insertAfter: true, apply, revisions, cancellationToken)),
        new OperationRegistration(
            "delete-block",
            ["target", "expect-text"],
            [],
            [],
                    Tracked(
                        "delete-block",
                        ["target"],
                        TrackClassParagraphBlock,
                        "tracked-paragraph-delete",
                        "Suggest/Require emit deleted paragraph text as w:del for simple paragraph targets; table/block or complex shapes warn with W4002 or fail with E6002.",
                        ["expect-text"],
                        "Paragraphs And Blocks",
                        "Deletes the target block"),
            ExecuteDeleteBlock),
        new OperationRegistration(
            "set-style",
            ["target", "style"],
            [],
            [],
                    Tracked(
                        "set-style",
                        ["target", "style"],
                        TrackClassParagraphProperty,
                        "tracked-style",
                        "Suggest/Require emit paragraph property revisions with w:pPrChange.",
                        [],
                        "Paragraphs And Blocks",
                        "Sets paragraph style"),
            ExecuteSetStyle),
        new OperationRegistration(
            "set-content-control-text",
            ["target", "expect-text", "text"],
            [],
            [],
                    Tracked(
                        "set-content-control-text",
                        ["target", "text"],
                        TrackClassTextRun,
                        "tracked-content-control-text",
                        "Suggest/Require emit w:del/w:ins inside simple plain-text and guarded paragraph-only rich-text content controls while preserving wrappers, bindings, locks, and paragraph containers; complex content controls warn with W4002 or fail with E6002.",
                        ["expect-text"],
                        "Content Controls",
                        "Plain-text controls are supported; guarded simple rich-text controls are supported when safe; picture/group controls fail with kind-specific guidance"),
            ExecuteSetContentControlText),
        new OperationRegistration(
            "set-content-control-checkbox",
            ["target", "checked"],
            [],
            [],
                    PreserveOnly(
                        "set-content-control-checkbox",
                        ["target", "checked"],
                        "Checkbox content controls update state metadata, not a simple Word revision range.",
                        [],
                        "Content Controls",
                        "Updates checkbox state and displayed symbol"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetContentControlCheckbox(package, operation, apply, cancellationToken)),
        new OperationRegistration(
            "set-content-control-choice",
            ["target", "value", "display-text"],
            [],
            [],
                    PreserveOnly(
                        "set-content-control-choice",
                        ["target plus value or display-text"],
                        "Dropdown and combo-box content controls update list value metadata and display text together; generated revision markup is not modeled yet.",
                        [],
                        "Content Controls",
                        "Selects a dropdown/combo item"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetContentControlChoice(package, operation, apply, cancellationToken)),
        new OperationRegistration(
            "set-content-control-date",
            ["target", "value", "display-text"],
            [],
            [],
                    PreserveOnly(
                        "set-content-control-date",
                        ["target", "value"],
                        "Date content controls update date metadata and display text together; generated revision markup is not modeled yet.",
                        ["display-text"],
                        "Content Controls",
                        "Updates date value and visible text"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetContentControlDate(package, operation, apply, cancellationToken)),
        new OperationRegistration(
            "add-repeating-section-item",
            ["target", "source", "index", "text"],
            [],
            ["index"],
                    Unsupported(
                        "add-repeating-section-item",
                        ["target"],
                        "Repeating-section item insertion is not safely modeled yet; check/apply fails with E4315.",
                        ["source", "index", "text"],
                        "Content Controls",
                        "Recognized but fails with `E4315`"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteUnsupportedRepeatingSectionOperation(operation)),
        new OperationRegistration(
            "delete-repeating-section-item",
            ["target", "index"],
            [],
            ["index"],
                    Unsupported(
                        "delete-repeating-section-item",
                        ["target"],
                        "Repeating-section item deletion is not safely modeled yet; check/apply fails with E4315.",
                        ["index"],
                        "Content Controls",
                        "Recognized but fails with `E4315`"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteUnsupportedRepeatingSectionOperation(operation)),
        new OperationRegistration(
            "add-bookmark",
            ["target", "expect-text", "name"],
            [],
            [],
                    PreserveOnly(
                        "add-bookmark",
                        ["target", "name"],
                        "Bookmark creation adds anchor metadata; Word has no useful generated revision range for the bookmark markers.",
                        ["expect-text"],
                        "Bookmarks",
                        "Creates a guarded paragraph bookmark"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteAddBookmark(package, operation, apply, cancellationToken)),
        new OperationRegistration(
            "replace-bookmark-text",
            ["target", "text"],
            [],
            [],
                    Tracked(
                        "replace-bookmark-text",
                        ["target", "text"],
                        TrackClassTextRun,
                        "tracked-bookmark-text",
                        "Suggest/Require emit w:del/w:ins inside simple same-paragraph bookmark ranges while preserving bookmark markers; direct mode also supports guarded multi-paragraph and simple table-spanning text-slot replacements. Multi-paragraph/table-spanning tracked output or protected ranges warn with W4002 or fail with E6002.",
                        [],
                        "Bookmarks",
                        "Replaces a complete paragraph-bounded bookmark range; simple table-spanning ranges require one replacement line per visible text slot"),
            ExecuteReplaceBookmarkText),
        new OperationRegistration(
            "rename-bookmark",
            ["target", "name"],
            [],
            [],
                    PreserveOnly(
                        "rename-bookmark",
                        ["target", "name"],
                        "Bookmark rename changes anchor metadata; Word has no useful generated revision range for the name update.",
                        [],
                        "Bookmarks",
                        "Renames markers and same-story internal hyperlink anchors when unambiguous"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteRenameBookmark(package, operation, apply, cancellationToken)),
        new OperationRegistration(
            "delete-bookmark",
            ["target"],
            [],
            [],
                    PreserveOnly(
                        "delete-bookmark",
                        ["target"],
                        "Bookmark deletion removes anchor metadata; Word has no useful generated revision range for the marker removal.",
                        [],
                        "Bookmarks",
                        "Removes complete unreferenced bookmark markers, preserving content"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteDeleteBookmark(package, operation, apply, cancellationToken)),
        new OperationRegistration(
            "add-comment",
            ["target", "expect-text", "anchor-text", "text", "author", "initials", "date", "occurrence"],
            [],
            ["occurrence"],
                    PreserveOnly(
                        "add-comment",
                        ["target", "text"],
                        "Comments are already review markup, so adding a comment does not create an additional tracked edit. Optional anchor-text selects one normalized text span inside the target paragraph; use occurrence when the span is repeated.",
                        ["expect-text", "anchor-text", "occurrence", "author", "initials", "date"],
                        "Comments",
                        "Anchors a new comment to a modeled paragraph, or to one selected text span inside it"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteAddComment(package, operation, options, apply, cancellationToken)),
        new OperationRegistration(
            "set-comment-text",
            ["target", "text"],
            [],
            [],
                    Tracked(
                        "set-comment-text",
                        ["target", "text"],
                        TrackClassTextRun,
                        "tracked-comment-text",
                        "Suggest/Require emit w:del/w:ins inside simple paragraph-only comment bodies while preserving comment metadata; complex comment bodies warn with W4002 or fail with E6002.",
                        [],
                        "Comments",
                        "Replaces one comment body"),
            ExecuteSetCommentText),
        new OperationRegistration(
            "resolve-comment",
            ["target"],
            [],
            [],
                    PreserveOnly(
                        "resolve-comment",
                        ["target"],
                        "Comment resolution changes review metadata, not visible document text.",
                        [],
                        "Comments",
                        "Creates or updates modern resolution metadata for basic comments"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetCommentResolved(package, operation, resolved: true, apply, cancellationToken)),
        new OperationRegistration(
            "reopen-comment",
            ["target"],
            [],
            [],
                    PreserveOnly(
                        "reopen-comment",
                        ["target"],
                        "Comment reopening changes review metadata, not visible document text.",
                        [],
                        "Comments",
                        "Clears modern resolution metadata for basic comments"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetCommentResolved(package, operation, resolved: false, apply, cancellationToken)),
        new OperationRegistration(
            "delete-comment",
            ["target"],
            [],
            [],
                    PreserveOnly(
                        "delete-comment",
                        ["target"],
                        "Comment deletion removes review markup, not a separate generated tracked edit.",
                        [],
                        "Comments",
                        "Removes body, range/reference markers, and matching extension records"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteDeleteComment(package, operation, apply, cancellationToken)),
        new OperationRegistration(
            "add-comment-reply",
            ["target", "text", "author", "initials", "date"],
            [],
            [],
                    PreserveOnly(
                        "add-comment-reply",
                        ["target", "text"],
                        "Threaded comment replies are review metadata, so adding a reply does not create an additional tracked edit.",
                        ["author", "initials", "date"],
                        "Comments",
                        "Adds a modern threaded reply under a comment"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteAddCommentReply(package, operation, options, apply, cancellationToken)),
        new OperationRegistration(
            "delete-comment-reply",
            ["target"],
            [],
            [],
                    PreserveOnly(
                        "delete-comment-reply",
                        ["target"],
                        "Threaded comment reply deletion removes review metadata, not a separate generated tracked edit.",
                        [],
                        "Comments",
                        "Removes a leaf threaded reply"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteDeleteCommentReply(package, operation, apply, cancellationToken)),
        new OperationRegistration(
            "set-field-dirty",
            ["target", "dirty"],
            ["dirty"],
            [],
                    PreserveOnly(
                        "set-field-dirty",
                        ["target", "dirty"],
                        "Field dirty flags are field metadata and have no useful generated visible revision representation.",
                        [],
                        "Fields",
                        "`target` can be a field ID or `all`"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetFieldFlag(package, operation, "dirty", "dirty", apply, cancellationToken)),
        new OperationRegistration(
            "set-field-lock",
            ["target", "locked"],
            ["locked"],
            [],
                    PreserveOnly(
                        "set-field-lock",
                        ["target", "locked"],
                        "Field lock flags are field metadata and have no useful generated visible revision representation.",
                        [],
                        "Fields",
                        "`target` can be a field ID or `all`"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetFieldFlag(package, operation, "locked", "fldLock", apply, cancellationToken)),
        new OperationRegistration(
            "set-field-code",
            ["target", "expect-code", "code"],
            [],
            [],
                    PreserveOnly(
                        "set-field-code",
                        ["target", "code"],
                        "Field codes are instruction metadata; generated revisions for field instructions are not modeled yet.",
                        ["expect-code"],
                        "Fields",
                        "Simple `w:fldSimple` fields only"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetFieldCode(package, operation, apply, cancellationToken)),
        new OperationRegistration(
            "set-field-result",
            ["target", "expect-result", "text"],
            [],
            [],
                    Tracked(
                        "set-field-result",
                        ["target", "text"],
                        TrackClassTextRun,
                        "tracked-field-result",
                        "Suggest/Require emit w:del/w:ins inside simple w:fldSimple cached result text while preserving the field instruction; direct mode also supports simple same-paragraph complex field result runs. Complex-field tracked output or unsafe topologies warn with W4002 or fail with E6002/E4313.",
                        ["expect-result"],
                        "Fields",
                        "Simple `w:fldSimple` cached result or validated simple same-paragraph complex result"),
            ExecuteSetFieldResult),
        new OperationRegistration(
            "refresh-field-result",
            ["target", "expect-code", "expect-result"],
            [],
            [],
                    PreserveOnly(
                        "refresh-field-result",
                        ["target"],
                        "Field refresh updates cached result text from modeled document state for REF/PAGEREF/NOTEREF bookmark fields and QUOTE literal fields; unsupported refresh types return categorized E4313 diagnostics. Generated revision markup for the refresh is not modeled yet.",
                        ["expect-code", "expect-result"],
                        "Fields",
                        "Limited refresh for simple REF/PAGEREF/NOTEREF bookmark fields and simple QUOTE literal fields"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteRefreshFieldResult(package, operation, apply, cancellationToken)),
        new OperationRegistration(
            "set-hyperlink-target",
            ["target", "uri", "anchor", "tooltip", "target-frame", "history"],
            ["history"],
            [],
                    PreserveOnly(
                        "set-hyperlink-target",
                        ["target plus uri or anchor"],
                        "Hyperlink target updates modify relationship or anchor metadata, not visible text.",
                        ["tooltip", "target-frame", "history"],
                        "Hyperlinks",
                        "Updates external URI or internal bookmark anchor"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetHyperlinkTarget(package, operation, apply, cancellationToken)),
        new OperationRegistration(
            "set-hyperlink-text",
            ["target", "text"],
            [],
            [],
                    Tracked(
                        "set-hyperlink-text",
                        ["target", "text"],
                        TrackClassTextRun,
                        "tracked-hyperlink-text",
                        "Suggest/Require emit w:del/w:ins inside the hyperlink wrapper for simple display text while preserving the relationship or anchor; protected or complex hyperlink content warns with W4002 or fails with E6002.",
                        [],
                        "Hyperlinks",
                        "Updates visible hyperlink text"),
            ExecuteSetHyperlinkText),
        new OperationRegistration(
            "insert-hyperlink-after",
            ["target", "text", "uri", "anchor", "tooltip", "target-frame", "history"],
            ["history"],
            [],
                    Tracked(
                        "insert-hyperlink-after",
                        ["target", "text plus uri or anchor"],
                        TrackClassTextRun,
                        "tracked-hyperlink-insert",
                        "Suggest/Require emit the inserted hyperlink display text as w:ins inside the hyperlink wrapper while preserving relationship or anchor metadata; text with tabs or line breaks warns with W4002 or fails with E6002.",
                        ["tooltip", "target-frame", "history"],
                        "Hyperlinks",
                        "Inserts a new hyperlink paragraph after the target"),
            ExecuteInsertHyperlinkAfter),
        new OperationRegistration(
            "remove-hyperlink",
            ["target"],
            [],
            [],
                    PreserveOnly(
                        "remove-hyperlink",
                        ["target"],
                        "Hyperlink removal changes wrapper and relationship metadata while preserving display text.",
                        [],
                        "Hyperlinks",
                        "Removes hyperlink markup and preserves display runs"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteRemoveHyperlink(package, operation, apply, cancellationToken)),
        new OperationRegistration(
            "set-cell",
            ["target", "expect-text", "expect-row-count", "expect-column-count", "text", "force"],
            ["force"],
            ["expect-row-count", "expect-column-count"],
                    Tracked(
                        "set-cell",
                        ["target", "text"],
                        TrackClassTextRun,
                        "tracked-cell-simple",
                        "Targets can be visual-grid cell IDs or merge-group IDs. Suggest/Require emit w:del/w:ins for simple text-only cells, including compatible multi-paragraph and horizontally merged cells; vertical-merge continuations, force, or complex cells warn with W4002 or fail with E6002.",
                        ["expect-text", "expect-row-count", "expect-column-count", "force"],
                        "Tables",
                        "Replaces one modeled cell by visual cell ID or merge-group ID"),
            ExecuteSetCell),
        new OperationRegistration(
            "set-cell-shading",
            ["target", "expect-fill", "fill", "clear"],
            ["clear"],
            [],
                    Tracked(
                        "set-cell-shading",
                        ["target plus fill or clear true"],
                        TrackClassCellProperty,
                        "tracked-cell-shading",
                        "Sets or clears w:tcPr/w:shd fill on a visual-grid cell ID or merge-group ID. Suggest/Require emit cell property revisions with w:tcPrChange while preserving previous cell properties; vertical-merge continuations fail with E4301.",
                        ["expect-fill", "clear"],
                        "Tables",
                        "Sets or clears `w:tcPr/w:shd` fill"),
            ExecuteSetCellShading),
        new OperationRegistration(
            "set-table-style",
            ["target", "expect-style", "style"],
            [],
            [],
                    Tracked(
                        "set-table-style",
                        ["target", "style"],
                        TrackClassTableProperty,
                        "tracked-table-style",
                        "Suggest/Require emit table property revisions with w:tblPrChange while preserving previous table properties.",
                        ["expect-style"],
                        "Tables",
                        "Updates `w:tblStyle`"),
            ExecuteSetTableStyle),
        new OperationRegistration(
            "set-table-metadata",
            ["target", "expect-caption", "expect-description", "caption", "description"],
            [],
            [],
                    PreserveOnly(
                        "set-table-metadata",
                        ["target plus caption or description"],
                        "Table caption and description updates are table metadata, not visible document text.",
                        ["expect-caption", "expect-description"],
                        "Tables",
                        "Sets or clears table caption/description"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetTableMetadata(package, operation, apply, cancellationToken)),
        new OperationRegistration(
            "set-row-header",
            ["target", "expect-header", "header"],
            ["expect-header", "header"],
            [],
                    Tracked(
                        "set-row-header",
                        ["target", "header"],
                        TrackClassRowProperty,
                        "tracked-row-header",
                        "Suggest/Require emit row property revisions with w:trPrChange while preserving previous row properties.",
                        ["expect-header"],
                        "Tables",
                        "Sets or clears the repeating-header flag"),
            ExecuteSetRowHeader),
        new OperationRegistration(
            "append-row",
            ["target", "expect-row-count", "expect-column-count", "cell"],
            [],
            ["expect-row-count", "expect-column-count"],
                    Tracked(
                        "append-row",
                        ["target, repeated cell"],
                        TrackClassRowStructure,
                        "tracked-row-insert",
                        "Direct mode appends by cloning the last row shape when the table has a consistent visual grid and the last row does not contain vertical merge cells. Suggest/Require emit row insertion revisions with w:trPr/w:ins for simple rectangular tables; visual-grid or other complex shapes warn with W4002 or fail with E6002.",
                        ["expect-row-count", "expect-column-count"],
                        "Tables",
                        "Appends by cloning the last row shape when the visual grid is consistent"),
            ExecuteAppendRow),
        new OperationRegistration(
            "insert-row-before",
            ["target", "expect-row-count", "expect-column-count", "expect-cell-count", "cell", "force"],
            ["force"],
            ["expect-row-count", "expect-column-count", "expect-cell-count"],
                    Tracked(
                        "insert-row-before",
                        ["target, repeated cell"],
                        TrackClassRowStructure,
                        "tracked-row-insert",
                        "Direct mode clones the target row shape for consistent visual-grid tables when the insertion boundary does not cross an active vertical merge chain. Suggest/Require emit row insertion revisions with w:trPr/w:ins for simple rectangular tables; force, visual-grid, or other complex shapes warn with W4002 or fail with E6002.",
                        ["expect-row-count", "expect-column-count", "expect-cell-count", "force"],
                        "Tables",
                        "Inserts before a row by cloning the target row shape when safe"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteInsertRow(package, operation, options, insertAfter: false, apply, revisions, cancellationToken)),
        new OperationRegistration(
            "insert-row-after",
            ["target", "expect-row-count", "expect-column-count", "expect-cell-count", "cell", "force"],
            ["force"],
            ["expect-row-count", "expect-column-count", "expect-cell-count"],
                    Tracked(
                        "insert-row-after",
                        ["target, repeated cell"],
                        TrackClassRowStructure,
                        "tracked-row-insert",
                        "Direct mode clones the target row shape for consistent visual-grid tables when the insertion boundary does not cross an active vertical merge chain. Suggest/Require emit row insertion revisions with w:trPr/w:ins for simple rectangular tables; force, visual-grid, or other complex shapes warn with W4002 or fail with E6002.",
                        ["expect-row-count", "expect-column-count", "expect-cell-count", "force"],
                        "Tables",
                        "Inserts after a row by cloning the target row shape when safe"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteInsertRow(package, operation, options, insertAfter: true, apply, revisions, cancellationToken)),
        new OperationRegistration(
            "delete-row",
            ["target", "expect-row-count", "expect-column-count", "expect-cell-count", "expect-contains", "force"],
            ["force"],
            ["expect-row-count", "expect-column-count", "expect-cell-count"],
                    Tracked(
                        "delete-row",
                        ["target"],
                        TrackClassRowStructure,
                        "tracked-row-delete",
                        "Direct mode deletes rows in consistent visual-grid tables and promotes the next vertical-merge continuation when deleting a merge root. Suggest/Require emit row deletion revisions with w:trPr/w:del for simple rectangular tables; force, visual-grid, or other complex shapes warn with W4002 or fail with E6002.",
                        ["expect-row-count", "expect-column-count", "expect-cell-count", "expect-contains", "force"],
                        "Tables",
                        "Deletes a row; direct mode can promote the next vertical-merge continuation"),
            ExecuteDeleteRow),
        new OperationRegistration(
            "append-column",
            ["target", "expect-row-count", "expect-column-count", "cell", "force"],
            ["force"],
            ["expect-row-count", "expect-column-count"],
                    Unsupported(
                        "append-column",
                        ["target, repeated cell"],
                        "Table-column transforms are not safely modeled yet; check/apply fails with E4316.",
                        ["expect-row-count", "expect-column-count", "force"],
                        "Tables",
                        "Recognized but fails with `E4316`"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteUnsupportedColumnOperation(operation)),
        new OperationRegistration(
            "insert-column-before",
            ["target", "column", "expect-row-count", "expect-column-count", "expect-cell-count", "cell", "force"],
            ["force"],
            ["column", "expect-row-count", "expect-column-count", "expect-cell-count"],
                    Unsupported(
                        "insert-column-before",
                        ["target", "column, repeated cell"],
                        "Table-column transforms are not safely modeled yet; check/apply fails with E4316.",
                        ["expect-row-count", "expect-column-count", "expect-cell-count", "force"],
                        "Tables",
                        "Recognized but fails with `E4316`"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteUnsupportedColumnOperation(operation)),
        new OperationRegistration(
            "insert-column-after",
            ["target", "column", "expect-row-count", "expect-column-count", "expect-cell-count", "cell", "force"],
            ["force"],
            ["column", "expect-row-count", "expect-column-count", "expect-cell-count"],
                    Unsupported(
                        "insert-column-after",
                        ["target", "column, repeated cell"],
                        "Table-column transforms are not safely modeled yet; check/apply fails with E4316.",
                        ["expect-row-count", "expect-column-count", "expect-cell-count", "force"],
                        "Tables",
                        "Recognized but fails with `E4316`"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteUnsupportedColumnOperation(operation)),
        new OperationRegistration(
            "delete-column",
            ["target", "column", "expect-row-count", "expect-column-count", "expect-cell-count", "expect-contains", "force"],
            ["force"],
            ["column", "expect-row-count", "expect-column-count", "expect-cell-count"],
                    Unsupported(
                        "delete-column",
                        ["target", "column"],
                        "Table-column transforms are not safely modeled yet; check/apply fails with E4316.",
                        ["expect-row-count", "expect-column-count", "expect-cell-count", "expect-contains", "force"],
                        "Tables",
                        "Recognized but fails with `E4316`"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteUnsupportedColumnOperation(operation)),
        new OperationRegistration(
            "replace-image",
            ["target", "asset", "expect-content-type", "alt"],
            [],
            [],
                    PreserveOnly(
                        "replace-image",
                        ["target", "asset"],
                        "Image replacement updates DrawingML and package media; generated drawing-level revision markup is not modeled yet.",
                        ["expect-content-type", "alt"],
                        "Images",
                        "Replaces media bytes and preserves supported drawing layout"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteReplaceImage(package, operation, options, apply, cancellationToken)),
        new OperationRegistration(
            "insert-image-after",
            ["target", "asset", "expect-content-type", "width", "height", "alt"],
            [],
            [],
                    PreserveOnly(
                        "insert-image-after",
                        ["target", "asset"],
                        "Image insertion creates DrawingML and package media; generated drawing-level revision markup is not modeled yet.",
                        ["expect-content-type", "width", "height", "alt"],
                        "Images",
                        "Inserts an inline image paragraph after a paragraph target"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteInsertImageAfter(package, operation, options, apply, cancellationToken)),
        new OperationRegistration(
            "set-image-alt",
            ["target", "expect-content-type", "alt"],
            [],
            [],
                    PreserveOnly(
                        "set-image-alt",
                        ["target", "alt"],
                        "Image alt-text updates DrawingML metadata, not visible document text.",
                        ["expect-content-type"],
                        "Images",
                        "Updates DrawingML description"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetImageAlt(package, operation, apply, cancellationToken)),
        new OperationRegistration(
            "set-image-metadata",
            ["target", "expect-content-type", "alt", "title", "name"],
            [],
            [],
                    PreserveOnly(
                        "set-image-metadata",
                        ["target plus alt, title, or name"],
                        "Image title/name/alt updates DrawingML metadata, not visible document text.",
                        ["expect-content-type"],
                        "Images",
                        "Updates DrawingML `docPr` metadata"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetImageMetadata(package, operation, apply, cancellationToken)),
        new OperationRegistration(
            "set-image-size",
            ["target", "expect-content-type", "width", "height"],
            [],
            [],
                    PreserveOnly(
                        "set-image-size",
                        ["target plus width or height"],
                        "Image size updates DrawingML layout metadata, not visible document text.",
                        ["expect-content-type"],
                        "Images",
                        "Updates DrawingML extents"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetImageSize(package, operation, apply, cancellationToken)),
        new OperationRegistration(
            "set-image-wrap",
            ["target", "expect-content-type", "mode", "dist-top", "dist-bottom", "dist-left", "dist-right"],
            [],
            [],
                    PreserveOnly(
                        "set-image-wrap",
                        ["target plus mode or a distance field"],
                        "Image wrapping updates DrawingML layout metadata, not visible document text.",
                        ["expect-content-type"],
                        "Images",
                        "Anchored images only"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetImageWrap(package, operation, apply, cancellationToken)),
        new OperationRegistration(
            "set-image-position",
            ["target", "expect-content-type", "horizontal-relative", "horizontal-offset", "horizontal-align", "vertical-relative", "vertical-offset", "vertical-align"],
            [],
            [],
                    PreserveOnly(
                        "set-image-position",
                        ["target plus relative, offset, or align field"],
                        "Image position updates DrawingML layout metadata, not visible document text.",
                        ["expect-content-type"],
                        "Images",
                        "Anchored images only"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetImagePosition(package, operation, apply, cancellationToken)),
        new OperationRegistration(
            "set-image-crop",
            ["target", "expect-content-type", "left-percent", "top-percent", "right-percent", "bottom-percent"],
            [],
            [],
                    PreserveOnly(
                        "set-image-crop",
                        ["target plus one crop percentage"],
                        "Image crop updates DrawingML layout metadata, not visible document text.",
                        ["expect-content-type"],
                        "Images",
                        "Updates DrawingML crop percentages"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetImageCrop(package, operation, apply, cancellationToken)),
        new OperationRegistration(
            "delete-image",
            ["target", "expect-content-type"],
            [],
            [],
                    PreserveOnly(
                        "delete-image",
                        ["target"],
                        "Image deletion removes DrawingML and package media; generated drawing-level revision markup is not modeled yet.",
                        ["expect-content-type"],
                        "Images",
                        "Deletes the modeled image"),
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteDeleteImage(package, operation, apply, cancellationToken)),
        new OperationRegistration(
            "set-section-columns",
            ["target", "expect-columns", "expect-orientation", "count"],
            [],
            ["expect-columns", "count"],
                    Tracked(
                        "set-section-columns",
                        ["target", "count"],
                        TrackClassSectionProperty,
                        "tracked-section-columns",
                        "Suggest/Require emit section property revisions with w:sectPrChange while preserving previous section properties and references; existing section property revisions fall back or fail instead of being replaced.",
                        ["expect-columns", "expect-orientation"],
                        "Sections",
                        "Column count must be 1 through 4"),
            ExecuteSetSectionColumns),
        new OperationRegistration(
            "set-section-orientation",
            ["target", "expect-columns", "expect-orientation", "orientation"],
            [],
            ["expect-columns"],
                    Tracked(
                        "set-section-orientation",
                        ["target", "orientation"],
                        TrackClassSectionProperty,
                        "tracked-section-orientation",
                        "Suggest/Require emit section property revisions with w:sectPrChange while preserving previous page size, section properties, and references; existing section property revisions fall back or fail instead of being replaced.",
                        ["expect-columns", "expect-orientation"],
                        "Sections",
                        "`orientation` is `portrait` or `landscape`"),
            ExecuteSetSectionOrientation)
    ];
    private static readonly IReadOnlyDictionary<string, OperationRegistration> OperationsByName =
        OperationRegistrations.ToDictionary(registration => registration.Name, StringComparer.Ordinal);

    private static readonly HashSet<string> PreservesFieldFreshness = ["set-field-result", "refresh-field-result"];

    internal static IReadOnlyList<OperationRegistration> AllOperations => OperationRegistrations;

    internal static IReadOnlyList<DocxPatchOperationInfo> CatalogOperations { get; } =
        OperationRegistrations.Select(registration => registration.Catalog).ToArray();

    internal static bool IsKnownOperation(string operationName)
    {
        return OperationsByName.ContainsKey(operationName);
    }

    internal static bool TryGetPatchOperationHandler(string operationName, out PatchOperationHandler? handler)
    {
        if (OperationsByName.TryGetValue(operationName, out OperationRegistration? registration))
        {
            handler = registration.Handler;
            return true;
        }

        handler = null;
        return false;
    }

    internal static bool SupportsTrackedChangeOutput(string operationName)
    {
        return OperationsByName.TryGetValue(operationName, out OperationRegistration? registration) &&
            registration.Catalog.GeneratesTrackedChanges;
    }

    internal static string TrackChangesSupportValue(string operationName)
    {
        return OperationsByName.TryGetValue(operationName, out OperationRegistration? registration)
            ? registration.Catalog.TrackChangesSupport
            : "unclassified";
    }

    internal static bool MarksFieldsDirtyAfterEdit(string operationName)
    {
        return !PreservesFieldFreshness.Contains(operationName);
    }

    private static IReadOnlyList<DocxDiagnostic>? TryExecuteOperation(
        OoxmlPackage package,
        DocxPatchOperation operation,
        DocxEditOptions options,
        bool apply,
        List<string> generatedRevisionIds,
        CancellationToken cancellationToken)
    {
        return OperationsByName.TryGetValue(operation.OperationName, out OperationRegistration? registration)
            ? registration.Handler(package, operation, options, apply, generatedRevisionIds, cancellationToken)
            : null;
    }

    private static DocxPatchOperationInfo PreserveOnly(
        string name,
        IReadOnlyList<string> requiredFields,
        string rationale,
        IReadOnlyList<string> optionalFields,
        string category,
        string description)
    {
        return new DocxPatchOperationInfo
        {
            Name = name,
            RequiredFields = requiredFields,
            OptionalFields = optionalFields,
            TrackChangesSupportClass = TrackClassPreserveOnly,
            TrackChangesSupport = "preserve-only",
            TrackChangesNote = $"{rationale} {PreserveOnlyTrackChangesNote}",
            Category = category,
            Description = description
        };
    }

    private static DocxPatchOperationInfo Tracked(
        string name,
        IReadOnlyList<string> requiredFields,
        string supportClass,
        string support,
        string note,
        IReadOnlyList<string> optionalFields,
        string category,
        string description)
    {
        return new DocxPatchOperationInfo
        {
            Name = name,
            RequiredFields = requiredFields,
            OptionalFields = optionalFields,
            TrackChangesSupportClass = supportClass,
            TrackChangesSupport = support,
            TrackChangesNote = note,
            Category = category,
            Description = description
        };
    }

    private static DocxPatchOperationInfo Unsupported(
        string name,
        IReadOnlyList<string> requiredFields,
        string note,
        IReadOnlyList<string> optionalFields,
        string category,
        string description)
    {
        return new DocxPatchOperationInfo
        {
            Name = name,
            RequiredFields = requiredFields,
            OptionalFields = optionalFields,
            TrackChangesSupportClass = TrackClassUnsupported,
            TrackChangesSupport = "unsupported",
            TrackChangesNote = note,
            Category = category,
            Description = description
        };
    }
}
