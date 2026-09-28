// Single canonical patch-operation table.
// Parser field checks (DocxPatchParser), help/catalog entries (DocxHelp), and
// engine dispatch (Execute) are all projections of the field tables in
// OperationRegistrations below: adding an operation means writing its Execute*
// implementation and adding the one row that names its fields, catalog entry,
// and dispatch binding. OperationRegistryTests keeps the views in agreement:
// every catalog name is parser-accepted, every parser-accepted field is
// catalogued, and every source position survives through FieldValues.
// This file is a second part of DocxPatchEngine (declared partial there) so the
// dispatch table can reference the private Execute* methods without widening
// their visibility; parser and catalog see only the internal projections.
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

/// <summary>Value shape of one patch field.</summary>
internal enum FieldValueKind
{
    /// <summary>Free text (heredoc-capable).</summary>
    Text,
    /// <summary>Must be true or false.</summary>
    Boolean,
    /// <summary>Must be an integer.</summary>
    Integer
}

/// <summary>One accepted patch field: its name, value shape, repetition, and requiredness.</summary>
/// <param name="Name">Field name as written in the patch file.</param>
/// <param name="Kind">Value shape enforced by the parser.</param>
/// <param name="Repeatable">Whether the field may repeat; every occurrence is kept in file order.</param>
/// <param name="Required">Whether the field is required on every invocation.</param>
/// <param name="AllowEmpty">Whether a present-but-empty value is accepted (for example text deletion); absence still fails requiredness.</param>
internal sealed record OperationFieldDefinition(
    string Name,
    FieldValueKind Kind,
    bool Repeatable,
    bool Required,
    bool AllowEmpty = false);

/// <summary>Canonical registry row for one patch operation (see file remarks).</summary>
/// <param name="Name">Operation name as written in the patch file.</param>
/// <param name="Fields">Every accepted field, with value shape, repetition, and requiredness.</param>
/// <param name="RequireOneOf">Field groups of which at least one member is required; value rules stay in operation code.</param>
/// <param name="TrackSupportClass">Machine-readable track-change class.</param>
/// <param name="TrackSupport">Human track-change support label.</param>
/// <param name="TrackNote">Suggest/Require behavior note.</param>
/// <param name="Category">Editing area group.</param>
/// <param name="Description">What the operation does.</param>
/// <param name="Handler">Engine implementation. Narrower Execute* signatures are adapted with lambdas; shared implementations bind their extra arguments here.</param>
internal sealed record OperationRegistration(
    string Name,
    OperationFieldDefinition[] Fields,
    string[][] RequireOneOf,
    string TrackSupportClass,
    string TrackSupport,
    string TrackNote,
    string Category,
    string Description,
    PatchOperationHandler Handler)
{
    /// <summary>Whether the operation is intrinsic review or annotation markup that is permitted under Require without generated revisions.</summary>
    public bool IsAnnotation { get; init; }
    /// <summary>Field groups of which at most one member may carry a value; execution rejects combinations with E4205 or E4202.</summary>
    public string[][] ExclusiveGroups { get; init; } = [];
    /// <summary>Executable example patches, minimal first and guarded second; empty when the operation has no curated example.</summary>
    public string[] Examples { get; init; } = [];
}

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
        Tracked(
            "replace-text",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-text", FieldValueKind.Text, Repeatable: false, Required: false),
                new("find", FieldValueKind.Text, Repeatable: false, Required: true),
                new("with", FieldValueKind.Text, Repeatable: false, Required: true, AllowEmpty: true),
                new("preserve-runs", FieldValueKind.Boolean, Repeatable: false, Required: false),
                new("occurrence", FieldValueKind.Integer, Repeatable: false, Required: false),
            ],
            [],
            TrackClassTextRun,
            "tracked-simple",
            "Suggest/Require emit tracked w:del/w:ins for supported simple text-only matches; unsupported shapes warn with W4002 or fail with E6002.",
            "Paragraphs And Blocks",
            "Replaces matching text inside one target; an ambiguous find without occurrence fails, occurrence N selects one match, occurrence all replaces every match",
            ExecuteReplaceText) with
        {
            Examples =
            [
                """
                # Minimal replace-text.
                docxpatch 1

                op replace-text
                target M.P0001
                find Alpha
                with Omega
                end
                """,
                """
                # Guarded replace-text: expect-text must match before editing.
                docxpatch 1

                op replace-text
                target M.P0001
                expect-text <<<
                Alpha Beta
                >>>
                find Alpha
                with Omega
                end
                """,
            ],
        },
        Tracked(
            "replace-paragraph",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-text", FieldValueKind.Text, Repeatable: false, Required: false),
                new("style", FieldValueKind.Text, Repeatable: false, Required: false),
                new("text", FieldValueKind.Text, Repeatable: false, Required: true, AllowEmpty: true),
            ],
            [],
            TrackClassParagraphBlock,
            "tracked-paragraph",
            "Suggest/Require emit whole-paragraph w:del/w:ins for simple text replacements and add w:pPrChange when a compatible style change is included; complex shapes warn with W4002 or fail with E6002.",
            "Paragraphs And Blocks",
            "Replaces the paragraph text, optionally setting style",
            ExecuteReplaceParagraph) with
        {
            Examples =
            [
                """
                # Minimal replace-paragraph.
                docxpatch 1

                op replace-paragraph
                target M.P0001
                text <<<
                New paragraph text
                >>>
                end
                """,
                """
                # Guarded replace-paragraph: expect-text must match before editing.
                docxpatch 1

                op replace-paragraph
                target M.P0001
                expect-text <<<
                Old paragraph text
                >>>
                text <<<
                New paragraph text
                >>>
                end
                """,
            ],
        },
        Tracked(
            "insert-before",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("style", FieldValueKind.Text, Repeatable: false, Required: false),
                new("copy-paragraph-properties", FieldValueKind.Boolean, Repeatable: false, Required: false),
                new("text", FieldValueKind.Text, Repeatable: false, Required: true),
            ],
            [],
            TrackClassParagraphBlock,
            "tracked-paragraph-insert",
            "Suggest/Require emit inserted paragraph text as w:ins when the inserted text has no tabs or line breaks; unsupported shapes warn with W4002 or fail with E6002.",
            "Paragraphs And Blocks",
            "Inserts a paragraph/block before the target",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteInsertBlock(package, operation, options, insertAfter: false, apply, revisions, cancellationToken)) with
        {
            Examples =
            [
                """
                # Minimal insert-before.
                docxpatch 1

                op insert-before
                target M.P0001
                text <<<
                Inserted paragraph
                >>>
                end
                """,
                """
                # Insert-before with style.
                docxpatch 1

                op insert-before
                target M.P0001
                style Normal
                text <<<
                Inserted paragraph
                >>>
                end
                """,
            ],
        },
        Tracked(
            "insert-after",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("style", FieldValueKind.Text, Repeatable: false, Required: false),
                new("copy-paragraph-properties", FieldValueKind.Boolean, Repeatable: false, Required: false),
                new("text", FieldValueKind.Text, Repeatable: false, Required: true),
            ],
            [],
            TrackClassParagraphBlock,
            "tracked-paragraph-insert",
            "Suggest/Require emit inserted paragraph text as w:ins when the inserted text has no tabs or line breaks; unsupported shapes warn with W4002 or fail with E6002.",
            "Paragraphs And Blocks",
            "Inserts a paragraph/block after the target",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteInsertBlock(package, operation, options, insertAfter: true, apply, revisions, cancellationToken)) with
        {
            Examples =
            [
                """
                # Minimal insert-after.
                docxpatch 1

                op insert-after
                target M.P0001
                text <<<
                Inserted paragraph
                >>>
                end
                """,
                """
                # Insert-after with style.
                docxpatch 1

                op insert-after
                target M.P0001
                style Normal
                text <<<
                Inserted paragraph
                >>>
                end
                """,
            ],
        },
        Tracked(
            "delete-block",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-text", FieldValueKind.Text, Repeatable: false, Required: false),
            ],
            [],
            TrackClassParagraphBlock,
            "tracked-paragraph-delete",
            "Suggest/Require emit deleted paragraph text as w:del for simple paragraph targets; table/block or complex shapes warn with W4002 or fail with E6002.",
            "Paragraphs And Blocks",
            "Deletes the target block",
            ExecuteDeleteBlock) with
        {
            Examples =
            [
                """
                # Minimal delete-block.
                docxpatch 1

                op delete-block
                target M.P0001
                end
                """,
                """
                # Guarded delete-block: expect-text must match before deleting.
                docxpatch 1

                op delete-block
                target M.P0001
                expect-text <<<
                Remove this paragraph
                >>>
                end
                """,
            ],
        },
        Tracked(
            "set-style",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("style", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-style", FieldValueKind.Text, Repeatable: false, Required: false),
            ],
            [],
            TrackClassParagraphProperty,
            "tracked-style",
            "Suggest/Require emit paragraph property revisions with w:pPrChange.",
            "Paragraphs And Blocks",
            "Sets paragraph style",
            ExecuteSetStyle) with
        {
            Examples =
            [
                """
                # Minimal set-style.
                docxpatch 1

                op set-style
                target M.P0001
                style Normal
                end
                """,
                """
                # Guarded set-style: expect-style must match before changing.
                docxpatch 1

                op set-style
                target M.P0001
                style Heading 2
                expect-style Normal
                end
                """,
            ],
        },
        Tracked(
            "set-content-control-text",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-text", FieldValueKind.Text, Repeatable: false, Required: false),
                new("text", FieldValueKind.Text, Repeatable: false, Required: true),
            ],
            [],
            TrackClassTextRun,
            "tracked-content-control-text",
            "Suggest/Require emit w:del/w:ins inside simple plain-text and guarded paragraph-only rich-text content controls while preserving wrappers, bindings, locks, and paragraph containers; complex content controls warn with W4002 or fail with E6002.",
            "Content Controls",
            "Plain-text controls are supported; guarded simple rich-text controls are supported when safe; picture/group controls fail with kind-specific guidance",
            ExecuteSetContentControlText) with
        {
            Examples =
            [
                """
                # Minimal set-content-control-text.
                docxpatch 1

                op set-content-control-text
                target M.CC0001
                text <<<
                New control text
                >>>
                end
                """,
                """
                # Guarded set-content-control-text: expect-text must match before editing.
                docxpatch 1

                op set-content-control-text
                target M.CC0001
                expect-text <<<
                Old control text
                >>>
                text <<<
                New control text
                >>>
                end
                """,
            ],
        },
        PreserveOnly(
            "set-content-control-checkbox",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("checked", FieldValueKind.Boolean, Repeatable: false, Required: true),
            ],
            [],
            "Checkbox content controls update state metadata, not a simple Word revision range.",
            "Content Controls",
            "Updates checkbox state and displayed symbol",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetContentControlCheckbox(package, operation, apply, cancellationToken)),
        PreserveOnly(
            "set-content-control-choice",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("value", FieldValueKind.Text, Repeatable: false, Required: false),
                new("display-text", FieldValueKind.Text, Repeatable: false, Required: false),
            ],
            [["value", "display-text"]],
            "Dropdown and combo-box content controls update list value metadata and display text together; generated revision markup is not modeled yet.",
            "Content Controls",
            "Selects a dropdown/combo item",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetContentControlChoice(package, operation, apply, cancellationToken)) with { ExclusiveGroups = [["value", "display-text"]] },
        PreserveOnly(
            "set-content-control-date",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("value", FieldValueKind.Text, Repeatable: false, Required: true),
                new("display-text", FieldValueKind.Text, Repeatable: false, Required: false),
            ],
            [],
            "Date content controls update date metadata and display text together; generated revision markup is not modeled yet.",
            "Content Controls",
            "Updates date value and visible text",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetContentControlDate(package, operation, apply, cancellationToken)),
        Unsupported(
            "add-repeating-section-item",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("source", FieldValueKind.Text, Repeatable: false, Required: false),
                new("index", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("text", FieldValueKind.Text, Repeatable: false, Required: false),
            ],
            [],
            "Repeating-section item insertion is not safely modeled yet; check/apply fails with E4315.",
            "Content Controls",
            "Recognized but fails with `E4315`",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteUnsupportedRepeatingSectionOperation(operation)),
        Unsupported(
            "delete-repeating-section-item",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("index", FieldValueKind.Integer, Repeatable: false, Required: false),
            ],
            [],
            "Repeating-section item deletion is not safely modeled yet; check/apply fails with E4315.",
            "Content Controls",
            "Recognized but fails with `E4315`",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteUnsupportedRepeatingSectionOperation(operation)),
        PreserveOnly(
            "add-bookmark",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-text", FieldValueKind.Text, Repeatable: false, Required: false),
                new("name", FieldValueKind.Text, Repeatable: false, Required: true),
            ],
            [],
            "Bookmark creation adds anchor metadata; Word has no useful generated revision range for the bookmark markers.",
            "Bookmarks",
            "Creates a guarded paragraph bookmark",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteAddBookmark(package, operation, apply, cancellationToken)) with
        {
            Examples =
            [
                """
                # Minimal add-bookmark.
                docxpatch 1

                op add-bookmark
                target M.P0001
                name MarkName
                end
                """,
                """
                # Guarded add-bookmark: expect-text must match before editing.
                docxpatch 1

                op add-bookmark
                target M.P0001
                expect-text <<<
                Marked paragraph
                >>>
                name MarkName
                end
                """,
            ],
        },
        Tracked(
            "replace-bookmark-text",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("text", FieldValueKind.Text, Repeatable: false, Required: true),
            ],
            [],
            TrackClassTextRun,
            "tracked-bookmark-text",
            "Suggest/Require emit w:del/w:ins inside simple same-paragraph bookmark ranges while preserving bookmark markers; direct mode also supports guarded multi-paragraph and simple table-spanning text-slot replacements. Multi-paragraph/table-spanning tracked output or protected ranges warn with W4002 or fail with E6002.",
            "Bookmarks",
            "Replaces a complete paragraph-bounded bookmark range; simple table-spanning ranges require one replacement line per visible text slot",
            ExecuteReplaceBookmarkText),
        PreserveOnly(
            "rename-bookmark",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("name", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-name", FieldValueKind.Text, Repeatable: false, Required: false),
            ],
            [],
            "Bookmark rename changes anchor metadata; Word has no useful generated revision range for the name update.",
            "Bookmarks",
            "Renames markers and same-story internal hyperlink anchors when unambiguous",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteRenameBookmark(package, operation, apply, cancellationToken)),
        PreserveOnly(
            "delete-bookmark",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
            ],
            [],
            "Bookmark deletion removes anchor metadata; Word has no useful generated revision range for the marker removal.",
            "Bookmarks",
            "Removes complete unreferenced bookmark markers, preserving content",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteDeleteBookmark(package, operation, apply, cancellationToken)),
        PreserveOnly(
            "add-comment",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-text", FieldValueKind.Text, Repeatable: false, Required: false),
                new("anchor-text", FieldValueKind.Text, Repeatable: false, Required: false),
                new("text", FieldValueKind.Text, Repeatable: false, Required: true),
                new("author", FieldValueKind.Text, Repeatable: false, Required: false),
                new("initials", FieldValueKind.Text, Repeatable: false, Required: false),
                new("date", FieldValueKind.Text, Repeatable: false, Required: false),
                new("occurrence", FieldValueKind.Integer, Repeatable: false, Required: false),
            ],
            [],
            "Comments are already review markup, so adding a comment does not create an additional tracked edit. Optional anchor-text selects one normalized text span inside the target paragraph; use occurrence when the span is repeated.",
            "Comments",
            "Anchors a new comment to a modeled paragraph, or to one selected text span inside it",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteAddComment(package, operation, options, apply, cancellationToken), isAnnotation: true) with
        {
            Examples =
            [
                """
                # Minimal add-comment.
                docxpatch 1

                op add-comment
                target M.P0001
                text <<<
                Review note
                >>>
                end
                """,
                """
                # Guarded add-comment: expect-text must match before editing.
                docxpatch 1

                op add-comment
                target M.P0001
                expect-text <<<
                Reviewed paragraph
                >>>
                text <<<
                Review note
                >>>
                end
                """,
            ],
        },
        Tracked(
            "set-comment-text",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("text", FieldValueKind.Text, Repeatable: false, Required: true),
            ],
            [],
            TrackClassTextRun,
            "tracked-comment-text",
            "Suggest/Require emit w:del/w:ins inside simple paragraph-only comment bodies while preserving comment metadata; complex comment bodies warn with W4002 or fail with E6002.",
            "Comments",
            "Replaces one comment body",
            ExecuteSetCommentText),
        PreserveOnly(
            "resolve-comment",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
            ],
            [],
            "Comment resolution changes review metadata, not visible document text.",
            "Comments",
            "Creates or updates modern resolution metadata for basic comments",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetCommentResolved(package, operation, resolved: true, apply, cancellationToken), isAnnotation: true),
        PreserveOnly(
            "reopen-comment",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
            ],
            [],
            "Comment reopening changes review metadata, not visible document text.",
            "Comments",
            "Clears modern resolution metadata for basic comments",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetCommentResolved(package, operation, resolved: false, apply, cancellationToken), isAnnotation: true),
        PreserveOnly(
            "delete-comment",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
            ],
            [],
            "Comment deletion removes review markup, not a separate generated tracked edit.",
            "Comments",
            "Removes body, range/reference markers, and matching extension records",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteDeleteComment(package, operation, apply, cancellationToken), isAnnotation: true),
        PreserveOnly(
            "add-comment-reply",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("text", FieldValueKind.Text, Repeatable: false, Required: true),
                new("author", FieldValueKind.Text, Repeatable: false, Required: false),
                new("initials", FieldValueKind.Text, Repeatable: false, Required: false),
                new("date", FieldValueKind.Text, Repeatable: false, Required: false),
            ],
            [],
            "Threaded comment replies are review metadata, so adding a reply does not create an additional tracked edit.",
            "Comments",
            "Adds a modern threaded reply under a comment",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteAddCommentReply(package, operation, options, apply, cancellationToken), isAnnotation: true),
        PreserveOnly(
            "delete-comment-reply",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
            ],
            [],
            "Threaded comment reply deletion removes review metadata, not a separate generated tracked edit.",
            "Comments",
            "Removes a leaf threaded reply",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteDeleteCommentReply(package, operation, apply, cancellationToken), isAnnotation: true),
        PreserveOnly(
            "set-field-dirty",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("dirty", FieldValueKind.Boolean, Repeatable: false, Required: true),
            ],
            [],
            "Field dirty flags are field metadata and have no useful generated visible revision representation.",
            "Fields",
            "`target` can be a field ID or `all`",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetFieldFlag(package, operation, "dirty", "dirty", apply, cancellationToken)),
        PreserveOnly(
            "set-field-lock",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("locked", FieldValueKind.Boolean, Repeatable: false, Required: true),
            ],
            [],
            "Field lock flags are field metadata and have no useful generated visible revision representation.",
            "Fields",
            "`target` can be a field ID or `all`",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetFieldFlag(package, operation, "locked", "fldLock", apply, cancellationToken)),
        PreserveOnly(
            "set-field-code",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-code", FieldValueKind.Text, Repeatable: false, Required: false),
                new("code", FieldValueKind.Text, Repeatable: false, Required: true),
            ],
            [],
            "Field codes are instruction metadata; generated revisions for field instructions are not modeled yet.",
            "Fields",
            "Simple `w:fldSimple` fields only",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetFieldCode(package, operation, apply, cancellationToken)),
        Tracked(
            "set-field-result",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-result", FieldValueKind.Text, Repeatable: false, Required: false),
                new("text", FieldValueKind.Text, Repeatable: false, Required: true),
            ],
            [],
            TrackClassTextRun,
            "tracked-field-result",
            "Suggest/Require emit w:del/w:ins inside simple w:fldSimple cached result text while preserving the field instruction; direct mode also supports simple same-paragraph complex field result runs. Complex-field tracked output or unsafe topologies warn with W4002 or fail with E6002/E4313.",
            "Fields",
            "Simple `w:fldSimple` cached result or validated simple same-paragraph complex result",
            ExecuteSetFieldResult),
        PreserveOnly(
            "refresh-field-result",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-code", FieldValueKind.Text, Repeatable: false, Required: false),
                new("expect-result", FieldValueKind.Text, Repeatable: false, Required: false),
            ],
            [],
            "Field refresh updates cached result text from modeled document state for REF/PAGEREF/NOTEREF bookmark fields and QUOTE literal fields; unsupported refresh types return categorized E4313 diagnostics. Generated revision markup for the refresh is not modeled yet.",
            "Fields",
            "Limited refresh for simple REF/PAGEREF/NOTEREF bookmark fields and simple QUOTE literal fields",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteRefreshFieldResult(package, operation, apply, cancellationToken)),
        PreserveOnly(
            "set-hyperlink-target",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("uri", FieldValueKind.Text, Repeatable: false, Required: false),
                new("anchor", FieldValueKind.Text, Repeatable: false, Required: false),
                new("tooltip", FieldValueKind.Text, Repeatable: false, Required: false),
                new("target-frame", FieldValueKind.Text, Repeatable: false, Required: false),
                new("history", FieldValueKind.Boolean, Repeatable: false, Required: false),
            ],
            [["uri", "anchor"]],
            "Hyperlink target updates modify relationship or anchor metadata, not visible text.",
            "Hyperlinks",
            "Updates external URI or internal bookmark anchor",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetHyperlinkTarget(package, operation, apply, cancellationToken)) with { ExclusiveGroups = [["uri", "anchor"]] },
        Tracked(
            "set-hyperlink-text",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("text", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-text", FieldValueKind.Text, Repeatable: false, Required: false),
            ],
            [],
            TrackClassTextRun,
            "tracked-hyperlink-text",
            "Suggest/Require emit w:del/w:ins inside the hyperlink wrapper for simple display text while preserving the relationship or anchor; protected or complex hyperlink content warns with W4002 or fails with E6002.",
            "Hyperlinks",
            "Updates visible hyperlink text",
            ExecuteSetHyperlinkText),
        Tracked(
            "insert-hyperlink-after",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("text", FieldValueKind.Text, Repeatable: false, Required: true),
                new("uri", FieldValueKind.Text, Repeatable: false, Required: false),
                new("anchor", FieldValueKind.Text, Repeatable: false, Required: false),
                new("tooltip", FieldValueKind.Text, Repeatable: false, Required: false),
                new("target-frame", FieldValueKind.Text, Repeatable: false, Required: false),
                new("history", FieldValueKind.Boolean, Repeatable: false, Required: false),
            ],
            [["uri", "anchor"]],
            TrackClassTextRun,
            "tracked-hyperlink-insert",
            "Suggest/Require emit the inserted hyperlink display text as w:ins inside the hyperlink wrapper while preserving relationship or anchor metadata; text with tabs or line breaks warns with W4002 or fails with E6002.",
            "Hyperlinks",
            "Inserts a new hyperlink paragraph after the target",
            ExecuteInsertHyperlinkAfter) with { ExclusiveGroups = [["uri", "anchor"]] },
        PreserveOnly(
            "remove-hyperlink",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
            ],
            [],
            "Hyperlink removal changes wrapper and relationship metadata while preserving display text.",
            "Hyperlinks",
            "Removes hyperlink markup and preserves display runs",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteRemoveHyperlink(package, operation, apply, cancellationToken)),
        Tracked(
            "set-cell",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-text", FieldValueKind.Text, Repeatable: false, Required: false),
                new("expect-row-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("expect-column-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("text", FieldValueKind.Text, Repeatable: false, Required: true, AllowEmpty: true),
                new("force", FieldValueKind.Boolean, Repeatable: false, Required: false),
            ],
            [],
            TrackClassTextRun,
            "tracked-cell-simple",
            "Targets can be visual-grid cell IDs or merge-group IDs. Suggest/Require emit w:del/w:ins for simple text-only cells, including compatible multi-paragraph and horizontally merged cells; vertical-merge continuations, force, or complex cells warn with W4002 or fail with E6002.",
            "Tables",
            "Replaces one modeled cell by visual cell ID or merge-group ID",
            ExecuteSetCell) with
        {
            Examples =
            [
                """
                # Minimal set-cell.
                docxpatch 1

                op set-cell
                target M.T0001.R01.C01
                text <<<
                New cell text
                >>>
                end
                """,
                """
                # Guarded set-cell: expect-text must match before editing.
                docxpatch 1

                op set-cell
                target M.T0001.R01.C01
                expect-text <<<
                Old cell text
                >>>
                text <<<
                New cell text
                >>>
                end
                """,
            ],
        },
        Tracked(
            "set-cell-shading",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-fill", FieldValueKind.Text, Repeatable: false, Required: false),
                new("fill", FieldValueKind.Text, Repeatable: false, Required: false),
                new("clear", FieldValueKind.Boolean, Repeatable: false, Required: false),
            ],
            [["fill", "clear"]],
            TrackClassCellProperty,
            "tracked-cell-shading",
            "Sets or clears w:tcPr/w:shd fill on a visual-grid cell ID or merge-group ID. Suggest/Require emit cell property revisions with w:tcPrChange while preserving previous cell properties; vertical-merge continuations fail with E4301.",
            "Tables",
            "Sets or clears `w:tcPr/w:shd` fill",
            ExecuteSetCellShading) with
        {
            ExclusiveGroups = [["fill", "clear"]],
            Examples =
            [
                """
                # Minimal set-cell-shading.
                docxpatch 1

                op set-cell-shading
                target M.T0001.R01.C01
                fill 4472C4
                end
                """,
                """
                # Guarded set-cell-shading: expect-fill must match before editing.
                docxpatch 1

                op set-cell-shading
                target M.T0001.R01.C01
                expect-fill auto
                fill 4472C4
                end
                """,
            ],
        },
        Tracked(
            "set-table-style",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-style", FieldValueKind.Text, Repeatable: false, Required: false),
                new("style", FieldValueKind.Text, Repeatable: false, Required: true),
            ],
            [],
            TrackClassTableProperty,
            "tracked-table-style",
            "Suggest/Require emit table property revisions with w:tblPrChange while preserving previous table properties.",
            "Tables",
            "Updates `w:tblStyle`",
            ExecuteSetTableStyle),
        PreserveOnly(
            "set-table-metadata",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-caption", FieldValueKind.Text, Repeatable: false, Required: false),
                new("expect-description", FieldValueKind.Text, Repeatable: false, Required: false),
                new("caption", FieldValueKind.Text, Repeatable: false, Required: false),
                new("description", FieldValueKind.Text, Repeatable: false, Required: false),
            ],
            [["caption", "description"]],
            "Table caption and description updates are table metadata, not visible document text.",
            "Tables",
            "Sets or clears table caption/description",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetTableMetadata(package, operation, apply, cancellationToken)),
        Tracked(
            "set-row-header",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-header", FieldValueKind.Boolean, Repeatable: false, Required: false),
                new("header", FieldValueKind.Boolean, Repeatable: false, Required: true),
            ],
            [],
            TrackClassRowProperty,
            "tracked-row-header",
            "Suggest/Require emit row property revisions with w:trPrChange while preserving previous row properties.",
            "Tables",
            "Sets or clears the repeating-header flag",
            ExecuteSetRowHeader),
        Tracked(
            "append-row",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-row-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("expect-column-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("cell", FieldValueKind.Text, Repeatable: true, Required: true),
            ],
            [],
            TrackClassRowStructure,
            "tracked-row-insert",
            "Direct mode appends by cloning the last row shape when the table has a consistent visual grid and the last row does not contain vertical merge cells. Suggest/Require emit row insertion revisions with w:trPr/w:ins for simple rectangular tables; visual-grid or other complex shapes warn with W4002 or fail with E6002.",
            "Tables",
            "Appends by cloning the last row shape when the visual grid is consistent",
            ExecuteAppendRow),
        Tracked(
            "insert-row-before",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-row-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("expect-column-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("expect-cell-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("cell", FieldValueKind.Text, Repeatable: true, Required: true),
                new("force", FieldValueKind.Boolean, Repeatable: false, Required: false),
            ],
            [],
            TrackClassRowStructure,
            "tracked-row-insert",
            "Direct mode clones the target row shape for consistent visual-grid tables when the insertion boundary does not cross an active vertical merge chain. Suggest/Require emit row insertion revisions with w:trPr/w:ins for simple rectangular tables; force, visual-grid, or other complex shapes warn with W4002 or fail with E6002.",
            "Tables",
            "Inserts before a row by cloning the target row shape when safe",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteInsertRow(package, operation, options, insertAfter: false, apply, revisions, cancellationToken)),
        Tracked(
            "insert-row-after",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-row-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("expect-column-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("expect-cell-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("cell", FieldValueKind.Text, Repeatable: true, Required: true),
                new("force", FieldValueKind.Boolean, Repeatable: false, Required: false),
            ],
            [],
            TrackClassRowStructure,
            "tracked-row-insert",
            "Direct mode clones the target row shape for consistent visual-grid tables when the insertion boundary does not cross an active vertical merge chain. Suggest/Require emit row insertion revisions with w:trPr/w:ins for simple rectangular tables; force, visual-grid, or other complex shapes warn with W4002 or fail with E6002.",
            "Tables",
            "Inserts after a row by cloning the target row shape when safe",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteInsertRow(package, operation, options, insertAfter: true, apply, revisions, cancellationToken)),
        Tracked(
            "delete-row",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-row-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("expect-column-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("expect-cell-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("expect-contains", FieldValueKind.Text, Repeatable: false, Required: false),
                new("force", FieldValueKind.Boolean, Repeatable: false, Required: false),
            ],
            [],
            TrackClassRowStructure,
            "tracked-row-delete",
            "Direct mode deletes rows in consistent visual-grid tables and promotes the next vertical-merge continuation when deleting a merge root. Suggest/Require emit row deletion revisions with w:trPr/w:del for simple rectangular tables; force, visual-grid, or other complex shapes warn with W4002 or fail with E6002.",
            "Tables",
            "Deletes a row; direct mode can promote the next vertical-merge continuation",
            ExecuteDeleteRow),
        Unsupported(
            "append-column",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-row-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("expect-column-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("cell", FieldValueKind.Text, Repeatable: true, Required: true),
                new("force", FieldValueKind.Boolean, Repeatable: false, Required: false),
            ],
            [],
            "Table-column transforms are not safely modeled yet; check/apply fails with E4316.",
            "Tables",
            "Recognized but fails with `E4316`",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteUnsupportedColumnOperation(operation)),
        Unsupported(
            "insert-column-before",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("column", FieldValueKind.Integer, Repeatable: false, Required: true),
                new("expect-row-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("expect-column-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("expect-cell-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("cell", FieldValueKind.Text, Repeatable: true, Required: true),
                new("force", FieldValueKind.Boolean, Repeatable: false, Required: false),
            ],
            [],
            "Table-column transforms are not safely modeled yet; check/apply fails with E4316.",
            "Tables",
            "Recognized but fails with `E4316`",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteUnsupportedColumnOperation(operation)),
        Unsupported(
            "insert-column-after",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("column", FieldValueKind.Integer, Repeatable: false, Required: true),
                new("expect-row-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("expect-column-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("expect-cell-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("cell", FieldValueKind.Text, Repeatable: true, Required: true),
                new("force", FieldValueKind.Boolean, Repeatable: false, Required: false),
            ],
            [],
            "Table-column transforms are not safely modeled yet; check/apply fails with E4316.",
            "Tables",
            "Recognized but fails with `E4316`",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteUnsupportedColumnOperation(operation)),
        Unsupported(
            "delete-column",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("column", FieldValueKind.Integer, Repeatable: false, Required: true),
                new("expect-row-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("expect-column-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("expect-cell-count", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("expect-contains", FieldValueKind.Text, Repeatable: false, Required: false),
                new("force", FieldValueKind.Boolean, Repeatable: false, Required: false),
            ],
            [],
            "Table-column transforms are not safely modeled yet; check/apply fails with E4316.",
            "Tables",
            "Recognized but fails with `E4316`",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteUnsupportedColumnOperation(operation)),
        PreserveOnly(
            "replace-image",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("asset", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-content-type", FieldValueKind.Text, Repeatable: false, Required: false),
                new("alt", FieldValueKind.Text, Repeatable: false, Required: false),
            ],
            [],
            "Image replacement updates DrawingML and package media; generated drawing-level revision markup is not modeled yet.",
            "Images",
            "Replaces media bytes and preserves supported drawing layout",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteReplaceImage(package, operation, options, apply, cancellationToken)),
        PreserveOnly(
            "insert-image-after",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("asset", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-content-type", FieldValueKind.Text, Repeatable: false, Required: false),
                new("width", FieldValueKind.Text, Repeatable: false, Required: false),
                new("height", FieldValueKind.Text, Repeatable: false, Required: false),
                new("alt", FieldValueKind.Text, Repeatable: false, Required: false),
            ],
            [],
            "Image insertion creates DrawingML and package media; generated drawing-level revision markup is not modeled yet.",
            "Images",
            "Inserts an inline image paragraph after a paragraph target",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteInsertImageAfter(package, operation, options, apply, cancellationToken)),
        PreserveOnly(
            "set-image-alt",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-content-type", FieldValueKind.Text, Repeatable: false, Required: false),
                new("alt", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-alt", FieldValueKind.Text, Repeatable: false, Required: false),
            ],
            [],
            "Image alt-text updates DrawingML metadata, not visible document text.",
            "Images",
            "Updates DrawingML description",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetImageAlt(package, operation, apply, cancellationToken)),
        PreserveOnly(
            "set-image-metadata",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-content-type", FieldValueKind.Text, Repeatable: false, Required: false),
                new("alt", FieldValueKind.Text, Repeatable: false, Required: false),
                new("title", FieldValueKind.Text, Repeatable: false, Required: false),
                new("name", FieldValueKind.Text, Repeatable: false, Required: false),
                new("expect-alt", FieldValueKind.Text, Repeatable: false, Required: false),
                new("expect-title", FieldValueKind.Text, Repeatable: false, Required: false),
                new("expect-name", FieldValueKind.Text, Repeatable: false, Required: false),
            ],
            [["alt", "title", "name"]],
            "Image title/name/alt updates DrawingML metadata, not visible document text.",
            "Images",
            "Updates DrawingML `docPr` metadata",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetImageMetadata(package, operation, apply, cancellationToken)),
        PreserveOnly(
            "set-image-size",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-content-type", FieldValueKind.Text, Repeatable: false, Required: false),
                new("width", FieldValueKind.Text, Repeatable: false, Required: false),
                new("height", FieldValueKind.Text, Repeatable: false, Required: false),
            ],
            [["width", "height"]],
            "Image size updates DrawingML layout metadata, not visible document text.",
            "Images",
            "Updates DrawingML extents",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetImageSize(package, operation, apply, cancellationToken)),
        PreserveOnly(
            "set-image-wrap",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-content-type", FieldValueKind.Text, Repeatable: false, Required: false),
                new("mode", FieldValueKind.Text, Repeatable: false, Required: false),
                new("dist-top", FieldValueKind.Text, Repeatable: false, Required: false),
                new("dist-bottom", FieldValueKind.Text, Repeatable: false, Required: false),
                new("dist-left", FieldValueKind.Text, Repeatable: false, Required: false),
                new("dist-right", FieldValueKind.Text, Repeatable: false, Required: false),
            ],
            [["mode", "dist-top", "dist-bottom", "dist-left", "dist-right"]],
            "Image wrapping updates DrawingML layout metadata, not visible document text.",
            "Images",
            "Anchored images only",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetImageWrap(package, operation, apply, cancellationToken)),
        PreserveOnly(
            "set-image-position",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-content-type", FieldValueKind.Text, Repeatable: false, Required: false),
                new("horizontal-relative", FieldValueKind.Text, Repeatable: false, Required: false),
                new("horizontal-offset", FieldValueKind.Text, Repeatable: false, Required: false),
                new("horizontal-align", FieldValueKind.Text, Repeatable: false, Required: false),
                new("vertical-relative", FieldValueKind.Text, Repeatable: false, Required: false),
                new("vertical-offset", FieldValueKind.Text, Repeatable: false, Required: false),
                new("vertical-align", FieldValueKind.Text, Repeatable: false, Required: false),
            ],
            [["horizontal-relative", "horizontal-offset", "horizontal-align", "vertical-relative", "vertical-offset", "vertical-align"]],
            "Image position updates DrawingML layout metadata, not visible document text.",
            "Images",
            "Anchored images only",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetImagePosition(package, operation, apply, cancellationToken)),
        PreserveOnly(
            "set-image-crop",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-content-type", FieldValueKind.Text, Repeatable: false, Required: false),
                new("left-percent", FieldValueKind.Text, Repeatable: false, Required: false),
                new("top-percent", FieldValueKind.Text, Repeatable: false, Required: false),
                new("right-percent", FieldValueKind.Text, Repeatable: false, Required: false),
                new("bottom-percent", FieldValueKind.Text, Repeatable: false, Required: false),
            ],
            [["left-percent", "top-percent", "right-percent", "bottom-percent"]],
            "Image crop updates DrawingML layout metadata, not visible document text.",
            "Images",
            "Updates DrawingML crop percentages",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteSetImageCrop(package, operation, apply, cancellationToken)),
        PreserveOnly(
            "delete-image",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-content-type", FieldValueKind.Text, Repeatable: false, Required: false),
            ],
            [],
            "Image deletion removes DrawingML and package media; generated drawing-level revision markup is not modeled yet.",
            "Images",
            "Deletes the modeled image",
            (package, operation, options, apply, revisions, cancellationToken) => ExecuteDeleteImage(package, operation, apply, cancellationToken)),
        Tracked(
            "set-section-columns",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-columns", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("expect-orientation", FieldValueKind.Text, Repeatable: false, Required: false),
                new("count", FieldValueKind.Integer, Repeatable: false, Required: true),
            ],
            [],
            TrackClassSectionProperty,
            "tracked-section-columns",
            "Suggest/Require emit section property revisions with w:sectPrChange while preserving previous section properties and references; existing section property revisions fall back or fail instead of being replaced.",
            "Sections",
            "Column count must be 1 through 4",
            ExecuteSetSectionColumns),
        Tracked(
            "set-section-orientation",
            [
                new("target", FieldValueKind.Text, Repeatable: false, Required: true),
                new("expect-columns", FieldValueKind.Integer, Repeatable: false, Required: false),
                new("expect-orientation", FieldValueKind.Text, Repeatable: false, Required: false),
                new("orientation", FieldValueKind.Text, Repeatable: false, Required: true),
            ],
            [],
            TrackClassSectionProperty,
            "tracked-section-orientation",
            "Suggest/Require emit section property revisions with w:sectPrChange while preserving previous page size, section properties, and references; existing section property revisions fall back or fail instead of being replaced.",
            "Sections",
            "`orientation` is `portrait` or `landscape`",
            ExecuteSetSectionOrientation)
    ];
    private static readonly IReadOnlyDictionary<string, OperationRegistration> OperationsByName =
        OperationRegistrations.ToDictionary(registration => registration.Name, StringComparer.Ordinal);


    // Operations that allocate wordprocessingML w:id values outside revision allocation.
    // If a future operation allocates w:id values without going through AllocateRevisionIds,
    // add its name here so later revision allocations rescan.
    private static readonly HashSet<string> WordIdAllocatingOperations = ["add-comment", "add-comment-reply", "add-bookmark"];
    private static readonly HashSet<string> PreservesFieldFreshness = ["set-field-result", "refresh-field-result"];

    internal static IReadOnlyList<OperationRegistration> AllOperations => OperationRegistrations;

    internal static IReadOnlyList<DocxPatchOperationInfo> CatalogOperations { get; } =
        OperationRegistrations.Select(BuildCatalogEntry).ToArray();

    private static DocxPatchOperationInfo BuildCatalogEntry(OperationRegistration registration)
    {
        var grouped = new HashSet<string>(registration.RequireOneOf.SelectMany(static group => group), StringComparer.Ordinal);
        return new DocxPatchOperationInfo
        {
            Name = registration.Name,
            RequiredFields = registration.Fields.Where(static field => field.Required).Select(static field => field.Name).ToArray(),
            RequiredAlternatives = registration.RequireOneOf.Select(static group => (IReadOnlyList<string>)group.ToArray()).ToArray(),
            ExclusiveAlternatives = registration.ExclusiveGroups.Select(static group => (IReadOnlyList<string>)group.ToArray()).ToArray(),
            Examples = registration.Examples.ToArray(),
            OptionalFields = registration.Fields.Where(field => !field.Required && !grouped.Contains(field.Name)).Select(static field => field.Name).ToArray(),
            RepeatableFields = registration.Fields.Where(static field => field.Repeatable).Select(static field => field.Name).ToArray(),
            BooleanFields = registration.Fields.Where(static field => field.Kind == FieldValueKind.Boolean).Select(static field => field.Name).ToArray(),
            IntegerFields = registration.Fields.Where(static field => field.Kind == FieldValueKind.Integer).Select(static field => field.Name).ToArray(),
            EmptyAllowedFields = registration.Fields.Where(static field => field.AllowEmpty).Select(static field => field.Name).ToArray(),
            TrackChangesSupportClass = registration.TrackSupportClass,
            TrackChangesSupport = registration.TrackSupport,
            TrackChangesNote = registration.TrackNote,
            Category = registration.Category,
            Description = registration.Description
        };
    }

    internal static bool IsAnnotationOperation(string operationName)
    {
        return OperationsByName.TryGetValue(operationName, out OperationRegistration? registration) &&
            registration.IsAnnotation;
    }

    internal static bool SupportsTrackedChangeOutput(string operationName)
    {
        return OperationsByName.TryGetValue(operationName, out OperationRegistration? registration) &&
            registration.TrackSupportClass is not TrackClassPreserveOnly and not TrackClassUnsupported;
    }

    internal static string TrackChangesSupportValue(string operationName)
    {
        return OperationsByName.TryGetValue(operationName, out OperationRegistration? registration)
            ? registration.TrackSupport
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

    private static OperationRegistration PreserveOnly(
        string name,
        OperationFieldDefinition[] fields,
        string[][] requireOneOf,
        string rationale,
        string category,
        string description,
        PatchOperationHandler handler,
        bool isAnnotation = false)
    {
        return new OperationRegistration(
            name,
            fields,
            requireOneOf,
            TrackClassPreserveOnly,
            "preserve-only",
            $"{rationale} {PreserveOnlyTrackChangesNote}" + (isAnnotation ? " Declared annotation operations stay permitted under Require." : ""),
            category,
            description,
            handler)
        {
            IsAnnotation = isAnnotation
        };
    }

    private static OperationRegistration Tracked(
        string name,
        OperationFieldDefinition[] fields,
        string[][] requireOneOf,
        string supportClass,
        string support,
        string note,
        string category,
        string description,
        PatchOperationHandler handler)
    {
        return new OperationRegistration(
            name,
            fields,
            requireOneOf,
            supportClass,
            support,
            note,
            category,
            description,
            handler);
    }

    private static OperationRegistration Unsupported(
        string name,
        OperationFieldDefinition[] fields,
        string[][] requireOneOf,
        string note,
        string category,
        string description,
        PatchOperationHandler handler)
    {
        return new OperationRegistration(
            name,
            fields,
            requireOneOf,
            TrackClassUnsupported,
            "unsupported",
            note,
            category,
            description,
            handler);
    }
}
