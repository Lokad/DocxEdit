using System.Text;

namespace Lokad.DocxEdit;

public sealed record DocxCommandCatalog
{
    public string ToolName { get; init; } = "docxedit";
    public string Summary { get; init; } = "read and patch .docx documents";
    public IReadOnlyList<DocxCommandInfo> Commands { get; init; } = [];
    public IReadOnlyList<string> Examples { get; init; } = [];
    public IReadOnlyList<DocxPatchOperationInfo> PatchOperations { get; init; } = [];
}

public sealed record DocxCommandInfo
{
    public string Name { get; init; } = string.Empty;
    public string Category { get; init; } = string.Empty;
    public string Summary { get; init; } = string.Empty;
    public string Usage { get; init; } = string.Empty;
    public string Description { get; init; } = string.Empty;
    public IReadOnlyList<DocxOptionInfo> Options { get; init; } = [];
    public IReadOnlyList<DocxOutputFieldInfo> OutputFields { get; init; } = [];
    public IReadOnlyList<string> PrivacyNotes { get; init; } = [];
    public IReadOnlyList<DocxHelpSection> Notes { get; init; } = [];
    public IReadOnlyList<string> Examples { get; init; } = [];
}

public sealed record DocxOptionInfo(string Syntax, string Description);

public sealed record DocxOutputFieldInfo(string Name, string Description);

public sealed record DocxHelpSection(string Heading, IReadOnlyList<string> Lines);

public sealed record DocxPatchOperationInfo
{
    public string Name { get; init; } = string.Empty;
    public IReadOnlyList<string> RequiredFields { get; init; } = [];
    public IReadOnlyList<string> OptionalFields { get; init; } = [];
    public string Description { get; init; } = string.Empty;
    public string TrackChangesSupportClass { get; init; } = "unsupported";
    public string TrackChangesSupport { get; init; } = "unsupported";
    public string TrackChangesNote { get; init; } = "Suggest applies directly with W4001; Require fails with E6001.";

    public bool GeneratesTrackedChanges =>
        TrackChangesSupportClass is not "preserve-only" and not "unsupported";

    public string RenderSummary()
    {
        string required = RequiredFields.Count == 0 ? string.Empty : string.Join("/", RequiredFields);
        string optional = OptionalFields.Count == 0 ? string.Empty : $", optional {string.Join(", ", OptionalFields)}";
        return string.IsNullOrWhiteSpace(required)
            ? $"{Name}{optional}"
            : $"{Name} {required}{optional}";
    }
}

public static class DocxHelp
{
    private const string PreserveOnlyTrackChangesNote =
        "Existing tracked-change markup is preserved, but this operation does not create new revision markup; Suggest applies directly with W4001 and Require fails with E6001.";

    private const string TrackClassTextRun = "text-run";
    private const string TrackClassParagraphBlock = "paragraph-block";
    private const string TrackClassParagraphProperty = "paragraph-property";
    private const string TrackClassPreserveOnly = "preserve-only";
    private const string TrackClassUnsupported = "unsupported";

    private const string PatchExamples =
        """
        docxpatch 1

        op replace-text
        target M.P0004
        expect-text <<<
        current text
        >>>
        find <<<
        current
        >>>
        with <<<
        new text
        >>>
        end

        op set-cell
        target M.T0001.R02.C03
        expect-text <<<
        current cell text
        >>>
        expect-row-count 4
        expect-column-count 3
        text <<<
        updated cell text
        >>>
        end

        op insert-after
        target M.P0004
        text <<<
        inserted paragraph text
        >>>
        end

        op replace-image
        target M.I0001
        asset chart.png
        alt Chart after update
        end

        op add-comment
        target M.P0004
        expect-text <<<
        reviewed paragraph text
        >>>
        text <<<
        review note
        >>>
        author Reviewer
        end

        op delete-row
        target M.T0001.R02
        expect-contains row text
        end
        """;

    public static DocxCommandCatalog Catalog { get; } = BuildCatalog();

    public static bool TryGetCommand(string name, out DocxCommandInfo command)
    {
        command = Catalog.Commands.FirstOrDefault(command =>
            string.Equals(command.Name, name, StringComparison.Ordinal)) ?? new DocxCommandInfo();
        return !string.IsNullOrEmpty(command.Name);
    }

    public static bool TryGetPatchOperation(string name, out DocxPatchOperationInfo operation)
    {
        operation = Catalog.PatchOperations.FirstOrDefault(operation =>
            string.Equals(operation.Name, name, StringComparison.Ordinal)) ?? new DocxPatchOperationInfo();
        return !string.IsNullOrEmpty(operation.Name);
    }

    public static string RenderOverview()
    {
        var builder = new StringBuilder();
        builder.Append(Catalog.ToolName).Append(" - ").Append(Catalog.Summary).AppendLine();
        builder.AppendLine();
        builder.AppendLine("Usage:");
        builder.Append("  ").Append(Catalog.ToolName).AppendLine(" <command> [options]");
        AppendCommandGroup(builder, "Read / explore", "read");
        AppendCommandGroup(builder, "Patch", "patch");
        builder.AppendLine();
        builder.AppendLine("Help:");
        builder.AppendLine("  help dump|context|changes|validate|check|apply|patch");
        builder.AppendLine();
        builder.AppendLine("Examples:");
        foreach (string example in Catalog.Examples)
        {
            builder.Append("  ").AppendLine(example);
        }

        return builder.ToString();
    }

    public static bool TryRenderTopic(string topic, out string text)
    {
        if (string.Equals(topic, "patch", StringComparison.Ordinal))
        {
            text = RenderPatchHelp();
            return true;
        }

        if (TryGetCommand(topic, out DocxCommandInfo command))
        {
            text = RenderCommandHelp(command);
            return true;
        }

        text = string.Empty;
        return false;
    }

    public static string RenderTopic(string topic)
    {
        if (TryRenderTopic(topic, out string text))
        {
            return text;
        }

        throw new ArgumentException($"Unknown help topic '{topic}'.", nameof(topic));
    }

    public static string RenderPatchTrackChangesSupportTable()
    {
        var builder = new StringBuilder();
        builder.AppendLine("operation | support class | support value | behavior");
        builder.AppendLine("--- | --- | --- | ---");
        foreach (DocxPatchOperationInfo operation in Catalog.PatchOperations)
        {
            builder
                .Append(EscapeMarkdownTableCell(operation.Name))
                .Append(" | ")
                .Append(EscapeMarkdownTableCell(operation.TrackChangesSupportClass))
                .Append(" | ")
                .Append(EscapeMarkdownTableCell(operation.TrackChangesSupport))
                .Append(" | ")
                .Append(EscapeMarkdownTableCell(operation.TrackChangesNote))
                .AppendLine();
        }

        return builder.ToString();
    }

    private static DocxPatchOperationInfo PreserveOnly(
        string name,
        IReadOnlyList<string> requiredFields,
        string rationale,
        IReadOnlyList<string>? optionalFields = null)
    {
        return new DocxPatchOperationInfo
        {
            Name = name,
            RequiredFields = requiredFields,
            OptionalFields = optionalFields ?? [],
            TrackChangesSupportClass = TrackClassPreserveOnly,
            TrackChangesSupport = "preserve-only",
            TrackChangesNote = $"{rationale} {PreserveOnlyTrackChangesNote}"
        };
    }

    private static DocxPatchOperationInfo Tracked(
        string name,
        IReadOnlyList<string> requiredFields,
        string supportClass,
        string support,
        string note,
        IReadOnlyList<string>? optionalFields = null)
    {
        return new DocxPatchOperationInfo
        {
            Name = name,
            RequiredFields = requiredFields,
            OptionalFields = optionalFields ?? [],
            TrackChangesSupportClass = supportClass,
            TrackChangesSupport = support,
            TrackChangesNote = note
        };
    }

    private static DocxPatchOperationInfo Unsupported(
        string name,
        IReadOnlyList<string> requiredFields,
        string note,
        IReadOnlyList<string>? optionalFields = null)
    {
        return new DocxPatchOperationInfo
        {
            Name = name,
            RequiredFields = requiredFields,
            OptionalFields = optionalFields ?? [],
            TrackChangesSupportClass = TrackClassUnsupported,
            TrackChangesSupport = "unsupported",
            TrackChangesNote = note
        };
    }

    private static void AppendCommandGroup(StringBuilder builder, string title, string category)
    {
        builder.AppendLine();
        builder.AppendLine(title + ":");
        foreach (DocxCommandInfo command in Catalog.Commands.Where(command =>
                     string.Equals(command.Category, category, StringComparison.Ordinal)))
        {
            builder.Append("  ").Append(command.Name.PadRight(10)).AppendLine(command.Summary);
        }
    }

    private static string RenderCommandHelp(DocxCommandInfo command)
    {
        var builder = new StringBuilder();
        builder.AppendLine(command.Usage);
        builder.AppendLine();
        builder.AppendLine(command.Description);
        AppendFields(builder, command.OutputFields);
        AppendLines(builder, "Privacy notes:", command.PrivacyNotes);
        foreach (DocxHelpSection section in command.Notes)
        {
            AppendLines(builder, section.Heading + ":", section.Lines);
        }

        AppendOptions(builder, command.Options);
        AppendLines(builder, "Examples:", command.Examples, indent: "  ");
        return builder.ToString();
    }

    private static string RenderPatchHelp()
    {
        var builder = new StringBuilder();
        builder.AppendLine(PatchExamples);
        builder.AppendLine();
        builder.AppendLine("Supported operations:");
        foreach (DocxPatchOperationInfo operation in Catalog.PatchOperations)
        {
            builder.Append("  ").AppendLine(operation.RenderSummary());
        }

        builder.AppendLine();
        builder.AppendLine("Track-change support:");
        builder.Append(RenderPatchTrackChangesSupportTable());

        builder.AppendLine();
        builder.AppendLine("Selectors may use explicit IDs, heading:\"Text\", heading:2:\"Text\", text:\"contained text\", bookmark:\"Name\", or content-control:\"TagOrAlias\" for paragraph targets.");
        builder.AppendLine("delete-row expect-contains checks that the target row's final visible text contains the supplied value.");
        builder.AppendLine("add-comment-reply and delete-comment-reply are recognized but fail with E4314 until threaded comment metadata is safely modeled.");
        builder.AppendLine("add-repeating-section-item and delete-repeating-section-item are recognized but fail with E4315 until repeating-section subtree edits are safely modeled.");
        builder.AppendLine("append-column, insert-column-before, insert-column-after, and delete-column are recognized but fail with E4316 until table-column transforms are safely modeled.");
        builder.AppendLine("set-table-style, set-table-metadata, and set-row-header update table properties with explicit guards.");
        builder.AppendLine("replace-image alt updates the image DrawingML description while replacing the media bytes.");
        builder.AppendLine("set-image-metadata updates image docPr alt/title/name without replacing media bytes.");
        builder.AppendLine("set-image-size updates DrawingML extents without replacing media bytes.");
        builder.AppendLine("set-image-wrap updates anchored DrawingML wrap mode and wrap distances.");
        builder.AppendLine("set-image-position updates anchored DrawingML relative positions, offsets, and alignments.");
        builder.AppendLine("set-image-crop updates DrawingML a:srcRect crop percentages without replacing media bytes.");
        builder.AppendLine("Linked images and unsupported drawing shapes are reported as diagnostics and preserved.");
        builder.AppendLine("Track changes are controlled by check/apply --track-changes off|preserve|suggest|require.");
        builder.AppendLine("Unsupported fields are rejected. expect-hash and preserve-size are not supported.");
        return builder.ToString();
    }

    private static string EscapeMarkdownTableCell(string value)
    {
        return value
            .Replace("\r", " ", StringComparison.Ordinal)
            .Replace("\n", " ", StringComparison.Ordinal)
            .Replace("|", "\\|", StringComparison.Ordinal);
    }

    private static void AppendFields(StringBuilder builder, IReadOnlyList<DocxOutputFieldInfo> fields)
    {
        if (fields.Count == 0)
        {
            return;
        }

        builder.AppendLine();
        builder.AppendLine("JSON output includes:");
        foreach (DocxOutputFieldInfo field in fields)
        {
            builder.Append("  ").Append(field.Name.PadRight(14)).AppendLine(field.Description);
        }
    }

    private static void AppendOptions(StringBuilder builder, IReadOnlyList<DocxOptionInfo> options)
    {
        if (options.Count == 0)
        {
            return;
        }

        builder.AppendLine();
        builder.AppendLine("Options:");
        foreach (DocxOptionInfo option in options)
        {
            builder.Append("  ").Append(option.Syntax.PadRight(28)).AppendLine(option.Description);
        }
    }

    private static void AppendLines(
        StringBuilder builder,
        string heading,
        IReadOnlyList<string> lines,
        string indent = "")
    {
        if (lines.Count == 0)
        {
            return;
        }

        builder.AppendLine();
        builder.AppendLine(heading);
        foreach (string line in lines)
        {
            builder.Append(indent).AppendLine(line);
        }
    }

    private static DocxCommandCatalog BuildCatalog()
    {
        return new DocxCommandCatalog
        {
            Commands =
            [
                new()
                {
                    Name = "read",
                    Category = "read",
                    Summary = "Produce an agent-friendly structural view of a .docx",
                    Usage = "docxedit read input.docx [--summary] [--view final|original|markup] [--max-text <chars>] [--json] [--diagnostics <path>] [--strict]",
                    Description = "Produce an agent-friendly structural view of a .docx.",
                    Options =
                    [
                        new("--summary", "Print compact counts without broad document text"),
                        new("--headers-footers", "Include header/footer stories"),
                        new("--all-stories", "Include all modeled stories"),
                        new("--view final|original|markup", "Text view for tracked insert/delete text"),
                        new("--max-text N", "Maximum text per rendered field"),
                        new("--json", "Print the result object as JSON"),
                        new("--diagnostics path", "Write diagnostics JSON"),
                        new("--strict", "Return 3 when warnings are present")
                    ],
                    PrivacyNotes =
                    [
                        "Use --summary, or DocxTextRenderer.RenderReadSummary, when only structure counts are needed."
                    ],
                    OutputFields =
                    [
                        new("Bookmarks", "Structured bookmark ranges with name, OOXML ID, story, part, start/end targets, completeness, and duplicate-name candidate IDs"),
                        new("ContentControls", "Structured content controls with kind, tag, alias, placeholder/data-binding metadata, hierarchy IDs, safe-edit status, duplicate tag/alias candidate IDs, lock, story, part, containing target, text length, checkbox state, dropdown/combo item metadata, repeating-section metadata, and date settings"),
                        new("Fields", "Structured fields with kind, parsed type, code, containing target, result length, nesting depth, bookmark/hyperlink dependencies, safe-edit status, dirty/lock flags, and completeness"),
                        new("Hyperlinks", "Structured hyperlinks with relationship ID, relationship part, target mode, URI, URI scheme validation, anchor resolution flags, target part, containing target, and broken relationship flag"),
                        new("Images", "Structured image instances with layout kind, relationship ID, containing target, size, alt text, wrap mode, wrap distances, anchor positioning, crop percentages, and media part"),
                        new("Tables", "Structured table, row, and cell metadata including style, grid columns, header rows, omitted columns, spans, merge groups, vertical merge roots, and nested tables")
                    ],
                    Notes =
                    [
                        new("List and style metadata",
                        [
                            "Paragraph records may include StyleId, StyleName, and resolved List metadata. List metadata reports concrete numbering ID, zero-based level, abstract numbering ID, format, level text, paragraph style link, source=style or source=style-inherited when numbering comes from style inheritance, visible labels, and structured label components for deterministic decimal, letter, roman, bullet, and nested lvlText patterns. Picture bullets and unsupported custom numbering formats are preserved and reported through diagnostics."
                        ]),
                        new("Bookmarks and content controls",
                        [
                            "Read and context output surface bookmark names/ranges and content-control metadata, including duplicate selector candidate IDs, placeholder/data-binding details, hierarchy IDs, safe-edit status, checkbox state, dropdown/combo item counts, repeating-section metadata, and date settings/value. Bookmark and content-control selectors can be used for safe paragraph targeting; use explicit IDs when selector diagnostics report multiple candidate targets. Patch operations can update plain-text content-control IDs, guarded rich-text content controls, checkbox state, dropdown/combo selections, date values, create guarded paragraph bookmarks, update bookmark names, remove unreferenced bookmark markers, and replace guarded paragraph-bounded bookmark ranges, including multi-run and multi-paragraph ranges, while preserving wrappers/markers; content-control edits reject controls whose w:lock value is present and not unlocked."
                        ]),
                        new("Fields",
                        [
                            "Read and context output surface simple and complex field metadata, including parsed field type, normalized field code, cached result text/length, nesting depth, REF/PAGEREF/NOTEREF bookmark dependencies, HYPERLINK URI/anchor dependencies, safe-edit status, dirty flags, and lock flags. Patch operations can set field dirty/lock flags on one existing field or on target all, update simple-field code/result caches, refresh simple REF/PAGEREF/NOTEREF cached results from unambiguous same-part bookmarks, and mark the document for field updates after other edits, but Word remains responsible for general field recalculation. Apply emits W5103 when a document containing fields was marked for Word-side refresh."
                        ]),
                        new("Hyperlinks",
                        [
                            "Read, outline, context, and dump --runs output surface hyperlink metadata. Relationship-backed links expose relationship ID, relationship part, target mode, URI or target part, scheme, validation status, and validation reason for unsupported schemes or relative/malformed targets. Patch URI targets must be absolute http, https, or mailto; relative, malformed, file/UNC-style, and unsafe-scheme targets are rejected. Internal anchors expose missing/duplicate anchor flags. Internal part links are preserved and emit W1023 because patch edits support external URI or anchor targets only. Target-frame/history metadata and broken relationship IDs are flagged. Patch operations can update URI/anchor targets, tooltip, target-frame, history, display text, insert hyperlinks, and remove hyperlink markup while preserving display runs."
                        ]),
                        new("Images",
                        [
                            "Read and media output surface inline and anchored image layout metadata, including DrawingML extent, docPr name/description/title, wrap mode, wrap distances, anchor relative positioning, relative height, overlap/aspect-lock flags, crop percentages from a:srcRect, and containing paragraph or cell target. Image byte replacement preserves existing drawing layout where supported, set-image-metadata updates docPr name/description/title, set-image-size updates extents, set-image-wrap updates anchored wrap mode/distances, set-image-position updates anchored positioning, and set-image-crop updates a:srcRect percentages. Linked images are not fetched or listed as editable images; VML, grouped drawings, charts, SmartArt, OLE objects, equations, and generic shapes are diagnostics-only preserve-only content."
                        ]),
                        new("Tables",
                        [
                            "Read and context output surface table style, caption, description, grid/header/merged/nested flags, row grid-before/grid-after/header/cant-split metadata, cell logical/physical column positions, merge group IDs, visual column ends, and vertical-merge root cells. Table metadata operations can set or clear caption/description values. Row operations reject visual-grid tables with gridSpan, gridBefore, gridAfter, or vertical merges unless a force mode is explicitly supported by that operation."
                        ])
                    ],
                    Examples =
                    [
                        "docxedit read report.docx [--view final|original|markup]",
                        "docxedit read report.docx --summary"
                    ]
                },
                new()
                {
                    Name = "outline",
                    Category = "read",
                    Summary = "Show headings, tables, images, sections, headers, footers",
                    Usage = "docxedit outline input.docx [--json] [--diagnostics <path>] [--strict]",
                    Description = "Show headings, tables, images, sections, headers, and footers.",
                    Options =
                    [
                        new("--headers-footers", "Include header/footer stories"),
                        new("--json", "Print the result object as JSON"),
                        new("--diagnostics path", "Write diagnostics JSON"),
                        new("--strict", "Return 3 when warnings are present")
                    ],
                    Notes =
                    [
                        new("Numbered headings",
                        [
                            "Heading lines include compact resolved list metadata when the heading paragraph has numbering."
                        ])
                    ]
                },
                new()
                {
                    Name = "find",
                    Category = "read",
                    Summary = "Find text and print stable edit targets",
                    Usage = "docxedit find input.docx \"text\" [--view final|original|markup] [--max-text <chars>] [--json] [--diagnostics <path>] [--strict]",
                    Description = "Find text and print stable edit targets.",
                    Options =
                    [
                        new("--headers-footers", "Include header/footer stories"),
                        new("--view final|original|markup", "Text view for tracked insert/delete text"),
                        new("--max-text N", "Maximum text per rendered field"),
                        new("--json", "Print the result object as JSON"),
                        new("--diagnostics path", "Write diagnostics JSON"),
                        new("--strict", "Return 3 when warnings are present")
                    ],
                    Notes =
                    [
                        new("Numbered paragraph matches",
                        [
                            "Paragraph matches include compact resolved list metadata when the matched paragraph has numbering."
                        ])
                    ]
                },
                new()
                {
                    Name = "dump",
                    Category = "read",
                    Summary = "Dump one target in detail",
                    Usage = "docxedit dump input.docx --id TARGET [options]",
                    Description = "Dump one target by stable ID. Paragraph and cell dumps print visible text. Comment body targets such as C001.C0001 or comment:3 print metadata only, never comment body text. With --runs, paragraph dumps include run-level markup metadata such as markup=inserted-run, markup=deleted-run, revision-id, author, timestamp-utc, comment-id, and comment range/reference markers without printing comment body text. With --json, the Runs array exposes the same run metadata as structured fields. Change IDs from `changes` identify markup records; run IDs from dump identify rendered run/marker lines and are not the same namespace. In JSON output, inspect Runs[] for structured run metadata.",
                    Options =
                    [
                        new("--id M.P0001", "Target ID from read, outline, find, or changes"),
                        new("--runs", "Include paragraph run lines and markup metadata"),
                        new("--view final|original|markup", "Text view for tracked insert/delete text"),
                        new("--max-text N", "Maximum text per rendered field"),
                        new("--json", "Print the result object as JSON"),
                        new("--diagnostics path", "Write diagnostics JSON"),
                        new("--strict", "Return 3 when warnings are present")
                    ],
                    OutputFields =
                    [
                        new("Runs", "Structured run metadata when --runs is used")
                    ],
                    Examples =
                    [
                        "docxedit dump report.docx --id M.P0004 --runs",
                        "docxedit dump report.docx --id M.P0004 --runs --view markup --max-text 200"
                    ]
                },
                new()
                {
                    Name = "context",
                    Category = "read",
                    Summary = "Show nearby structure around one target without broad text",
                    Usage = "docxedit context input.docx --id TARGET [options]",
                    Description = "Summarize nearby modeled structure around one target without broad document text. By default, --max-text is 0, so paragraph and cell text fields are present but empty. Comment anchors are surfaced as comment IDs and comment body IDs, and comment body targets such as C001.C0001 or comment:3 return metadata only. Increase --max-text only when short snippets are needed.",
                    Options =
                    [
                        new("--id M.P0001", "Target ID from read, outline, find, or changes"),
                        new("--radius N", "Number of same-kind neighbors to include"),
                        new("--headers-footers", "Include header/footer stories"),
                        new("--view final|original|markup", "Text view when --max-text is greater than 0"),
                        new("--max-text N", "Maximum text per paragraph/cell; default 0"),
                        new("--json", "Print the result object as JSON"),
                        new("--diagnostics path", "Write diagnostics JSON"),
                        new("--strict", "Return 3 when warnings are present")
                    ],
                    PrivacyNotes =
                    [
                        "The default metadata-only context profile uses MaxText=0.",
                        "Comment body targets do not print comment body text; use changes --include-comment-text only when a bounded snippet is explicitly needed."
                    ],
                    Examples =
                    [
                        "docxedit context report.docx --id M.P0004",
                        "docxedit context report.docx --id M.T0001.R02.C03 --radius 1 --max-text 80",
                        "docxedit context report.docx --id M.P0004 --json"
                    ]
                },
                new()
                {
                    Name = "styles",
                    Category = "read",
                    Summary = "List paragraph, character, and table styles",
                    Usage = "docxedit styles input.docx [--json] [--diagnostics <path>] [--strict]",
                    Description = "List paragraph, character, and table styles. Output includes inheritance links such as based-on, next, linked, and style-level numbering defaults when present.",
                    Options =
                    [
                        new("--json", "Print the result object as JSON"),
                        new("--diagnostics path", "Write diagnostics JSON"),
                        new("--strict", "Return 3 when warnings are present")
                    ]
                },
                new()
                {
                    Name = "media",
                    Category = "read",
                    Summary = "List embedded images",
                    Usage = "docxedit media input.docx [--extract <dir>] [--json] [--diagnostics <path>] [--strict]",
                    Description = "List embedded images.",
                    Options =
                    [
                        new("--extract dir", "Extract embedded image parts to a directory"),
                        new("--json", "Print the result object as JSON"),
                        new("--diagnostics path", "Write diagnostics JSON"),
                        new("--strict", "Return 3 when warnings are present")
                    ]
                },
                new()
                {
                    Name = "validate",
                    Category = "read",
                    Summary = "Validate package and WordprocessingML invariants",
                    Usage = "docxedit validate input.docx [--profile structural|package] [--max-diagnostics <count>] [--json] [--diagnostics <path>] [--strict]",
                    Description = "Validate package-level XML roots and, with the structural profile, WordprocessingML invariants: paired bookmark/comment ranges, duplicate semantic selectors, commentsExtended consistency, complex field balance/result containment, content-control metadata, tracked revision markup, paragraph style and numbering references, settings updateFields, header/footer references, section properties, drawing relationships, image target/content-type checks, drawing geometry, basic table shape, and table visual-grid consistency. Package profile limits validation to package/XML root checks. Diagnostics are capped and report E9199 or W9199 when omitted.",
                    Options =
                    [
                        new("--profile structural|package", "Validation profile; structural is the default"),
                        new("--max-diagnostics N", "Maximum diagnostics to return; default 500"),
                        new("--json", "Print the result object as JSON"),
                        new("--diagnostics path", "Write diagnostics JSON"),
                        new("--strict", "Return 3 when warnings are present")
                    ],
                    OutputFields =
                    [
                        new("Diagnostics", "Validation errors and warnings with part names and stable diagnostic codes")
                    ],
                    Examples =
                    [
                        "docxedit validate report.docx",
                        "docxedit validate report.docx --json"
                    ]
                },
                new()
                {
                    Name = "changes",
                    Category = "read",
                    Summary = "List tracked-change and comment markup; comment text is opt-in",
                    Usage = "docxedit changes input.docx [options]",
                    Description = "List tracked-change and comment markup. By default this does not print revision text or comment body text. Records include change IDs, type, story, part, normalized parent type, target, revision/comment metadata, text length, child element count, and comment anchor targets when known. Use --include-comment-text only when short comment body snippets are needed.",
                    Options =
                    [
                        new("--include-comment-text", "Include explicit comment body snippets in comment summaries and comment body records"),
                        new("--max-comment-text N", "Maximum comment body snippet length when --include-comment-text is used; default 240"),
                        new("--json", "Print the result object as JSON"),
                        new("--diagnostics path", "Write diagnostics JSON"),
                        new("--strict", "Return 3 when warnings are present")
                    ],
                    OutputFields =
                    [
                        new("Summary", "Counts by change type"),
                        new("GroupSummary", "Counts by group=story|part|author|target and type"),
                        new("TargetSummary", "Compact per-target type rollups"),
                        new("CommentSummary", "Compact per-comment anchor/type rollups; may include TextSnippet only when requested"),
                        new("Changes", "Individual change records; may include CommentTextSnippet only when requested")
                    ],
                    PrivacyNotes =
                    [
                        "The default changes profile is private-text-free: it reports metadata, IDs, text lengths, child counts, targets, stories, and parts without revision or comment body text.",
                        "Comment body text appears only when IncludeCommentText is set or --include-comment-text is passed. Keep MaxCommentText low for agent context."
                    ],
                    Notes =
                    [
                        new("Target notes",
                        [
                            "target-status explains whether a record is targeted, linked by a comment anchor, or targetless. target-source distinguishes exact ancestor matches from adjacent-range heuristics. target-reason classifies targetless records, and nearest-target is context only, not exact ownership. paired-change-id links range starts/ends that share a revision or comment ID.",
                            "target=unknown means the markup is not inside or adjacent to a modeled paragraph/table/section target. Comment body records may expose comment-anchor-target and comment-reference-target when the main-story anchor/reference can be correlated by comment ID."
                        ]),
                        new("Text output",
                        [
                            "Text output includes type counts, group summaries, target summaries, comment summaries, and individual records. Individual records include parent=<kind> when the immediate WordprocessingML parent can be normalized. Modern Word comment metadata appears as para-id, parent-para-id, root-para-id, is-reply, resolved, comment-para-id, comment-parent-para-id, comment-root-para-id, comment-is-reply, and comment-resolved when commentsExtended data is present. With --include-comment-text, comment-summary and comment body records add comment-text-length, comment-text, and comment-text-truncated when applicable."
                        ]),
                        new("Timestamp notes",
                        [
                            "TimestampUtc and CommentTimestampUtc are nullable UTC ISO-8601 values. Raw JSON uses UTC values such as +00:00. Some JSON consumers may display parsed date values in local time, so inspect the raw JSON string when the serialized timestamp offset matters."
                        ])
                    ],
                    Examples =
                    [
                        "docxedit changes report.docx",
                        "docxedit changes report.docx --json",
                        "docxedit changes report.docx --include-comment-text --max-comment-text 120"
                    ]
                },
                new()
                {
                    Name = "check",
                    Category = "patch",
                    Summary = "Validate a .docxpatch file without writing output",
                    Usage = "docxedit check input.docx edits.docxpatch [options]",
                    Description = "Validate a patch against an input document without writing an output file. Use check before apply to verify selectors, guards, assets, and track-change constraints. Text output includes one operation line per patch operation and affected row/cell lines for table operations.",
                    Options =
                    [
                        new("--track-changes off|preserve|suggest|require", "Tracked-change handling mode"),
                        new("--author name", "Non-empty author used for generated revisions; defaults to docxedit"),
                        new("--timestamp-utc instant", "Timestamp normalized to UTC for generated revisions"),
                        new("--json", "Print the result object as JSON"),
                        new("--report path", "Write operation report JSON"),
                        new("--diagnostics path", "Write diagnostics JSON"),
                        new("--strict", "Return 3 when warnings are present")
                    ],
                    Examples =
                    [
                        "docxedit check report.docx edits.docxpatch",
                        "docxedit check report.docx edits.docxpatch --track-changes require --author Agent --timestamp-utc 2026-01-01T00:00:00Z"
                    ]
                },
                new()
                {
                    Name = "apply",
                    Category = "patch",
                    Summary = "Apply a .docxpatch file and write a new .docx",
                    Usage = "docxedit apply input.docx edits.docxpatch --output output.docx [options]",
                    Description = "Apply a patch and write a new .docx. The input is never modified in place. Text output includes one operation line per patch operation, generated revision IDs when tracked markup is created, and affected row/cell lines for table operations.",
                    Options =
                    [
                        new("--output path, -o path", "Output .docx path"),
                        new("--track-changes off|preserve|suggest|require", "Tracked-change handling mode"),
                        new("--author name", "Non-empty author used for generated revisions; defaults to docxedit"),
                        new("--timestamp-utc instant", "Timestamp normalized to UTC for generated revisions"),
                        new("--json", "Print the result object as JSON"),
                        new("--report path", "Write operation report JSON"),
                        new("--diagnostics path", "Write diagnostics JSON"),
                        new("--strict", "Return 3 when warnings are present")
                    ],
                    Notes =
                    [
                        new("Track-change modes",
                        [
                            "off       Apply direct edits",
                            "preserve  Preserve existing markup while applying direct edits",
                            "suggest   Generate tracked replacements when supported, otherwise warn and edit directly",
                            "require   Generate tracked replacements and fail unsupported tracked shapes"
                        ])
                    ],
                    Examples =
                    [
                        "docxedit apply report.docx edits.docxpatch --output report.edited.docx",
                        "docxedit apply report.docx edits.docxpatch --output report.edited.docx --track-changes preserve"
                    ]
                }
            ],
            Examples =
            [
                "docxedit read report.docx [--view final|original|markup]",
                "docxedit read report.docx --summary",
                "docxedit dump report.docx --id M.P0004 --runs",
                "docxedit context report.docx --id M.P0004",
                "docxedit media report.docx --extract media",
                "docxedit validate report.docx",
                "docxedit changes report.docx",
                "docxedit check report.docx edits.docxpatch",
                "docxedit apply report.docx edits.docxpatch --output report.edited.docx"
            ],
            PatchOperations =
            [
                Tracked(
                    "replace-text",
                    ["target", "find", "with"],
                    TrackClassTextRun,
                    "tracked-simple",
                    "Suggest/Require emit tracked w:del/w:ins for supported simple text-only matches; unsupported shapes warn with W4002 or fail with E6002.",
                    ["expect-text", "preserve-runs", "occurrence"]),
                Tracked(
                    "replace-paragraph",
                    ["target", "text"],
                    TrackClassParagraphBlock,
                    "tracked-paragraph",
                    "Suggest/Require emit whole-paragraph w:del/w:ins for simple text replacements and add w:pPrChange when a compatible style change is included; complex shapes warn with W4002 or fail with E6002.",
                    ["expect-text", "style"]),
                Tracked(
                    "insert-before",
                    ["target", "text"],
                    TrackClassParagraphBlock,
                    "tracked-paragraph-insert",
                    "Suggest/Require emit inserted paragraph text as w:ins when the inserted text has no tabs or line breaks; unsupported shapes warn with W4002 or fail with E6002.",
                    ["style", "copy-paragraph-properties"]),
                Tracked(
                    "insert-after",
                    ["target", "text"],
                    TrackClassParagraphBlock,
                    "tracked-paragraph-insert",
                    "Suggest/Require emit inserted paragraph text as w:ins when the inserted text has no tabs or line breaks; unsupported shapes warn with W4002 or fail with E6002.",
                    ["style", "copy-paragraph-properties"]),
                Tracked(
                    "delete-block",
                    ["target"],
                    TrackClassParagraphBlock,
                    "tracked-paragraph-delete",
                    "Suggest/Require emit deleted paragraph text as w:del for simple paragraph targets; table/block or complex shapes warn with W4002 or fail with E6002.",
                    ["expect-text"]),
                Tracked(
                    "set-style",
                    ["target", "style"],
                    TrackClassParagraphProperty,
                    "tracked-style",
                    "Suggest/Require emit paragraph property revisions with w:pPrChange."),
                PreserveOnly(
                    "set-content-control-text",
                    ["target", "text"],
                    "Content-control text replacement must preserve the wrapper, bindings, and locks; generated revision markup inside the wrapper is not modeled yet.",
                    ["expect-text"]),
                PreserveOnly(
                    "set-content-control-checkbox",
                    ["target", "checked"],
                    "Checkbox content controls update state metadata, not a simple Word revision range."),
                PreserveOnly(
                    "set-content-control-choice",
                    ["target plus value or display-text"],
                    "Dropdown and combo-box content controls update list value metadata and display text together; generated revision markup is not modeled yet."),
                PreserveOnly(
                    "set-content-control-date",
                    ["target", "value"],
                    "Date content controls update date metadata and display text together; generated revision markup is not modeled yet.",
                    ["display-text"]),
                Unsupported(
                    "add-repeating-section-item",
                    ["target"],
                    "Repeating-section item insertion is not safely modeled yet; check/apply fails with E4315.",
                    ["source", "index", "text"]),
                Unsupported(
                    "delete-repeating-section-item",
                    ["target"],
                    "Repeating-section item deletion is not safely modeled yet; check/apply fails with E4315.",
                    ["index"]),
                PreserveOnly(
                    "add-bookmark",
                    ["target", "name"],
                    "Bookmark creation adds anchor metadata; Word has no useful generated revision range for the bookmark markers.",
                    ["expect-text"]),
                PreserveOnly(
                    "replace-bookmark-text",
                    ["target", "text"],
                    "Bookmark text replacement must preserve range anchors; generated revision markup across bookmark boundaries is not modeled yet."),
                PreserveOnly(
                    "rename-bookmark",
                    ["target", "name"],
                    "Bookmark rename changes anchor metadata; Word has no useful generated revision range for the name update."),
                PreserveOnly(
                    "delete-bookmark",
                    ["target"],
                    "Bookmark deletion removes anchor metadata; Word has no useful generated revision range for the marker removal."),
                PreserveOnly(
                    "add-comment",
                    ["target", "text"],
                    "Comments are already review markup, so adding a comment does not create an additional tracked edit.",
                    ["expect-text", "author", "initials", "date"]),
                PreserveOnly(
                    "set-comment-text",
                    ["target", "text"],
                    "Comment body edits modify review markup; generated tracked revisions inside comments are not modeled yet."),
                PreserveOnly(
                    "resolve-comment",
                    ["target"],
                    "Comment resolution changes review metadata, not visible document text."),
                PreserveOnly(
                    "reopen-comment",
                    ["target"],
                    "Comment reopening changes review metadata, not visible document text."),
                PreserveOnly(
                    "delete-comment",
                    ["target"],
                    "Comment deletion removes review markup, not a separate generated tracked edit."),
                Unsupported(
                    "add-comment-reply",
                    ["target", "text"],
                    "Threaded comment replies are not safely modeled yet; check/apply fails with E4314.",
                    ["author", "initials", "date"]),
                Unsupported(
                    "delete-comment-reply",
                    ["target"],
                    "Threaded comment replies are not safely modeled yet; check/apply fails with E4314."),
                PreserveOnly(
                    "set-field-dirty",
                    ["target or all", "dirty"],
                    "Field dirty flags are field metadata and have no useful generated visible revision representation."),
                PreserveOnly(
                    "set-field-lock",
                    ["target or all", "locked"],
                    "Field lock flags are field metadata and have no useful generated visible revision representation."),
                PreserveOnly(
                    "set-field-code",
                    ["target", "code"],
                    "Field codes are instruction metadata; generated revisions for field instructions are not modeled yet.",
                    ["expect-code"]),
                PreserveOnly(
                    "set-field-result",
                    ["target", "text"],
                    "Field result replacement must preserve field topology; generated revision markup for field results is not modeled yet.",
                    ["expect-result"]),
                PreserveOnly(
                    "refresh-field-result",
                    ["target"],
                    "Field refresh updates cached result text from modeled document state; generated revision markup for the refresh is not modeled yet.",
                    ["expect-code", "expect-result"]),
                PreserveOnly(
                    "set-hyperlink-target",
                    ["target plus uri or anchor"],
                    "Hyperlink target updates modify relationship or anchor metadata, not visible text.",
                    ["tooltip", "target-frame", "history"]),
                Tracked(
                    "set-hyperlink-text",
                    ["target", "text"],
                    TrackClassTextRun,
                    "tracked-hyperlink-text",
                    "Suggest/Require emit w:del/w:ins inside the hyperlink wrapper for simple display text while preserving the relationship or anchor; protected or complex hyperlink content warns with W4002 or fails with E6002."),
                PreserveOnly(
                    "insert-hyperlink-after",
                    ["target", "text plus uri or anchor"],
                    "Hyperlink insertion creates visible text plus relationship metadata; generated revision markup for the combined shape is not modeled yet.",
                    ["tooltip", "target-frame", "history"]),
                PreserveOnly(
                    "remove-hyperlink",
                    ["target"],
                    "Hyperlink removal changes wrapper and relationship metadata while preserving display text."),
                Tracked(
                    "set-cell",
                    ["target", "text"],
                    TrackClassTextRun,
                    "tracked-cell-simple",
                    "Suggest/Require emit w:del/w:ins for simple text-only cells, including compatible multi-paragraph cells; force or complex cells warn with W4002 or fail with E6002.",
                    ["expect-text", "expect-row-count", "expect-column-count", "force"]),
                PreserveOnly(
                    "set-table-style",
                    ["target", "style"],
                    "Table style updates are table properties; generated w:tblPrChange output is not modeled yet.",
                    ["expect-style"]),
                PreserveOnly(
                    "set-table-metadata",
                    ["target plus caption or description"],
                    "Table caption and description updates are table metadata, not visible document text.",
                    ["expect-caption", "expect-description"]),
                PreserveOnly(
                    "set-row-header",
                    ["target", "header"],
                    "Repeating-row header updates are row properties; generated w:trPrChange output is not modeled yet.",
                    ["expect-header"]),
                PreserveOnly(
                    "append-row",
                    ["target plus repeated cell"],
                    "Table row insertion revisions are not modeled yet.",
                    ["expect-row-count", "expect-column-count"]),
                PreserveOnly(
                    "insert-row-before",
                    ["target plus repeated cell"],
                    "Table row insertion revisions are not modeled yet.",
                    ["expect-row-count", "expect-column-count", "expect-cell-count", "force"]),
                PreserveOnly(
                    "insert-row-after",
                    ["target plus repeated cell"],
                    "Table row insertion revisions are not modeled yet.",
                    ["expect-row-count", "expect-column-count", "expect-cell-count", "force"]),
                PreserveOnly(
                    "delete-row",
                    ["target"],
                    "Table row deletion revisions are not modeled yet.",
                    ["expect-row-count", "expect-column-count", "expect-cell-count", "expect-contains", "force"]),
                Unsupported(
                    "append-column",
                    ["target plus repeated cell"],
                    "Table-column transforms are not safely modeled yet; check/apply fails with E4316.",
                    ["expect-row-count", "expect-column-count", "force"]),
                Unsupported(
                    "insert-column-before",
                    ["target", "column plus repeated cell"],
                    "Table-column transforms are not safely modeled yet; check/apply fails with E4316.",
                    ["expect-row-count", "expect-column-count", "expect-cell-count", "force"]),
                Unsupported(
                    "insert-column-after",
                    ["target", "column plus repeated cell"],
                    "Table-column transforms are not safely modeled yet; check/apply fails with E4316.",
                    ["expect-row-count", "expect-column-count", "expect-cell-count", "force"]),
                Unsupported(
                    "delete-column",
                    ["target", "column"],
                    "Table-column transforms are not safely modeled yet; check/apply fails with E4316.",
                    ["expect-row-count", "expect-column-count", "expect-cell-count", "expect-contains", "force"]),
                PreserveOnly(
                    "replace-image",
                    ["target", "asset"],
                    "Image replacement updates DrawingML and package media; generated drawing-level revision markup is not modeled yet.",
                    ["expect-content-type", "alt"]),
                PreserveOnly(
                    "insert-image-after",
                    ["target", "asset"],
                    "Image insertion creates DrawingML and package media; generated drawing-level revision markup is not modeled yet.",
                    ["expect-content-type", "width", "height", "alt"]),
                PreserveOnly(
                    "set-image-alt",
                    ["target", "alt"],
                    "Image alt-text updates DrawingML metadata, not visible document text.",
                    ["expect-content-type"]),
                PreserveOnly(
                    "set-image-metadata",
                    ["target plus alt, title, or name"],
                    "Image title/name/alt updates DrawingML metadata, not visible document text.",
                    ["expect-content-type"]),
                PreserveOnly(
                    "set-image-size",
                    ["target plus width or height"],
                    "Image size updates DrawingML layout metadata, not visible document text.",
                    ["expect-content-type"]),
                PreserveOnly(
                    "set-image-wrap",
                    ["target plus mode or distance"],
                    "Image wrapping updates DrawingML layout metadata, not visible document text.",
                    ["expect-content-type"]),
                PreserveOnly(
                    "set-image-position",
                    ["target plus relative, offset, or align"],
                    "Image position updates DrawingML layout metadata, not visible document text.",
                    ["expect-content-type"]),
                PreserveOnly(
                    "set-image-crop",
                    ["target plus one crop percentage"],
                    "Image crop updates DrawingML layout metadata, not visible document text.",
                    ["expect-content-type"]),
                PreserveOnly(
                    "delete-image",
                    ["target"],
                    "Image deletion removes DrawingML and package media; generated drawing-level revision markup is not modeled yet.",
                    ["expect-content-type"]),
                PreserveOnly(
                    "set-section-columns",
                    ["target", "count"],
                    "Section column updates are section properties; generated w:sectPrChange output is not modeled yet.",
                    ["expect-columns", "expect-orientation"]),
                PreserveOnly(
                    "set-section-orientation",
                    ["target", "orientation"],
                    "Section orientation updates are section properties; generated w:sectPrChange output is not modeled yet.",
                    ["expect-columns", "expect-orientation"])
            ]
        };
    }
}
