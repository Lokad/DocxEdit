using System.Text;

namespace Lokad.DocxEdit;

/// <summary>
/// Machine-readable command surface: commands, examples, and patch operations.
/// </summary>
public sealed record DocxCommandCatalog
{
    /// <summary>Invoked tool name.</summary>
    public string ToolName { get; init; } = "docxedit";

    /// <summary>One-line tool summary.</summary>
    public string Summary { get; init; } = "read and patch .docx documents";

    /// <summary>Inspect commands (read, find, check, apply, …).</summary>
    public IReadOnlyList<DocxCommandInfo> Commands { get; init; } = [];

    /// <summary>End-to-end example invocations.</summary>
    public IReadOnlyList<string> Examples { get; init; } = [];

    /// <summary>Patch operations with fields and track-change support.</summary>
    public IReadOnlyList<DocxPatchOperationInfo> PatchOperations { get; init; } = [];
}

/// <summary>
/// One inspect/check/apply command: usage, options, outputs, and privacy notes.
/// </summary>
public sealed record DocxCommandInfo
{
    /// <summary>Command name (for example <c>dump</c>).</summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>Command group (inspect, validate, or edit).</summary>
    public string Category { get; init; } = string.Empty;

    /// <summary>One-line command summary.</summary>
    public string Summary { get; init; } = string.Empty;

    /// <summary>Canonical usage line.</summary>
    public string Usage { get; init; } = string.Empty;

    /// <summary>Longer command description.</summary>
    public string Description { get; init; } = string.Empty;

    /// <summary>Minimum positional arguments.</summary>
    public int MinPositionals { get; init; }

    /// <summary>Maximum positional arguments.</summary>
    public int MaxPositionals { get; init; }

    /// <summary>Accepted flags and values.</summary>
    public IReadOnlyList<DocxOptionInfo> Options { get; init; } = [];

    /// <summary>Output fields callers can rely on.</summary>
    public IReadOnlyList<DocxOutputFieldInfo> OutputFields { get; init; } = [];

    /// <summary>Privacy handling notes for this command.</summary>
    public IReadOnlyList<string> PrivacyNotes { get; init; } = [];

    /// <summary>Long-form help sections.</summary>
    public IReadOnlyList<DocxHelpSection> Notes { get; init; } = [];

    /// <summary>Example invocations.</summary>
    public IReadOnlyList<string> Examples { get; init; } = [];
}

/// <summary>
/// One accepted flag or value.
/// </summary>
/// <param name="Syntax">Flag syntax as shown in usage.</param>
/// <param name="Description">What the flag does.</param>
public sealed record DocxOptionInfo(string Syntax, string Description)
{
    /// <summary>Canonical accepted spellings, aliases included. Parse validation matches these exact strings; help and usage prose never affects parsing.</summary>
    public required IReadOnlyList<string> Flags { get; init; }
}

/// <summary>
/// One output field callers can rely on.
/// </summary>
/// <param name="Name">Field name.</param>
/// <param name="Description">What the field carries.</param>
public sealed record DocxOutputFieldInfo(string Name, string Description);

/// <summary>
/// One help section: a heading plus body lines.
/// </summary>
/// <param name="Heading">Section heading.</param>
/// <param name="Lines">Section body lines.</param>
public sealed record DocxHelpSection(string Heading, IReadOnlyList<string> Lines);

/// <summary>
/// One patch operation: fields, prose, and track-change support.
/// </summary>
public sealed record DocxPatchOperationInfo
{
    /// <summary>Operation name (for example <c>replace-text</c>).</summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>Editing area group (for example <c>Tables</c>).</summary>
    public string Category { get; init; } = string.Empty;

    /// <summary>Required patch fields: real field names, each required on every invocation.</summary>
    public IReadOnlyList<string> RequiredFields { get; init; } = [];

    /// <summary>Optional patch fields: real field names.</summary>
    public IReadOnlyList<string> OptionalFields { get; init; } = [];

    /// <summary>Field groups of which at least one member is required. Every member is a real field name.</summary>
    public IReadOnlyList<IReadOnlyList<string>> RequiredAlternatives { get; init; } = [];

    /// <summary>Field groups of which at most one member may carry a value; every other combination fails patch validation.</summary>
    public IReadOnlyList<IReadOnlyList<string>> ExclusiveAlternatives { get; init; } = [];
    /// <summary>Fields that must be true or false.</summary>
    public IReadOnlyList<string> BooleanFields { get; init; } = [];
    /// <summary>Executable example patches, minimal first and guarded second.</summary>
    public IReadOnlyList<string> Examples { get; init; } = [];
    /// <summary>Accepted explicit target ID kinds as lowercase words; empty skips document-independent kind validation.</summary>
    public IReadOnlyList<string> AcceptedTargets { get; init; } = [];
    /// <summary>Unit contracts as field=hint tokens; the hint kind precedes the colon and the accepted range or set follows it.</summary>
    public IReadOnlyList<string> UnitHints { get; init; } = [];
    /// <summary>Closed value sets as field=value1|value2 tokens.</summary>
    public IReadOnlyList<string> AllowedValues { get; init; } = [];
    /// <summary>Omitted-field effective values as field=value tokens.</summary>
    public IReadOnlyList<string> FieldDefaults { get; init; } = [];
    /// <summary>Fields that must be integers greater than 0, plus the literal occurrence value all for replace-text.</summary>
    public IReadOnlyList<string> IntegerFields { get; init; } = [];
    /// <summary>Fields that may repeat; every occurrence is kept in file order. Every other field must appear at most once per operation; repeating one fails patch parsing.</summary>
    public IReadOnlyList<string> RepeatableFields { get; init; } = [];

    /// <summary>Required fields that also accept a present-but-empty value (for example text deletion).</summary>
    public IReadOnlyList<string> EmptyAllowedFields { get; init; } = [];

    /// <summary>What the operation does. Markdown-flavored; may contain inline code spans.</summary>
    public string Description { get; init; } = string.Empty;

    /// <summary>Machine-readable track-change class (drives <c>GeneratesTrackedChanges</c>).</summary>
    public string TrackChangesSupportClass { get; init; } = "unsupported";

    /// <summary>Human track-change support label.</summary>
    public string TrackChangesSupport { get; init; } = "unsupported";

    /// <summary>Suggest/Require behavior note.</summary>
    public string TrackChangesNote { get; init; } = "Suggest applies directly with W4001; Require fails with E6001.";

    /// <summary>Whether the operation can emit tracked-change markup.</summary>
    public bool GeneratesTrackedChanges =>
        TrackChangesSupportClass is not "preserve-only" and not "unsupported";

    /// <summary>Renders a one-line summary of the operation.</summary>
    public string RenderSummary()
    {
        var required = new List<string>(RequiredFields.Select(MarkRepeatable));
        var head = new List<string>();
        if (required.Count != 0)
        {
            head.Add(string.Join("/", required));
        }
        head.AddRange(RequiredAlternatives.Select(group => string.Join("|", group)));
        string requiredText = string.Join(" ", head);
        string optional = OptionalFields.Count == 0 ? string.Empty : $", optional {string.Join(", ", OptionalFields.Select(MarkRepeatable))}";
        return string.IsNullOrWhiteSpace(requiredText)
            ? $"{Name}{optional}"
            : $"{Name} {requiredText}{optional}";
    }

    private string MarkRepeatable(string field) =>
        RepeatableFields.Contains(field, StringComparer.Ordinal) ? field + "+" : field;
}

/// <summary>
/// Built-in help: catalog lookup plus text rendering for the CLI.
/// </summary>
public static class DocxHelp
{
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

    /// <summary>Shared command catalog instance.</summary>
    public static DocxCommandCatalog Catalog { get; } = BuildCatalog();

    /// <summary>Tries to look up a command by name.</summary>
    public static bool TryGetCommand(string name, out DocxCommandInfo command)
    {
        command = Catalog.Commands.FirstOrDefault(command =>
            string.Equals(command.Name, name, StringComparison.Ordinal)) ?? new DocxCommandInfo();
        return !string.IsNullOrEmpty(command.Name);
    }

    /// <summary>Tries to look up a patch operation by name.</summary>
    public static bool TryGetPatchOperation(string name, out DocxPatchOperationInfo operation)
    {
        operation = Catalog.PatchOperations.FirstOrDefault(operation =>
            string.Equals(operation.Name, name, StringComparison.Ordinal)) ?? new DocxPatchOperationInfo();
        return !string.IsNullOrEmpty(operation.Name);
    }

    /// <summary>Renders the top-level help overview.</summary>
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
        builder.Append("  help ").AppendLine(string.Join("|", Catalog.Commands.Select(command => command.Name).Concat(["patch"])));
        builder.AppendLine();
        builder.AppendLine("Examples:");
        foreach (string example in Catalog.Examples)
        {
            builder.Append("  ").AppendLine(example);
        }

        builder.AppendLine("Pipes:");
        builder.AppendLine("  - stands for stdin/stdout: the input .docx, the patch file, and --output accept -");
        builder.AppendLine();
        builder.AppendLine("Exit codes:");
        builder.AppendLine("  0 success (warnings allowed; inspect diagnostics)");
        builder.AppendLine("  1 unsuccessful result; text mode prints error diagnostics to stderr");
        builder.AppendLine("  2 invalid command-line usage");
        builder.AppendLine("  3 successful result with warnings under --strict");
        builder.AppendLine("  4 unhandled exception");

        return builder.ToString();
    }

    private const string UnsupportedSupportClass = "unsupported";

    /// <summary>Tries to render one help topic by name.</summary>
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

        if (TryGetPatchOperation(topic, out DocxPatchOperationInfo patchOperation))
        {
            text = RenderPatchOperationHelp(patchOperation);
            return true;
        }

        text = string.Empty;
        return false;
    }

    /// <summary>Renders one help topic by name.</summary>
    public static string RenderTopic(string topic)
    {
        if (TryRenderTopic(topic, out string text))
        {
            return text;
        }

        throw new ArgumentException($"Unknown help topic '{topic}'.", nameof(topic));
    }

    /// <summary>Renders the patch track-change support table.</summary>
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

    /// <summary>Renders the patch operation reference tables grouped by editing area.</summary>
    public static string RenderPatchOperationTables()
    {
        var builder = new StringBuilder();
        bool first = true;
        foreach (string category in DocxPatchEngine.CatalogOperations.Select(static operation => operation.Category).Distinct())
        {
            if (!first)
            {
                builder.AppendLine();
            }

            first = false;
            builder.Append("### ").AppendLine(category);
            builder.AppendLine();
            builder.AppendLine("| Operation | Required fields | Optional fields | Notes |");
            builder.AppendLine("| --- | --- | --- | --- |");
            foreach (DocxPatchOperationInfo operation in DocxPatchEngine.CatalogOperations.Where(operation => operation.Category == category))
            {
                string optional = RenderOperationFields(operation, required: false);
                string optionalCell = optional.Length == 0 ? "| " : optional + " | ";
                builder.Append("| `").Append(operation.Name).Append("` | ")
                    .Append(RenderOperationFields(operation, required: true)).Append(" | ")
                    .Append(optionalCell)
                    .Append(EscapeMarkdownTableCell(operation.Description))
                    .Append(" |")
                    .AppendLine();
            }
        }

        return builder.ToString();

        static string RenderOperationFields(DocxPatchOperationInfo operation, bool required)
        {
            var tokens = (required ? operation.RequiredFields : operation.OptionalFields)
                .Select(field => operation.RepeatableFields.Contains(field, StringComparer.Ordinal) ? field + "+" : field)
                .ToList();
            if (required)
            {
                tokens.AddRange(operation.RequiredAlternatives.Select(group => string.Join("|", group)));
            }

            var knownFields = new HashSet<string>(operation.RequiredFields
                .Concat(operation.OptionalFields)
                .Concat(operation.RepeatableFields)
                .Concat(operation.RequiredAlternatives.SelectMany(group => group)), StringComparer.Ordinal);
            return EscapeMarkdownTableCell(RenderFieldDescriptors(tokens, knownFields));
        }

        static string RenderFieldDescriptors(IEnumerable<string> descriptors, HashSet<string> knownFields)
        {
            return string.Join(", ", descriptors.Select(descriptor => RenderFieldDescriptor(descriptor, knownFields)));
        }

        static string RenderFieldDescriptor(string descriptor, HashSet<string> knownFields)
        {
            var builder = new StringBuilder();
            int index = 0;
            while (index < descriptor.Length)
            {
                if (IsFieldWordChar(descriptor[index]))
                {
                    int start = index;
                    while (index < descriptor.Length && IsFieldWordChar(descriptor[index]))
                    {
                        index++;
                    }

                    string word = descriptor.Substring(start, index - start);
                    builder.Append(knownFields.Contains(word) ? "`" + word + "`" : word);
                }
                else
                {
                    builder.Append(descriptor[index]);
                    index++;
                }
            }

            return builder.ToString();
        }

        static bool IsFieldWordChar(char value)
        {
            return char.IsLetterOrDigit(value) || value == '_' || value == '-';
        }
    }

    private static void AppendCommandGroup(StringBuilder builder, string title, string category)
    {
        builder.AppendLine();
        builder.AppendLine(title + ":");
        List<DocxCommandInfo> group = Catalog.Commands.Where(command =>
            string.Equals(command.Category, category, StringComparison.Ordinal)).ToList();
        int width = Math.Max(10, group.Select(static command => command.Name.Length).DefaultIfEmpty(0).Max() + 1);
        foreach (DocxCommandInfo command in group)
        {
            builder.Append("  ").Append(command.Name.PadRight(width)).AppendLine(command.Summary);
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
        foreach (DocxPatchOperationInfo operation in Catalog.PatchOperations.Where(static operation => operation.TrackChangesSupportClass != UnsupportedSupportClass))
        {
            builder.Append("  ").AppendLine(operation.RenderSummary());
        }

        builder.AppendLine("Recognized but unimplemented operations (check/apply fail; see notes):");
        foreach (DocxPatchOperationInfo operation in Catalog.PatchOperations.Where(static operation => operation.TrackChangesSupportClass == UnsupportedSupportClass))
        {
            builder.Append("  ").AppendLine(operation.RenderSummary());
        }

        builder.AppendLine();
        builder.AppendLine("Track-change support:");
        builder.Append(RenderPatchTrackChangesSupportTable());

        builder.AppendLine();
        builder.AppendLine("Selectors may use explicit IDs, heading:\"Text\", heading:2:\"Text\", text:\"contained text\", bookmark:\"Name\", or content-control:\"TagOrAlias\" for paragraph targets.");
        builder.AppendLine("add-comment targets a modeled paragraph; optional anchor-text selects one direct text span inside it, and occurrence disambiguates repeated anchor text.");
        builder.AppendLine("delete-row expect-contains checks that the target row's final visible text contains the supplied value.");
        builder.AppendLine("add-comment-reply creates modern commentsExtended/commentsIds metadata; delete-comment-reply removes leaf replies and fails with E4314 when deletion would change child thread topology.");
        builder.AppendLine("add-repeating-section-item and delete-repeating-section-item are recognized but fail with E4315 until repeating-section subtree edits are safely modeled.");
        builder.AppendLine("append-column, insert-column-before, insert-column-after, and delete-column are recognized but fail with E4316 until table-column transforms are safely modeled.");
        builder.AppendLine("set-table-style, set-table-metadata, set-row-header, and set-cell-shading update table properties with explicit guards.");
        builder.AppendLine("replace-image alt updates the image DrawingML description while replacing the media bytes.");
        builder.AppendLine("set-image-metadata updates image docPr alt/title/name without replacing media bytes.");
        builder.AppendLine("set-image-size updates DrawingML extents without replacing media bytes.");
        builder.AppendLine("set-image-wrap updates anchored DrawingML wrap mode and wrap distances.");
        builder.AppendLine("set-image-position updates anchored DrawingML relative positions, offsets, and alignments.");
        builder.AppendLine("set-image-crop updates DrawingML a:srcRect crop percentages without replacing media bytes.");
        builder.AppendLine("Linked images and unsupported drawing shapes are reported as diagnostics and preserved.");
        builder.AppendLine("Track changes are controlled by check/apply --track-changes off|preserve|suggest|require.");
        builder.AppendLine("Tracked edits preserve unrelated existing markup, allow adjacent revisions, and reject overlaps with tracked insert/delete, move, custom XML revision, comment, bookmark, content-control, field, or hyperlink boundaries.");
        builder.AppendLine("Unsupported fields are rejected. expect-hash and preserve-size are not supported.");
        builder.AppendLine("Target lifetime: explicit paragraph/table/row/cell/section IDs bind to the input snapshot for one patch, so an earlier insert or delete never renumbers a later explicit ID; a deleted target fails instead of editing a neighbour, and newly inserted blocks are not addressable by pre-discovered IDs in the same patch. Semantic selectors resolve live; guards evaluate sequentially.");
        return builder.ToString();
    }

    /// <summary>Renders focused help for one patch operation from the shared catalog.</summary>
    private static string RenderPatchOperationHelp(DocxPatchOperationInfo operation)
    {
        var builder = new StringBuilder();
        builder.Append("op ").Append(operation.Name).Append(" (").Append(operation.Category).AppendLine(")");
        builder.AppendLine();
        builder.AppendLine(operation.Description);
        builder.Append("Required fields: ").AppendLine(RenderTopicFields(operation, required: true));
        builder.Append("Optional fields: ").AppendLine(RenderTopicFields(operation, required: false));
        if (operation.EmptyAllowedFields.Count != 0)
        {
            builder.Append("Empty values allowed: ").AppendLine(string.Join(", ", operation.EmptyAllowedFields));
        }
        if (operation.ExclusiveAlternatives.Count != 0)
        {
            builder.Append("Exclusive fields (at most one): ").AppendLine(string.Join(", ", operation.ExclusiveAlternatives.Select(static group => string.Join("|", group))));
        }
        if (operation.FieldDefaults.Count != 0)
        {
            builder.Append("Defaults: ").AppendLine(string.Join(", ", operation.FieldDefaults));
        }
        if (operation.AllowedValues.Count != 0)
        {
            builder.Append("Allowed values: ").AppendLine(string.Join(", ", operation.AllowedValues));
        if (operation.UnitHints.Count != 0)
        {
            builder.Append("Units: ").AppendLine(string.Join(", ", operation.UnitHints));
        }
        }
        builder.Append("Track-change support: ").Append(operation.TrackChangesSupportClass).Append(" / ").Append(operation.TrackChangesSupport).Append(" - ").AppendLine(operation.TrackChangesNote);
        foreach (string example in operation.Examples)
        {
            builder.AppendLine();
            builder.AppendLine(example);
        }
        return builder.ToString();
    }

    private static string RenderTopicFields(DocxPatchOperationInfo operation, bool required)
    {
        var tokens = new List<string>((required ? operation.RequiredFields : operation.OptionalFields)
            .Select(field => operation.RepeatableFields.Contains(field, StringComparer.Ordinal) ? field + "+" : field));
        if (required)
        {
            tokens.AddRange(operation.RequiredAlternatives.Select(group => string.Join("|", group)));
        }

        return tokens.Count == 0 ? "none" : string.Join(", ", tokens);
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
        int width = Math.Max(14, fields.Select(static field => field.Name.Length).DefaultIfEmpty(0).Max() + 1);
        foreach (DocxOutputFieldInfo field in fields)
        {
            builder.Append("  ").Append(field.Name.PadRight(width)).AppendLine(field.Description);
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
            builder.Append("  ").Append(option.Syntax.PadRight(Math.Max(28, options.Select(static option => option.Syntax.Length).DefaultIfEmpty(0).Max() + 1))).AppendLine(option.Description);
        }
    }

    private static void AppendLines(
        StringBuilder builder,
        string heading,
        IReadOnlyList<string> lines)
    {
        AppendLines(builder, heading, lines, indent: "");
    }

    private static void AppendLines(
        StringBuilder builder,
        string heading,
        IReadOnlyList<string> lines,
        string indent)
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
                    MinPositionals = 1,
                    MaxPositionals = 1,
                    Category = "read",
                    Summary = "Produce an agent-friendly structural view of a .docx",
                    Usage = "docxedit read input.docx [--summary] [--headers-footers] [--view final|original|markup] [--max-text <chars>] [--json] [--compact] [--diagnostics <path>] [--strict]",
                    Description = "Produce an agent-friendly structural view of a .docx.",
                    Options =
                    [
                        new("--summary", "Print compact counts without broad document text") { Flags = ["--summary"] },
                        new("--headers-footers", "Include header/footer stories") { Flags = ["--headers-footers"] },
                        new("--view final|original|markup", "Text view for tracked insert/delete text") { Flags = ["--view"] },
                        new("--max-text N", "Maximum text per rendered field") { Flags = ["--max-text"] },
                        new("--json", "Print the result object as JSON") { Flags = ["--json"] },
                        new("--compact", "Print JSON without indentation") { Flags = ["--compact"] },
                        new("--diagnostics path", "Write diagnostics JSON") { Flags = ["--diagnostics"] },
                        new("--strict", "Return 3 when warnings are present") { Flags = ["--strict"] }
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
                            "Paragraph records may include StyleId, StyleName, and resolved List metadata. List metadata reports concrete numbering ID, zero-based level, abstract numbering ID, format, level text, paragraph style link, source=style or source=style-inherited when numbering comes from style inheritance, visible labels, and structured label components for deterministic decimal, letter, roman, bullet, and nested lvlText patterns. List counters follow the selected text view for block-level inserted/deleted numbered paragraphs. Picture bullets, unsupported custom numbering formats, and tracked numbering property revisions are preserved and reported through diagnostics."
                        ]),
                        new("Bookmarks and content controls",
                        [
                            "Read and context output surface bookmark names/ranges and content-control metadata, including duplicate selector candidate IDs, placeholder/data-binding details, hierarchy IDs, safe-edit status/reason, checkbox state, dropdown/combo item counts, repeating-section metadata, and date settings/value. Bookmark and content-control selectors can be used for safe paragraph targeting; use explicit IDs when selector diagnostics report multiple candidate targets. Patch operations can update plain-text content-control IDs, guarded rich-text content controls, checkbox state, dropdown/combo selections, date values, create guarded paragraph bookmarks, update bookmark names, remove unreferenced bookmark markers, and replace guarded paragraph-bounded bookmark ranges, including multi-run, multi-paragraph, and simple table-spanning text-slot ranges, while preserving wrappers/markers/table structure; content-control edits reject controls whose w:lock value is present and not unlocked, reject picture controls with image/media guidance, and reject group controls with guidance to target editable child controls."
                        ]),
                        new("Fields",
                        [
                            "Read and context output surface simple and complex field metadata, including parsed field type, normalized field code, parsed arguments/switches, refresh policy/reason, cached result text/length, nesting depth, REF/PAGEREF/NOTEREF bookmark dependencies, HYPERLINK URI/anchor dependencies, safe-edit status, dirty flags, and lock flags. Patch operations can set field dirty/lock flags on one existing field or on target all, update simple-field code/result caches, refresh simple REF/PAGEREF/NOTEREF cached results from unambiguous same-part bookmarks, refresh simple QUOTE literal fields, and mark the document for field updates after other edits, but Word remains responsible for general field recalculation. Unsupported refresh diagnostics classify layout, document-property, formula, mail-merge, date/time, conditional, and external-state fields. Apply emits W5103 when a document containing fields was marked for Word-side refresh."
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
                            "Read and context output surface table style, caption, description, grid/header/merged/nested flags, row grid-before/grid-after/header/cant-split metadata, cell logical/physical column positions, merge group IDs, visual column ends, and vertical-merge root cells. set-cell targets visual-grid cell IDs or merge-group IDs; visual columns inside a horizontal span resolve to the spanning cell while vertical-merge continuation cells are rejected. Table metadata operations can set or clear caption/description values. Direct row operations can clone consistent visual-grid row shapes and delete safe vertical-merge rows; generated tracked row revisions remain limited to simple rectangular tables."
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
                    MinPositionals = 1,
                    MaxPositionals = 1,
                    Category = "read",
                    Summary = "Show headings, tables, images, sections, headers, footers",
                    Usage = "docxedit outline input.docx [--headers-footers] [--view final|original|markup] [--max-text <chars>] [--json] [--compact] [--diagnostics <path>] [--strict]",
                    Description = "Show headings, tables, images, sections, headers, and footers.",
                    Options =
                    [
                        new("--headers-footers", "Include header/footer stories") { Flags = ["--headers-footers"] },
                        new("--view final|original|markup", "Text view for tracked insert/delete text") { Flags = ["--view"] },
                        new("--max-text N", "Maximum text per rendered field") { Flags = ["--max-text"] },
                        new("--json", "Print the result object as JSON") { Flags = ["--json"] },
                        new("--compact", "Print JSON without indentation") { Flags = ["--compact"] },
                        new("--diagnostics path", "Write diagnostics JSON") { Flags = ["--diagnostics"] },
                        new("--strict", "Return 3 when warnings are present") { Flags = ["--strict"] }
                    ],
                    Notes =
                    [
                        new("Numbered headings",
                        [
                            "Heading lines include compact resolved list metadata when the heading paragraph has numbering. The selected text view controls whether block-level inserted or deleted headings participate in the outline and list-label counters."
                        ])
                    ]
                },
                new()
                {
                    Name = "find",
                    MinPositionals = 2,
                    MaxPositionals = 2,
                    Category = "read",
                    Summary = "Find text and print stable edit targets",
                    Usage = "docxedit find input.docx \"text\" [--headers-footers] [--view final|original|markup] [--max-text <chars>] [--json] [--compact] [--diagnostics <path>] [--strict]",
                    Description = "Find text and print stable edit targets.",
                    Options =
                    [
                        new("--headers-footers", "Include header/footer stories") { Flags = ["--headers-footers"] },
                        new("--view final|original|markup", "Text view for tracked insert/delete text") { Flags = ["--view"] },
                        new("--max-text N", "Maximum text per rendered field") { Flags = ["--max-text"] },
                        new("--json", "Print the result object as JSON") { Flags = ["--json"] },
                        new("--compact", "Print JSON without indentation") { Flags = ["--compact"] },
                        new("--diagnostics path", "Write diagnostics JSON") { Flags = ["--diagnostics"] },
                        new("--strict", "Return 3 when warnings are present") { Flags = ["--strict"] }
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
                    MinPositionals = 1,
                    MaxPositionals = 1,
                    Category = "read",
                    Summary = "Dump one target in detail",
                    Usage = "docxedit dump input.docx --id M.P0001 [--runs] [--view final|original|markup] [--max-text <chars>] [--json] [--compact] [--diagnostics <path>] [--strict]",
                    Description = "Dump one target by stable ID. Paragraph and cell dumps print visible text. Comment body targets such as C001.C0001 or comment:3 print metadata only, never comment body text. With --runs, paragraph dumps include run-level markup metadata such as markup=inserted-run, markup=deleted-run, revision-id, author, timestamp-utc, comment-id, and comment range/reference markers without printing comment body text. For targets that own tracked markup records, dump appends a privacy-safe changes block with change IDs, types, parent type, revision metadata, and child element counts; this is how property revisions on paragraph, table, row, cell, and section targets can be inspected without raw OOXML. With --json, the Runs array exposes run metadata as structured fields. Change IDs from `changes` identify markup records; run IDs from dump identify rendered run/marker lines and are not the same namespace. In JSON output, inspect Runs[] for structured run metadata.",
                    Options =
                    [
                        new("--id M.P0001", "Target ID from read, outline, find, or changes") { Flags = ["--id"] },
                        new("--runs", "Include paragraph run lines and markup metadata") { Flags = ["--runs"] },
                        new("--view final|original|markup", "Text view for tracked insert/delete text") { Flags = ["--view"] },
                        new("--max-text N", "Maximum text per rendered field") { Flags = ["--max-text"] },
                        new("--json", "Print the result object as JSON") { Flags = ["--json"] },
                        new("--compact", "Print JSON without indentation") { Flags = ["--compact"] },
                        new("--diagnostics path", "Write diagnostics JSON") { Flags = ["--diagnostics"] },
                        new("--strict", "Return 3 when warnings are present") { Flags = ["--strict"] }
                    ],
                    OutputFields =
                    [
                        new("Text", "Rendered target text plus target-scoped change metadata when present"),
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
                    MinPositionals = 1,
                    MaxPositionals = 1,
                    Category = "read",
                    Summary = "Show nearby structure around one target without broad text",
                    Usage = "docxedit context input.docx --id M.P0001 [--headers-footers] [--radius <count>] [--view final|original|markup] [--max-text <chars>] [--json] [--compact] [--diagnostics <path>] [--strict]",
                    Description = "Summarize nearby modeled structure around one target without broad document text. By default, --max-text is 0, so paragraph and cell text fields are present but empty. Comment anchors are surfaced as comment IDs, comment body IDs, para IDs, durable IDs, reply IDs, resolved IDs, and parent/root para IDs; comment body targets such as C001.C0001 or comment:3 return metadata only. Increase --max-text only when short snippets are needed.",
                    Options =
                    [
                        new("--id M.P0001", "Target ID from read, outline, find, or changes") { Flags = ["--id"] },
                        new("--radius N", "Number of same-kind neighbors to include") { Flags = ["--radius"] },
                        new("--headers-footers", "Include header/footer stories") { Flags = ["--headers-footers"] },
                        new("--view final|original|markup", "Text view when --max-text is greater than 0") { Flags = ["--view"] },
                        new("--max-text N", "Maximum text per paragraph/cell; default 0") { Flags = ["--max-text"] },
                        new("--json", "Print the result object as JSON") { Flags = ["--json"] },
                        new("--compact", "Print JSON without indentation") { Flags = ["--compact"] },
                        new("--diagnostics path", "Write diagnostics JSON") { Flags = ["--diagnostics"] },
                        new("--strict", "Return 3 when warnings are present") { Flags = ["--strict"] }
                    ],
                    PrivacyNotes =
                    [
                        "The default metadata-only context profile uses MaxText=0.",
                        "Comment body targets do not print comment body text; use changes --include-comment-text only when a bounded snippet is explicitly needed.",
                        "Threaded-comment metadata is exposed as IDs and topology fields, not as comment body text."
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
                    Name = "capabilities",
                    MinPositionals = 1,
                    MaxPositionals = 1,
                    Category = "read",
                    Summary = "Show supported, conditional, and unsupported edits for one target",
                    Usage = "docxedit capabilities input.docx --id M.P0001 [--track-changes <mode>] [--json] [--compact] [--diagnostics <path>] [--strict]",
                    Description = "Describe supported, conditional, and unsupported edits for one paragraph, content-control, cell, or merge-group target under an effective track-change policy. Each operation carries an actionable reason, a help topic, and a safe alternative when one exists. Output is metadata-only guidance: capabilities can change after an edit, and check remains authoritative for a patch.",
                    Options =
                    [
                        new("--id M.P0001", "Target ID from read, outline, find, or changes") { Flags = ["--id"] },
                        new("--track-changes off|preserve|suggest|require", "Effective track-change policy the capabilities describe") { Flags = ["--track-changes"] },
                        new("--json", "Print the result object as JSON") { Flags = ["--json"] },
                        new("--compact", "Print JSON without indentation") { Flags = ["--compact"] },
                        new("--diagnostics path", "Write diagnostics JSON") { Flags = ["--diagnostics"] },
                        new("--strict", "Return 3 when warnings are present") { Flags = ["--strict"] }
                    ],
                    PrivacyNotes =
                    [
                        "Capabilities output carries no document body text; reasons name only short OOXML boundary, kind, and lock values."
                    ],
                    OutputFields =
                    [
                        new("TargetId", "Requested target echoed byte-identical"),
                        new("Capabilities", "Resolved target kind and story plus per-operation support, reason, help topic, and alternative")
                    ],
                    Examples =
                    [
                        "docxedit capabilities report.docx --id M.P0004",
                        "docxedit capabilities report.docx --id M.T0001.R02.C03 --track-changes require --json"
                    ]
                },
                new()
                {
                    Name = "template",
                    MinPositionals = 1,
                    MaxPositionals = 1,
                    Category = "read",
                    Summary = "Print a guarded patch template for one target",
                    Usage = "docxedit template input.docx --id M.P0001 [--track-changes <mode>] [--json] [--compact] [--diagnostics <path>] [--strict]",
                    Description = "Print a guarded patch template for one paragraph, content-control, cell, or merge-group target. The active block passes check under the requested policy and changes nothing visible; remaining supported operations follow as commented blocks. Templates are generated from target-specific capabilities, so unsupported operations are omitted and conditional ones carry their condition.",
                    Options =
                    [
                        new("--id M.P0001", "Target ID from read, outline, find, or changes") { Flags = ["--id"] },
                        new("--track-changes off|preserve|suggest|require", "Effective track-change policy the template is generated under") { Flags = ["--track-changes"] },
                        new("--json", "Print the result object as JSON") { Flags = ["--json"] },
                        new("--compact", "Print JSON without indentation") { Flags = ["--compact"] },
                        new("--diagnostics path", "Write diagnostics JSON") { Flags = ["--diagnostics"] },
                        new("--strict", "Return 3 when warnings are present") { Flags = ["--strict"] }
                    ],
                    PrivacyNotes =
                    [
                        "Templates embed the current target text in expect-text guards; they are working material for one target, not broad document text."
                    ],
                    OutputFields =
                    [
                        new("TargetId", "Requested target echoed byte-identical"),
                        new("Template", "Patch template text with one check-clean starter block and commented examples"),
                        new("Capabilities", "Target capabilities the template was generated from")
                    ],
                    Examples =
                    [
                        "docxedit template report.docx --id M.P0004",
                        "docxedit template report.docx --id M.T0001.R02.C03 --track-changes require"
                    ]
                },
                new()
                {
                    Name = "styles",
                    MinPositionals = 1,
                    MaxPositionals = 1,
                    Category = "read",
                    Summary = "List paragraph, character, and table styles",
                    Usage = "docxedit styles input.docx [--json] [--compact] [--diagnostics <path>] [--strict]",
                    Description = "List paragraph, character, and table styles. Output includes inheritance links such as based-on, next, linked, and style-level numbering defaults when present.",
                    Options =
                    [
                        new("--json", "Print the result object as JSON") { Flags = ["--json"] },
                        new("--compact", "Print JSON without indentation") { Flags = ["--compact"] },
                        new("--diagnostics path", "Write diagnostics JSON") { Flags = ["--diagnostics"] },
                        new("--strict", "Return 3 when warnings are present") { Flags = ["--strict"] }
                    ]
                },
                new()
                {
                    Name = "media",
                    MinPositionals = 1,
                    MaxPositionals = 1,
                    Category = "read",
                    Summary = "List embedded images",
                    Usage = "docxedit media input.docx [--extract <dir>] [--json] [--compact] [--diagnostics <path>] [--strict]",
                    Description = "List embedded images.",
                    Options =
                    [
                        new("--extract dir", "Extract embedded image parts to a directory") { Flags = ["--extract"] },
                        new("--json", "Print the result object as JSON") { Flags = ["--json"] },
                        new("--compact", "Print JSON without indentation") { Flags = ["--compact"] },
                        new("--diagnostics path", "Write diagnostics JSON") { Flags = ["--diagnostics"] },
                        new("--strict", "Return 3 when warnings are present") { Flags = ["--strict"] }
                    ]
                },
                new()
                {
                    Name = "validate",
                    MinPositionals = 1,
                    MaxPositionals = 1,
                    Category = "read",
                    Summary = "Validate package and WordprocessingML invariants",
                    Usage = "docxedit validate input.docx [--profile structural|package] [--max-diagnostics <count>] [--json] [--compact] [--diagnostics <path>] [--strict]",
                    Description = "Validate package-level XML roots and, with the structural profile, WordprocessingML invariants: paired bookmark/comment ranges, duplicate semantic selectors, commentsExtended consistency, complex field balance/result containment, content-control metadata, tracked revision markup, paragraph style and numbering references, settings updateFields, header/footer references, section properties, drawing relationships, image target/content-type checks, drawing geometry, basic table shape, and table visual-grid consistency. This is layered DocxEdit structural validation, not full ISO/IEC 29500 schema validation; no strict schema profile is exposed. Package profile limits validation to package/XML root checks. Diagnostics are capped and report E9199 or W9199 when omitted.",
                    Options =
                    [
                        new("--profile structural|package", "Validation profile; structural is the default") { Flags = ["--profile"] },
                        new("--max-diagnostics N", "Maximum diagnostics to return; default 500") { Flags = ["--max-diagnostics"] },
                        new("--json", "Print the result object as JSON") { Flags = ["--json"] },
                        new("--compact", "Print JSON without indentation") { Flags = ["--compact"] },
                        new("--diagnostics path", "Write diagnostics JSON") { Flags = ["--diagnostics"] },
                        new("--strict", "Return 3 when warnings are present") { Flags = ["--strict"] }
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
                    MinPositionals = 1,
                    MaxPositionals = 1,
                    Category = "read",
                    Summary = "List tracked-change and comment markup; comment text is opt-in",
                    Usage = "docxedit changes input.docx [--operation-report <path>] [--include-comment-text] [--max-comment-text <chars>] [--json] [--compact] [--diagnostics <path>] [--strict]",
                    Description = "List tracked-change and comment markup. By default this does not print revision text or comment body text. Records include change IDs, type, story, part, normalized parent type, target, revision/comment metadata, text length, child element count, and comment anchor targets when known. Use --operation-report with a check/apply report to annotate generated revisions with the originating operation. Use --include-comment-text only when short comment body snippets are needed.",
                    Options =
                    [
                        new("--operation-report path", "Annotate generated revisions with operation index/name/target from a check/apply report JSON (use the report from the run that produced the scanned document)") { Flags = ["--operation-report"] },
                        new("--include-comment-text", "Include explicit comment body snippets in comment summaries and comment body records") { Flags = ["--include-comment-text"] },
                        new("--max-comment-text N", "Maximum comment body snippet length when --include-comment-text is used; default 240") { Flags = ["--max-comment-text"] },
                        new("--json", "Print the result object as JSON") { Flags = ["--json"] },
                        new("--compact", "Print JSON without indentation") { Flags = ["--compact"] },
                        new("--diagnostics path", "Write diagnostics JSON") { Flags = ["--diagnostics"] },
                        new("--strict", "Return 3 when warnings are present") { Flags = ["--strict"] }
                    ],
                    OutputFields =
                    [
                        new("Summary", "Counts by change type"),
                        new("GroupSummary", "Counts by group=story|part|author|target and type"),
                        new("TargetSummary", "Compact per-target type rollups"),
                        new("CommentSummary", "Compact per-comment anchor/type rollups; may include TextSnippet only when requested"),
                        new("Changes", "Individual change records; may include OperationIndex/OperationName/OperationTarget when --operation-report is supplied, and CommentTextSnippet only when requested")
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
                            "Text output includes type counts, group summaries, target summaries, comment summaries, and individual records. Individual records include parent=<kind> when the immediate WordprocessingML parent can be normalized. With --operation-report, records whose revision-id appears in the report add operation-index, operation-name, and operation-target. Modern Word comment metadata appears as para-id, parent-para-id, root-para-id, durable-id, is-reply, resolved, comment-para-id, comment-parent-para-id, comment-root-para-id, comment-durable-id, comment-is-reply, and comment-resolved when commentsExtended or commentsIds data is present. With --include-comment-text, comment-summary and comment body records add comment-text-length, comment-text, and comment-text-truncated when applicable."
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
                        "docxedit changes report.edited.docx --operation-report apply-report.json",
                        "docxedit changes report.docx --include-comment-text --max-comment-text 120"
                    ]
                },
                new()
                {
                    Name = "catalog",
                    MinPositionals = 0,
                    MaxPositionals = 0,
                    Category = "read",
                    Summary = "Print the machine-readable command and patch-operation catalog",
                    Usage = "docxedit catalog [--json] [--compact]",
                    Description = "Print the structured docxedit surface: commands with options and output fields, plus patch operations with required/optional fields and track-change support. Text output prints the overview and the operation support table; --json prints the full catalog object.",
                    Options =
                    [
                        new("--json", "Print the catalog object as JSON") { Flags = ["--json"] },
                        new("--compact", "Print JSON without indentation") { Flags = ["--compact"] },
                    ],
                    Examples =
                    [
                        "docxedit catalog",
                        "docxedit catalog --json",
                    ]
                },
                new()
                {
                    Name = "version",
                    MinPositionals = 0,
                    MaxPositionals = 0,
                    Category = "read",
                    Summary = "Print the docxedit version",
                    Usage = "docxedit version [--json] [--compact]",
                    Description = "Print the library version, target framework, and patch-operation count. Text output prints one line; --json prints the version object.",
                    Options =
                    [
                        new("--json", "Print the version object as JSON") { Flags = ["--json"] },
                        new("--compact", "Print JSON without indentation") { Flags = ["--compact"] },
                    ],
                    Examples =
                    [
                        "docxedit version",
                    ]
                },
                new()
                {
                    Name = "lint",
                    MinPositionals = 1,
                    MaxPositionals = 1,
                    Category = "patch",
                    Summary = "Validate patch shape without loading a document",
                    Usage = "docxedit lint edits.docxpatch [--json] [--compact] [--diagnostics <path>] [--strict]",
                    Description = "Validate patch syntax, required fields, alternative groups, and exclusive fields without loading a document. Lint failures always predict check failures; lint success leaves document-dependent targets, guards, assets, and shapes to check.",
                    Options =
                    [
                        new("--json", "Print the result object as JSON") { Flags = ["--json"] },
                        new("--compact", "Print JSON without indentation") { Flags = ["--compact"] },
                        new("--diagnostics path", "Write diagnostics JSON") { Flags = ["--diagnostics"] },
                        new("--strict", "Return 3 when warnings are present") { Flags = ["--strict"] }
                    ],
                    OutputFields =
                    [
                        new("OperationCount", "Number of parsed operations; zero when parsing failed"),
                        new("Operations", "Parsed operation indexes and names in file order")
                    ],
                    Examples =
                    [
                        "docxedit lint edits.docxpatch",
                        "docxedit lint edits.docxpatch --json"
                    ]
                },
                new()
                {
                    Name = "check",
                    MinPositionals = 2,
                    MaxPositionals = 2,
                    Category = "patch",
                    Summary = "Validate a .docxpatch file without writing output",
                    Usage = "docxedit check input.docx edits.docxpatch [--track-changes <mode>] [--author <name>] [--timestamp-utc <instant>] [--json] [--compact] [--report <path>] [--diagnostics <path>] [--max-preview-chars <n>] [--strict]",
                    Description = "Validate a patch against an input document without writing an output file. Use check before apply to verify selectors, guards, assets, and track-change constraints. Text output includes one operation line per patch operation and affected row/cell lines for table operations, including visual-grid, merge-group, and nested-table metadata when relevant.",
                    Options =
                    [
                        new("--track-changes off|preserve|suggest|require", "Tracked-change handling mode") { Flags = ["--track-changes"] },
                        new("--author name", "Non-empty author used for generated revisions; defaults to docxedit") { Flags = ["--author"] },
                        new("--timestamp-utc instant", "Timestamp in ISO-8601 format normalized to UTC for generated revisions (e.g. 2026-01-01T00:00:00Z)") { Flags = ["--timestamp-utc"] },
                        new("--json", "Print the result object as JSON") { Flags = ["--json"] },
                        new("--compact", "Print JSON without indentation") { Flags = ["--compact"] },
                        new("--report path", "Write operation report JSON") { Flags = ["--report"] },
                        new("--diagnostics path", "Write diagnostics JSON") { Flags = ["--diagnostics"] },
                        new("--max-preview-chars n", "Bounded before/after preview text per operation report; 0 disables previews") { Flags = ["--max-preview-chars"] },
                        new("--strict", "Return 3 when warnings are present") { Flags = ["--strict"] }
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
                    MinPositionals = 2,
                    MaxPositionals = 2,
                    Category = "patch",
                    Summary = "Apply a .docxpatch file and write a new .docx",
                    Usage = "docxedit apply input.docx edits.docxpatch --output output.docx [--track-changes <mode>] [--author <name>] [--timestamp-utc <instant>] [--json] [--compact] [--report <path>] [--diagnostics <path>] [--max-preview-chars <n>] [--strict]",
                    Description = "Apply a patch and write a new .docx. The input is never modified in place. Text output includes one operation line per patch operation, generated revision IDs when tracked markup is created, and affected row/cell lines for table operations, including visual-grid, merge-group, and nested-table metadata when relevant.",
                    Options =
                    [
                        new("--output path, -o path", "Output .docx path") { Flags = ["--output", "-o"] },
                        new("--track-changes off|preserve|suggest|require", "Tracked-change handling mode") { Flags = ["--track-changes"] },
                        new("--author name", "Non-empty author used for generated revisions; defaults to docxedit") { Flags = ["--author"] },
                        new("--timestamp-utc instant", "Timestamp in ISO-8601 format normalized to UTC for generated revisions (e.g. 2026-01-01T00:00:00Z)") { Flags = ["--timestamp-utc"] },
                        new("--json", "Print the result object as JSON") { Flags = ["--json"] },
                        new("--compact", "Print JSON without indentation") { Flags = ["--compact"] },
                        new("--report path", "Write operation report JSON") { Flags = ["--report"] },
                        new("--diagnostics path", "Write diagnostics JSON") { Flags = ["--diagnostics"] },
                        new("--max-preview-chars n", "Bounded before/after preview text per operation report; 0 disables previews") { Flags = ["--max-preview-chars"] },
                        new("--strict", "Return 3 when warnings are present") { Flags = ["--strict"] }
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
                "docxedit outline report.docx --view markup",
                "docxedit dump report.docx --id M.P0004 --runs",
                "docxedit context report.docx --id M.P0004",
                "docxedit media report.docx --extract media",
                "docxedit validate report.docx",
                "docxedit changes report.docx",
                "docxedit check report.docx edits.docxpatch",
                "docxedit apply report.docx edits.docxpatch --output report.edited.docx"
            ],
            PatchOperations = DocxPatchEngine.CatalogOperations,
        };
    }
}
