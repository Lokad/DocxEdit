using System.Text;
using Lokad.DocxEdit.Ooxml;

namespace Lokad.DocxEdit;

// Shared by focused operation help, the JSON catalog, and help equations.
// Keep the high-impact syntax differences beside the latex field, so callers
// do not need repository documentation to author a valid expression.
internal static class EquationHelp
{
    internal static IReadOnlyList<DocxHelpSection> InputNotes { get; } =
    [
        new("Equation input",
        [
            @"Bounded LaTeX-like expressions, not full TeX. Omit $...$, \(...\), and document wrappers. Unknown commands fail with E4205.",
            @"Groups/scripts/fractions/roots: {a+b}, x_i^2, \frac{a}{b}, \sqrt{x}, \sqrt[3]{x}. Unicode and named symbols such as \alpha, \pi, \leq are supported.",
            @"Sums/products/integrals REQUIRE a braced body: \sum_{i=1}^{n}{i^2}, \prod_{i=1}^{n}{i}, \int_0^1{x^2}. Limits are optional; \sum_{i=1}^{n} i^2 is invalid.",
            @"Matrices: \begin{pmatrix}a&b\\c&d\end{pmatrix}. Environments: matrix, pmatrix, bmatrix, cases. Use & for cells, \\ for rows; equal cell counts per row.",
            @"Delimiters/text/styles: \left(\frac{x}{y}\right), \text{if }x\geq0, \mathrm{d}, \mathbf{x}, \mathit{x}. Ordinary whitespace is ignored; \text{...} preserves spaces.",
            "Literal backslashes in patch files; JSON/shell escaping is separate. Multi-line field: latex <<<, expression lines, >>> on separate lines.",
            $"Limits: {LatexMath.MaxSourceLength} source characters, {LatexMath.MaxNestingDepth} parser nesting levels, {LatexMath.MaxMatrixCells} matrix cells. Full command list and workflow: docxedit help equations."
        ])
    ];

    internal static IReadOnlyList<DocxHelpSection> TargetNotes { get; } =
    [
        new("Equation targets",
        [
            "Use E IDs from read/outline (M.E0001, H001.E0001, F001.E0001), or @name bound by insert-equation as name. IDs bind to the patch input snapshot.",
            "For expect-hash, copy ContentHash from read input.docx --json or Equation.ContentHash from dump input.docx --id M.E0001 --json. Preview Text is not round-trip LaTeX. See docxedit help equations."
        ])
    ];

    internal static string Render()
    {
        var builder = new StringBuilder();
        builder.AppendLine("Equations - native editable Word Office Math (OMML)");
        builder.AppendLine();
        builder.AppendLine("Patch operations (one field per line):");
        builder.AppendLine("  insert-equation target <paragraph/table> latex <expression> [placement after|inline] [as name]");
        builder.AppendLine("    after (default): a new display paragraph; inline: append to a paragraph only.");
        builder.AppendLine("  replace-equation target <E ID/@name> latex <expression> [expect-hash <ContentHash>]");
        builder.AppendLine("    Replaces the whole equation and keeps its inline/display placement.");
        builder.AppendLine("  delete-equation target <E ID/@name> [expect-hash <ContentHash>]");
        builder.AppendLine("    Removes the equation and retains the containing Word paragraph.");
        foreach (DocxHelpSection section in TargetNotes.Concat(InputNotes))
        {
            builder.AppendLine().Append(section.Heading).AppendLine(":");
            foreach (string line in section.Lines) builder.Append("  ").AppendLine(line);
        }
        builder.AppendLine();
        builder.AppendLine("Additional supported syntax:");
        builder.AppendLine(@"  \dfrac and \tfrac alias \frac without size overrides. Upright names: \sin \cos \tan \log \ln \exp.");
        builder.AppendLine(@"  \left/\right delimiters: ( ) [ ] | \| \{ \} \langle \rangle; . means no delimiter.");
        builder.AppendLine(@"  Escapes: \{ \} \_ \% \# \& \$ \|. Spacing: \, \: \; and backslash-space all insert a space.");
        builder.AppendLine(@"  \text{...} has literal content (no nested braces/commands); escaped braces, backslash, and _ % # & $ work inside it.");
        builder.AppendLine("  Empty matrix cells are allowed. Macros, packages, raw OMML input, and partial equation edits are unsupported.");
        builder.AppendLine();
        builder.AppendLine("Named symbols (case-sensitive; literal Unicode also works):");
        foreach (string[] commands in LatexMath.SymbolCommands.Order(StringComparer.Ordinal).Chunk(10))
            builder.Append("  ").AppendLine(string.Join(" ", commands.Select(command => "\\" + command)));
        builder.AppendLine();
        builder.AppendLine("Minimal patch (save as equation.docxpatch):");
        builder.AppendLine("""
            docxpatch 1
            op insert-equation
            target M.P0001
            latex \sum_{i=1}^{n}{i^2}
            end
            """);
        builder.AppendLine();
        builder.AppendLine("Check and apply:");
        builder.AppendLine("  docxedit check input.docx equation.docxpatch");
        builder.AppendLine("  docxedit apply input.docx equation.docxpatch --output output.docx");
        builder.AppendLine("Off edits directly; Suggest warns W4001; Require fails E6001 (no tracked equation revisions).");
        builder.AppendLine("Protected equation markup/wrappers fail E4305. Use capabilities for the target; check is authoritative.");
        builder.AppendLine("read/dump --json expose native Omml, omitted if the complete XML exceeds --max-text; it is not decompiled to LaTeX.");
        builder.AppendLine("Use --headers-footers for header/footer discovery. Legacy Equation Editor OLE objects remain preserve-only.");
        return builder.ToString();
    }
}
