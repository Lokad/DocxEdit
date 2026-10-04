# Native Word equations

DocxEdit discovers and edits whole Office Math equations (OMML). The output is
editable with Word's equation editor. Conversion and editing run without Word or
external packages; Word is used only by the optional rendering smoke test.

## Discovery and extraction

`read`, `outline`, `dump`, and `context` expose equation IDs such as `M.E0001`.
Header/footer equations use `H001.E0001`/`F001.E0001`; include those stories with
`--headers-footers`. Equations inside table cells are also discovered. IDs count
physical `m:oMath` elements, including hidden revisions, so changing the text view
does not renumber surviving targets. Legacy Equation Editor OLE objects remain
preserve-only objects, outside this model.

The library's `DocxReadResult.Equations` and `DocxDumpResult.Equation` include:

- `Id`, story, part, containing paragraph/cell, and `IsDisplay`.
- `Text`: concatenated math text for preview; it is **not** a lossless formula.
- `Omml`: standalone native `m:oMath` XML for extraction. If the complete XML exceeds
  `MaxText`, this is null; increase the limit to extract it. It is never truncated
  into malformed XML. It retains stored revision markup regardless of the selected
  text view. `MaxText = 0` suppresses both text and OMML.
- `ContentHash`: a SHA-256 content guard, ignoring namespace prefixes/declarations,
  indentation outside text, and internal patch bookkeeping. Native formatting is
  included. Word may normalize formatting on save, requiring a fresh guard.

For example, `docxedit dump input.docx --id M.E0001 --json --max-text 100000`
returns the native equation in the `Equation.Omml` property.
Paragraph/cell plain text continues to exclude math text; use the equation list
and containing IDs to distinguish formulas from prose. `find` searches prose,
not mathematical structure. `W1011` reports whole-equation support and the absence
of internal-run/tracked-equation editing.

## Operations

See the [patch fields and example](patch-format.md#equations).
`insert-equation` accepts paragraph/table anchors. Its default `placement after`
creates a display equation in a new paragraph; `placement inline` appends an inline
equation at the end of a paragraph. Insertion into arbitrary character offsets or
table cells is not yet supported. Existing equations inside cells can be replaced
or deleted by their equation IDs.

`replace-equation` changes only the selected equation's content, retaining its
inline/display position. `delete-equation` removes that equation and an empty
display wrapper, retaining the Word paragraph and its layout properties. Both
accept `expect-hash` copied from discovery. A mismatch fails with `E3201`.

The `as` field binds a newly created equation for later operations in the same
patch. Explicit IDs bind to the input snapshot, so deleting an earlier equation
cannot redirect a later operation to its neighbor. Reports provide final IDs;
deleted targets have no final ID.

These operations do not generate tracked equation revisions. Track mode Off
applies directly; Suggest applies directly with `W4001`; Require fails with
`E6001`. Equations carrying protected anchors/revisions, inside content controls,
complex fields, or unsupported wrappers are inspectable but cannot be rewritten.
Ordinary text replacement cannot silently remove equations, including forced
`set-cell` and container rewrites. Run-preserving `replace-text` may edit prose
on either side, but cannot match across the intervening equation.

`capabilities` reports equation support under the selected policy. Equation
`template` output contains commented guarded examples: native OMML is not
decompiled to LaTeX, so there is no automatic unchanged-expression starter.
Raw OMML import and partial equation editing are outside this initial surface.

## LaTeX-like expression syntax

Use a math expression without `$`, `\(`, or document wrappers. Braces group
expressions. Ordinary whitespace is ignored; `\text{...}` retains literal spaces.
Patch fields trim surrounding whitespace; heredocs retain it. This is a bounded
expression language, not a TeX engine: macros, packages, file access, and arbitrary
TeX commands are unavailable. Unknown commands and malformed input fail with
`E4205`, the `latex` field's patch line, and a character position in the expression.

| Construct | Example |
| --- | --- |
| Subscript and superscript | `x_i^2`, `x_{i+1}^{n-1}` |
| Fraction | `\frac{a+b}{c}` (`\dfrac` and `\tfrac` are aliases without size overrides) |
| Root | `\sqrt{x}`, `\sqrt[3]{x}` |
| Greek | `\alpha`, `\beta`, `\theta`, `\pi`, `\Sigma`, `\Omega` |
| Relations/operators | `\pm`, `\times`, `\cdot`, `\leq`, `\geq`, `\neq`, `\approx`, `\infty` |
| Scalable delimiters | `\left(\frac{x}{y}\right)`, `\left\langle x\right\rangle`, `\left\{x\right.` |
| Sum/product/integral | `\sum_{i=1}^{n}{i^2}`, `\prod_{i=1}^{n}{i}`, `\int_0^1{x^2}\,\mathrm{d}x` |
| Matrix | `\begin{pmatrix}a&b\\c&d\end{pmatrix}` |
| Cases | `\begin{cases}x&\text{if }x\geq0\\-x&\text{otherwise}\end{cases}` |
| Literal text/styles | `\text{for all }`, `\mathrm{d}`, `\mathbf{x}`, `\mathit{x}` |
| Upright names | `\sin`, `\cos`, `\tan`, `\log`, `\ln`, `\exp` |

Sums, products and integrals require a **braced body**, including a single symbol:
`\sum_{i=1}^{n}{i}`. This makes the scope explicit. Limits are optional.
Matrix environments are `matrix`, `pmatrix`, `bmatrix`, and `cases`; rows must have
equal cell counts, with `&` between cells and `\\` between rows. Empty cells are
allowed. `\,`, `\:`, `\;`, and backslash-space insert a space, without TeX's
different spacing widths. Escaped braces and `_ % # & $ |` are supported.

The full symbol vocabulary is defined in
[`LatexMath.cs`](../src/Lokad.DocxEdit/Ooxml/LatexMath.cs); literal Unicode symbols
also work. Source is limited to 32,768 characters, parser nesting to 64 levels,
and matrices to 1,024 cells. These limits fail before document mutation.

## Word/PDF smoke test

On Windows with desktop Word installed, run:

```powershell
pwsh -NoProfile -File tools/Test-EquationsWithWord.ps1
```

This builds the CLI, creates a synthetic formula sheet, replaces an equation and
deletes another, checks the edit patch, exports both documents to PDF through local
Word, saves Word round-trip copies, validates those copies, and compares native
equation counts. Each Word worker has a 60-second deadline. Outputs and logs go
under ignored `artifacts/equation-smoke/`. Existing user documents are never opened.
Inspect both PDFs to verify fractions, radical bars, limit placement, matrices,
inline baselines, and missing glyphs. This is a bounded compatibility smoke test,
not a guarantee of identical rendering across Word versions or installed fonts.

The implementation follows [Microsoft's Office Math model](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.officemath).
