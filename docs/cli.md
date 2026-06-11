# CLI

Run locally with:

```powershell
dotnet run --project src/DocxEdit.Cli/DocxEdit.Cli.csproj -- <command> [options]
```

## Read And Explore

- `read input.docx [--headers-footers] [--all-stories] [--view final|original|markup] [--max-text N] [--json] [--diagnostics path] [--strict]`
- `outline input.docx [--headers-footers] [--json] [--diagnostics path] [--strict]`
- `find input.docx "text" [--headers-footers] [--view final|original|markup] [--max-text N] [--json] [--diagnostics path] [--strict]`
- `dump input.docx --id M.P0001 [--runs] [--view final|original|markup] [--max-text N] [--json] [--diagnostics path] [--strict]`
- `styles input.docx [--json] [--diagnostics path] [--strict]`
- `media input.docx [--extract dir] [--json] [--diagnostics path] [--strict]`
- `changes input.docx [--json] [--diagnostics path] [--strict]`

`changes` lists existing tracked-change and comment markup without printing revision or comment text. It reports counts, IDs, type, story, part, target, revision/comment metadata, text length, and child element count.
Read text views are `final` (default), `original`, and `markup`. Markup view includes inserted and deleted text with lightweight `[+text+]` and `[-text-]` markers.

## Patch

- `check input.docx edits.docxpatch [--track-changes mode] [--author name] [--timestamp-utc instant] [--json] [--report path] [--diagnostics path] [--strict]`
- `apply input.docx edits.docxpatch -o output.docx [--track-changes mode] [--author name] [--timestamp-utc instant] [--json] [--report path] [--diagnostics path] [--strict]`

Track-change modes are `off`, `preserve`, `suggest`, and `require`. `replace-text` generates simple `w:del`/`w:ins` tracked-change markup with the selected author and timestamp under `suggest` or `require`. `require` fails unsupported operations; `suggest` warns and applies unsupported operations directly.

Exit codes:

- `0`: success
- `1`: operation failed
- `2`: invalid CLI usage
- `3`: strict mode saw warnings or errors
- `4`: unexpected CLI exception
