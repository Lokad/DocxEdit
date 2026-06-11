# Patch Format

Patch files start with `docxpatch 1` and then one or more operation blocks.

```text
docxpatch 1

op replace-text
target M.P0001
expect-text <<<
current text
>>>
find current
with updated
end
```

Fields are line based. Heredocs use `<<<` and close with `>>>`. Repeated fields are preserved for operations such as table row edits.

## Selectors

Explicit IDs remain the most stable selectors:

- `M.P0001`: main paragraph
- `H001.P0001` / `F001.P0001`: header/footer paragraph
- `M.T0001`, `M.T0001.R02`, `M.T0001.R02.C03`: table, row, cell
- `H001.T0001`, `H001.T0001.R02`, `H001.T0001.R02.C03`: header table, row, cell
- `F001.T0001`, `F001.T0001.R02`, `F001.T0001.R02.C03`: footer table, row, cell
- `M.I0001`: inline main-document image
- `M.S0001`: main-document section

Paragraph operations also support:

- `heading:"Exact heading"`
- `heading:2:"Exact heading"`
- `text:"contained paragraph text"`
- `bookmark:"BookmarkName"`
- `content-control:"TagOrAlias"`

Ambiguous selectors fail with `E1202`; use a more specific selector or an explicit ID.
Parsed selectors that match no target fail with `E1201` and include nearby target IDs without
including nearby paragraph text.

## Supported Operations

- `replace-text`: `target`, `find`, `with`, optional `expect-text`, `preserve-runs`, `occurrence`. Under `TrackChangesMode.Suggest` or `Require`, simple text-only replacements are emitted as tracked `w:del`/`w:ins` markup with the configured author and timestamp.
- `replace-paragraph`: `target`, `text`, optional `expect-text`, `style`
- `insert-before`, `insert-after`: `target`, `text`, optional `style`
- `delete-block`: `target`, optional `expect-text`
- `set-style`: `target`, `style`
- `set-cell`: `target`, `text`, optional `expect-text`, `expect-row-count`, `expect-column-count`, `force`
- `append-row`: `target`, repeated `cell`, optional `expect-row-count`, `expect-column-count`
- `insert-row-before`, `insert-row-after`: `target`, repeated `cell`, optional `expect-row-count`, `expect-column-count`, `expect-cell-count`, `force`
- `delete-row`: `target`, optional `expect-row-count`, `expect-column-count`, `expect-cell-count`, `expect-contains`, `force`
- `replace-image`: `target`, `asset`, optional `expect-content-type`, `alt`, `preserve-size`
- `insert-image-after`: `target`, `asset`, optional `expect-content-type`, `width`, `height`, `alt`, `caption`
- `set-image-alt`: `target`, `alt`, optional `expect-content-type`
- `delete-image`: `target`, optional `expect-content-type`
- `set-section-columns`: `target`, `count`, optional `expect-columns`, `expect-orientation`
- `set-section-orientation`: `target`, `orientation`, optional `expect-columns`, `expect-orientation`

Explicit header/footer paragraph IDs can be used for paragraph text/style edits, block insertion/deletion, and image insertion after the paragraph. Explicit header/footer table IDs can be used for simple table edits and as block insertion/deletion anchors.

Parsed but not fully implemented fields include `preserve-size` and `caption`. `expect-hash` is rejected.
