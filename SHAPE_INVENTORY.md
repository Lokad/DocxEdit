# DocxEdit OOXML Shape Inventory

This inventory is public and synthetic. It records supported and intentionally
unsupported WordprocessingML shapes without private document details.

## Comments

- `comments.xml` comment bodies are modeled through `changes`, `dump`, and
  `context` without exposing body text by default.
- `commentsExtended.xml` `w15:commentEx` records are modeled for para IDs,
  parent/root para IDs, reply classification, and resolved state.
- `commentsIds` and full threaded-comment reply bodies are not safely modeled.
  Reply operations are recognized and fail with `E4314`.
- Comment anchors, references, body IDs, duplicate IDs, orphan anchors, and
  orphan extension records are validated.
- Comment creation, body replacement, resolution toggles, reopening, and deletion
  are supported for basic comments in supported document stories.

## Content Controls

- Plain-text controls support guarded text replacement while preserving `w:sdt`
  and `w:sdtPr`.
- Rich-text controls support guarded paragraph-only replacement when
  `expect-text` matches and protected boundaries are absent.
- Date controls support metadata/display updates.
- Dropdown and combo box controls support guarded selection by value or display
  text.
- Checkbox controls support checked-state updates and displayed symbol updates.
- Picture, group, repeating-section, and repeating-section-item controls are
  surfaced as metadata and safe-edit diagnostics. Repeating-section item edit
  operations are recognized and fail with `E4315`.
- Tags, aliases, placeholders, data bindings, lock values, list items, checkbox
  symbols, date metadata, parent/child relationships, duplicate semantic
  selectors, and safe-edit status are surfaced.

## Bookmarks

- Complete paragraph-bounded bookmark ranges are modeled and can be renamed,
  deleted when unreferenced, and replaced when protected boundaries are absent.
- Same-paragraph replacements support ranges spanning multiple direct run
  siblings. Same-container multi-paragraph replacements are supported.
- Whole-paragraph bookmark creation is supported with `expect-text`, duplicate
  name checks, and protected-boundary checks.
- Duplicate bookmark names are surfaced with candidate IDs so agents can switch
  to explicit selectors.
- Table-spanning, cross-story, and protected-boundary bookmark replacement remains
  unsupported.
- Hidden Word bookmarks and incomplete ranges are surfaced as metadata and
  validation diagnostics where pairing is malformed.

## Fields

- Simple `w:fldSimple` fields expose normalized code, parsed type, cached result
  text/length, dirty/lock flags, dependencies, and safe-edit status.
- Complex begin/separate/end fields expose normalized code, parsed type, cached
  result text/length, nesting depth, dependencies, dirty/lock flags,
  completeness, and safe-edit status.
- Nested complex fields are modeled; outer cached results include visible nested
  field results.
- REF/PAGEREF/NOTEREF dependencies are surfaced. HYPERLINK field URI/anchor
  dependencies are surfaced.
- Simple field code and cached result replacement are supported with guards.
- Simple REF-style cached results can be refreshed from one unambiguous
  same-part, same-paragraph bookmark.
- General field recalculation, TOC rebuilding, formulas, DOCPROPERTY, MERGEFIELD,
  and complex-field code/result rewriting remain Word-side or unsupported
  workflows. DocxEdit can mark fields dirty and request Word-side update.
- Malformed field boundaries, duplicate separators, orphan instruction text,
  invalid field-char types, invalid dirty/lock flags, and invalid result
  containment are validated.
