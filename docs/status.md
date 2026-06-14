# Status

This is a pre-release editor. It supports useful structural reads and a focused set of edits, but it is not a complete WordprocessingML implementation.

Implemented areas include:

- safe ZIP/package loading without filesystem extraction;
- stable IDs for paragraphs, tables, cells, images, sections, headers, and footers;
- resolved paragraph numbering/list metadata, including abstract numbering IDs, formats,
  level text, style-linked list sources, start/suffix metadata, visible labels, and
  structured label components for deterministic decimal, letter, roman, bullet, and
  nested `lvlText` patterns, stable across final/original/markup run-level tracked
  text views and view-aware for block-level inserted/deleted numbered paragraphs;
- merged/nested/styled table and row/grid read metadata, including merge groups and
  vertical-merge root cells;
- styles, media, outline, find, dump, and tracked-change markup summaries;
- bookmark and content-control selectors for safe paragraph targeting;
- patch operations for paragraph text, blocks, styles, simple main/header/footer tables,
  inline images and image metadata, and basic sections;
- no-content tracked-change and comment markup viewing through `changes`;
- target-scoped `dump` change summaries for tracked markup records, including
  property revisions, without raw OOXML or revision text;
- opt-in bounded comment body snippets through `changes`;
- comment creation on modeled paragraphs or selected direct text spans inside one
  paragraph, plus comment body editing and deletion
  through explicit patch operations keyed by paragraph IDs, `comment:<id>`, or
  comment body IDs;
- comment anchor/context metadata, privacy-safe threaded-comment context
  topology, and metadata-only comment body `dump`/`context` targets;
- comment root/reply and resolution metadata from `commentsExtended.xml`, durable
  comment IDs from `commentsIds.xml`, plus
  `resolve-comment` and `reopen-comment` workflows that create or update modern extension records for
  basic comments, simple threaded reply add/delete workflows, and explicit
  `E4314` diagnostics when reply deletion would change child thread topology;
- bookmark and content-control read/context metadata with selector guidance, including
  duplicate selector candidate IDs, placeholder/data-binding metadata, hierarchy IDs,
  safe-edit status, checkbox state, dropdown/combo item counts, repeating-section
  metadata, and date settings;
- lock-aware safe bookmark/content-control patch operations for plain-text content
  controls, guarded rich-text content controls, checkbox state, dropdown/combo
  selections, date values, guarded paragraph bookmark creation, bookmark
  rename/delete, and guarded paragraph-bounded bookmark ranges, including
  multi-run and multi-paragraph ranges;
- explicit `E4315` diagnostics for unsupported repeating-section item insertion
  and deletion attempts;
- simple and complex field read/context metadata, including parsed field type,
  cached result text/length, nesting depth, bookmark/hyperlink dependencies,
  safe-edit status, explicit dirty/lock flag patch operations for one field or all modeled fields,
  simple-field code/result patching, simple REF/PAGEREF/NOTEREF cached-result
  refresh from unambiguous bookmarks, field-update marking after edits, and `W5103`
  diagnostics for Word-side refresh;
- hyperlink read/context/dump metadata for external links, internal anchors, relationship
  IDs/parts/target modes, broken relationship IDs, URI scheme validation,
  missing/duplicate internal anchors, unsupported internal part links, target frames,
  history flags, and aggregate diagnostics for invalid URI/anchor states;
- hyperlink patch operations for target URI/anchor updates, tooltip/frame/history updates,
  display text updates, insertion, and unlinking while preserving display runs;
- check/apply operation reports with affected row/cell summaries for table edits,
  visual-grid/merge/nested-table metadata, and generated revision IDs for apply
  operations that create tracked markup;
- table property patch operations for table style, caption/description metadata,
  row repeating-header flags, and cell shading fills;
- table-cell text replacement by visual-grid cell ID or merge-group ID while
  preserving cell properties and horizontal spans;
- direct row insertion/deletion for consistent visual-grid tables by cloning row
  shape metadata, with vertical-merge root deletion promoting the next
  continuation when spans match;
- explicit `E4316` diagnostics for unsupported table-column insertion and deletion
  attempts;
- inline and anchored image metadata with layout kind, size, wrap mode, wrap distances,
  anchor positioning, aspect-lock, crop percentages, alt text, and containing target;
- safe image patch operations for media replacement, alt/title/name metadata, extents,
  anchored wrap mode/distances, anchored positioning, and crop percentages;
- tracked-change output for simple text replacement, whole-paragraph replacement,
  paragraph insertion/deletion, paragraph style changes, simple plain-text
  content controls, guarded paragraph-only rich-text content controls, simple
  same-paragraph bookmark text, simple paragraph-only comment body text, simple
  field result text, simple hyperlink display text and insertion, and simple
  table-cell text replacement, plus table style, row-header, and cell-shading
  property revisions,
  simple table row insertion/deletion structure revisions, and section
  column/orientation property revisions, with author, timestamp, and revision IDs;
- operation-level track-change capability metadata in the shared help catalog,
  including tracked and preserve-only operation classifications, with `W4001` and
  `E6001` diagnostics that report the catalog support value;
- operation-specific preserve-only rationales in the shared help catalog so agents
  can distinguish metadata/review-markup operations from tracked shapes that are
  simply not modeled yet;
- tracked-capable target-shape diagnostics through `W4002` and `E6002` that report
  the operation name, target ID, catalog support value, and unsupported-shape reason;
- machine-readable track-change diagnostic metadata that separates missing revision
  representation from unsupported target shape and direct-edit fallback from
  require-mode failure;
- generated revision metadata normalization, including trimmed non-empty authors,
  invariant UTC timestamps, and early `E6003` rejection before output is written;
- structural/package validation profiles through `validate`, including known part roots, paired
  ranges, complex field balance/result-containment/flag consistency,
  content-control metadata consistency, missing paragraph style-reference warnings,
  numbering reference warnings,
  settings metadata consistency, comment body/anchor consistency, comment extension consistency, drawing relationships,
  header/footer section-reference consistency,
  section property consistency,
  tracked revision metadata and basic revision nesting,
  image relationship target/content-type checks, drawing property ID uniqueness,
  drawing extent/crop geometry, duplicate semantic selector warnings, basic table
  shape, table visual-grid consistency, and capped validation diagnostics;
- post-edit validation for touched XML parts before writing output.

Known limits include exotic/custom visual numbering expansion beyond deterministic
decimal, letter, roman, bullet, and nested `lvlText` labels/components, previous
numbering-state reconstruction for paragraph property revisions flagged with
`W1026`, arbitrary multi-paragraph or protected-boundary comment selections and
threaded comment workflows beyond simple reply add/delete, advanced
bookmark/content-control editing beyond plain-text and guarded rich-text controls,
checkbox toggles, dropdown/combo selections, date values, guarded paragraph
bookmark creation, bookmark rename/delete, and guarded paragraph-bounded bookmark
ranges, field result recalculation and complex-field
edit workflows beyond dirty/lock flags, simple-field code/result edits, and simple
REF-style bookmark refresh,
advanced hyperlink edit workflows beyond target URI/anchor/text/tooltip/frame/history,
relative/UNC/file hyperlink target support, advanced floating-image edit operations beyond
size/wrap/position/crop metadata, linked images and VML/grouped/chart/SmartArt/OLE
drawing shapes beyond diagnostics-only preservation, complex
merged/nested table edit transformations beyond visual cell/merge-root text,
safe visual-grid row cloning/deletion, simple tracked row, and metadata edits,
table-column transforms beyond explicit unsupported diagnostics,
full tracked-change edit coverage, and full
ISO/IEC 29500 schema validation beyond DocxEdit's layered structural invariants.

## OOXML Shape Inventory

This inventory is public and synthetic. It records supported and intentionally
unsupported WordprocessingML shapes without private document details.

### Comments

- `comments.xml` comment bodies are modeled through `changes`, `dump`, and
  `context` without exposing body text by default.
- `commentsExtended.xml` `w15:commentEx` records are modeled for para IDs,
  parent/root para IDs, reply classification, and resolved state.
- `commentsIds.xml` durable IDs are modeled and validated. Simple reply creation
  and leaf reply deletion are supported; deeper threaded-comment workflows remain
  guarded.
- `context` annotates anchored comments and simple threaded replies with comment
  IDs, body IDs, para IDs, durable IDs, reply IDs, resolved IDs, and parent/root
  para IDs without exposing body text.
- Comment anchors, references, body IDs, duplicate IDs, orphan anchors, and
  orphan extension records are validated.
- Comment creation, body replacement, resolution toggles, reopening, and deletion
  are supported for basic comments in supported document stories. Creation can
  target a whole paragraph or one direct text span inside a paragraph. Simple
  threaded replies can be added and leaf replies can be deleted.

### Content Controls

- Plain-text controls support guarded text replacement while preserving `w:sdt`
  and `w:sdtPr`.
- Rich-text controls support guarded paragraph-only replacement when
  `expect-text` matches and protected boundaries are absent.
- Date controls support metadata/display updates.
- Dropdown and combo box controls support guarded selection by value or display
  text.
- Checkbox controls support checked-state updates and displayed symbol updates.
- Picture, group, repeating-section, and repeating-section-item controls are
  surfaced as metadata with safe-edit reasons. Picture-control edit attempts fail
  with image/media guidance; group-control edit attempts fail with guidance to
  target editable child controls. Repeating-section item edit operations are
  recognized and fail with `E4315`; this remains intentional until subtree
  cloning/deletion can preserve content-control IDs, bindings, and section
  boundaries.
- Tags, aliases, placeholders, data bindings, lock values, list items, checkbox
  symbols, date metadata, parent/child relationships, duplicate semantic
  selectors, safe-edit status, and safe-edit reason are surfaced.

### Bookmarks

- Complete paragraph-bounded bookmark ranges are modeled and can be renamed,
  deleted when unreferenced, and replaced when protected boundaries are absent.
- Same-paragraph replacements support ranges spanning multiple direct run
  siblings. Same-container multi-paragraph replacements are supported. Simple
  table-spanning replacements are supported when table structure can be preserved
  and replacement text provides one line per visible text slot. Tracked output is
  available for simple same-paragraph replacements only.
- Whole-paragraph bookmark creation is supported with `expect-text`, duplicate
  name checks, and protected-boundary checks.
- Duplicate bookmark names are surfaced with candidate IDs so agents can switch
  to explicit selectors.
- Cross-story and protected-boundary bookmark replacement remains unsupported;
  unsupported table-spanning shapes fail before writing output.
- Hidden Word bookmarks and incomplete ranges are surfaced as metadata and
  validation diagnostics where pairing is malformed.

### Fields

- Simple `w:fldSimple` fields expose normalized code, parsed type, cached result
  text/length, dirty/lock flags, dependencies, and safe-edit status.
- Complex begin/separate/end fields expose normalized code, parsed type, cached
  result text/length, parsed arguments/switches, refresh policy/reason, nesting
  depth, dependencies, dirty/lock flags, completeness, and safe-edit status.
- Nested complex fields are modeled; outer cached results include visible nested
  field results.
- REF/PAGEREF/NOTEREF dependencies are surfaced. HYPERLINK field URI/anchor
  dependencies are surfaced.
- Simple field code and cached result replacement are supported with guards.
  Simple same-paragraph complex-field result replacement is supported when the
  begin/separate/end result topology is validated.
- Simple REF-style cached results can be refreshed from one unambiguous
  same-part, same-paragraph bookmark. Simple `QUOTE` cached results can be
  refreshed from literal field-code arguments.
- General field recalculation, TOC rebuilding, PAGE pagination, formulas,
  DOCPROPERTY, MERGEFIELD, IF, DATE/TIME, complex-field code rewriting, and
  unsafe complex-field result topologies remain Word-side or unsupported workflows
  with explicit `E4313` refresh/edit diagnostics. DocxEdit can mark fields dirty
  and request Word-side update.
- Malformed field boundaries, duplicate separators, orphan instruction text,
  invalid field-char types, invalid dirty/lock flags, and invalid result
  containment are validated.
