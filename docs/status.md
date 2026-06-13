# Status

This is a pre-release editor. It supports useful structural reads and a focused set of edits, but it is not a complete WordprocessingML implementation.

Implemented areas include:

- safe ZIP/package loading without filesystem extraction;
- stable IDs for paragraphs, tables, cells, images, sections, headers, and footers;
- resolved paragraph numbering/list metadata, including abstract numbering IDs, formats,
  level text, style-linked list sources, start/suffix metadata, visible labels, and
  structured label components for deterministic decimal, letter, roman, bullet, and
  nested `lvlText` patterns, stable across final/original/markup run-level tracked
  text views;
- merged/nested/styled table and row/grid read metadata, including merge groups and
  vertical-merge root cells;
- styles, media, outline, find, dump, and tracked-change markup summaries;
- bookmark and content-control selectors for safe paragraph targeting;
- patch operations for paragraph text, blocks, styles, simple main/header/footer tables,
  inline images and image metadata, and basic sections;
- no-content tracked-change and comment markup viewing through `changes`;
- opt-in bounded comment body snippets through `changes`;
- comment creation on modeled paragraphs, plus comment body editing and deletion
  through explicit patch operations keyed by paragraph IDs, `comment:<id>`, or
  comment body IDs;
- comment anchor/context metadata and metadata-only comment body `dump`/`context`
  targets;
- comment root/reply and resolution metadata from `commentsExtended.xml`, plus
  `resolve-comment` and `reopen-comment` workflows that create or update modern extension records for
  basic comments, and explicit `E4314` diagnostics for unsupported threaded reply
  operations;
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
- check/apply operation reports with affected row/cell summaries for table edits;
- table property patch operations for table style, caption/description metadata,
  and row repeating-header flags;
- explicit `E4316` diagnostics for unsupported table-column insertion and deletion
  attempts;
- inline and anchored image metadata with layout kind, size, wrap mode, wrap distances,
  anchor positioning, aspect-lock, crop percentages, alt text, and containing target;
- safe image patch operations for media replacement, alt/title/name metadata, extents,
  anchored wrap mode/distances, anchored positioning, and crop percentages;
- tracked-change output for simple text replacement, whole-paragraph replacement,
  paragraph insertion/deletion, paragraph style changes, and simple table-cell text
  replacement with author, timestamp, and revision IDs;
- operation-level track-change capability metadata in the shared help catalog,
  including tracked and preserve-only operation classifications, with `W4001` and
  `E6001` diagnostics that report the catalog support value;
- structural/package validation profiles through `validate`, including known part roots, paired
  ranges, complex field balance/result-containment/flag consistency,
  content-control metadata consistency, missing paragraph style-reference warnings,
  numbering reference warnings,
  settings metadata consistency, comment body/anchor consistency, comment extension consistency, drawing relationships,
  header/footer section-reference consistency,
  section property consistency,
  image relationship target/content-type checks, drawing property ID uniqueness,
  drawing extent/crop geometry, duplicate semantic selector warnings, basic table
  shape, table visual-grid consistency, and capped validation diagnostics;
- post-edit validation for touched XML parts before writing output.

Known limits include exotic/custom visual numbering expansion beyond deterministic
decimal, letter, roman, bullet, and nested `lvlText` labels/components, selected-range comment
creation and full threaded comment models, advanced
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
merged/nested table edit transformations beyond simple cell/row and metadata edits,
table-column transforms beyond explicit unsupported diagnostics,
full tracked-change edit coverage, and full
ISO/IEC 29500 schema validation beyond DocxEdit's layered structural invariants.
