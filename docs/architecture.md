# Architecture

DocxEdit is a stream-first .NET library plus a thin CLI harness, with no NuGet dependencies beyond
the platform libraries. Data flows one way: streams enter, a buffered package is inspected or edited
in memory, and results or a new package leave.

## Data flow

1. DocxEditor (src/Lokad.DocxEdit/Public/DocxEditor.cs) accepts input streams plus an options object
   and returns result records; it never exposes the package. The patch path adds a patch reader:
   parse (DocxPatchParser) produces DocxPatch, and DocxPatchEngine executes it against a disposable
   in-memory package.
2. OoxmlPackage (src/Lokad.DocxEdit/Ooxml/OoxmlPackage.cs) owns ZIP IO and part bytes: load
   validates paths, relationships, content types, and quotas; mutations mark touched parts; save
   writes parts back.
3. Scanners (src/Lokad.DocxEdit/Model/Docx*Scanner.cs) build focused read models for paragraphs,
   tables, images, sections, styles, and tracked-change markup. Validators
   (DocxPackageValidator*.cs) check structural rules and report diagnostics without throwing for
   content.
4. Renderers (src/Lokad.DocxEdit/Rendering/TextRenderers.cs, Public/DocxTextRenderer.cs) format
   result records as agent-readable text; JSON consumers use the records directly.
5. The CLI (src/Lokad.DocxEdit.Cli/Program.cs) maps flags to options, publishes file outputs
   atomically, and owns stdout and stderr discipline. DocxHelp and CommandCatalog own command and
   patch metadata; the docs directory mirrors it for readers.
Ownership rule: each package invariant and each accepted command or patch field has exactly one
owner above. The operation registry (DocxPatchOperationRegistry.cs) is the single table that the
parser, catalog, and engine dispatch all project from.

## Component invariants

- Target identity (scanners, resolvers, and the change map via Model/DocxStoryBlocks.cs): IDs
  enumerate physical document order and resolve to the same element across read, dump, context,
  find, changes, and patch resolution. They are stable for the same bytes, not permanently stable.
  Within one patch, explicit paragraph/table/row/cell/section IDs bind to the input snapshot:
  an earlier insert or delete never renumbers a later explicit ID, a deleted target fails
  instead of editing a neighbour, and newly inserted blocks are not addressable by
  pre-discovered IDs in the same patch. Semantic selectors (text:, heading:, ...)
  keep resolving live against current content. Guards still evaluate sequentially, so
  dependent replacements can assert state produced by a preceding operation.
- Check and apply parity (DocxPatchEngine.cs): both paths execute the same mutations against the
  disposable package; only apply saves. Check reports keep GeneratedRevisionIds empty by design.
- Report coordinates (DocxPatchEngine.cs): AffectedTargets carry historical identity plus lifetime: Coordinate names input or operation-time identity, FinalId carries the live final ID matching a fresh read and stays absent for deleted objects, so a historical ordinal that coincides with a survivor never reads as live. CreatedTargetIds carry live final IDs only; vanished creations, including comments with reused numbers, are absent. XML snapshot and creation marks stay internal. Result aliases (@name) always resolve to the current element. Re-read the output for final coordinates; check and apply agree on every per-operation report.
- Range-structure safety (DocxPatchEngine.cs): deleting a block or row that would newly orphan a
  healthy bookmark, comment range, field pair, move range, custom-XML range, or table interior is refused (E4305) before publication; removing a
  whole range stays allowed. Post-edit validation reports only violations the patch introduced.
- Revision identity (DocxPatchEngine.cs): one shared allocator per execution hands out revision IDs;
  comment and bookmark id writes invalidate it so later revisions rescan above them.
- Part roles (Model/DocxPartRoles.cs): known parts are discovered through package relationships,
  never by fixed path (fixed paths are used only when creating parts). Shared header and footer
  parts are scanned once.
- Quotas, cancellation, and ownership: load and save enforce size quotas at ZIP, XML, asset, and
  edited-package boundaries; cancellation is honored between parts, operations, and scans; inputs
  and outputs are disposed exactly when the matching Leave flag is false. Patch readers stay callerowned; opened asset streams are engine-owned.
- Malformed input: scanners tolerate missing parts and skip unrecognized entries, while malformed
  document data throws a document exception at an explicit boundary that the editor converts into a
  failed result with diagnostics. Validators report content problems as diagnostics.
- Images: magic bytes decide the content type; hints and extensions must agree; PNG and JPEG
  structure is validated within bounds.

## Processing model

The public surface is stream-first: callers hand over readable streams (seekable or not) and receive
result objects or a written package. Inside, the package is fully buffered: every part lives in
memory, nothing borrows the input after load returns, and non-seekable inputs are copied once under
quota. Cancellation is checked between parts, operations, and scans, not mid-read.

## Preservation scope

Edits preserve package parts, not ZIP bytes: unknown safe parts round-trip, part order on save is
normalized by entry name, compression is re-applied, and entry timestamps record the save. Two saves
of equal payloads therefore differ byte-for-byte; compare part payloads or rendered text (see SPEC
23.2). Structural validation covers roots, revision markup, bookmark, comment, and field pairing,
table shape, styles and numbering references, and comment consistency; it does not judge layout,
pagination, or Word rendering fidelity.

## Current limitations

These hold today and are not roadmap promises: Word open and save compatibility is checked only by
the opt-in Office project, never by default runs; layout and pagination are out of scope; there are
no checked-in golden files (fixtures assert readback values); archive bytes are not reproducible;
the CLI is invoked from source, not installed as a tool. Breaking public API changes may still occur
before 1.0 when they make the supported surface clearer or safer.
