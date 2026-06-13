# Validation

Validation currently has three layers.

## Public API

The public API normalizes expected package, ZIP, XML, and relationship failures into diagnostics. Programmer errors such as null arguments and unreadable streams still throw.

## Patch Apply

`apply` validates touched XML before saving:

- `[Content_Types].xml` root;
- relationship part root and internal targets;
- main document root and `w:body`;
- touched header and footer roots.
- touched PNG/JPEG media parts are non-empty.

This is structural validation, not full OOXML schema validation.

## Structural Package Validation

`DocxEditor.Validate` and `docxedit validate input.docx` run validation without
editing the document. The result includes the selected validation profile, package
part names, the main document part name, and diagnostics with part names when a
problem is tied to a specific OOXML part.

Supported profiles:

- `structural` is the default. It runs package/XML root checks plus the
  WordprocessingML invariants below.
- `package` limits validation to package/XML root checks after the package has been
  loaded safely.

Current checks include:

- expected roots for known WordprocessingML parts such as the main document, styles,
  numbering, settings, comments, commentsExtended, headers, and footers;
- paired bookmark and comment range start/end IDs;
- comment body, anchor, and reference ID consistency;
- modern `commentsExtended.xml` paraId consistency;
- complex field begin/separate/end balance, instruction-text containment,
  cached-result containment, and dirty/lock flag validity;
- DrawingML `a:blip` relationship references, image relationship target parts, and
  image media content types;
- duplicate DrawingML `wp:docPr` IDs within a Word part;
- DrawingML extent and crop geometry;
- basic table row/cell shape;
- table visual-grid consistency for `gridSpan`, `gridBefore`, `gridAfter`,
  `vMerge` continuations, and declared `tblGrid` width.

This remains layered internal validation rather than full ISO/IEC 29500 schema
validation. A strict schema profile is not exposed until an actual schema validator
bridge exists. The current profiles are designed to catch common corruption and
relationship mistakes with stable `E91xx` diagnostics.

## Public Edit Cases

Tracked public cases live under `edit-cases/cases/` as JSON manifests. A case defines synthetic
WordprocessingML body XML or a non-private fixture under `edit-cases/fixtures/`, a
`.docxpatch` payload, and expected readback values. The harness generates synthetic
`.docx` inputs at runtime when body XML is supplied, so most fixtures remain reviewable text.

Supported manifest extras include:

- `input.headerXml` and `input.footerXml` for header/footer stories;
- `input.fixture` for a checked-in non-private `.docx` under `edit-cases/fixtures/`;
- `assets` values referenced from patches as `{{asset:name.png}}`;
- `applyOptions.trackChanges`, `author`, and `timestampUtc`;
- expectations for `paragraphs`, `allStoryParagraphs`, `tableCells`, `sections`,
  `images`, and `changeSummary`.

The CLI resolves image operation assets from local file paths. Relative paths are
resolved from the current working directory; public manifests use harness-generated
absolute asset paths.

Run one case:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/CheckDocxCase.ps1 -Case basic-replace
```

Run every public case:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/ValidateDocxCases.ps1
```

Artifacts are written under ignored `artifacts/edit-cases/`.

## Office Compatibility

Optional Microsoft Word automation tests are gated by `DOCXEDIT_ENABLE_OFFICE_TESTS=1`.
On Windows with Word installed, the Office test project creates a synthetic `.docx`, applies
a patch, opens and saves the generated output through Word COM automation, then re-reads it.

## Private Cases

Private inputs must live under ignored `private-cases/` and must never be committed. The private harness rejects cases outside `private-cases/`, git-tracked private files, non-ignored private files, and unsafe case IDs.

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/CheckPrivateCase.ps1 -Case local-case.docx -ValidateOnly
```

The harness writes sanitized artifacts under ignored `artifacts/private-edit/` and prints only structure counts, tracked-change counts, diagnostics count, and status. It must not print or store private document text.

An ignored private manifest may sit next to the input and provide aggregate-only
expectations:

```json
{
  "id": "local-case",
  "input": "local-case.docx",
  "expected": {
    "Structure": { "Paragraphs": 10, "Tables": 2 },
    "Changes": { "Total": 3, "ByType": { "inserted-run": 2 } }
  }
}
```

Expected values are compared to sanitized counts only. Text content is neither printed
nor stored in the summary artifact.

## Agent Challenges

Agent challenges live under `agent-challenges/challenges/`. They are usability probes
for fresh Codex agents, not deterministic regression cases. They launch `codex exec`
against an ignored run directory containing a private input copy, a local `docxedit.ps1`
wrapper, and a generated prompt.

List probes:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/RunAgentChallenge.ps1 -List
```

Preview a prompt without launching Codex:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/RunAgentChallenge.ps1 -Challenge markup-inventory -DryRun
```

Prepare the ignored run directory without launching Codex:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/RunAgentChallenge.ps1 -Challenge markup-inventory -PrepareOnly
```

Run a probe:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/RunAgentChallenge.ps1 -Challenge markup-inventory
```

Artifacts are written under ignored `artifacts/agent-challenges/`. The runner captures
Codex JSONL events, the final response, post-run `docxedit` read/changes checks for
generated outputs, and heuristic signals such as whether the agent used `docxedit`,
`changes`, `check`, `apply`, or forbidden raw OOXML inspection. These artifacts may
contain private text if the evaluated agent printed it, so do not commit or publish them.
