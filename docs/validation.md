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

This is structural validation, not full OOXML schema validation.

## Public Edit Cases

Tracked public cases live under `edit-cases/cases/` as JSON manifests. A case defines synthetic
WordprocessingML body XML, a `.docxpatch` payload, and expected readback values. The harness
generates `.docx` inputs at runtime, so fixtures remain reviewable text.

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
