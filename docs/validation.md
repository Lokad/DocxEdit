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

## Private Cases

Private inputs must live under ignored `private-cases/` and must never be committed. The private harness rejects cases outside `private-cases/`, git-tracked private files, non-ignored private files, and unsafe case IDs.

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/CheckPrivateCase.ps1 -Case local-case.docx -ValidateOnly
```

The harness writes sanitized artifacts under ignored `artifacts/private-edit/` and prints only structure counts, tracked-change counts, diagnostics count, and status. It must not print or store private document text.
