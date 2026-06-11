# Architecture

DocxEdit is organized around a small dependency-free library plus a CLI harness.

## Package Layer

`OoxmlPackage` loads ZIP entries into memory, normalizes package paths, parses content types and relationships, rejects unsafe paths and missing internal targets, enforces size limits, and preserves unknown safe parts. Mutations mark touched parts for post-edit validation.

## Model Layer

Scanners build focused read models for paragraphs, tables, cells, images, sections, styles, and existing tracked-change markup. Text reads use a final-view approximation: inserted text is visible, deleted and move-from text is skipped.

## Patch Layer

The parser produces structured operation blocks. The engine resolves selectors, validates guards, applies supported edits to `XDocument` parts, updates package parts and relationships, and runs touched-part validation before save.

## Public API And CLI

`DocxEditor` is stream-first and returns result objects with success flags, diagnostics, structured data, and operation reports. The CLI is a thin local harness over the public API with text and JSON modes.
