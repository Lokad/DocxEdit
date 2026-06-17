# Architecture

DocxEdit is organized around a small dependency-free library plus a CLI harness.

## Package Layer

`OoxmlPackage` loads ZIP entries into memory, normalizes package paths, parses content types and relationships, rejects unsafe paths and missing internal targets, enforces size limits, and preserves unknown safe parts. Mutations mark touched parts for post-edit validation.

## Model Layer

Scanners build focused read models for paragraphs, tables, cells, images, sections, styles, and existing tracked-change markup. Text reads support final, original, and lightweight markup views for tracked inserted/deleted text.

## Patch Layer

The parser produces structured operation blocks. The engine resolves selectors, validates guards, applies supported edits to `XDocument` parts, updates package parts and relationships, and runs touched-part validation before save.

## Public API And CLI

`DocxEditor` is stream-first and returns result objects with success flags, diagnostics, structured data, and operation reports. The CLI is a thin local harness over the public API with text and JSON modes.

## Compatibility

DocxEdit is pre-1.0. Public result models prefer init-only properties over long
positional records so future metadata can be added as optional properties instead of
changing constructor and deconstruction shapes. Breaking public API changes may still
occur before 1.0 when they make the supported surface clearer or safer.
