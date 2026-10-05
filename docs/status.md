# Status

DocxEdit is an early-stage editor. It supports useful structural reads and a focused set of edits, but it is not a complete WordprocessingML implementation.

## Where To Look

Feature-by-feature inventories rot, so each source below is either generated or
verified by hand when touched. The patch tables in patch-format.md are frozen
copies of generator output — refresh them by hand after registry changes:

- `docxedit help <topic>` (`dump`, `context`, `changes`, `validate`, `check`,
  `apply`, `patch`, and every other command) renders the current command
  surface, options, and patch-operation fields from the shared command catalog.
- `docxedit catalog --json` emits the same catalog machine-readably: commands,
  options, output fields, and patch operations with track-change support.
- [cli.md](cli.md) documents the agent workflow: discovery, target IDs, privacy
  defaults, check/apply, and exit codes.
- [patch-format.md](patch-format.md) documents the `.docxpatch` DSL, including
  the operation/field tables and the track-change support matrix generated from
  the operation registry.
- [diagnostics.md](diagnostics.md) catalogs every diagnostic code the library emits.
- [validation.md](validation.md) describes the structural validation profiles.

## Current Guarantees

[Lokad.DocxEdit 0.1.0](https://www.nuget.org/packages/Lokad.DocxEdit/0.1.0)
is published on NuGet. The package targets `net10.0` and contains the library;
the development CLI is available from source.

Repository documentation describes current source, including unreleased changes.
See [CHANGELOG.md](../CHANGELOG.md) for the published and unreleased changes.
These guarantees hold for current source:

- Stream-first library with no NuGet dependencies beyond the .NET platform libraries.
- Repeatable `.docxpatch` edits against discovered target IDs (`read`, `outline`, `find`, `dump`, or `context`): same bytes scan to the same IDs,
  which are positional rather than permanently stable (SPEC 8.2, SPEC 23.2).
- `check` before `apply`: patches validate selectors, guards, assets, and
  track-change constraints without writing output.
- Privacy-safe inspection: shape and metadata before text (`read --summary`,
  metadata-only `context`, text-free `changes` by default).
- Safe package handling: no filesystem extraction, size limits, and post-edit
  validation of touched parts before writing output.
- No Word automation and no pixel-perfect layout editing; Word remains
  responsible for general field recalculation and pagination.

## Non-Goals

Explicit non-goals live with the implementation specification in `SPEC.md`
(Key design principle, Explicit non-goals); behavior outside them fails with
stable diagnostics instead of guessing.
