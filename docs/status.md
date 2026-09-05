# Status

This is a pre-release editor. It supports useful structural reads and a focused set of edits, but it is not a complete WordprocessingML implementation.

## Where To Look

Feature-by-feature inventories rot; the sources below do not, because they are
generated from or tested against the code:

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

## Stable Guarantees

These hold across releases and are safe to rely on in prose:

- Stream-first library with no NuGet dependencies beyond the .NET platform libraries.
- Deterministic `.docxpatch` edits against stable target IDs discovered through
  `read`, `outline`, `find`, `dump`, or `context`.
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
