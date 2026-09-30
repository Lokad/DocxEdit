Test layers stay separated by what they prove, not by topic.

JSON edit cases prove harness shape and diagnostic codes.
Unit tests prove preserved bytes and neighboring content.
Word tests prove open and save round trips.

Protected spans keep both layers for endpoints and interiors.
Alias cases keep both layers for composition and refusals.
Image cases keep both layers for validation and round trips.
Bookmark cases keep both layers for ranges and bindings.
Table cases keep both layers for grids and merges.

No cases were removed for looking similar.
New layered cases need new assertions to justify them.

Overlap audit 2026-10-01: sampled JSON cases (basic-replace, alias-compose,
ambiguous-selector-refusal, tracked-replace) against unit coverage across the
205-case corpus plus the tracked-change matrix. Overlap is topical, not
evidential: JSON cases assert harness boundaries (diagnostic codes/counts and
message fragments, applySuccess, public readback lists such as paragraphs and
changeSummary counts); unit tests assert library precision (preserved bytes,
XML fragments, neighbor objects, typed IDs, per-operation report shapes). No
JSON case asserts wire bytes or XML internals, so none duplicates a unit
test's oracle. Retained without deletions per the rule above.
