Test layers stay separated by what they prove, not by topic.

JSON edit cases prove harness shape and diagnostic codes.
Unit tests prove preserved bytes and neighboring content.
Word tests prove open and save round trips.

Creation tests cover package validity, all paper/orientation combinations, subsequent
patch editing, section layout readback, stream ownership, and command publication.
The opt-in creation Word test checks that empty documents retain their section layout
after Word opens and saves each combination.

Protected spans keep both layers for endpoints and interiors.
Alias cases keep both layers for composition and refusals.
Image cases keep both layers for validation and round trips.
Bookmark cases keep both layers for ranges and bindings.
Table cases keep both layers for grids and merges.

Cases stay unless a new assertion shows one layer adds no distinct evidence; none were removed for looking similar.
New layered cases need new assertions to justify them.

Overlap note 2026-10-01, rechecked after the report-lifetime and image
follow-ups: four JSON cases (basic-replace, alias-compose,
ambiguous-selector-refusal, tracked-replace) were sampled against unit
coverage plus the tracked-change matrix, not the whole corpus. The sampled
overlap is topical rather than evidential: the sampled JSON cases assert
harness boundaries (diagnostic codes and counts, message fragments,
applySuccess, public readback lists such as paragraphs and changeSummary
counts), while unit tests assert library precision (preserved bytes, XML
fragments, neighbor objects, typed IDs, per-operation report shapes). The
sample does not prove the absence of duplication elsewhere, so it is kept
as a sampling method with worked examples rather than a blanket claim.
Overlapping cases stay unless a new assertion shows one layer adds no
distinct evidence.
