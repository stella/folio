---
"@stll/folio-core": patch
---

Bound DOCX parsing and layout against crafted-input resource exhaustion: clamp
table `gridSpan`/column counts, section column count, and page dimensions;
floor the default tab stop; cap kinsoku/line-break rule lists (stored as sets)
and run language tags; add an iteration cap to cross-run hyphenation; replace
the conformance root-tag regex with a linear, non-backtracking scanner; cap
per-element xmlns declarations; enforce an incremental size budget while
building grouped-drawing SVG previews; make style-numbered list resume O(1) per
paragraph; and guard encrypted-DOCX parsing with DIFAT cycle detection plus a
`spinCount` ceiling.
