---
"@stll/folio-core": patch
"@stll/docx-core": patch
"@stll/folio-agents": patch
---

Bound version-comparison, agent, and validation paths against crafted-input
resource exhaustion: a single aggregate LCS cell budget is now shared across all
document stories; move detection dequeues in O(1) instead of `Array.shift`;
`diffWordSegments` caps its DP matrix and falls back to a whole-string diff;
`ensureParaIds` and `docx-core`'s `validateDocxPackage` enforce entry-count and
uncompressed-size limits before reading; note-paragraph patching builds a linear
offset index instead of rescanning per id; agent whole-word search uses a bounded
boundary window; `suggest_changes` enforces an aggregate operation-text budget;
and tracked vertical cell split refuses a stored continuation whose `gridSpan`
exceeds one column.
