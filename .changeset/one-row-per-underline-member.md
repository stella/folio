---
"@stll/folio-core": patch
---

Read an underline member's pattern, weight and stroke count off one row.

The display list held three parallel total tables keyed by `ST_Underline`, so `wavyDouble`'s `wavy` pattern and its second stroke were one decision spelled in two places, and a new member needed three edits. `UNDERLINE_STROKES` states a row per member and `underlinePattern`, `underlineWeight` and `underlineStrokeCount` read it, so no consumer changes.

The DOM's `text-decoration-thickness` now comes from that row's weight rather than from a second list of the heavy members, which is what kept the page and the editor honest about which members Word draws heavy.
