---
"@stll/folio-core": patch
---

fix(core): derive single-line height from real font hhea metrics

Single-line height per font is now derived from the font's real `hhea`
metrics `(ascent + |descent| + lineGap) / unitsPerEm` — the value Word uses —
instead of hand-transcribed constants, several of which dropped the line gap
or were otherwise wrong. Corrects 9 fonts (measured against Word): Palatino
Linotype (was 31% short), Book Antiqua (17%), Cambria (8%), Century Gothic
(6%), Times New Roman (4%), Arial (3%), Trebuchet MS (2%), Consolas (1%), and
Lucida Console (14% tall). A shared derivation and a discriminated `hhea` /
`legacy` representation make it structurally impossible to silently drop a
term again. CJK fonts are unchanged (their line height is an East-Asian
layout concern, not a run-font hhea ratio).
