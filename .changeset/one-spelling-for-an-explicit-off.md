---
"@stll/folio-core": patch
---

Write one spelling of an `ST_OnOff` value: `0` for an off, the bare element for an on.

`w:hideMark` wrote its explicit off as `w:val="off"` while every other on/off element in the package wrote `0`, and `w:updateFields` wrote its on as `w:val="true"` while its neighbour in the same part wrote the bare element. All three spellings mean the same thing, so a package that mixes them only makes every byte-level comparison argue about which one it is looking at.

`scripts/on-off-spelling.test.ts` holds the tree to the one spelling, with no allowlist. Reading is unchanged: `parseOnOffValue` and `parseBooleanElement` take all six spellings, and a captured element still replays the bytes it arrived as.
