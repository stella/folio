---
"@stll/docx-core": minor
---

Replace `ThemeColorSlot` with `ThemeColor`, generated from `ST_ThemeColor`, and
carry a token outside it as `{kind: "unrecognised", raw}` rather than dropping
it. `ColorValue.themeColor` now holds either; `themeColorSlot` resolves one to a
theme slot, through the `w:clrSchemeMapping` key for the mapped members.
