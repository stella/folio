---
"@stll/folio-core": patch
---

Read `w:themeColor`, `w:themeFill` and `a:schemeClr/@val` through the generated
enumerations, so `hyperlink`, `followedHyperlink`, `dark1`, `light1`, `dark2`,
`light2` and `none` survive a save instead of being dropped at parse time.
`w:shd` keeps its pattern colour's theme reference too.
