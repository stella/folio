---
"@stll/folio-core": patch
---

Keep an optional paragraph property that a style supplies out of the paragraph's own `w:pPr`. The editor holds the effective value, and the save path read it as authored, so `w:bidi`, `w:widowControl`, `w:suppressAutoHyphens`, `w:kinsoku`, `w:overflowPunct`, `w:snapToGrid`, `w:pageBreakBefore`, `w:contextualSpacing`, `w:ind`, `w:pBdr`, `w:shd`, `w:tabs` and `w:outlineLvl` came back as direct formatting on any document whose styles define them: the paragraph stopped following its style, and an unset tri-state became an explicit value. A table with no `w:tblBorders` no longer gets one synthesised from its first bordered cell for the same reason. A run carrying an embedded object, picture or shape keeps its `w:rPr`.
