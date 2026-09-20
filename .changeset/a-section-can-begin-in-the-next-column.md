---
"@stll/folio-core": minor
---

Lay out a `nextColumn` section break. `ST_SectionMark` has five members and the layout modelled four, so `w:type="nextColumn"` reached the paginator through a cast as a value its switch did not handle: the section fell through to the switch's silent default, started no section at all (its page numbering and header references never took effect) and resumed below the outgoing content instead of in the next column.

`SectionBreakBlock["type"]` is now `SectionStart`, the model's `ST_SectionMark`, so the bridge hands the parsed value through and a member the layout does not handle is a compile error rather than a cast. Per §17.18.77 a `nextColumn` section begins in the next column of the region it shares with the outgoing section; when there is no such column, because the section is single-column or because the incoming one redefines the column geometry, it begins in place like `continuous`, and when the region's last column is already in use it opens the next page. The measure pass mirrors the same decision, so the width and column a block is measured against match the one it is painted in.

`normalizeSectionBreakType` still reads an absent `w:type` as `nextPage` (§17.6.22) but no longer passes a value outside the enumeration through as though it were one: that is a bug in whatever produced it, and it now panics instead of laying out as something else. `SECTION_BREAK_TYPES` is total over the enumeration, checked against the committed schema graph, and the Yjs v3 migration therefore keeps a `nextColumn` section rather than dropping it. Which members the insert commands offer is unchanged.
