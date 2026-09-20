---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Keep a bookmark boundary inside the inline content control that held it.

`CT_SdtContentRun` reaches `w:bookmarkStart` and `w:bookmarkEnd` through `EG_RunLevelElts > EG_RangeMarkupElements`, so a marker inside `w:sdtContent` is markup Word writes. folio lifted it out to a sibling of the control. That is not a re-spelling: a bookmark whose extent was the control's content came back starting before the control, so a `REF` field or a link to it resolved to a different range, and a marker in the middle of the content split one control into two carrying the same `w:id`, `w:tag` and data binding.

`InlineSdt["content"]` gains `BookmarkStart` and `BookmarkEnd`, and `INLINE_SDT_CONTENT` admits them: the admission map is bound to the content type, so the parser, the serializer and the editor's save filter all follow from the one decision. The boundary rides the editor as the inline atom it already was, inside the control's `inline*` node, and the pairing pass looks inside the control, so a range that opens inside and closes outside keeps both halves instead of being deleted as an orphan. A revision covering a whole control that holds a marker still hoists to `w:ins > w:sdt`.

The other range markers are still lifted: a `w:commentRangeStart` or a `w:moveFromRangeStart` inside the control is a marker the control does not own, and the pairing passes read it as a paragraph-level sibling.
