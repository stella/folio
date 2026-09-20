---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

One `ListState`, resolved from the numbering definitions

The toolbar's list state is a union in `@stll/folio-core/prosemirror`, read by
both adapters, in place of three declarations that never imported one another
and had already drifted. It replaces the `isInList` flag, which restated
`type !== "none"`, and the `numId` that meant nothing on the empty arm.

Which kind of list a paragraph is in now comes from its level's `w:numFmt`
rather than from its numbering id. `numId === 1` meant bullets only in a
document Folio had created itself, so an imported bulleted list read as
numbered in every toolbar. The list commands keep minting their own instances
and say which kind they are creating, for a document that defines no numbering
yet.

`SelectionState` carries the resolved `listState`, and `SelectionContext`
replaces `inList` / `listType` / `listLevel` with it.
