---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Keep the transparent wrapper a tracked change was authored around inside the
revision. `TrackedRunContent` held no `w:bdo`/`w:dir` and no inline `w:sdt`, so
the parser lifted one out to a sibling and
`<w:ins><w:bdo>x</w:bdo></w:ins>` saved as `<w:ins/><w:bdo>x</w:bdo>`: `x` was
no longer inserted, and accepting the revision kept it exactly as rejecting it
did. `TrackedRunContent` and `InlineSdt["content"]` now admit both wrappers, a
single admission map bound to those content types decides what each wrapper
keeps, and the serializer carries the revision's disposition through the
wrapper so a `w:del` still writes `w:delText` around it. The opposite authored
order, a revision inside the wrapper, is unchanged.
