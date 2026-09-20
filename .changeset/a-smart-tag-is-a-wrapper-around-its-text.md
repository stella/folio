---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Keep a `w:smartTag` and a run-level `w:customXml` as the wrappers they are.
folio spliced a smart tag's children into the paragraph and kept no wrapper, so
the tag, its namespace and its properties were gone on the first save; and it
captured a run-level `w:customXml` whole, so the wrapper came back but every run
inside it was opaque bytes the editor could not touch. Both are `InlineWrapper`
kinds now — `smartTag` and `customXml`, each carrying `element`, an optional
`uri` and the `w:smartTagPr` / `w:customXmlPr` verbatim — so their content is
parsed by the same run-level walk `w:bdo` and `w:dir` take, nests with them in
either order, and comes back through the editor on the same `inlineWrapper`
mark. The properties are part of the mark's stack key, so two adjacent tags that
differ only in their properties stay two tags; they are replayed only when they
are structurally the element they claim to be, because a paste from outside the
editor can put any string on a mark.
