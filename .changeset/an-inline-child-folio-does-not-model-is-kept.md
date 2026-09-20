---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Keep every inline child folio does not model, at its source position. One walk serves a paragraph, the four run-level tracked-change wrappers, `w:bdo`/`w:dir` and an inline content control, and it now goes through the shared child dispatcher over a handler map the compiler makes total. `ParagraphContent` gains a `preservedInline` member holding the captured markup, so `w:permStart`, `w:proofErr`, `w:customXml`, the eight custom-XML revision ranges and `w:subDoc` survive a save and the editor round trip.

Inside a tracked change the position is the point: markup lifted out of a `w:ins` is markup the reviewer no longer accepts or rejects with the change, so the capture sits inside the wrapper in the model, in the editor and in the saved part. `w:customXml` also keeps the text it puts on the line.

A bare OMML element is now read by namespace rather than by falling off the end of a switch, and `m:oMathPara` keeps its display form.
