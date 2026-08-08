---
"@stll/folio-core": patch
---

Keep cursive words joined when a run boundary splits them mid-word. Browsers shape across an inline box boundary only while no shaping-relevant property changes, so a bold, italic, resized or refaced run inside an Arabic word stopped shaping there and the word rendered as isolated letter forms; Word joins straight through the same boundary. The painter now emits a zero-width joiner on each side of such a boundary, each in its own span so run text nodes keep their exact ProseMirror offsets. Colour and underline boundaries (tracked changes, comment anchors) already shaped correctly and are left untouched. Joining classes are generated from the pinned Unicode Character Database rather than hand-listed, so combining marks are classified correctly.
