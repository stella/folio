---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Write an explicit off for every `CT_OnOff` element. `serializeOnOffElement` is the
one writer: absent writes nothing, an on writes the bare element, and an off
writes `w:val="0"`, which is what cancels an inherited on. The row, cell, table,
control and paragraph-mark readers keep the three states apart as well, and a
control that states nothing keeps stating nothing through the editor.
