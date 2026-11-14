---
"@stll/docx-core": minor
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Choosing a value in a dropdown, combo box, date or checkbox content control changes the control in the canonical session. The control's value (`w:lastValue`, `w:fullDate`, `w14:checked`) and its shown content change together as one undoable edit, suggesting mode tracks the shown content, and the value survives save and reopen. A locked control refuses the change and the editor says so; a refused pick no longer reports success. The new `setContentControl` operation sets a block-level or inline control's properties by its id, with an exact inverse, and is sequenced like other operations that address the paragraphs a record holds.
