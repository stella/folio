---
"@stll/folio-core": minor
---

Header, footer and note editors now take Enter, Shift+Enter, Tab, Backspace and Delete at a paragraph's edge, and cut, as the body editor does: each is one journalled change that undo reverts in one step and that suggesting mode records. A page break in a header, footer or note is refused, since only the main story takes one.
