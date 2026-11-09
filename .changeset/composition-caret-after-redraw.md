---
"@stll/folio-core": patch
---

In the canonical session, text being composed with an input method right after a pasted run, or at a paragraph's start after a redraw, now shows at the caret while it is composed. After an edit, the editor sets the browser's caret again, in the text before the editor's caret, so a composition starts there.
