---
"@stll/folio-react": patch
---

Typing in the React editor commits React updates less often: about 2 commits per keystroke instead of 5 on a long contract. Caret and selection geometry, overlay rectangles, toolbar context and the outline keep the value they hold when a keystroke produces an equal one, so an unchanged value no longer re-renders and commits. Each commit made while the editor has focus reads the editor's whole document structure to save the selection, so fewer commits make typing in long documents faster.
