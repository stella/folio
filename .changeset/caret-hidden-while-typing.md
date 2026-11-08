---
"@stll/folio-react": patch
"@stll/folio-core": patch
---

A keystroke in the React editor no longer re-renders the editor to hide the caret while the typed text lays out again: the selection overlay is hidden directly and shown again with the next caret. `getFolioCaretViewportRect` returns `null` while the caret is hidden this way, as it did when the caret was cleared.
