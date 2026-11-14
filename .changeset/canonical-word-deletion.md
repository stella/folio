---
"@stll/folio-core": patch
---

A canonical session now deletes by word and by line: Ctrl or Alt with Backspace or Delete (and Cmd on macOS) inside text deletes the range the browser measures, in editing and in suggesting mode, across runs and in table cells; over a selection or at a paragraph's edge the modified key deletes or joins as the plain key does.
