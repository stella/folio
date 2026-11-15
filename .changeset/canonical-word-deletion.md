---
"@stll/folio-core": patch
---

A canonical session now deletes by word and by line, following the platform convention: Ctrl or Alt with Backspace or Delete (and Cmd on macOS) inside text deletes the range the browser measures, in editing and in suggesting mode, across runs, beside bookmarks, fields and links, and in table cells; bookmark and range markers at the range's edges stay. Over a selection or at a paragraph's edge the modified key deletes or joins as the plain key does. Every deleting input is either applied by the session or refused before the browser changes the page.
