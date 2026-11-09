---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Dragging a table's column border, right edge or row border now resizes the table in the experimental canonical session, in React and Vue, instead of being refused. A resize is a command (`columnResizeCommand`, `rightEdgeResizeCommand`, `rowResizeCommand`) that the hosts run through the editor; the session applies it as the grid and row and cell formatting edits it makes, records it as a revision while suggesting, and undoes it.
