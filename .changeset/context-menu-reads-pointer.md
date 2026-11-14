---
"@stll/folio-core": patch
"@stll/folio-react": patch
---

The context menu's Accept Change and Reject Change act on the tracked change under the pointer, also when the right-click lands inside a range selection that starts outside it. `ContextMenuSnapshot` carries that change as `trackedChange` in place of `cursorInTrackedChange`.
