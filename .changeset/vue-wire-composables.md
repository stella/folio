---
"@stll/folio-vue": minor
---

Wire ported composables and real pickers into `DocxEditor.vue`, replacing the
shell-time PORT-BLOCKED stubs. Pages-area pointer gestures (`usePagesPointer`:
multi-click, drag-select, table quick-insert button, hyperlink popup, page
indicator), image actions + selection overlay (`useImageActions` +
`ImageSelectionOverlay`), right-click menus (`useContextMenus`), hyperlink
management (`useHyperlinkManagement`), the table context toolbar (`TableToolbar`
with its border/fill/more pickers), and the comment sidebar (`UnifiedSidebar`
now renders controlled comments via `useCommentSidebarItems`) are functional.
The toolbar image group (`ImageWrapDropdown`/`ImageTransformDropdown`) and the
Insert Table dialog's style gallery (`TableStyleGallery`) are un-stubbed. 2
props (`enableWheelZoom`, `onSave`) and 1 ref member (`scrollToParaId`) move
from `deferredInVue` to `paired`.
