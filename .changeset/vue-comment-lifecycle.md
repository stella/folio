---
"@stll/folio-vue": minor
---

Wire the comment + tracked-change lifecycle into `DocxEditor.vue`. The sidebar
now populates from the document: comments are seeded from
`document.package.document.comments` on load, and tracked-change cards derive
from core's `extractTrackedChanges(editorState)` (the Vue `TrackedChangeEntry`
now re-exports core's type, removing the optional-field mismatch that left
cards empty). Two new composables built on core's comment ops drive mutations:
`useCommentManagement` (add reply, resolve/unresolve, delete, tracked-change
reply, and accept/reject by revision id via `findChangeRange` + range-based
`acceptChange`/`rejectChange`) and `useCommentLifecycle` (floating add-comment
button, pending-highlight range, submit/cancel). Every mutation fires
`onCommentsChange` (and a `comments-change` emit). `comments` and
`onCommentsChange` move from `deferredInVue` to `paired`.
