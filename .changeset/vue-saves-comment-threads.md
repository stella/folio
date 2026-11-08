---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

The Vue editor saves the comment threads its sidebar shows: added comments, replies, resolutions and deleted threads now reach the saved document, and `getDocument()` returns them, as in React. Both editors write their threads through one helper, `withHostCommentThreads`.
