---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

A comment that a pending suggestion proposes is no longer saved on its own. The saved package holds back the suggested text, and now the comment's thread too, so reopening the file no longer shows a card with nothing to point at. Accepting the suggestion saves both. `withHostCommentThreads` takes the comment ids the editor's document anchors, and leaves out a thread that the editor anchors but the saved document does not.
