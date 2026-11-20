---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
"@stll/folio-agents": patch
---

Comment ids are allocated per document instead of from one counter for the whole process: the same edits to the same document give the same comment ids and reply paragraph ids, whatever other documents allocated, and an id the document holds is never reissued. `createCommentIdAllocator()` returns an independent allocator, and `createReply` takes the document's allocator.
