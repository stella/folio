---
"@stll/folio-core": patch
---

The experimental canonical session's save keeps the source XML of every paragraph an edit did not touch in more documents:

- When `word/document.xml` starts with a byte order mark, the body is no longer written whole from the model, which could drop formatting the model does not carry.
- A structural edit in a body that holds comment anchors is now spliced into the source part instead of rewriting the whole body, as long as the spliced body carries exactly the serialized body's comment anchors.
