---
"@stll/folio-core": patch
---

A comment thread now goes with the content it is anchored to: rejecting the tracked insertion it covers, accepting the deletion of that content or deleting it outright removes the thread from `getComments()` and from the saved `comments.xml`, replies included, instead of leaving a thread about nothing. A definition the source package already left unanchored is kept. A block inserted inside a comment range that crosses into or out of a table (or past empty paragraphs) is now covered live, as the saved package already covered it.
