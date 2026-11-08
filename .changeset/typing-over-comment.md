---
"@stll/folio-core": patch
---

Typing over exactly the words a comment covers keeps the typed words inside the comment, and the characters typed next join them until the caret moves. Before, the typed words landed after the comment's end: typing directly left the comment with no text and the save dropped it, and while tracking changes, accepting the replacement did the same.
