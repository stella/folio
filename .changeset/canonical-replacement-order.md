---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

In the experimental canonical session, a tracked replacement now reads as a redline: the struck text comes first and the new text follows it, as the default editor saves it. This holds for typing, pasting and pressing Enter over a selection, and for public replacements. Text replacing a selection that ends inside a content control stays inside that control instead of splitting it, and replacing exactly the words a comment covers keeps the comment on the new text.
