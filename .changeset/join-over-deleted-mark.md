---
"@stll/docx-core": patch
---

Let a suggested join of a paragraph whose mark a pending deletion already took apply: the two already read as one paragraph, so the suggestion keeps that deletion and records only the properties the join gives, instead of refusing it.
