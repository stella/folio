---
"@stll/docx-core": patch
---

A tracked deletion of a paragraph that holds only empty containers, and whose mark a pending deletion already took, records nothing and is refused, as for an empty paragraph, instead of recording a deletion no reviewer sees.
