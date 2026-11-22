---
"@stll/docx-core": patch
---

Plan text ranges across range boundaries standing between paragraphs as one run of paragraphs: replacing or deleting such a range joins them with the boundaries where they meet, directly or once its tracked deletion is accepted, and a direct deletion's caret lands in the paragraph past them.
