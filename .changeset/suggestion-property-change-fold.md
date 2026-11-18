---
"@stll/docx-core": patch
---

Refuse a suggestion's tracked paragraph property change, join or split that would fold into a pending property change belonging to another author or suggestion. A paragraph holds one property change and a suggestion is rejected by its id, so taking the record over discarded the other change when the suggestion was rejected or saved as rejected.
