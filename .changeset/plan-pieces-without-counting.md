---
"@stll/docx-core": patch
---

Planning a tracked edit applies each piece once when it takes no new ids, instead of first counting the ids it needs, so tracked typing in long documents spends less time planning.
