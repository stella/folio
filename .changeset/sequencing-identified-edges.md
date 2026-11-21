---
"@stll/docx-core": patch
---

Sequencing refuses two more concurrent pairs whose orders give a record different ids. A paragraph restoration conflicts with an edit at an edge of its span when it moves where a revision or content control starts or ends there. A direct deletion conflicts with a concurrent direct insertion that supplies new ids when the deletion removes content up to or across the insertion's point.
