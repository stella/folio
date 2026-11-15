---
"@stll/docx-core": patch
---

A direct deletion whose range holds a note's own reference mark keeps the mark, and now records that in its sequenced effect (`keptMark` in a new `deleteRange` effect). A concurrent edit at a position the deletion closes up to its start, such as typing just after the deleted text, lands after the kept mark, as it does when the edit applies first. A direct format inside such a deletion records the mark on the records it cuts, so the deletion rebased over the format joins the pieces the mark keeps, as one record, as in the other order.
