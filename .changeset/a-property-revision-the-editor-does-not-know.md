---
"@stll/docx-core": patch
"@stll/folio-core": patch
"@stll/folio-vue": patch
---

Resolve `w:tblPrExChange` like every other tracked property revision, from one list of them.

The element round-tripped but the editor did not know it: accept, reject and the tracked-change list each carried their own list of four or five change elements, and none of them named the fifth. Accepting every change left the revision on the row, so the document said a formatting change was still pending after the reviewer had resolved it.

The set is now written down once. `PROPERTY_REVISION_KINDS` is the model's census of the change elements that store a complete previous property set, and one site table says where each one lives, how it resolves and what a reader calls it. The carrier reader, the accept/reject command, the tracked-change list, the comparison's scopes and the Vue sidebar's labels are each total over it, so a revision the model gains is a compile error at every one of those rather than a branch nobody wrote.

Two revisions the list had already lost come back with it: a paragraph's `w:pPrChange` was read from an attr the schema does not declare, and `w:sectPrChange` was never listed at all.

Accepting a `w:tblPrExChange` drops the record and keeps the row's current exceptions; rejecting it restores the stored ones wholesale, including restoring their absence.
