---
"@stll/docx-core": patch
"@stll/folio-core": patch
"@stll/folio-agents": patch
---

Document operations now honour a content control's `w:lock` in every mode. Under `contentLocked` or `sdtContentLocked`, an operation that changes text, formatting or blocks inside the control is skipped. Under `sdtLocked` or `sdtContentLocked`, an operation that removes the control is skipped, including deleting the paragraph that holds it. Either way the operation reports the new reason `contentControlLock` and changes nothing. A document comparison still shows changes inside locked controls.

`@stll/docx-core/model` exports the rule: `contentControlLockAllows(lock, change)` decides whether a lock allows editing a control's content or removing the control, and `lockedControlViolation` applies that decision to the stretch a change touches.
