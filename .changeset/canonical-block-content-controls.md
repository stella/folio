---
"@stll/folio-core": patch
---

The experimental canonical session now opens a document with block content controls. Text inside a control edits like any other paragraph, and the control keeps its edges through typing, Enter and joins. The session honours a control's `w:lock` on every editing path: a change that edits inside a content-locked control, or removes a control-locked one, refuses with the `contentControlLock` rule. Accepting or rejecting holds back a revision that would break a lock, which stays pending, and resolves the rest. A batch of document operations skips the operation a lock refuses, with the reason `contentControlLock`, and applies the others.
