---
"@stll/folio-core": minor
---

Preserve authored run properties through document comparison and tracked revision resolution. Retain text in the selected AlternateContent branch during parsing and serialization.

Preserve direct paragraph indentation and inline tab/break controls through comparison and reviewed views. Insert operations can retain hard breaks with `lineBreakMode: "inline"`.

Preserve inherited table run formatting when deleting content. Standalone hard page breaks can be inserted as tracked content with `hardPageBreak`.

Carry concrete list references and import target numbering definitions. Rebind conflicting definitions through tracked paragraph changes so accepting and rejecting retain the corresponding list formatting.

Preserve complete table formatting, empty field results, and untouched drawing geometry through document saves. Explicit image edits invalidate stale editable drawing captures.

Reconcile field, picture, and page-break atoms through mapped review positions. Report atom-only changes and verify both accepted and rejected content, importing picture media without overwriting existing resources.

Track section property changes through accepted and rejected views, and report unsupported section topology explicitly.
