---
"@stll/folio-core": patch
---

A `splitBlock` inside an inline content control is now refused in both modes. It used to break the control in two within one paragraph, and rejecting the tracked form joined the paragraph to the next one.
