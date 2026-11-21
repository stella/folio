---
"@stll/docx-core": patch
---

A split that keeps the paragraph's identity on its first half gives that half the surviving mark's properties and leaves the paragraph's own on the second, so a split of a paragraph whose mark is pending deletion reads, once accepted, as the merged paragraph split directly.
