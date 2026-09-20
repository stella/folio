---
"@stll/folio-core": patch
---

Keep a section break on the paragraph that ends the section.

A `w:sectPr` inside a `w:pPr` says the section ends at that paragraph's mark. ProseMirror copies a node's attrs when a command splits it, so pressing Enter inside a section-ending paragraph produced two nodes over the same `_sectionProperties` object and the save wrote the break twice: a section nobody added, whose `w:sectPr` repeated the real one's `w:rsidSect` and so claimed its revision history as well.

The from-leg now assigns the break by reference identity and position — among the paragraphs holding one `_sectionProperties` object, only the last in document order writes it, and the object is never cloned. That is where the rule can be total: Enter, a paste, an AI edit and a split merged in from another client each reach the model by their own route, and only the projection sees them all.

Joining is the same rule read backwards. A join consumes one paragraph mark and keeps the other, so Backspace at the start of a section-ending paragraph leaves a merged paragraph that still ends the section; ProseMirror keeps the first node's attrs, so the break is now carried to the mark that survived. A mark deleted outright still takes its section with it.
