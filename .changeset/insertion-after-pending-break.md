---
"@stll/folio-core": patch
---

On the canonical engine, a paragraph inserted after one whose break is pending removal goes right after it only where both it and the paragraph its join leaves keep words once accepted. Where the paragraph its join leaves has no words left (a last-paragraph deletion took it over), or the paragraph itself keeps none and reads as one paragraph with the one its join leaves, it goes after that paragraph with its properties, so accepting gives the result of the same edits applied directly.
