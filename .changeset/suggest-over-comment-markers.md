---
"@stll/folio-core": patch
---

A suggestion replacing a paragraph's text together with the markers of a comment covering only words inside it records the replacement in the canonical session instead of refusing: the comment stays on the deleted words until the suggestion is accepted, which removes it as the direct edit does, and rejecting restores it.
