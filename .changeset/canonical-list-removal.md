---
"@stll/folio-core": patch
---

Taking a paragraph out of a list in a canonical session removes its own numbering when its style chain does not number it, so it keeps no numbering override it never had; the explicit cancellation is written only for a paragraph its style numbers.
