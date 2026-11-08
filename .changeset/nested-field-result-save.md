---
"@stll/folio-core": patch
---

A complex field nested in another field's result, such as a citation inside a citation, no longer loses its result text on save. A save that writes a paragraph from the model, after a paragraph was added or removed anywhere in the document, wrote the nested field without its own separator and end, so the reopened document read its result as an instruction and showed nothing. The nested field's structure is now kept with the result it sits in, at any depth.
