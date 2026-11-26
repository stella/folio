---
"@stll/folio-core": minor
---

Typing, pasting and commenting in a text box's paragraphs in a canonical session apply in the box, save there and undo exactly; no position in a text box's text refuses under the `publicOps.unsupportedInline` gap any more. A position that names no text, such as the edge of a text box, a point between blocks, or one inside an indivisible inline element or a character, now refuses as an invalid request (`{ condition: "invalidRequest" }`) instead of as a capability gap.
