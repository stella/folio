---
"@stll/docx-core": patch
---

Rejecting a paragraph's inserted mark together with the pieces of a change it split no longer moves a bookmark or range boundary kept between them into a content control that follows: it stays where the rejected pieces were, as when each is rejected in turn.
