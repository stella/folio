---
"@stll/docx-core": patch
---

A paragraph mark before a table resolves in a container that ends with a block content control holding its final paragraph, as it does where the container ends with that paragraph itself: a tracked or suggested paragraph inserted before the table applies, and rejecting it restores the original rather than being refused.
