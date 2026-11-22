---
"@stll/docx-core": patch
---

A direct deletion of a paragraph with range boundaries before a block that offers no paragraph to run into leaves them in the next paragraph after that block, as accepting the tracked deletion does, instead of refusing.
