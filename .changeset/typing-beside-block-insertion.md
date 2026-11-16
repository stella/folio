---
"@stll/docx-core": patch
---

Typing, inline content insertions and inline record splits and joins in a paragraph pass a concurrent paragraph insertion before or after that paragraph, direct or tracked, in either sequence order, rather than being refused.
