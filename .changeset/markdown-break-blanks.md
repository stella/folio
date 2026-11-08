---
"@stll/folio-core": patch
---

Markdown output keeps the blanks next to a line break and keeps consecutive line breaks in one paragraph: a space or tab that starts a line after a break, or ends one before it, is written as a character reference, and an empty line between breaks uses the backslash break, so reading the Markdown back gives the paragraph's text.
