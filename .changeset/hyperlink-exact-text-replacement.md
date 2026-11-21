---
"@stll/docx-core": patch
---

Text replacing exactly a hyperlink's text keeps the link, edited directly or tracked: the new text goes in a copy of the link, and accepting the tracked replacement gives the direct edit. A range reaching past the link, or text typed at its edge, still goes beside it.
