---
"@stll/folio-core": patch
---

The XML safety check no longer rejects a part that holds U+FEFF past its start (a zero-width no-break space, as document variables written by word processors hold): it is a legal XML character. A byte order mark still counts only at the very start, and a declaration behind it is still refused.
