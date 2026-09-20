---
"@stll/folio-core": patch
---

Read through a bidirectional wrapper when exporting markdown. `w:bdo` and
`w:dir` say how their text is laid out, not what it is, and both inline
renderers narrowed by a switch whose default contributed nothing: a paragraph
whose runs sat inside one exported as an empty line, in the pipe-table path and
the HTML-cell path alike. The wrapper is the ordinary way to write a
right-to-left run, so the loss fell entirely on right-to-left documents.
