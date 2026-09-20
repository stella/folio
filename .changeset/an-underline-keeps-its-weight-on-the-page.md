---
"@stll/folio-core": patch
---

Stroke every underline member at the weight it is drawn with. The display list thickened nothing, so `thick` and the seven `*Heavy` members painted on the page and in the PDF as the plain rules they are heavy versions of, while the editor drew them twice as thick: the same run came out two ways. Weight is now a total table over `ST_Underline`, and the heavy multiple is one constant both the stroke and the CSS `text-decoration-thickness` derive from.

Two members are no longer approximated by what CSS can spell. `wavyDouble` draws two waves rather than two straight rules, and `words` underlines the words and skips the spaces between them: the display list carries one advance per code point, so it can place a span per word where a `text-decoration` cannot.
