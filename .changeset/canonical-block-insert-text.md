---
"@stll/folio-core": patch
---

Block insertion through public operations in the experimental canonical session now handles multi-line text, `formattingScope`, `lineBreakMode: "inline"` and `hardPageBreak` as the default engine does. Each line becomes its own paragraph, blank lines are dropped and reported as `splitMultilineText`, and later lines are body paragraphs unless every paragraph is formatted alike. Inline mode keeps tabs and line breaks in one paragraph, and an authored page break stands in for the first paragraph's text.
