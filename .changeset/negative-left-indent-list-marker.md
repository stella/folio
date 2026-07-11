---
"@stll/folio-core": patch
---

Fix list marker position for paragraphs with a negative left indent and a hanging indent (`w:ind w:left="-180" w:hanging="360"`). The marker was painted at the left indent instead of `left - hanging`, shifting it one hanging-indent (e.g. 18pt) too far right and mis-indenting continuation lines. The negative left indent is now realized by the line's own `margin-left`, and the marker's remaining negative offset rides on its `margin-left`; positive and zero left-indent lists are unchanged.
