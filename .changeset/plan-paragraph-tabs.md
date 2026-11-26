---
"@stll/folio-core": minor
---

A source item of a template plan can give its paragraph left tab stops of its own (`tabStops`). A line of a label, a tab and a value takes the stops of the template's style, which are made for other text, so the values of a block of such lines started at different positions; a plan can now set one stop behind the widest label. The stops are measured from the left margin and are added to the stops of the paragraph's style. A template inspection reports the left indent each style gives (`indentLeft`), so a caller knows where a label in that style starts.
