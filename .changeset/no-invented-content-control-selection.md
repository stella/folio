---
"@stll/folio-core": patch
---

Never write a content-control selection nobody made. A block-level dropdown or
combo box whose `w:sdtPr` was not replayed verbatim — every control an editor
command built, and every control on the rebuild path — had its `@w:lastValue`
recovered from the body's display text, so a control still showing its
placeholder was saved as selected, and a displayText shared by two list items
selected the first of them. `properties.dropdownLastValue` is now the only
record of a selection; the schema's empty-string default keeps "never
selected", "cleared" and "selected" distinguishable.
