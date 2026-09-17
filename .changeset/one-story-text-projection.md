---
"@stll/folio-core": patch
---

Read a story's text through one projection whether or not it is open for editing. A loaded story used to run a separate walk that glued words together across a tab or a hard break, showed tracked-deleted text, and ran table cells and paragraphs into each other, so the same footnote read differently before and after it was loaded. Header and footer text came from a third walk that saw only plain runs, silently dropping fields, hyperlinks, tabs and breaks. All of them now read the document model's own walk, and a result-less PAGE field reports what the document holds rather than the placeholder a save writes.
