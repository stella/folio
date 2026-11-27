---
"@stll/docx-core": minor
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

Pasting a chart, an embedded object or kept block markup no longer refuses. A paste made before the editor finished reading the copied package's parts waits for them; a part outside the main document's folder is copied into it; and a paste into a header, footer or note editor carries the parts into that story's own part. Content from outside the editor pastes as its preview picture or the text it holds, or is left out, and hosts receive each substitution as a structured warning through the new `onCanonicalPasteWarnings` prop. The package model now also keeps each header's and footer's part relationships in `annotationRelationships`, keyed `header:<rId>` or `footer:<rId>`; removing a header or footer removes its entry, and undo restores it.
