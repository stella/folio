---
"@stll/folio-core": minor
"@stll/folio-agents": minor
---

Classification change: the AI snapshot, `getContent()` and the agent tools classify a numbered heading as `kind: "heading"` instead of `"listItem"`. A paragraph with an outline level or a built-in heading style is a heading whether or not it is numbered; its number (`1.`, `2.1.`) is its `displayLabel` and its numbering level its `listLevel`. A paragraph whose numbering shows no marker (a `w:vanish` level, or `w:numId="0"` cancelling its style's numbering) is now `kind: "paragraph"`, keeps its `listLevel`, and no longer carries the hidden marker as its `displayLabel`. `getContentAsText()` prints a numbered heading as `[id] (h2) 1. Scope`. Code that read `kind === "listItem"` to find numbered paragraphs should read `listLevel` instead.

New row fields: `read_document` and `read_section` rows carry `displayLabel`, `headingLevel` and `listLevel` when the block has them (absent fields are omitted), so a model is given the numbers the document shows.
