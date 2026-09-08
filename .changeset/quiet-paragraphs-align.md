---
"@stll/folio-core": minor
"@stll/folio-agents": patch
---

Compare and track direct paragraph alignment changes while preserving the distinction between direct formatting and style inheritance.
Paragraph replacement operations can now clear a direct paragraph style with `null`; paragraph insertion and property schemas expose their existing style and list clear values consistently.
Unstamped multi-paragraph insert batches now reserve revision IDs for synthesized paragraph-property changes, so later batches cannot reuse an existing ID.
Tracked paragraph insertion receipts include synthesized paragraph-property revisions, so targeted acceptance and rejection resolve the whole operation.
Accepting or independently resolving suggested paragraphs at the end of a container keeps every final paragraph mark resolvable.
