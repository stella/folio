---
"@stll/docx-core": minor
"@stll/folio-core": minor
"@stll/folio-agents": patch
---

Report a tracked operation over a paragraph mark pending insertion as the new `pendingParagraphMarkInsertion` skip reason; one over a mark pending deletion stays `pendingParagraphMarkDeletion`, so the two are never reported alike.
