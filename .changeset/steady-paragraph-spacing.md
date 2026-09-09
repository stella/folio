---
"@stll/folio-core": minor
"@stll/folio-agents": patch
---

Compare, inspect, insert, and track the complete direct paragraph-spacing cluster while preserving absent attributes, explicit zero and false values, and style inheritance.
Line values and line rules retain independent direct-formatting provenance, so saving one no longer materializes an inherited counterpart.
Paragraph-format operations now refuse a second unresolved serializable `w:pPrChange`; editor-only suggestion histories remain independently rejectable and save to at most one such child per paragraph.
