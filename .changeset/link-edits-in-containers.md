---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Setting and removing a hyperlink in the canonical session now works in paragraphs that hold content controls, direction and smart tag wrappers, equations and move markers. A content control or wrapper the range cuts stays one record and the link goes inside it; an equation in the range stays in place, unlinked, between two pieces of the link, which a later link removal treats as one. A suggested link change that would copy text inside a content control or wrapper is refused under the `linkSuggestionShape` rule.
