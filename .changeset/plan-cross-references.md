---
"@stll/folio-core": minor
---

The block manifest reports a paragraph's cross-reference fields (`crossReferences`): each `REF` field whose result stands in the paragraph's text, by where that result sits, with its code, the bookmark it names, the paragraph that bookmark marks and whether it holds all of that paragraph's content. A plan can then keep such a field as a field, with a reference over its span to that paragraph, instead of its result as text. Verification predicts no text for a reference without a number switch whose target holds a field that numbers itself, since that number is counted only when the document is laid out.
