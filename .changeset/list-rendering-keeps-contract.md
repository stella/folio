---
"@stll/folio-core": patch
"@stll/docx-core": patch
---

A list normalization replaces the paragraphs whose rendering it changed through `withParagraphRenderings`, a docx-core helper that refuses a replacement differing in anything but its list rendering and the formatting fields that rendering derives, and keeps a document known to meet the operations contract known to meet it. The next edit's operation no longer checks the whole document against the contract.
