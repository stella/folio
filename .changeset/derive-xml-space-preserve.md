---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

Derive `xml:space="preserve"` from the text instead of storing it. `TextContent.preserveSpace` is gone: whether `<w:t>` needs the attribute is a pure function of its text, and a stored copy of a derived fact only drifts — the editor lost it whenever two adjacent runs merged, because ProseMirror has nowhere to carry it. Both serializers now call the same `requiresXmlSpacePreserve`, which `@stll/docx-core` exports.
