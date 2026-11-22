---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

Refuse a tracked join, a tracked deletion, a tracked row operation and a tracked terminal table insertion that would record over a paragraph mark already under review as `pendingParagraphMarkDeletion` or `pendingParagraphMarkInsertion`, by what the mark has pending, instead of `revisionConflict`; the canonical session reports them under the public reasons of the same names.
