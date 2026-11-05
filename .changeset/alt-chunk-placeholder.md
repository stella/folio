---
"@stll/folio-core": minor
"@stll/folio-react": patch
"@stll/folio-vue": minor
---

A `w:altChunk` is no longer invisible: the page shows a tinted, non-editable block with the chunk's own text when it has any, otherwise a localised "Content not displayed" marker, and saving still keeps the chunk byte for byte. `useDocxEditor` gains an `altChunkMarkerText` option carrying that marker, and `ParagraphBlock` gains `fixedPosition` for a block that exists only in layout.
