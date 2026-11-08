---
"@stll/folio-core": patch
---

Encrypted documents are read only when their key descriptors declare the 16-byte AES block and their compound-file FAT lists each sector once, within the sectors the file holds. PDF export decodes a PNG of at most 64 megapixels, and at most 256 megapixels of PNG per document; a larger raster is refused before it is inflated.
