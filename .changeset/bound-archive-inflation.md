---
"@stll/folio-core": minor
"@stll/docx-core": patch
---

Bound archive inflation when opening documents. A document now fails to open when one of its parts inflates past the size the archive declares for it, or when a markup part or the package as a whole exceeds the compression-ratio limit; `maxCompressionRatio` in the archive and unzip options adjusts that limit.
