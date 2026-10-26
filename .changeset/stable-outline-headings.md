---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": patch
---

`HeadingInfo` carries an `id` (the paragraph's `paraId`, else its position), which outline items now use instead of the heading's position, and `resolveOutlineHeadings` places collected headings on a layout's pages while keeping unchanged entries. The React outline recollects headings on each completed layout instead of on document history changes and timers, so an edit that moves no heading keeps the same heading list and page numbers.
