---
"@stll/folio-core": patch
---

Escape theme-font, page-background and chapter-separator attributes on save;
narrow theme-font references and `w:shd` colours at parse; drop the attached
template reference from saved packages; convert only referenced media, with a
package-wide decode budget; preflight XML resource limits at unzip and bound
extracted text; bound `xmlns` declaration values; narrow run hyperlink targets
in the painter; let the save path materialize a header/footer added through
`createEmptyHeaderFooter`.
