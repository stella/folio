---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

Paint the editor's pages from the display list, behind a renderer option.

`buildDisplayList` now takes the page furniture it previously could only
report: page borders, watermarks, footnote bodies, header and footer stories,
and the package's own embedded font faces. A construct that is supplied is
painted; one the document has but the caller withheld is still reported; one
the document does not have is neither. `layoutDocxHeadless` produces all of it,
so an export paints the pages an editor paints rather than bare bodies.

The DOM backend places every code point at the advance the layout engine
measured instead of letting inline layout advance it, so the two backends agree
on glyph positions to within the browser's 1/64 px layout quantum. Cursively
joined clusters stay in one box, because only shaping can choose a positional
form; the advances inside such a cluster are the shaper's.

`pageRenderer` on the React and Vue editors selects which renderer paints the
pages, defaulting to the existing painter. Only painting is swapped: page
shells, virtualization, the fingerprint comparison that skips an unchanged page
and the painted event stay shared, so incremental repaint behaves the same
under either renderer.
