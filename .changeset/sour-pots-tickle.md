---
"@stll/folio-core": minor
---

Add a painter-neutral display list and a native PDF backend that consumes it.

`buildDisplayList` turns a laid-out document into an ordered list of paint
primitives per page: glyph runs carrying the advances layout was decided on,
filled and stroked rects, lines, images, clip, rotate and opacity groups, link
annotations and a heading outline. Two backends consume it and nothing else:
`renderDisplayListToDom` paints it into DOM elements, and `writePdf` writes a
PDF with subset TrueType faces, PNG and JPEG images, vector borders and
shading, hyperlinks and an outline. A backend that reads layout data the
display list does not carry now fails a dependency-cruiser rule rather than a
review.

`exportDocxToPdf` composes the whole chain without a browser:
`layoutDocxHeadless` paginates a package through the measurement seam, and
`installHeadlessMeasureProvider` supplies that seam from parsed font binaries
instead of a canvas. Output is deterministic: `timestamp` is required rather
than defaulted, so two exports of one document are byte-identical.

A font source supplies every binary that carries part of a face, not one, and
both measurement and embedding resolve each code point to the binary that
covers it. Families are routinely shipped split by script, so a Czech, Slovak
or Polish document needs two subsets of one family in the same paragraph;
resolving per face rather than per code point would paint an empty box for
every character outside whichever subset was chosen. A code point no supplied
binary can encode is reported in `unencodable` rather than painted silently,
and `strictGlyphCoverage` turns it into a failure for a caller who would
rather not ship the page at all.

The existing layout painter is unchanged and still paints the editor. The
display-list DOM backend is additive in this release.
