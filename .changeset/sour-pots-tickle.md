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

The existing layout painter is unchanged and still paints the editor. The
display-list DOM backend is additive in this release.
