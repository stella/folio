---
"@stll/folio-core": minor
---

Build a display page when it is painted, and let a run say which document its
positions belong to.

The editor's display-list renderer built the whole document's list on every
layout run, then painted the three pages on screen from it. Measured on a
ninety-one-page document, that cost 49.6 ms per painted page against the
painter's 0.567 ms: work proportional to the document where the page container
does work proportional to the screen. `createDisplayListBuilder` computes what
belongs to the document once and each page when it is asked for, which brings
the same measurement to 0.43 ms per painted page.

`DisplayGlyphRun.pmRange` now carries the story it addresses, so a header, a
footer and a note run can carry one at all. A page is not one document: the same
position means a different character in the body, in each header and footer
part, and in each note, so a range that did not name its story could only be
used for the body. The producer used to drop the others and the painter used to
strip them; both now travel, named, and the DOM backend writes them onto the
element that paints the run.
