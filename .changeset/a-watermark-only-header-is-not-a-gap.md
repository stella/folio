---
"@stll/folio-core": patch
---

Stop reporting a watermark-only header as a story that never arrived.

Word puts a watermark in a header part that holds nothing else, so that
header converts to no paintable content and the display list reported it as
"the header part this page selects was not among the supplied stories" on
every page of every watermarked document. The watermark painted from that same
part proves it reached the producer and was read, so it is no longer reported
as a gap; a header that names content the builder really did not get still is.
