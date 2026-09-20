---
"@stll/folio-core": patch
---

Keep one equality over a paragraph's stated numbering.

The marker tier (`prosemirror/listMarker.ts`) and the layout tier (`layout-bridge/convert/toFlowBlocks.ts`) each carried a private copy of the same two helpers, byte-identical bodies under different names. `isListNumPr` and `sameListNumPr` now live once beside `ListNumPr`, and both copies are gone: an equality that decides whether a list renumbers is not a thing to have two of.
