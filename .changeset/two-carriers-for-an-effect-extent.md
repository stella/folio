---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Write a `wp:effectExtent` back on the element that authored it, so a wrap child's own reservation survives a rebuild.

`CT_Inline` and `CT_Anchor` declare a `wp:effectExtent`, and so do `CT_WrapSquare` and `CT_WrapTopBottom`. They are two values: the drawing's is the object's own effect reservation, the wrap child's is the reservation the text flow is computed against. folio read only the drawing's, into `Image.padding`, and wrote it back there, so a wrap child's own reservation round-tripped an untouched document on the strength of its captured bytes and was gone the moment anything forced the serializer.

`ImageWrap` gains `effectExtentSlots`, the `distanceSlots` shape one element over: `drawing` and `wrapChild`, each holding the element's four sides. The value in force stays where its consumers read it, on `Image.padding`. `resolveEffectExtents` decides the rebuild the way `resolveWrapDistances` decides the insets — each reservation goes back on the element that stated it while the drawing's is unmoved, and once an editor has resized it the rebuild states the value in force on the drawing alone rather than keeping a wrap reservation computed against a shape that is no longer there. The slots ride through the editor as `wrapEffectExtentSlots` on the image, shape and text-box nodes.

A shape and a text box have never had a reservation of their own — the rebuild wrote `l="0" t="0" r="0" b="0"` on every one of them — and now keep the one they were authored with.
