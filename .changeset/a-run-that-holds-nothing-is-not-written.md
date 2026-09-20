---
"@stll/folio-core": patch
---

Stop writing a run the next parse drops, so a package that holds a text box is stable from the first save.

Whether a run is kept was decided in three places. The parser drops a run that holds no payload, the consolidator dropped a run that held none, and the serializer wrote one anyway. A text box makes the three disagree: the shape is claimed by a second pass over the paragraph, so between the two passes its run is legitimately empty. The consolidator dropped that carrier, which cost the run its `w:rPr`, and when the carrier survived, the save wrote `<w:r><w:rPr…/></w:r>` for it and the next parse dropped that run, so save 2 differed from save 1.

`runHoldsPayload` now owns the question and the parser's keep rule, the hyperlink walk, the consolidator and the serializer all ask it. The consolidator keeps a payload-less run as the boundary its comment always claimed it was, the serializer writes no run that holds nothing, and a link admits a run on the same terms a paragraph does.

Two defects the text-box pass hid behind that are fixed with it: it enriches a paragraph before markers carried over from the body are put in front of its content, so the positions it walks are the ones its own `w:p` produced; and filling a carrier now advances its cursor, so a second text box in the same paragraph no longer inserts itself in front of the first.
