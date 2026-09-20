---
"@stll/folio-core": minor
---

Write `w:bdo`/`w:dir` back around the text the editor still says they hold.

The projection lifts a transparent inline wrapper out of the paragraph's content tree and records the nesting on the `inlineWrapper` mark of the leaves it held. The save leg ignored that mark, so a wrapper only survived where the source paragraph's markup was replayed and an edited span lost it.

`fromProseDoc` now cuts the paragraph's inline sequence into maximal groups of equal stack before it builds runs, and closes the wrappers around each group. A revision stays outermost — `w:ins > w:bdo > w:hyperlink > w:r` — because folio already writes a revision outside the hyperlink it spans, the parse leg is revision-owned, and accepting or rejecting one is a range operation over the revision's own content. A group with nothing left in it writes no wrapper, so a wrapper whose text was deleted or rejected disappears with it.

A paragraph that was not edited keeps its authored markup, including a wrapper the author put outside a revision: selective save replays its bytes. Rebuilt from the editor, that order is written the canonical way round.
