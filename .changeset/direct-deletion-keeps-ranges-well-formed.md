---
"@stll/docx-core": patch
---

A direct deletion that covers one end of a comment range or a tracked move range, and not the other, keeps that end where it stands, so the range shrinks to the deletion's edge and the document can still be saved. A deletion that covers both ends removes the whole range; a range that was already malformed is left to the rest of the deletion.
