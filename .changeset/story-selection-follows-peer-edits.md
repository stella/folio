---
"@stll/docx-core": patch
"@stll/folio-core": minor
---

In a canonical room, an open header, footer or note editor keeps its caret and selection where its text went when a peer edits the story: the session holds the selection each such editor shows (`CanonicalSession.recordStorySelection` and `storySelection`) and moves it over every peer's edit by paragraph and offset, as it moves the body's. `transformPosition` follows inline content inserted in a position's paragraph, and the run splits and joins beside it, by offset instead of giving up on the paragraph.
