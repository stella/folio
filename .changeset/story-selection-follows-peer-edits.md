---
"@stll/docx-core": patch
"@stll/folio-core": minor
---

In a canonical room, an open header, footer or note editor keeps its caret and selection where its text went when a peer edits the story: the session holds the selection each such editor shows (`CanonicalSession.recordStorySelection` and `storySelection`) and moves it over every peer's edit by paragraph and offset, as it moves the body's. `transformPosition` maps a position past zero-width children, such as a caret just after a note's own mark, over a paragraph split or direct deletion that does not cut at its offset.
