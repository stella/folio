---
"@stll/folio-core": minor
---

A review resolution that removes the paragraph a pending suggestion's record names no longer throws. A proposed split follows its mark in saved text, a proposed paragraph stands beside the saved paragraph next to it, and a record whose words no saved paragraph holds is marked `retired: "anchorRemoved"`; loading it reports it stale with that reason.
