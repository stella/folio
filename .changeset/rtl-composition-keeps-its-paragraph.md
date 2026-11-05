---
"@stll/folio-core": patch
---

Composing right-to-left text in a left-to-right paragraph no longer inserts it twice. The paragraph's direction now changes once the composition has ended, rather than during it, which also ended the composition early and, with Track Changes on, left later edits untracked.
