---
"@stll/docx-core": patch
---

A table, row or cell with a pending tracked formatting change can now take another formatting change. In suggesting mode the change joins the pending record, which still holds the formatting from before both, so rejecting restores the original and accepting keeps the latest. An untracked change keeps the pending record. Before, any further change to that element was refused until the pending one was accepted or rejected.
