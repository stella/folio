---
"@stll/folio-core": patch
---

Start each vertical caret step from where the caret actually is. ArrowUp and ArrowDown remembered the visual line and column of the previous step and only discarded them on a key the editor view handled, so a caret moved by anything else (a click, a find result, an agent edit) kept stepping from the line it had left: the next ArrowDown skipped a visual line and snapped back to the earlier column. The remembered step is now bound to the position it settled on, which keeps it where it is needed — a soft-wrap boundary belongs to both the line it ends and the line it starts — and re-resolves it everywhere else. A caret after a line-edge space also measures its column through the painted caret geometry, so a vertical step keeps the column the caret is drawn at instead of the one before the space.
