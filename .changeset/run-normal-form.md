---
"@stll/docx-core": patch
---

A direct deletion joins the runs that meet at its point, and a direct format the runs that meet at its edges, when they can merge as the parser would consolidate them, so a format lying inside a concurrent direct deletion leaves the same runs in either order. Runs meeting elsewhere, authored boundaries among them, are left alone. The joins are part of the edit and its inverse splits them first, so undo restores the structure exactly; tracked edits and an inverse's own deletions and formats join nothing. `runSeams` lists where a paragraph's runs could merge.
