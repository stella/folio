---
"@stll/docx-core": patch
---

Concurrent tracked changes no longer diverge: a concurrent record cut at a tracked change's edge or inside its range (another tracked edge, an insertion point, a format's end, an inline split or join, a paragraph split), a direct deletion reaching one of its edges and not the other, and two overlapping tracked deletions (nested, equal or crossing) are refused with `unsupportedPair` in either order. Tracked changes apart from every concurrent cut are rebased as before.
