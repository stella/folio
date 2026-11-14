---
"@stll/docx-core": patch
---

Concurrent tracked changes no longer diverge: a tracked change whose edge meets a concurrent record cut (another tracked edge, an insertion point, a format's end, an inline split or join, a paragraph split), a direct deletion reaching one of its edges and not the other, and two overlapping tracked deletions (nested, equal or crossing) are refused with `unsupportedPair` in either order. Tracked changes apart from every concurrent cut are rebased as before.
