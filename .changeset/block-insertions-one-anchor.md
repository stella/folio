---
"@stll/docx-core": patch
---

A direct paragraph insertion after a paragraph whose inserted mark a tracked insertion after it took over goes in before that inserted break: its last new paragraph takes the inserted mark and the anchor's mark is its own again; the tracked paragraph holding the anchor's former mark records its property change against the new holder of the inserted mark. Concurrent direct and tracked insertions after the same paragraph are ordered direct first in either sequence order, so after the last paragraph of a story both orders leave the same document, the tracked paragraphs last.
