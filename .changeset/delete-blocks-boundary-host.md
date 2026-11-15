---
"@stll/docx-core": patch
---

Deleting paragraphs directly reports the paragraph after them as modified when it takes the range boundaries, such as bookmarks, that the deleted paragraphs held. The modified paragraphs are read off the blocks the deletion writes back, so every paragraph it changes is reported.
