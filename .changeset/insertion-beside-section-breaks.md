---
"@stll/docx-core": patch
---

A paragraph put after one that ends a section now applies as a tracked change wherever its suggestion does, and rejecting every change gives the original back: an added paragraph left with nothing goes from before the block-level bookmarks after it, which stay where they stood; an added paragraph opening a section before trailing bookmarks goes as it does with the paragraphs joining it set aside; a mark before a table resolves where the body ends past bookmarks, or in a block another producer ended it with; a paragraph with content of its own whose added mark joins it into an emptied added paragraph keeps its properties; and a mark before a table that has nothing to anchor its removal stays pending in the proof as other blocked marks do. A suggested deletion whose accepted join would run words into a last paragraph keeping another reviewer's pending property change is refused, as one restyling that paragraph is.
