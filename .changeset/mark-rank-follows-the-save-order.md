---
"@stll/folio-core": patch
"@stll/folio-react": patch
---

Nest a leaf's marks in the order the save leg nests its elements.

Mark rank was the key order of `MARK_EXTENSIONS`; element nesting is decided in `extractParagraphContent`. Nothing held the two together and they had drifted: `hyperlink` was registered before `insertion`/`deletion`, so the editor DOM read `<a><span class="docx-insertion">` while the save wrote `w:ins > w:hyperlink > w:r`, and a rule or a walk written against one nesting was written against a document the other leg does not produce.

`MARK_NESTING_ORDER` now states the order once, outermost first, and `StarterKit` registers from it; the record keeps its job of naming the marks that exist and building them. Only the marks the save leg gives an element of its own are ranked by it: a comment range around everything a leaf produces, then the revision, then the transparent wrapper the revision takes inside it, then the link and its runs. The rest are run properties with no element in OOXML, so the save leg ranks them against nothing and they stay where presentation put them. A test serializes a leaf under a comment, a revision, a wrapper, a link and bold through both legs and holds their nesting to each other.

The change is visible in the ProseMirror layer, where an inserted or deleted link is now inside the change's span: the adapter's link colour would paint over the redline, so the anchor inherits it, and the display modes that drop the change colour hand the link colour back.
