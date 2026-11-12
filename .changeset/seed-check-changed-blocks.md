---
"@stll/folio-core": patch
"@stll/docx-core": patch
---

The canonical session judges an edit's result by the blocks the edit replaced: the blocks it kept were judged before, and the comment check runs again only when the comments, the stories, or the order, kind and id of the replaced blocks' comment anchors change. docx-core exports `commentAnchorSequence`, the anchors a run of blocks adds as the comment check reads them.
