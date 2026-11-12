---
"@stll/folio-core": patch
---

In a canonical session, a public `insertAfterBlock` or `insertBeforeBlock` beside a paragraph whose break a pending join removes applies instead of being refused. The batch reads the document as it stood before the batch: a paragraph put after one that reads as part of the next goes after the paragraph the join leaves, even when the same batch splits that paragraph, and one put before it goes before the join's first paragraph; either copies the properties of the paragraph the join leaves. A suggestion reads the document as it stands and goes beside the paragraph it names. A paragraph inserted right before or after one the same batch splits applies with the split instead of being refused as overlapping it. The default engine reads pending joins through the same module.
