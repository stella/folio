---
"@stll/folio-core": patch
"@stll/docx-core": patch
---

The canonical session applies the public `insertSignatureTable` operation, directly or as a tracked change, and inserts a table from inside a table beside the outermost one, touching it, as the default engine does, including where no paragraph stands on that side: at the start or end of the body, or beside another table. A table inserted beside a paragraph whose break is pending deletion, where no pending join reads that side, goes beside the paragraph as it stands, as in the default engine. A table insertion can now name a table to go before or after (`beforeTable`, `afterTable`), and it travels to collaborators. A table operation in a story that ends with a table no longer refuses for want of a final paragraph it never had.
