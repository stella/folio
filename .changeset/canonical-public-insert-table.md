---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A canonical session applies a public `insertTable` beside a paragraph instead of refusing it: the table is built from the cell texts as the default engine builds it and placed as a paragraph beside that anchor is, a pending join included. It goes between the paragraph and the block next to it, and after a container's last paragraph it brings the empty paragraph the container ends with. Beside a table or any block other than a paragraph it stays refused under `publicOps.tableStructure`. The editor intent that inserts a table takes `paragraphAfter: "none"` to leave out the empty paragraph it otherwise writes after the table.
