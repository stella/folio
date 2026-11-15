---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Accepting a tracked deletion of a whole paragraph no longer leaves its bookmarks or move ranges standing between blocks. Every bookmark, move range boundary and comment anchor the paragraph held goes to the nearest paragraph that stays, as deleting the paragraph directly leaves them, and the edit reports that paragraph as changed. Before a table with no paragraph ahead of it, the paragraph stays holding them. The editor resolves paragraphs the same way. `paragraphRemains` is no longer exported from `@stll/docx-core/ops`.
