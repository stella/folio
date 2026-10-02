---
"@stll/folio-cli": minor
---

`suggest_changes` (MCP and `folio suggest`) takes `replaceAll` (every match in the body and tables, run formatting kept) and `addComment` (by `blockId`, or by a `quote` found in exactly one block) inside a batch. `describe_capability` returns a parameter outline, short guidance and an example; `detail: "full"` returns the full schema. `read_document` prints a table row as one `| [id] cell | [id] cell |` line, and `folio read` blocks in a table carry `tableCell`. A `suggest_changes` result reports `replaced` counts and `commentIds`, and the server instructions say a successful write needs no verification read.
