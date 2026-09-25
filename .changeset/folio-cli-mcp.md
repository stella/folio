---
"@stll/folio-cli": minor
---

Add `folio mcp`, a stdio MCP server over the same tool registry: each tool takes its folio-agents arguments plus a file envelope (`path`, `fileVersion` required on changes, `destination`, `overwrite`, `txId`, `allowRepack`, `mode`), every path must resolve inside an allowed `--root`, reads are bounded by block count, match count, and response size with version-bound cursors, and the operation schema and server rules are served as resources.
