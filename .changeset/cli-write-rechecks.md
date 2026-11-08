---
"@stll/folio-cli": patch
---

A transaction id retried after its earlier attempt was discarded now replays the retry's receipt. MCP writes re-check that the directory they commit into is still inside the allowed roots, and `folio render` re-checks its output directory, temporary file, and destination before the final rename.
