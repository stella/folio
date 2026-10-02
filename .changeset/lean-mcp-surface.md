---
"@stll/folio-cli": minor
---

`folio mcp` lists only the frequent tools (`read_document`, `find_text`, `read_comments`, `read_changes`, `suggest_changes`, `add_comment`) with compact schemas, and reaches the rest through `list_capabilities`, `describe_capability` and `invoke_capability`. Results carry only what the next call needs (`read_document` returns `[id] text` lines, a write returns the new `fileVersion` and what it produced), a failure is `{ error: { code, message, hint, retryable } }`, and arguments are read leniently with an `Input read:` note.
