---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

A revision can name the pending suggestion it belongs to: `TrackedChangeInfo.suggestionId`, set from a new `suggestionId` on an operation's revision stamp. A record of one suggestion never merges into a record of another suggestion or into an ordinary tracked change. The id is held in memory only. `withoutPendingSuggestions` returns a document with every pending suggestion rejected, and the experimental canonical session's save writes the file through it, so a pending suggestion never reaches the file as an ordinary tracked change. If a suggestion cannot be rejected, the save refuses instead of writing it.
