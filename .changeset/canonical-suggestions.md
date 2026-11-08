---
"@stll/docx-core": minor
"@stll/folio-core": minor
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

The experimental canonical session supports pending suggestions. A document operation batch in suggested mode records tracked changes that carry the operation's suggestion id, never folding into a suggestion already pending. The review surface lists them as suggestions, including formatting suggestions. Accepting one adopts it as ordinary tracked changes by the accepting author; rejecting one rejects its revisions. Both are journalled, so undo restores the suggestion, and the React and Vue editors route accept and reject to the session. The engine gains `setRevisionInfo`, which rewrites revision metadata with an exact inverse, `suggestionRevisionInfos` and `pendingSuggestionIdsIn`.
