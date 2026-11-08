---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Suggested batches can now carry `commentOnRange` and `commentOnBlock` in both engines, and a comment attached to any suggested operation now belongs to its suggestion. A proposed comment is listed with its suggestion (a new `"comment"` kind in `getSuggestions`; a suggestion that only proposes a comment applies as `"direct"`), is left out of saved files while pending, travels through `exportPendingSuggestions` and `loadPendingSuggestions`, is removed when the suggestion is rejected and stays as an ordinary comment when it is accepted. In the canonical session, a tracked or suggested edit's attached comment now covers the text the edit changed, as in the default engine, instead of the whole replaced range.
