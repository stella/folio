---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

In the canonical session, suggested batches can now carry `commentOnRange` and `commentOnBlock`, and a comment attached to any suggested operation now belongs to its suggestion. A proposed comment is listed with its suggestion (a new `"comment"` kind in `getSuggestions`; a suggestion that only proposes a comment applies as `"direct"`), is left out of saved files while pending, travels through the pending-suggestion export and restore, is removed when the suggestion is rejected and stays as an ordinary comment when it is accepted. A tracked or suggested edit's attached comment now covers the text the edit changed, as in the default engine, instead of the whole replaced range.
