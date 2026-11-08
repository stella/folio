---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

A refusal from the experimental canonical session now says why: `CanonicalSessionRefusalError` carries a `ground`, either `{ gap }` for a capability the session does not have yet or `{ rule }` for a rule the document states, starting with `contentControlLock`. The editor shows a rule refusal as its own translated message, "This part of the document is locked." in English, which hosts can translate through the editor messages. `CanonicalSessionRefusalError` no longer has a top-level `gap`; read `error.ground.gap`.
