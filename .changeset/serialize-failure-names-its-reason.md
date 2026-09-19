---
"@stll/folio-core": patch
---

Say why a comparison could not be saved. `CompareDocxSerializeError` carried the reason in `cause` and a constant sentence in `message`, so every distinct save failure read identically in a log, a report or a census and none of them could be told apart without a debugger. The message now names the underlying failure, and `cause` still carries it structured.
