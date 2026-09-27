---
"@stll/folio-agents": minor
---

An operation id names one operation for the document session. Resending a `suggest_changes` operation that already applied or queued, with the same id and content, no longer applies it again: the result lists it under the new `replayed` field with its original receipt. Reusing an id for a different operation fails the call. The reviewer bridge scopes the session to its reviewer; any other bridge is its own session.
