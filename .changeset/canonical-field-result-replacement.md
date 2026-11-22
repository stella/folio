---
"@stll/folio-core": patch
---

On the canonical engine, a replacement whose word diff places a change inside a field's result replaces the whole field with its text: the field goes as a tracked deletion (or directly) and the edited words come in as text. Changes that only meet a field at its edge stay word-level. The legacy engine already did this.
