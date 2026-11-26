---
"@stll/docx-core": patch
---

A package's id indexes hold less memory: what each record contributes is remembered in one weak map per context instead of a map per record, and a record whose ids are those of the one record it holds shares that record's list.
