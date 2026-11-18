---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Remove a container whose own content an accepted deletion took once the insertions it still holds are rejected, as accepting the deletion after rejecting them does. The acceptance marks those insertions with private resolution provenance that saved documents keep.
