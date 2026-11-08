---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

The experimental canonical session now opens documents whose text holds a literal tab character, and documents with a header, footer or note that holds no blocks. The session reads a literal tab as a tab element and gives an empty story an empty paragraph, in its own model only. A part whose story is unedited saves exactly as it was read; a story the user edits saves in the session's spelling. `replaceStoryBody` is exported from `@stll/docx-core/ops`.
