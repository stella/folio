---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Keep the list-number fields shown in a paragraph's list marker when the paragraph is saved. Such a field, and the tab after it, now stay in the paragraph content as preserved markup and are written back as they were read.

The reader now draws a list-number field in the list marker only when the field opens a numbered paragraph whose marker has text of its own. A list-number field that follows other text, or that stands in a bulleted paragraph, is shown inline where it stands.
