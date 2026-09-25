---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Keep the `mc:AlternateContent` around a DrawingML shape or text box on save, `mc:Fallback` included, while the shape is unedited, through the model and the editor. An edited shape is regenerated from the model without its Fallback, which would otherwise contradict the new Choice for consumers that read only the Fallback.
