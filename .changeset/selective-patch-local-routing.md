---
"@stll/folio-core": patch
---

Keep a paragraph edit local in documents whose text boxes the model represents differently from the source. The selective save no longer requires the whole `word/document.xml` to hold the same `w:p` elements as the model's serialization: each changed paragraph is located by its `w14:paraId`, or, when the source names none, by its position within its story (the main flow or the text boxes, `mc:Fallback` excluded). A typed character beside a text box with an `mc:Fallback`, or beside a VML text box, now saves only that paragraph instead of falling back to the full repack. The patch still refuses a paragraph it cannot place unambiguously, one inside `mc:AlternateContent`, one whose container or position changed, and a part whose WordprocessingML prefix is not `w`.
