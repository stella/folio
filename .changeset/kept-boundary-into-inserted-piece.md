---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Rejecting a link while text inserted inside it is pending, and then accepting that text, places the link's kept range boundary inside the piece of the link the text keeps, as accepting the text first does, whichever way other insertions beside it are resolved. A range boundary another rejection kept beside it stays where it stands. Saved review provenance keeps how many boundaries a seam kept.
