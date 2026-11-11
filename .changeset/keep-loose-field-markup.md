---
"@stll/folio-core": patch
---

Saving an edited paragraph now keeps field markup that folio cannot assemble into a field. Examples are an instruction with no field around it, a field whose start a tracked deletion holds, and a field end a link holds alone. Before, those `w:instrText` and `w:fldChar` runs were left out of the saved paragraph. The parser now keeps them as their source markup, so the editor and the experimental canonical session save them byte for byte, and the canonical session opens such documents.
