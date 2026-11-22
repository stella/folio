---
"@stll/docx-core": patch
---

Rejecting an insertion that keeps nothing leaves content made inside it with the join that insertion had to the record beyond it, as rejecting it after accepting that content does. It still cancels the join the content had toward the rejected insertion itself.
