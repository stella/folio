---
"@stll/docx-core": patch
---

Rejecting content made inside a pending insertion while another pending insertion stands between its pieces keeps the seam on that insertion: rejecting it later joins the pieces, or keeps its range markers in their cut, as rejecting both at once does.
