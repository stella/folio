---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Accepting an inserted paragraph break whose deletion is still pending leaves that deletion holding the split's cut, with when the split was made (`splitDate` on the paragraph mark, saved as its own ignorable review-provenance attribute, `folio:splitDate`). Accepting the deletion later rejoins the same records as accepting both at once. An accepted change at the edge of a deleted break's cut retires that cut as it retires an inserted break's, so accepting the change and the deletion together leaves the pieces apart, as accepting them one at a time does.
