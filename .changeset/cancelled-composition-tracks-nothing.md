---
"@stll/folio-core": patch
---

With Track Changes on, cancelling an IME composition records no changes. It used to record the paragraph's existing text as an insertion, so rejecting that change removed the text, or leave part of the cancelled text in the document. A committed composition records only the text it added.
