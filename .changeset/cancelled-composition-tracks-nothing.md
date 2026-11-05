---
"@stll/folio-core": patch
---

With Track Changes on, cancelling an IME composition leaves the document as it was. It used to record the paragraph's existing text as an insertion, so rejecting that change removed the text, leave part of the cancelled text in the document, or drop the revision of text typed earlier. A committed composition records only the text it added.
