---
"@stll/folio-core": patch
---

Every save entry that takes the bytes of an opened package (`repackDocx`, `attemptSelectiveSave`, the canonical save, `updateDocumentXml`, `updateXmlFile`, `updateMultipleFiles`, `addRelationship`, `addMedia`) now reads them with the main document at the standard path, as the reader does. Bytes handed over as the producer wrote them, with the main document at another path, save to one main document part named by the package relationships and the content types.
