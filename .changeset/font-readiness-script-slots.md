---
"@stll/folio-core": patch
---

Wait for the complex-script and East-Asian fonts before the first layout, not just the Latin ones.

The font-readiness gate collected only the `ascii` and `hAnsi` slots of a run's font. Word writes the Arabic and Hebrew face into `w:cs` and the CJK face into `w:eastAsia`, so a document that styles those scripts the standard way had its font ignored by the gate. The first layout then measured a fallback face while the painter later drew the real one once it loaded, and the two disagreed for exactly the scripts whose advances differ most from a Latin fallback.
