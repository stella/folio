---
"@stll/folio-core": minor
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Saving works for a picture copied into a header, footer or note whose id another binding in that part already holds. The picture is saved under a fresh id in that part, and the editor's document keeps its own. Before, the save failed with a `conflict` refusal.

- Saving during an input composition is now refused under the `composing` condition.
- `CanonicalSaveDiagnosticError` and `CanonicalImageRelationshipRefusalError` no longer carry a `gap`.
- The `save-serialization` gap is retired.
