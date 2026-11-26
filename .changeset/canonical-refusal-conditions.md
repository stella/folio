---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

A refusal can now name a condition of the session or the request, such as a read-only editor, a composition in progress, a document that is not open yet, a request naming a missing comment, or an exhausted id space. Its ground is `{ condition }` from the new `CANONICAL_CONDITION`, and the editor shows the new catalog message `editor.changeNotApplied` ("This change couldn't be made.") in every shipped locale.

- Section property and comment changes refused for such a condition report it. `CanonicalSectionPropertiesResult` and `CanonicalCommentResult` now carry `ground` instead of `gap`, and `CanonicalCommentError` carries `ground`.
- The `authority-routing`, `comment-model-edits`, `section-properties` and `direct-model-edits` gaps are retired.
- In Vue, `setDocument(doc)` now opens `doc`, as `loadDocument` does, instead of refusing.
- In React, a section change made before the editor is ready shows the catalog message instead of the internal one.
