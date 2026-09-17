---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Classify a drawing as `native`, `replayable` or `opaque` through the predicate the run serializer already uses, so a document is no longer opened read-only because a header carries a logo; only `opaque` content blocks editing, and `DocxCompatibility` gains a `drawings` list at `schemaVersion: 2`. Regenerating a picture now round-trips `a:graphicFrameLocks` and `wp:effectExtent`, and a rasterized shape group is marked `previewOnly` so the editor declines to resize it rather than replacing the group with one child picture. Shape drawings Folio cannot model — unmodeled effects and 3-D, `wpg:wgp` groups without a preview, a `w:pict` with no resolvable image, an `mc:AlternateContent` whose every branch declines — are preserved verbatim instead of dropped. Field results are no longer missing from the AI-facing block text, so a paragraph carrying a cross-reference reads as the text Word shows.
