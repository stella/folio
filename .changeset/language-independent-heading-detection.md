---
"@stll/folio-core": patch
---

Headings are recognised from the paragraph's effective `w:outlineLvl` and the style's built-in `w:name` rather than from an English style id. A localized Word writes `Nadpis1`, `berschrift2`, `Titre3`, `Nagwek4` or `Cmsor5` for the same built-in heading, so id matching found English output and nothing else: those paragraphs were missing from the outline sidebar and generated tables of contents, arrived at the AI snapshot as plain paragraphs, and exported to Markdown without `#`. One classifier (`docx/builtInStyles.ts`) now answers for every consumer, with `w:outlineLvl` 9 meaning body text rather than a tenth level, and the bilingual builder's language-list regex is gone. Documents folio creates carry outline levels on their heading styles, and define the `Quote` style a Markdown blockquote compiles to.
