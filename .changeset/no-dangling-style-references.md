---
"@stll/folio-core": minor
---

No document folio authors references a style it does not define. A comment or note reference mark carried `w:rStyle` pointing at `CommentReference`/`FootnoteReference`/`EndnoteReference` whether or not the package declared them, so the mark lost its superscript; the generic style set stopped at `Heading4` while the report builder accepts six levels and applies `TableGrid`; and `generateTOC` wrote `TOCHeading`/`TOC1` regardless of what the open document calls its TOC styles. The table of contents now takes its styles from the document through the built-in classifier, and writes no style id rather than a dangling one. A property test holds every authoring path to this.

**Breaking for direct callers of the TOC command:** `generateTOC` is now `generateTableOfContents({ title })` and `insertTableOfContentsInView(view, { title })`. The title was a hardcoded English "Table of Contents", which is wrong in every document that is not in English; this layer has no locale, so the host supplies the string.
