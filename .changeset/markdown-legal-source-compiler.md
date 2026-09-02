---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

Parse legal-source drafts as GFM markdown plus `@` directives with `marked`: clause bodies, list items, and table cells keep inline emphasis, links, and code spans; markdown lists and pipe tables outside a directive become real list and table blocks. `compileMarkdownToContent` and `sanitizeExternalUrl` move into `@stll/docx-core`, and `@stll/folio-core`'s `fromMarkdown` now wraps them.
