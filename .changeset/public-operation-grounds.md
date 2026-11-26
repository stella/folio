---
"@stll/folio-core": minor
---

A skipped document operation's `canonicalRefusal` is now the refusal's ground: a `gap`, a `rule` or a `condition`, and a stale document reports `staleRange`. Replacement text with `**bold**` or `***bold italic***` markers writes those words bold or italic, `replaceBlock` with `preserveFormatting: false` now applies, a join separator may hold a tab or a line feed, and inserted or replaced text that writes a note marker of its own is skipped as `protectedReference`. The `publicOps.comments` and `publicOps.unsupportedInline` gaps are retired.
