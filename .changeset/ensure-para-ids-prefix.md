---
"@stll/folio-core": patch
---

`ensureParaIds` now finds paragraphs by namespace, so a package that binds WordprocessingML to a prefix other than `w:` (or to the default namespace) gets its paragraph IDs instead of being reported `alreadyComplete` with none assigned. A part whose prefix bindings it cannot follow is refused with `EnsureParaIdsError`. The save path's text-level patches refuse non-conventional prefixes and fall back to a full repack, appended styles and footnote reference marks respect the part's own prefix, and the parser reads document, header, footer and note parts under a default namespace.
