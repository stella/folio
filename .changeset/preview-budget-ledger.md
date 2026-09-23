---
"@stll/folio-core": patch
---

Charge the package preview budget from a per-parse ledger of the previews the parse built instead of walking the whole document model. Previews are now charged in the order the parse builds them, so a package over its preview allowance may keep a different set of previews than before.
