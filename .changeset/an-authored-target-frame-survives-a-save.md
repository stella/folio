---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Keep `w:hyperlink/@w:tgtFrame` as the document authored it. Any frame name outside `_blank`, `_self`, `_parent` and `_top` was mapped to `_blank` at parse, so a saved file no longer said what the source said. The allow-list clamp now lives where a DOM anchor or a navigation is produced, in one owner (`anchorTargetAttrs`) that every rendered document, editor popover and `window.open` in core, React and Vue goes through, so the `target` and the `rel` that must accompany it are decided in a single place.
