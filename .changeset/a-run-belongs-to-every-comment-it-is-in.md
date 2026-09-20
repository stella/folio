---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Keep a run inside overlapping comment ranges in every comment it belongs to. The painted run advertised only the first id, so hover and active styling, and the sidebar anchor, answered for one comment and denied the other. A run in more than one range now also carries `data-comment-ids`, the whole membership, and the adapters read the painted anchors through one shared helper rather than the first id alone; a run in a single range paints exactly what it painted before.
