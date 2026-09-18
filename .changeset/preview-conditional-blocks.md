---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

The template fill preview acts on conditional blocks a host has ruled on: `TemplatePreviewValues.conditions` maps an `{% if %}` expression, or the bare field path its filter chain hangs off, to whether the block applies; `false` drops the span from the opener through its `{% endif %}`, `plain` mode also drops the tag lines of a block that does apply, and the paged layout drops the blocks those spans swallow whole so the pages paginate without them.
