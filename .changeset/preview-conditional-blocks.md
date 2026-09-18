---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

The template fill preview hides conditional blocks a host reports as not applying: `TemplatePreviewValues.conditions` maps an `{% if %}` expression to whether its block applies, `false` drops the span from the opener through its `{% endif %}`, and the plugin state exposes those spans as `hidden` beside its substitution entries.
