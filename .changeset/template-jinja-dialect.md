---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

Template directives follow the docxtpl dialect of Jinja.

`@stll/folio-core` now scans markers with `@stll/template-conditions` 0.4:
`{{ path | filter(...) }}`, `{% if %}` / `{% elif %}` / `{% else %}` /
`{% endif %}`, `{% for alias in path %}` / `{% endfor %}`,
`{{ clause("Name") }}`, `{{ num("key") }}`, `{{ ref("key") }}`, and the
`{{ loop.* }}` counters.

`DirectiveKind` renames accordingly (`each` → `for`, `endeach` → `endfor`,
`elseif` → `elif`, `index`/`count` → the single `loop` kind). A `for`
`DirectiveRange` carries the iterated array path in `expr` and the loop alias in
the new optional `alias` field.

The React and Vue overlays rename the kind-derived class suffixes to match:
`--each` → `--for`, `--endeach` → `--endfor`, `--elseif` → `--elif` on
`.folio-template-directive`, `--each` → `--for` on `.folio-template-band-rail`,
plus a new `.folio-template-directive--loop`. Closer hover hints read `endif` /
`endfor`.
