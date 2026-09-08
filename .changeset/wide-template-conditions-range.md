---
"@stll/folio-core": patch
---

Widen the `@stll/template-conditions` dependency to `>=0.4.0 <1.0.0`. folio-core consumes only the scanner surface, so a host monorepo that already provides template-conditions as a workspace package satisfies the range across 0.x minors instead of installing a second registry copy beside its own.
