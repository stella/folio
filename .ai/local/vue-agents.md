## Vue Adapter

- Keep the package a thin binding over folio-core managers. Shared behavior belongs in
  framework-neutral core code, not duplicated composables or components.
- Preserve SSR-safe imports: DOM-dependent work starts only after mount or behind an
  explicit browser boundary.
- Keep the Vue and React adapter contracts aligned; update parity checks with any
  intentional divergence.
