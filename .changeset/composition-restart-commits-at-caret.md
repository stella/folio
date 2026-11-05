---
"@stll/folio-core": patch
---

In the canonical session, a composition is no longer refused when the browser restarts it part way through, or when it starts at the beginning of a paragraph. It commits its final text once, at the caret.
