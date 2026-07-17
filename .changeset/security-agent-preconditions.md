---
"@stll/folio-core": patch
"@stll/folio-agents": patch
---

Harden agent operation integrity: `read_document`/`read_section`/`find_text`
now return a per-block text hash and `suggest_changes`/`add_comment` accept a
caller-supplied precondition, so a stale edit prepared against content the model
read earlier is detected instead of being stamped from a fresh apply-time
snapshot. Operation-mode and block-range lookups use own-property checks so
prototype keys (`__proto__`, `constructor`) can no longer crash the API.
