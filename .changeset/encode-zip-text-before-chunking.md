---
"@stll/utf8-boundaries": minor
"@stll/folio-nuxt": patch
"@stll/docx-core": patch
"@stll/folio-core": patch
"@stll/folio-agents": patch
"@stll/folio-cli": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Saving could corrupt astral characters such as emoji at 16 KiB boundaries inside document XML parts. Upgrade to keep them intact.
