---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Prevent untrusted DOCX assets from affecting the host page. Embedded fonts are
registered under per-document scoped family names (resolved through the font
resolver) so a document embedding a face named after a host UI family can no
longer shadow it page-wide, and watermark dialogs validate external image
targets against an http/https allowlist (with a defensive guard before emitting
an external relationship) so `file:`/UNC/other-scheme targets cannot be written
into exported documents.
