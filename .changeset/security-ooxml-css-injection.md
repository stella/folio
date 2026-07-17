---
"@stll/folio-core": patch
"@stll/folio-react": patch
---

Prevent CSS and OOXML injection from attacker-controlled document/collaboration
values. Colors are validated to a strict hex/`auto` format at a single
`colorResolver` choke point (closing themed table-fill and diagonal-border
`url()` injection and the pasted `data-bgcolor` path); comment `paraId`/`textId`
are validated to 8-hex at parse and XML-escaped on serialize; run/paragraph/
table/style color and theme attributes are XML-escaped and hex-validated; inline
and block SDT raw properties are replayed only when they are a single
well-formed `w:sdtPr`/`w:sdtEndPr` element (otherwise synthesized); remote
collaborator colors are validated before use and painted via `backgroundColor`
(not the `background` shorthand); and controlled comments are sanitized before
becoming editor state.
