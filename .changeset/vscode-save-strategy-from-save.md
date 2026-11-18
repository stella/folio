---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

Report the actual DOCX write path through shared save callbacks, expose scoped selective-write refusals separately from fidelity warnings, and use the canonical journal without a ProseMirror paragraph tracker in React and Vue. Remove the obsolete `markPackageChange` export; canonical callers use the journal and the legacy tracker remains available.

Align Vue’s observational tripwire callbacks with React: host exceptions propagate outside serializer fallback handling, after save acknowledgement.
