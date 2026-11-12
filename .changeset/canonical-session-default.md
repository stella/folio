---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

The canonical session is now the default editing engine in the React and Vue editors. An editor without `experimentalSession` runs the canonical session, which refuses an edit it cannot make yet with a notice. `experimentalSession="legacy"` (and the `experimentalSession` option of `useDocxEditor`) selects the previous ProseMirror-owned engine, which collaboration still requires. `experimentalSession` takes `EditorSession` (`"canonical" | "legacy"`) from `@stll/folio-core/types/canonicalCapabilities`, and `usesCanonicalSession` treats an unset session as canonical.
