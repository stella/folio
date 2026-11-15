---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

The React comments column measures its place beside the page again once the editor's scroll container mounts, and follows it as that container resizes, so the column no longer stays where the page was before the outline took its track. React and Vue place the column with one shared rule, `commentsColumnStart`.
