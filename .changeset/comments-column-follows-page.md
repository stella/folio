---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

The comments column is placed from the panel layout's numbers (the scroll area's width after the outline track, the widest page and the comments gutter) instead of a measured page, in React and Vue alike, so it stays a gap after the page whenever the outline takes or gives back its track. `computePanelLayout` returns the column start as `commentsColumnStart`, placed by the shared `commentsColumnStart` rule. The React column also watches the editor's scroll container once that container mounts.
