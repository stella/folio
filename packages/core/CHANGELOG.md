# @stll/folio-core

## 0.2.0

### Minor Changes

- [#54](https://github.com/stella/folio/pull/54) [`444c11b`](https://github.com/stella/folio/commit/444c11b6c165864bbea59e4ed54c498d1b6fa02e) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Readable template directive rails: gutter rails indent by nesting depth
  within a margin-confined budget, colour by directive kind (loop violet,
  condition teal, theme-aware), render quiet by default and emphasize the
  block containing the caret or under the pointer, split into per-page
  segments so no rail crosses a page header, footer, or gap, and pair
  opener/closer chips on hover with a hint of what a bare closer closes.
  Adds pure helpers `computeBlockDepths` (core) and gutter geometry
  measurement to the paged-layout range projection.

- [#57](https://github.com/stella/folio/pull/57) [`c40cb00`](https://github.com/stella/folio/commit/c40cb00390db3abf109c3764645a87451fb6a249) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add `@stll/folio-vue`, a Vue 3 editor adapter over `@stll/folio-core` that tracks
  the `@stll/folio-react` editor contract (`DocxEditor`, `DocxEditorProps`,
  `DocxEditorRef`, `renderAsync`).

  To share one framework-neutral base across adapters, the folio UI translation
  catalog moves into `@stll/folio-core/i18n/messages` (the React `messages` subpath
  re-exports it, so `@stll/folio-react/messages` is unchanged). Core also gains the
  helpers the adapters build on: `ClipboardManager`, `AutoSaveManager`,
  `resolveColorToHex`, the `WrapType` union, `docx` + `prosemirror/extensions`
  barrels, and a set of ported editor-engine helpers (comment/section-break/table/
  image commands, tracked-change extraction, visual-line navigation, image layout).

### Patch Changes

- [#56](https://github.com/stella/folio/pull/56) [`98022a7`](https://github.com/stella/folio/commit/98022a7873188b42e406b689dd5c4c9b33bb98b8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - fix(core): use a CJK font's line height for CJK text in non-CJK-font runs

- [#56](https://github.com/stella/folio/pull/56) [`98022a7`](https://github.com/stella/folio/commit/98022a7873188b42e406b689dd5c4c9b33bb98b8) Thanks [@jan-kubica](https://github.com/jan-kubica)! - fix(core): derive single-line height from real font hhea metrics

  Single-line height per font is now derived from the font's real `hhea`
  metrics `(ascent + |descent| + lineGap) / unitsPerEm` — the value Word uses —
  instead of hand-transcribed constants, several of which dropped the line gap
  or were otherwise wrong. Corrects 9 fonts (measured against Word): Palatino
  Linotype (was 31% short), Book Antiqua (17%), Cambria (8%), Century Gothic
  (6%), Times New Roman (4%), Arial (3%), Trebuchet MS (2%), Consolas (1%), and
  Lucida Console (14% tall). A shared derivation and a discriminated `hhea` /
  `legacy` representation make it structurally impossible to silently drop a
  term again. CJK fonts are unchanged (their line height is an East-Asian
  layout concern, not a run-font hhea ratio).

- [#55](https://github.com/stella/folio/pull/55) [`e44ba30`](https://github.com/stella/folio/commit/e44ba300c0978b9e2836bafcfe84cd06491e87bf) Thanks [@jan-kubica](https://github.com/jan-kubica)! - fix(core): apply table style paragraph spacing to cell paragraphs

  Table cells now inherit paragraph spacing (space-after, line spacing,
  contextual spacing) from the enclosing table style's `w:pPr` — and from the
  applicable `w:tblStylePr` conditional region (first row, banding, etc.) —
  instead of falling through to `docDefaults`. Per ECMA-376 §17.7.2 this table
  style layer sits between docDefaults and the cell paragraph's own style
  chain/direct formatting, so an explicit paragraph style or direct spacing on
  the paragraph still wins. Previously, cell paragraphs with neither a
  `w:pStyle` nor a direct `w:pPr` picked up the document's default paragraph
  spacing (e.g. Word's default ~10pt space-after and 1.15x line spacing)
  instead of the table style's typically compact spacing (e.g. `TableGrid`'s
  zero space-after / single line), inflating table row heights.

## 0.1.3

### Patch Changes

- Fix DOCX parsing and layout edge cases: smartTag-wrapped runs, EMF header
  previews, percent-suffixed table widths, CJK line breaks, off-page
  header/footer floats, and right-tab trailing width reservation.
