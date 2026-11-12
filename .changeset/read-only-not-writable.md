---
"@stll/docx-core": patch
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

A read-only model value no longer passes where the writable type is expected. Every record type of the model carries an optional `WritableRecord` field (exported from `@stll/docx-core/model` with `RecordField`, a record's own field names without it) whose `DeepReadonly` form TypeScript refuses where the writable form is expected, so handing a `ReadonlyDocument`, a style result or any part of them to a function that takes the writable type fails to compile. A writable value still passes where a read-only one is expected, and object literals need nothing extra. Readers that only read now take read-only values: the color, border, shading and theme readers (`resolveColor`, `resolveShadingColor`, `resolveColorToHex`, `generateThemeTintShadeMatrix`, `resolveThemeFont`, `resolveThemeFontRef`, `getMajorFont`, `getMinorFont`, `getThemeColor`, `getThemeFonts`, `statesNoBorder`, `cssBorderStyle`, `knownThemeColor`, `themeColorSlot`), every `theme` parameter and prop in core, React and Vue, the table context's `cellBorderColor`, the footnote properties dialog's inputs, the paragraph attribute `spacingExplicit` and the table attribute `_resolvedIndent`. `ExhaustiveFields` classifies a record's own fields only. Saving a paragraph copies its `spacingExplicit`, and a new document copies the theme it is given with `writableCopy`.
