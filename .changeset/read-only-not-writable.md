---
"@stll/docx-core": minor
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

A read-only model value no longer passes where the writable type is expected. Every record type of the model, inline records and index-signature records included, carries an optional `WritableRecord` field (exported from `@stll/docx-core/model`) whose `DeepReadonly` form TypeScript refuses where the writable form is expected, so handing a `ReadonlyDocument`, a style result or any part of them to a function that takes the writable type fails to compile. `DeepReadonly` hands a `Date` out as `ReadonlyDate`, a date without its setters. A writable value still passes where a read-only one is expected, and object literals need nothing extra.

Breaking for host code that lists a model type's fields: `keyof` of a model record now includes the mark, so `Record<keyof TextFormatting, …>`, `satisfies Record<keyof …>` and `Required<ParagraphFormatting>` stop compiling. Use `RecordField<TextFormatting>` (exported from `@stll/docx-core/model` and `@stll/folio-core/types/document`), a record's own field names without the mark, in place of `keyof`, and `{ [Field in RecordField<ParagraphFormatting>]-?: ParagraphFormatting[Field] }` in place of `Required<…>`. `ExhaustiveFields` classifies a record's own fields only. Host code that passes a read-only value it got from the editor to its own function taking the writable type now fails to compile: take the read-only type when the function only reads, and copy with `writableCopy` or `cloneDocumentWithParagraphPropertySources` when it writes.

Readers that only read now take read-only values: the color, border, shading and theme readers (`resolveColor`, `resolveShadingColor`, `resolveColorToHex`, `generateThemeTintShadeMatrix`, `resolveThemeFont`, `resolveThemeFontRef`, `getMajorFont`, `getMinorFont`, `getThemeColor`, `getThemeFonts`, `statesNoBorder`, `cssBorderStyle`, `knownThemeColor`, `themeColorSlot`), the `theme` parameters and props of core, React and Vue that only read it, the table context's `cellBorderColor`, the footnote properties dialog's inputs, the paragraph attribute `spacingExplicit` and the table attribute `_resolvedIndent`. Saving a paragraph copies its `spacingExplicit`.
