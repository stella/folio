/**
 * @stll/folio-vue dialogs
 *
 * Modal dialog components — find/replace, page setup, table/image properties,
 * image position, footnote/endnote properties, insert table, insert symbol, and
 * hyperlinks. Mirrors the React dialogs subpath.
 *
 * @example
 * ```ts
 * import { HyperlinkDialog, FindReplaceDialog } from '@stll/folio-vue/dialogs';
 * ```
 *
 * @packageDocumentation
 * @public
 */

export { default as FindReplaceDialog } from './FindReplaceDialog.vue';
export { default as FootnotePropertiesDialog } from './FootnotePropertiesDialog.vue';
export { default as HyperlinkDialog } from './HyperlinkDialog.vue';
export { default as ImagePositionDialog } from './ImagePositionDialog.vue';
export { default as ImagePropertiesDialog } from './ImagePropertiesDialog.vue';
export { default as InsertSymbolDialog } from './InsertSymbolDialog.vue';
export { default as InsertTableDialog } from './InsertTableDialog.vue';
export { default as PageSetupDialog } from './PageSetupDialog.vue';
export { default as TablePropertiesDialog } from './TablePropertiesDialog.vue';
