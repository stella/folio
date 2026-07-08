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

import type { Component } from "vue";

import FindReplaceDialogComponent from "./FindReplaceDialog.vue";
import FootnotePropertiesDialogComponent from "./FootnotePropertiesDialog.vue";
import HyperlinkDialogComponent from "./HyperlinkDialog.vue";
import ImagePositionDialogComponent from "./ImagePositionDialog.vue";
import ImagePropertiesDialogComponent from "./ImagePropertiesDialog.vue";
import InsertImageDialogComponent from "./InsertImageDialog.vue";
import InsertSymbolDialogComponent from "./InsertSymbolDialog.vue";
import InsertTableDialogComponent from "./InsertTableDialog.vue";
import PageSetupDialogComponent from "./PageSetupDialog.vue";
import PasteSpecialDialogComponent from "./PasteSpecialDialog.vue";
import SplitCellDialogComponent from "./SplitCellDialog.vue";
import TablePropertiesDialogComponent from "./TablePropertiesDialog.vue";
import WatermarkDialogComponent from "./WatermarkDialog.vue";

export const FindReplaceDialog: Component = FindReplaceDialogComponent;
export const FootnotePropertiesDialog: Component = FootnotePropertiesDialogComponent;
export const HyperlinkDialog: Component = HyperlinkDialogComponent;
export const ImagePositionDialog: Component = ImagePositionDialogComponent;
export const ImagePropertiesDialog: Component = ImagePropertiesDialogComponent;
export const InsertImageDialog: Component = InsertImageDialogComponent;
export const InsertSymbolDialog: Component = InsertSymbolDialogComponent;
export const InsertTableDialog: Component = InsertTableDialogComponent;
export const PageSetupDialog: Component = PageSetupDialogComponent;
export const PasteSpecialDialog: Component = PasteSpecialDialogComponent;
export const SplitCellDialog: Component = SplitCellDialogComponent;
export const TablePropertiesDialog: Component = TablePropertiesDialogComponent;
export const WatermarkDialog: Component = WatermarkDialogComponent;
