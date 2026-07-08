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

export const FindReplaceDialog = FindReplaceDialogComponent;
export const FootnotePropertiesDialog = FootnotePropertiesDialogComponent;
export const HyperlinkDialog = HyperlinkDialogComponent;
export const ImagePositionDialog = ImagePositionDialogComponent;
export const ImagePropertiesDialog = ImagePropertiesDialogComponent;
export const InsertImageDialog = InsertImageDialogComponent;
export const InsertSymbolDialog = InsertSymbolDialogComponent;
export const InsertTableDialog = InsertTableDialogComponent;
export const PageSetupDialog = PageSetupDialogComponent;
export const PasteSpecialDialog = PasteSpecialDialogComponent;
export const SplitCellDialog = SplitCellDialogComponent;
export const TablePropertiesDialog = TablePropertiesDialogComponent;
export const WatermarkDialog = WatermarkDialogComponent;
