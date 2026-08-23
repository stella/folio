/**
 * ProseMirror Schema for DOCX Editor
 *
 * Singleton ExtensionManager that builds the schema and initializes runtime.
 * Legacy code imports `schema` and commands from here; new code should use
 * ExtensionManager directly.
 */

import { ExtensionManager } from "../extensions/ExtensionManager";
import type { FolioCommandName } from "../extensions/types";
// oxlint-disable-next-line import/no-cycle -- singleton registry; StarterKit pulls in extensions that read back the singleton at runtime
import { createStarterKit } from "../extensions/StarterKit";

// Re-export type interfaces (used by toProseDoc, fromProseDoc, and other modules)
export type {
  HardBreakAttrs,
  TabAttrs,
  SymbolAttrs,
  ParagraphAttrs,
  ParagraphPropertyChangeAttrs,
  FieldAttrs,
  ImageAttrs,
  ImagePositionAttrs,
  MathAttrs,
  SdtAttrs,
  BlockSdtAttrs,
  ShapeAttrs,
  TableAttrs,
  TableRowAttrs,
  TableCellAttrs,
  TextBoxAttrs,
} from "./nodes";
export type {
  CharacterStyleAttrs,
  TextColorAttrs,
  UnderlineAttrs,
  StrikeAttrs,
  FontSizeAttrs,
  FontFamilyAttrs,
  LanguageAttrs,
  HighlightAttrs,
  CharacterSpacingAttrs,
  EmphasisMarkAttrs,
  TextEffectAttrs,
  FootnoteRefAttrs,
  CommentAttrs,
  TrackedChangeMarkAttrs,
  RunPropertyChangeMarkAttrs,
  RunFormattingOverrideAttrs,
  RunShadingAttrs,
  HyperlinkAttrs,
} from "./marks";

/**
 * Singleton ExtensionManager — builds schema + initializes runtime (plugins, commands, keymaps)
 */
const mgr = new ExtensionManager(createStarterKit());
mgr.buildSchema();
mgr.initializeRuntime();

const completeCommandNames = <const Names extends readonly FolioCommandName[]>(
  names: Names & (FolioCommandName extends Names[number] ? unknown : never),
): Names => names;

const requiredCommands = completeCommandNames([
  "toggleBold",
  "toggleItalic",
  "toggleUnderline",
  "toggleStrike",
  "toggleSuperscript",
  "toggleSubscript",
  "setTextColor",
  "clearTextColor",
  "setHighlight",
  "clearHighlight",
  "setFontSize",
  "clearFontSize",
  "setFontFamily",
  "clearFontFamily",
  "setUnderlineStyle",
  "setHyperlink",
  "removeHyperlink",
  "insertHyperlink",
  "setAlignment",
  "alignLeft",
  "alignCenter",
  "alignRight",
  "alignJustify",
  "setLineSpacing",
  "singleSpacing",
  "oneAndHalfSpacing",
  "doubleSpacing",
  "increaseIndent",
  "decreaseIndent",
  "setIndentLeft",
  "setIndentRight",
  "setIndentFirstLine",
  "toggleBulletList",
  "toggleNumberedList",
  "increaseListLevel",
  "decreaseListLevel",
  "removeList",
  "setSpaceBefore",
  "setSpaceAfter",
  "applyStyle",
  "clearStyle",
  "insertSectionBreak",
  "removeSectionBreak",
  "addTabStop",
  "removeTabStop",
  "toggleBidi",
  "setRtl",
  "setLtr",
  "setTabs",
  "generateTOC",
  "insertTable",
  "addRowAbove",
  "addRowBelow",
  "deleteRow",
  "addColumnLeft",
  "addColumnRight",
  "deleteColumn",
  "deleteTable",
  "selectTable",
  "selectRow",
  "selectColumn",
  "mergeCells",
  "splitCell",
  "setCellBorder",
  "setTableBorderPreset",
  "setTableBorders",
  "removeTableBorders",
  "setAllTableBorders",
  "setOutsideTableBorders",
  "setInsideTableBorders",
  "setCellVerticalAlign",
  "setCellMargins",
  "setCellTextDirection",
  "toggleNoWrap",
  "setRowHeight",
  "toggleHeaderRow",
  "distributeColumns",
  "autoFitContents",
  "setTableProperties",
  "applyTableStyle",
  "setCellFillColor",
  "setTableBorderColor",
  "setTableBorderWidth",
]);

for (const commandName of requiredCommands) {
  mgr.requireCommand(commandName);
}

export const singletonManager = mgr;
export const schema = mgr.getSchema();

/**
 * Export types for convenience
 */
export type DocxSchema = typeof schema;
export type DocxNode = ReturnType<typeof schema.node>;
export type DocxMark = ReturnType<typeof schema.mark>;
