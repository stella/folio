/**
 * Extension System Type Definitions
 *
 * Tiptap-style extension architecture for ProseMirror.
 * Three extension types:
 * - Extension: plugins, commands, keymaps (no schema)
 * - NodeExtension: adds a node spec to the schema
 * - MarkExtension: adds a mark spec to the schema
 */

import type { Schema, NodeSpec, MarkSpec } from "prosemirror-model";
import type { Plugin as PMPlugin, Command } from "prosemirror-state";
import type { TextColorAttrs } from "../schema";
import type {
  ParagraphAlignment,
  LineSpacingRule,
  TabStop,
  TabLeader,
  TabStopAlignment,
} from "../../types/document";
import type { TablePropertiesCommand } from "../../utils/tableOperations";
import type { ResolvedStyleAttrs } from "./core/ParagraphExtension";
import type { BorderPreset, TableBorderPreset } from "./nodes/TableExtension";

export type TableCellBorderCommandSpec = {
  style: string;
  size?: number;
  color?: { rgb: string };
};

export type TableBorderCommandSpec = {
  style: string;
  size: number;
  color: { rgb: string };
};

export type TableCellMarginsCommand = {
  top?: number;
  bottom?: number;
  left?: number;
  right?: number;
};

export type { TablePropertiesCommand } from "../../utils/tableOperations";

export type TableStyleCommand = {
  styleId: string;
  tableBorders?: Partial<
    Record<"top" | "bottom" | "left" | "right" | "insideH" | "insideV", TableCellBorderCommandSpec>
  >;
  conditionals?: Record<
    string,
    {
      backgroundColor?: string;
      borders?: Partial<
        Record<"top" | "bottom" | "left" | "right", TableCellBorderCommandSpec | null>
      >;
      bold?: boolean;
      color?: string;
    }
  >;
  look?: {
    firstRow?: boolean;
    lastRow?: boolean;
    firstCol?: boolean;
    lastCol?: boolean;
    noHBand?: boolean;
    noVBand?: boolean;
  };
};

// ============================================================================
// PRIORITY
// ============================================================================

export type ExtensionPriority = number;

export const Priority = {
  Highest: 0,
  High: 50,
  Default: 100,
  Low: 150,
  Lowest: 200,
} as const;

// ============================================================================
// CONTEXT & RUNTIME
// ============================================================================

export type ExtensionContext = {
  schema: Schema;
};

export type FolioCommandArguments = {
  toggleBold: [];
  toggleItalic: [];
  toggleUnderline: [];
  toggleStrike: [];
  toggleSuperscript: [];
  toggleSubscript: [];
  setTextColor: [attrs: TextColorAttrs];
  clearTextColor: [];
  setHighlight: [color: string];
  clearHighlight: [];
  setFontSize: [size: number];
  clearFontSize: [];
  setFontFamily: [fontName: string];
  clearFontFamily: [];
  setUnderlineStyle: [style: string, color?: TextColorAttrs];
  setHyperlink: [href: string, tooltip?: string];
  removeHyperlink: [];
  insertHyperlink: [text: string, href: string, tooltip?: string];
  setAlignment: [alignment: ParagraphAlignment];
  alignLeft: [];
  alignCenter: [];
  alignRight: [];
  alignJustify: [];
  setLineSpacing: [value: number, rule?: LineSpacingRule];
  singleSpacing: [];
  oneAndHalfSpacing: [];
  doubleSpacing: [];
  increaseIndent: [amount?: number];
  decreaseIndent: [amount?: number];
  setIndentLeft: [twips: number];
  setIndentRight: [twips: number];
  setIndentFirstLine: [twips: number, hanging?: boolean];
  toggleBulletList: [];
  toggleNumberedList: [];
  increaseListLevel: [];
  decreaseListLevel: [];
  removeList: [];
  setSpaceBefore: [twips: number];
  setSpaceAfter: [twips: number];
  applyStyle: [styleId: string, resolvedAttrs?: ResolvedStyleAttrs];
  clearStyle: [];
  insertSectionBreak: [breakType: "nextPage" | "continuous" | "oddPage" | "evenPage"];
  removeSectionBreak: [];
  addTabStop: [position: number, alignment?: TabStopAlignment, leader?: TabLeader];
  removeTabStop: [position: number];
  toggleBidi: [];
  setRtl: [];
  setLtr: [];
  setTabs: [tabs: TabStop[]];
  generateTOC: [];
  insertTable: [rows: number, cols: number];
  addRowAbove: [];
  addRowBelow: [];
  deleteRow: [];
  addColumnLeft: [];
  addColumnRight: [];
  deleteColumn: [];
  deleteTable: [];
  selectTable: [];
  selectRow: [];
  selectColumn: [];
  mergeCells: [];
  splitCell: [];
  setCellBorder: [
    side: "top" | "bottom" | "left" | "right" | "all",
    spec: TableCellBorderCommandSpec | null,
    clearOthers?: boolean,
  ];
  setTableBorderPreset: [preset: TableBorderPreset];
  setTableBorders: [preset: BorderPreset, borderSpec?: TableBorderCommandSpec];
  removeTableBorders: [];
  setAllTableBorders: [borderSpec?: TableBorderCommandSpec];
  setOutsideTableBorders: [borderSpec?: TableBorderCommandSpec];
  setInsideTableBorders: [borderSpec?: TableBorderCommandSpec];
  setCellVerticalAlign: [align: "top" | "center" | "bottom"];
  setCellMargins: [margins: TableCellMarginsCommand];
  setCellTextDirection: [direction: string | null];
  toggleNoWrap: [];
  setRowHeight: [height: number | null, rule?: "auto" | "atLeast" | "exact"];
  toggleHeaderRow: [];
  distributeColumns: [];
  autoFitContents: [];
  setTableProperties: [props: TablePropertiesCommand];
  applyTableStyle: [styleData: TableStyleCommand];
  setCellFillColor: [color: string | null];
  setTableBorderColor: [color: string];
  setTableBorderWidth: [size: number];
};

export type FolioCommandName = keyof FolioCommandArguments;
export type CommandFactory<Args extends readonly unknown[] = readonly unknown[]> = (
  ...args: Args
) => Command;
export type FolioCommandMap = Partial<{
  [Name in FolioCommandName]: CommandFactory<FolioCommandArguments[Name]>;
}>;

// oxlint-disable-next-line typescript/no-explicit-any -- public extension boundary; strict built-ins are intersected internally
export type CommandMap = Record<string, (...args: any[]) => Command>;
export type ExtensionCommandMap = CommandMap & FolioCommandMap;
export type KeyboardShortcutMap = Record<string, Command>;

export type ExtensionRuntime = {
  commands?: CommandMap;
  keyboardShortcuts?: KeyboardShortcutMap;
  plugins?: PMPlugin[];
};

// ============================================================================
// EXTENSION CONFIGS
// ============================================================================

export type ExtensionConfig = {
  name: string;
  priority: ExtensionPriority;
  options: Record<string, unknown>;
};

export type NodeExtensionConfig = {
  schemaNodeName: string;
  nodeSpec: NodeSpec;
} & ExtensionConfig;

export type MarkExtensionConfig = {
  schemaMarkName: string;
  markSpec: MarkSpec;
} & ExtensionConfig;

// ============================================================================
// EXTENSION INSTANCES
// ============================================================================

export type Extension = {
  type: "extension";
  config: ExtensionConfig;
  onSchemaReady: (ctx: ExtensionContext) => ExtensionRuntime;
};

export type NodeExtension = {
  type: "node";
  config: NodeExtensionConfig;
  onSchemaReady: (ctx: ExtensionContext) => ExtensionRuntime;
};

export type MarkExtension = {
  type: "mark";
  config: MarkExtensionConfig;
  onSchemaReady: (ctx: ExtensionContext) => ExtensionRuntime;
};

export type AnyExtension = Extension | NodeExtension | MarkExtension;

// ============================================================================
// DEFINITION TYPES (used by factory functions)
// ============================================================================

type ExtensionOptions = Record<string, unknown>;

type ExtensionDefinitionWithDefaults<TOptions extends ExtensionOptions> = {
  name: string;
  priority?: ExtensionPriority;
  defaultOptions: TOptions;
  onSchemaReady: (ctx: ExtensionContext, options: TOptions) => ExtensionRuntime;
};

type ExtensionDefinitionWithoutDefaults<TOptions extends ExtensionOptions = ExtensionOptions> = {
  name: string;
  priority?: ExtensionPriority;
  defaultOptions?: undefined;
  onSchemaReady: (ctx: ExtensionContext, options: Partial<TOptions>) => ExtensionRuntime;
};

export type ExtensionDefinition<TOptions extends ExtensionOptions = ExtensionOptions> =
  | ExtensionDefinitionWithDefaults<TOptions>
  | ExtensionDefinitionWithoutDefaults<TOptions>;

type NodeExtensionDefinitionWithDefaults<TOptions extends ExtensionOptions> = {
  name: string;
  priority?: ExtensionPriority;
  defaultOptions: TOptions;
  schemaNodeName: string;
  nodeSpec: NodeSpec | ((options: TOptions) => NodeSpec);
  onSchemaReady?: (ctx: ExtensionContext, options: TOptions) => ExtensionRuntime;
};

type NodeExtensionDefinitionWithoutDefaults<TOptions extends ExtensionOptions = ExtensionOptions> =
  {
    name: string;
    priority?: ExtensionPriority;
    defaultOptions?: undefined;
    schemaNodeName: string;
    nodeSpec: NodeSpec | ((options: Partial<TOptions>) => NodeSpec);
    onSchemaReady?: (ctx: ExtensionContext, options: Partial<TOptions>) => ExtensionRuntime;
  };

export type NodeExtensionDefinition<TOptions extends ExtensionOptions = ExtensionOptions> =
  | NodeExtensionDefinitionWithDefaults<TOptions>
  | NodeExtensionDefinitionWithoutDefaults<TOptions>;

type MarkExtensionDefinitionWithDefaults<TOptions extends ExtensionOptions> = {
  name: string;
  priority?: ExtensionPriority;
  defaultOptions: TOptions;
  schemaMarkName: string;
  markSpec: MarkSpec | ((options: TOptions) => MarkSpec);
  onSchemaReady?: (ctx: ExtensionContext, options: TOptions) => ExtensionRuntime;
};

type MarkExtensionDefinitionWithoutDefaults<TOptions extends ExtensionOptions = ExtensionOptions> =
  {
    name: string;
    priority?: ExtensionPriority;
    defaultOptions?: undefined;
    schemaMarkName: string;
    markSpec: MarkSpec | ((options: Partial<TOptions>) => MarkSpec);
    onSchemaReady?: (ctx: ExtensionContext, options: Partial<TOptions>) => ExtensionRuntime;
  };

export type MarkExtensionDefinition<TOptions extends ExtensionOptions = ExtensionOptions> =
  | MarkExtensionDefinitionWithDefaults<TOptions>
  | MarkExtensionDefinitionWithoutDefaults<TOptions>;
