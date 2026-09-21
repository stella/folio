/** Reserved-value decisions for run, paragraph, and table formatting. */

import type { ExhaustiveFields } from "../../packages/docx-core/src/model/exhaustiveFields";
import type {
  CellMargins,
  ConditionalFormatStyle,
  FloatingTableProperties,
  ParagraphFormatting,
  SpacingExplicit,
  TableBorders,
  TableCellBorders,
  TableCellFormatting,
  TableFormatting,
  TableGridChange,
  TableLook,
  TableMeasurement,
  TableRowFormatting,
  TabStop,
  TextFormatting,
} from "../../packages/docx-core/src/model/formatting";
import {
  NO_RESERVED_VALUE,
  notModelled,
  readerOwned,
  type ReservedValueDisposition,
  toggle,
  type UnionFields,
  unrepresentable,
} from "./disposition.ts";
import { RESERVED_VALUE_READERS } from "./readers.ts";

// ============================================================================
// RUN PROPERTIES (w:rPr)
// ============================================================================

export const TEXT_FORMATTING_RESERVED = {
  bold: toggle("w:b@val"),
  boldCs: toggle("w:bCs@val"),
  italic: toggle("w:i@val"),
  italicCs: toggle("w:iCs@val"),
  underline: NO_RESERVED_VALUE,
  strike: toggle("w:strike@val"),
  doubleStrike: toggle("w:dstrike@val"),
  vertAlign: readerOwned({
    slot: "w:vertAlign@val",
    sentinel: "baseline",
    reader: RESERVED_VALUE_READERS.runProperties,
  }),
  smallCaps: toggle("w:smallCaps@val"),
  allCaps: toggle("w:caps@val"),
  hidden: toggle("w:vanish@val"),
  noProof: toggle("w:noProof@val"),
  color: NO_RESERVED_VALUE,
  highlight: readerOwned({
    slot: "w:highlight@val",
    sentinel: "none",
    reader: RESERVED_VALUE_READERS.highlight,
  }),
  shading: NO_RESERVED_VALUE,
  fontSize: NO_RESERVED_VALUE,
  fontSizeCs: NO_RESERVED_VALUE,
  fontFamily: NO_RESERVED_VALUE,
  language: NO_RESERVED_VALUE,
  spacing: NO_RESERVED_VALUE,
  position: NO_RESERVED_VALUE,
  scale: NO_RESERVED_VALUE,
  kerning: NO_RESERVED_VALUE,
  effect: readerOwned({
    slot: "w:effect@val",
    sentinel: "none",
    reader: RESERVED_VALUE_READERS.runProperties,
  }),
  emphasisMark: readerOwned({
    slot: "w:em@val",
    sentinel: "none",
    reader: RESERVED_VALUE_READERS.runProperties,
  }),
  emboss: toggle("w:emboss@val"),
  imprint: toggle("w:imprint@val"),
  outline: toggle("w:outline@val"),
  shadow: toggle("w:shadow@val"),
  rtl: toggle("w:rtl@val"),
  cs: toggle("w:cs@val"),
  styleId: readerOwned({
    slot: "w:rStyle@val",
    sentinel: "unresolvable-styleid",
    reader: RESERVED_VALUE_READERS.styleChain,
    evidence: "unknown-styleid-falls-back-to-the-default-style",
  }),
  // Children replayed as the source wrote them; nothing interprets a spelling.
  preserved: NO_RESERVED_VALUE,
} satisfies Record<keyof TextFormatting, ReservedValueDisposition>;

export type ExhaustiveTextFormattingReserved = ExhaustiveFields<
  TextFormatting,
  keyof typeof TEXT_FORMATTING_RESERVED
>;

type Underline = NonNullable<TextFormatting["underline"]>;

export const UNDERLINE_RESERVED = {
  style: readerOwned({
    slot: "w:u@val",
    sentinel: "none",
    reader: RESERVED_VALUE_READERS.underline,
  }),
  color: NO_RESERVED_VALUE,
} satisfies Record<keyof Underline, ReservedValueDisposition>;

export type ExhaustiveUnderlineReserved = ExhaustiveFields<
  Underline,
  keyof typeof UNDERLINE_RESERVED
>;

type FontFamily = NonNullable<TextFormatting["fontFamily"]>;

export const FONT_FAMILY_RESERVED = {
  ascii: NO_RESERVED_VALUE,
  hAnsi: NO_RESERVED_VALUE,
  eastAsia: NO_RESERVED_VALUE,
  cs: NO_RESERVED_VALUE,
  hint: notModelled({
    slot: "w:rFonts@hint",
    sentinel: "default",
    reason:
      "`default` is a named ST_Hint member (use the run's own script), not an unset marker. folio narrows the attribute and round-trips it, but font selection never consults it, so all three members resolve the same glyph font.",
  }),
  asciiTheme: NO_RESERVED_VALUE,
  hAnsiTheme: NO_RESERVED_VALUE,
  eastAsiaTheme: NO_RESERVED_VALUE,
  csTheme: NO_RESERVED_VALUE,
} satisfies Record<keyof FontFamily, ReservedValueDisposition>;

export type ExhaustiveFontFamilyReserved = ExhaustiveFields<
  FontFamily,
  keyof typeof FONT_FAMILY_RESERVED
>;

type RunLanguage = NonNullable<TextFormatting["language"]>;

export const RUN_LANGUAGE_RESERVED = {
  val: NO_RESERVED_VALUE,
  eastAsia: NO_RESERVED_VALUE,
  bidi: NO_RESERVED_VALUE,
} satisfies Record<keyof RunLanguage, ReservedValueDisposition>;

export type ExhaustiveRunLanguageReserved = ExhaustiveFields<
  RunLanguage,
  keyof typeof RUN_LANGUAGE_RESERVED
>;

// ============================================================================
// PARAGRAPH PROPERTIES (w:pPr)
// ============================================================================

export const PARAGRAPH_FORMATTING_RESERVED = {
  alignment: NO_RESERVED_VALUE,
  bidi: toggle("w:bidi@val"),
  kinsoku: toggle("w:kinsoku@val"),
  overflowPunctuation: toggle("w:overflowPunct@val"),
  spaceBefore: NO_RESERVED_VALUE,
  spaceAfter: NO_RESERVED_VALUE,
  lineSpacing: readerOwned({
    slot: "w:spacing@line",
    sentinel: "240",
    reader: RESERVED_VALUE_READERS.lineSpacing,
    evidence: "spacing-line-unit-depends-on-linerule",
  }),
  lineSpacingRule: readerOwned({
    slot: "w:spacing@lineRule",
    sentinel: "auto",
    reader: RESERVED_VALUE_READERS.lineSpacing,
    evidence: "spacing-line-unit-depends-on-linerule",
  }),
  snapToGrid: toggle("w:snapToGrid@val"),
  beforeAutospacing: toggle("w:spacing@beforeAutospacing"),
  afterAutospacing: toggle("w:spacing@afterAutospacing"),
  spacingExplicit: NO_RESERVED_VALUE,
  indentLeft: NO_RESERVED_VALUE,
  indentRight: NO_RESERVED_VALUE,
  indentFirstLine: readerOwned({
    slot: "w:ind@firstLine|w:ind@hanging",
    sentinel: "both-present",
    reader: RESERVED_VALUE_READERS.paragraphProperties,
    evidence: "ind-hanging-wins-over-firstline",
  }),
  hangingIndent: readerOwned({
    slot: "w:ind@hanging",
    sentinel: "both-present",
    reader: RESERVED_VALUE_READERS.paragraphProperties,
    evidence: "ind-hanging-wins-over-firstline",
  }),
  borders: NO_RESERVED_VALUE,
  shading: NO_RESERVED_VALUE,
  tabs: NO_RESERVED_VALUE,
  keepNext: toggle("w:keepNext@val"),
  keepLines: toggle("w:keepLines@val"),
  widowControl: toggle("w:widowControl@val"),
  pageBreakBefore: toggle("w:pageBreakBefore@val"),
  contextualSpacing: toggle("w:contextualSpacing@val"),
  numPr: NO_RESERVED_VALUE,
  numPrFromStyle: NO_RESERVED_VALUE,
  // Was reader-owned, with `headingCollector` named as the reader although
  // every piece of sentinel logic lived in `docx/builtInStyles.ts`. The union
  // removed the question: a `w:val="9"` parses to `OutlineLevel`'s body-text
  // arm and a heading level is one of nine literal types, so no field holds
  // the number 9 and no consumer can read it as a tenth level.
  outlineLevel: unrepresentable({
    slot: "w:outlineLvl@val",
    sentinel: "9",
    carrier: "OutlineLevel",
    evidence: "outlinelvl-nine-is-body-text",
  }),
  styleId: readerOwned({
    slot: "w:pStyle@val",
    sentinel: "unresolvable-styleid",
    reader: RESERVED_VALUE_READERS.styleChain,
    evidence: "unknown-styleid-falls-back-to-the-default-style",
  }),
  frame: NO_RESERVED_VALUE,
  suppressLineNumbers: toggle("w:suppressLineNumbers@val"),
  suppressAutoHyphens: toggle("w:suppressAutoHyphens@val"),
  runProperties: NO_RESERVED_VALUE,
  runInWithNext: toggle("w:specVanish@val"),
  numberingChangeXml: NO_RESERVED_VALUE,
  numberingInsertionXml: NO_RESERVED_VALUE,
  // Children replayed as the source wrote them; nothing interprets a spelling.
  preserved: NO_RESERVED_VALUE,
  indentPreservedAttributes: NO_RESERVED_VALUE,
  spacingPreservedAttributes: NO_RESERVED_VALUE,
} satisfies Record<keyof ParagraphFormatting, ReservedValueDisposition>;

export type ExhaustiveParagraphFormattingReserved = ExhaustiveFields<
  ParagraphFormatting,
  keyof typeof PARAGRAPH_FORMATTING_RESERVED
>;

type ParagraphNumbering = NonNullable<ParagraphFormatting["numPr"]>;
type StyleParagraphNumbering = NonNullable<ParagraphFormatting["numPrFromStyle"]>;

/**
 * `w:numPr` on a paragraph and on its style state the same thing, so one
 * decision covers `numPr` and `numPrFromStyle`, and the map is keyed over both.
 *
 * The keys are every arm's, not the arms' intersection: `keyof` over a union
 * would be `"kind"` alone and let the two payload fields through with no
 * decision at all. That is also why there is no `ExhaustiveFields` alias here,
 * as there is none on the other union-keyed maps: over a union the alias
 * compares against the shared keys and would hold whatever it was given.
 */
export const PARAGRAPH_NUMBERING_RESERVED = {
  kind: NO_RESERVED_VALUE,
  // Was reader-owned, with `isNumberingReference` carrying a sentinel the model
  // still held. `ParagraphNumberingOverride` maps `w:numId w:val="0"` onto its
  // `none` arm at the parse boundary, so the only field that holds a `numId`
  // is `reference`, which the sentinel cannot reach.
  numId: unrepresentable({
    slot: "w:numId@val",
    sentinel: "0",
    carrier: "ParagraphNumberingOverride",
    evidence: "numid-zero-is-no-numbering",
  }),
  // Stays reader-owned: nine levels is a prose limit on an unfacetted
  // `ST_DecimalNumber`, so no arm of the union makes an out-of-range level
  // unrepresentable and one reader still has to drop it.
  ilvl: readerOwned({
    slot: "w:ilvl@val",
    sentinel: "9",
    reader: RESERVED_VALUE_READERS.paragraphProperties,
    evidence: "ilvl-outside-zero-to-eight-names-no-level",
  }),
} satisfies Record<
  UnionFields<ParagraphNumbering | StyleParagraphNumbering>,
  ReservedValueDisposition
>;

export const SPACING_EXPLICIT_RESERVED = {
  before: NO_RESERVED_VALUE,
  after: NO_RESERVED_VALUE,
} satisfies Record<keyof SpacingExplicit, ReservedValueDisposition>;

export type ExhaustiveSpacingExplicitReserved = ExhaustiveFields<
  SpacingExplicit,
  keyof typeof SPACING_EXPLICIT_RESERVED
>;

type ParagraphFrame = NonNullable<ParagraphFormatting["frame"]>;

export const PARAGRAPH_FRAME_RESERVED = {
  dropCap: readerOwned({
    slot: "w:framePr@dropCap",
    sentinel: "none",
    reader: RESERVED_VALUE_READERS.paragraphProperties,
  }),
  lines: NO_RESERVED_VALUE,
  width: NO_RESERVED_VALUE,
  height: NO_RESERVED_VALUE,
  hSpace: NO_RESERVED_VALUE,
  vSpace: NO_RESERVED_VALUE,
  hAnchor: NO_RESERVED_VALUE,
  vAnchor: NO_RESERVED_VALUE,
  x: NO_RESERVED_VALUE,
  y: NO_RESERVED_VALUE,
  xAlign: NO_RESERVED_VALUE,
  yAlign: NO_RESERVED_VALUE,
  wrap: readerOwned({
    slot: "w:framePr@wrap",
    sentinel: "auto|none",
    reader: RESERVED_VALUE_READERS.paragraphProperties,
  }),
  preservedAttributes: NO_RESERVED_VALUE,
} satisfies Record<keyof ParagraphFrame, ReservedValueDisposition>;

export type ExhaustiveParagraphFrameReserved = ExhaustiveFields<
  ParagraphFrame,
  keyof typeof PARAGRAPH_FRAME_RESERVED
>;

export const TAB_STOP_RESERVED = {
  position: NO_RESERVED_VALUE,
  alignment: readerOwned({
    slot: "w:tab@val",
    sentinel: "clear|num|bar",
    reader: RESERVED_VALUE_READERS.tabStops,
  }),
  leader: readerOwned({
    slot: "w:tab@leader",
    sentinel: "none",
    reader: RESERVED_VALUE_READERS.tabStops,
  }),
  preservedAttributes: NO_RESERVED_VALUE,
} satisfies Record<keyof TabStop, ReservedValueDisposition>;

export type ExhaustiveTabStopReserved = ExhaustiveFields<TabStop, keyof typeof TAB_STOP_RESERVED>;

// ============================================================================
// TABLE PROPERTIES (w:tblPr, w:trPr, w:tcPr)
// ============================================================================

export const TABLE_MEASUREMENT_RESERVED = {
  value: readerOwned({
    slot: "w:tblW@w|w:tcW@w|w:tblCellSpacing@w|w:tblInd@w|w:wBefore@w|w:wAfter@w",
    sentinel: "meaningless-under-auto",
    reader: RESERVED_VALUE_READERS.tableWidth,
    evidence: "table-width-auto-makes-the-number-meaningless",
  }),
  type: readerOwned({
    slot: "w:tblW@type|w:tcW@type|w:tblCellSpacing@type|w:tblInd@type|w:wBefore@type|w:wAfter@type",
    sentinel: "auto|nil",
    reader: RESERVED_VALUE_READERS.tableWidth,
    evidence: "table-width-auto-makes-the-number-meaningless",
  }),
} satisfies Record<keyof TableMeasurement, ReservedValueDisposition>;

export type ExhaustiveTableMeasurementReserved = ExhaustiveFields<
  TableMeasurement,
  keyof typeof TABLE_MEASUREMENT_RESERVED
>;

/** Every side is a `CT_Border`; `BORDER_SPEC_RESERVED` owns its slots. */
export const TABLE_BORDERS_RESERVED = {
  top: NO_RESERVED_VALUE,
  bottom: NO_RESERVED_VALUE,
  left: NO_RESERVED_VALUE,
  right: NO_RESERVED_VALUE,
  insideH: NO_RESERVED_VALUE,
  insideV: NO_RESERVED_VALUE,
} satisfies Record<keyof TableBorders, ReservedValueDisposition>;

export type ExhaustiveTableBordersReserved = ExhaustiveFields<
  TableBorders,
  keyof typeof TABLE_BORDERS_RESERVED
>;

export const TABLE_CELL_BORDERS_RESERVED = {
  ...TABLE_BORDERS_RESERVED,
  topLeftToBottomRight: NO_RESERVED_VALUE,
  topRightToBottomLeft: NO_RESERVED_VALUE,
} satisfies Record<keyof TableCellBorders, ReservedValueDisposition>;

export type ExhaustiveTableCellBordersReserved = ExhaustiveFields<
  TableCellBorders,
  keyof typeof TABLE_CELL_BORDERS_RESERVED
>;

export const CELL_MARGINS_RESERVED = {
  top: NO_RESERVED_VALUE,
  bottom: NO_RESERVED_VALUE,
  left: NO_RESERVED_VALUE,
  right: NO_RESERVED_VALUE,
} satisfies Record<keyof CellMargins, ReservedValueDisposition>;

export type ExhaustiveCellMarginsReserved = ExhaustiveFields<
  CellMargins,
  keyof typeof CELL_MARGINS_RESERVED
>;

/**
 * A `w:tblLook` flag is `ST_OnOff`, and folio now models all three of its
 * states. An explicit `0|false|off` is not absence: absence falls back to the
 * matching bit of `w:val`, an explicit off overrides that bit, and the two
 * select different conditional formats out of the table style. `w:val` itself
 * carries the mirror-image reserved value — a bit whose flag was stated means
 * nothing — so both sides are read through the one resolver.
 */
const tableLookFlag = (attribute: string): ReservedValueDisposition =>
  readerOwned({
    slot: `w:tblLook@${attribute}`,
    sentinel: "0|false|off",
    reader: RESERVED_VALUE_READERS.tableLook,
  });

export const TABLE_LOOK_RESERVED = {
  val: readerOwned({
    slot: "w:tblLook@val",
    sentinel: "superseded-by-flag",
    reader: RESERVED_VALUE_READERS.tableLook,
  }),
  firstColumn: tableLookFlag("firstColumn"),
  firstRow: tableLookFlag("firstRow"),
  lastColumn: tableLookFlag("lastColumn"),
  lastRow: tableLookFlag("lastRow"),
  noHBand: tableLookFlag("noHBand"),
  noVBand: tableLookFlag("noVBand"),
} satisfies Record<keyof TableLook, ReservedValueDisposition>;

export type ExhaustiveTableLookReserved = ExhaustiveFields<
  TableLook,
  keyof typeof TABLE_LOOK_RESERVED
>;

export const FLOATING_TABLE_PROPERTIES_RESERVED = {
  horzAnchor: NO_RESERVED_VALUE,
  vertAnchor: NO_RESERVED_VALUE,
  tblpX: NO_RESERVED_VALUE,
  tblpXSpec: NO_RESERVED_VALUE,
  tblpY: NO_RESERVED_VALUE,
  tblpYSpec: NO_RESERVED_VALUE,
  topFromText: NO_RESERVED_VALUE,
  bottomFromText: NO_RESERVED_VALUE,
  leftFromText: NO_RESERVED_VALUE,
  rightFromText: NO_RESERVED_VALUE,
} satisfies Record<keyof FloatingTableProperties, ReservedValueDisposition>;

export type ExhaustiveFloatingTablePropertiesReserved = ExhaustiveFields<
  FloatingTableProperties,
  keyof typeof FLOATING_TABLE_PROPERTIES_RESERVED
>;

export const TABLE_GRID_CHANGE_RESERVED = {
  id: NO_RESERVED_VALUE,
  columnWidths: NO_RESERVED_VALUE,
} satisfies Record<keyof TableGridChange, ReservedValueDisposition>;

export type ExhaustiveTableGridChangeReserved = ExhaustiveFields<
  TableGridChange,
  keyof typeof TABLE_GRID_CHANGE_RESERVED
>;

export const TABLE_FORMATTING_RESERVED = {
  width: NO_RESERVED_VALUE,
  justification: NO_RESERVED_VALUE,
  cellSpacing: NO_RESERVED_VALUE,
  indent: NO_RESERVED_VALUE,
  borders: NO_RESERVED_VALUE,
  cellMargins: NO_RESERVED_VALUE,
  layout: readerOwned({
    slot: "w:tblLayout@type",
    sentinel: "autofit",
    reader: RESERVED_VALUE_READERS.tableProperties,
    evidence: "table-width-auto-makes-the-number-meaningless",
  }),
  styleId: readerOwned({
    slot: "w:tblStyle@val",
    sentinel: "unresolvable-styleid",
    reader: RESERVED_VALUE_READERS.styleChain,
    evidence: "unknown-styleid-falls-back-to-the-default-style",
  }),
  look: NO_RESERVED_VALUE,
  shading: NO_RESERVED_VALUE,
  overlap: notModelled({
    slot: "w:tblOverlap@val",
    sentinel: "never",
    reason:
      "`never` keeps a floating table from sharing a band with another one. folio parses and round-trips the value, but the paginator places floating tables in document order and nothing reads it, so `never` and `overlap` lay out identically.",
  }),
  floating: NO_RESERVED_VALUE,
  bidi: toggle("w:bidiVisual@val"),
  caption: NO_RESERVED_VALUE,
  description: NO_RESERVED_VALUE,
  rowBandSize: NO_RESERVED_VALUE,
  columnBandSize: NO_RESERVED_VALUE,
  // Children replayed as the source wrote them; nothing interprets a spelling.
  preserved: NO_RESERVED_VALUE,
  gridSourceXml: NO_RESERVED_VALUE,
  gridChange: NO_RESERVED_VALUE,
  sourceXml: NO_RESERVED_VALUE,
} satisfies Record<keyof TableFormatting, ReservedValueDisposition>;

export type ExhaustiveTableFormattingReserved = ExhaustiveFields<
  TableFormatting,
  keyof typeof TABLE_FORMATTING_RESERVED
>;

export const TABLE_ROW_FORMATTING_RESERVED = {
  gridBefore: NO_RESERVED_VALUE,
  widthBefore: NO_RESERVED_VALUE,
  gridAfter: NO_RESERVED_VALUE,
  widthAfter: NO_RESERVED_VALUE,
  height: NO_RESERVED_VALUE,
  heightRule: readerOwned({
    slot: "w:trHeight@hRule",
    sentinel: "auto",
    reader: RESERVED_VALUE_READERS.tableRowProperties,
  }),
  header: toggle("w:tblHeader@val"),
  cantSplit: toggle("w:cantSplit@val"),
  justification: NO_RESERVED_VALUE,
  hidden: toggle("w:hidden@val"),
  conditionalFormat: NO_RESERVED_VALUE,
  preserved: NO_RESERVED_VALUE,
  sourceXml: NO_RESERVED_VALUE,
} satisfies Record<keyof TableRowFormatting, ReservedValueDisposition>;

export type ExhaustiveTableRowFormattingReserved = ExhaustiveFields<
  TableRowFormatting,
  keyof typeof TABLE_ROW_FORMATTING_RESERVED
>;

/** `w:cnfStyle` is a 12-bit pattern; every bit means itself. */
export const CONDITIONAL_FORMAT_STYLE_RESERVED = {
  firstRow: NO_RESERVED_VALUE,
  lastRow: NO_RESERVED_VALUE,
  firstColumn: NO_RESERVED_VALUE,
  lastColumn: NO_RESERVED_VALUE,
  oddHBand: NO_RESERVED_VALUE,
  evenHBand: NO_RESERVED_VALUE,
  oddVBand: NO_RESERVED_VALUE,
  evenVBand: NO_RESERVED_VALUE,
  nwCell: NO_RESERVED_VALUE,
  neCell: NO_RESERVED_VALUE,
  swCell: NO_RESERVED_VALUE,
  seCell: NO_RESERVED_VALUE,
} satisfies Record<keyof ConditionalFormatStyle, ReservedValueDisposition>;

export type ExhaustiveConditionalFormatStyleReserved = ExhaustiveFields<
  ConditionalFormatStyle,
  keyof typeof CONDITIONAL_FORMAT_STYLE_RESERVED
>;

export const TABLE_CELL_FORMATTING_RESERVED = {
  width: NO_RESERVED_VALUE,
  borders: NO_RESERVED_VALUE,
  margins: NO_RESERVED_VALUE,
  shading: NO_RESERVED_VALUE,
  verticalAlign: NO_RESERVED_VALUE,
  textDirection: NO_RESERVED_VALUE,
  gridSpan: readerOwned({
    slot: "w:gridSpan@val",
    sentinel: "1|0",
    reader: RESERVED_VALUE_READERS.tableCellProperties,
    evidence: "gridspan-one-is-no-span",
  }),
  // `@w:val` is optional with no XSD default, and an omitted one means
  // `continue`, so absence and the explicit token are the same reserved value.
  vMerge: readerOwned({
    slot: "w:vMerge@val",
    sentinel: "absent|continue",
    reader: RESERVED_VALUE_READERS.tableCellProperties,
    evidence: "vmerge-absent-means-continue",
  }),
  fitText: toggle("w:tcFitText@val"),
  noWrap: toggle("w:noWrap@val"),
  hideMark: toggle("w:hideMark@val"),
  conditionalFormat: NO_RESERVED_VALUE,
  preserved: NO_RESERVED_VALUE,
  sourceXml: NO_RESERVED_VALUE,
} satisfies Record<keyof TableCellFormatting, ReservedValueDisposition>;

export type ExhaustiveTableCellFormattingReserved = ExhaustiveFields<
  TableCellFormatting,
  keyof typeof TABLE_CELL_FORMATTING_RESERVED
>;
