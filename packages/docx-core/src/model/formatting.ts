/**
 * Text, Paragraph, and Table Formatting Types
 *
 * Properties that control how text, paragraphs, and table structures
 * are formatted in OOXML (w:rPr, w:pPr, w:tblPr, etc.).
 */

import type { ColorValue, BorderSpec, ShadingProperties } from "./colors";
import type {
  ParagraphAlignment,
  TableAlignment,
  TabStopAlignment,
  TextDirection,
} from "./ooxmlEnumerations.gen";
import type { PreservedMarkup } from "./preservedMarkup";
import type { OutlineLevel } from "./outlineLevel";
import type { ParagraphNumberingOverride } from "./paragraphNumbering";

// ============================================================================
// TEXT FORMATTING (Run Properties - rPr)
// ============================================================================

/**
 * Underline style options
 */
export type UnderlineStyle =
  | "none"
  | "single"
  | "words"
  | "double"
  | "thick"
  | "dotted"
  | "dottedHeavy"
  | "dash"
  | "dashedHeavy"
  | "dashLong"
  | "dashLongHeavy"
  | "dotDash"
  | "dashDotHeavy"
  | "dotDotDash"
  | "dashDotDotHeavy"
  | "wave"
  | "wavyHeavy"
  | "wavyDouble";

/**
 * Text effect animations
 */
export type TextEffect =
  | "none"
  | "blinkBackground"
  | "lights"
  | "antsBlack"
  | "antsRed"
  | "shimmer"
  | "sparkle";

/**
 * Emphasis mark type
 */
export type EmphasisMark = "none" | "dot" | "comma" | "circle" | "underDot";

/**
 * Complete text formatting properties (w:rPr)
 */
export type TextFormatting = {
  // Basic formatting
  /** Bold (w:b) */
  bold?: boolean;
  /** Bold complex script (w:bCs) */
  boldCs?: boolean;
  /** Italic (w:i) */
  italic?: boolean;
  /** Italic complex script (w:iCs) */
  italicCs?: boolean;

  // Underline & strikethrough
  /** Underline style and color (w:u) */
  underline?: {
    style: UnderlineStyle;
    color?: ColorValue;
  };
  /** Strikethrough (w:strike) */
  strike?: boolean;
  /** Double strikethrough (w:dstrike) */
  doubleStrike?: boolean;

  // Vertical alignment
  /** Superscript/subscript (w:vertAlign) */
  vertAlign?: "baseline" | "superscript" | "subscript";

  // Capitalization
  /** Small caps (w:smallCaps) */
  smallCaps?: boolean;
  /** All caps (w:caps) */
  allCaps?: boolean;

  // Visibility
  /** Hidden text (w:vanish) */
  hidden?: boolean;
  /** Exclude this run from spellchecking (w:noProof). */
  noProof?: boolean;

  // Colors and highlighting
  /** Text color (w:color) */
  color?: ColorValue;
  /** Highlight/background color (w:highlight) */
  highlight?:
    | "black"
    | "blue"
    | "cyan"
    | "darkBlue"
    | "darkCyan"
    | "darkGray"
    | "darkGreen"
    | "darkMagenta"
    | "darkRed"
    | "darkYellow"
    | "green"
    | "lightGray"
    | "magenta"
    | "none"
    | "red"
    | "white"
    | "yellow";
  /** Character shading (w:shd) */
  shading?: ShadingProperties;

  // Font properties
  /** Font size in half-points (w:sz) - e.g., 24 = 12pt */
  fontSize?: number;
  /** Font size complex script (w:szCs) */
  fontSizeCs?: number;
  /** Font family (w:rFonts) */
  fontFamily?: {
    ascii?: string;
    hAnsi?: string;
    eastAsia?: string;
    cs?: string;
    /** Script slot Word should prefer when selecting a glyph font (w:hint). */
    hint?: "default" | "eastAsia" | "cs";
    /** Theme font reference */
    asciiTheme?:
      | "majorAscii"
      | "majorHAnsi"
      | "majorEastAsia"
      | "majorBidi"
      | "minorAscii"
      | "minorHAnsi"
      | "minorEastAsia"
      | "minorBidi";
    hAnsiTheme?: string;
    eastAsiaTheme?: string;
    csTheme?: string;
  };
  /** Run language metadata (`w:lang`) used by spelling and line breaking. */
  language?: {
    val?: string;
    eastAsia?: string;
    bidi?: string;
  };

  // Spacing and position
  /** Character spacing in twips (w:spacing) */
  spacing?: number;
  /** Raised/lowered text position in half-points (w:position) */
  position?: number;
  /** Horizontal text scale percentage (w:w) */
  scale?: number;
  /** Kerning threshold in half-points (w:kern) */
  kerning?: number;

  // Effects
  /** Text effect animation (w:effect) */
  effect?: TextEffect;
  /** Emphasis mark (w:em) */
  emphasisMark?: EmphasisMark;
  /** Emboss effect (w:emboss) */
  emboss?: boolean;
  /** Imprint/engrave effect (w:imprint) */
  imprint?: boolean;
  /** Outline effect (w:outline) */
  outline?: boolean;
  /** Shadow effect (w:shadow) */
  shadow?: boolean;

  // Complex script
  /** Right-to-left text (w:rtl) */
  rtl?: boolean;
  /** Complex script formatting (w:cs) */
  cs?: boolean;

  // Style reference
  /** Character style ID (w:rStyle) */
  styleId?: string;

  /**
   * The `w:rPr` children folio does not model, in the order the schema
   * declares them.
   *
   * The sink records a capture's schema ordinal rather than a count of
   * modelled siblings, because the writer puts every child at the place the
   * declared list gives its name; see `containerChildren.ts`. One field serves
   * all four owners of a run property set — a run, the paragraph mark, and the
   * snapshot inside either one's `w:rPrChange` — because one reader fills it
   * and one writer empties it.
   *
   * It belongs to the element that was parsed, and to no other. A style's
   * captured bytes are not a run's direct formatting, so style resolution and
   * `mergeTextFormatting` drop it rather than inheriting it.
   */
  preserved?: PreservedMarkup;
};

// ============================================================================
// PARAGRAPH FORMATTING (Paragraph Properties - pPr)
// ============================================================================

/** Tab stop alignment (`w:tab/@w:val`), generated from `ST_TabJc`. */
export type { TabStopAlignment };

/**
 * A table's or a row's placement (`w:jc/@w:val`), generated from `ST_JcTable`.
 *
 * `start` and `end` are members in their own right, not spellings of `left`
 * and `right`: they name an edge of the table's own direction, which
 * `w:bidiVisual` sets, and the layout resolves them against it.
 */
export type { TableAlignment };

/**
 * A text flow (`w:textDirection/@w:val`), generated from `ST_TextDirection`.
 *
 * One type for the three places the format declares the element: a table cell,
 * a section and a paragraph. Twelve tokens for six flows; `TEXT_DIRECTION_FLOW_BY_TOKEN`
 * says which flow each token names.
 */
export type { TextDirection };

/**
 * Tab leader character
 */
export type TabLeader = "none" | "dot" | "hyphen" | "underscore" | "heavy" | "middleDot";

/**
 * Tab stop definition
 */
export type TabStop = {
  /** Position in twips from left margin */
  position: number;
  /** Alignment at tab stop */
  alignment: TabStopAlignment;
  /** Leader character */
  leader?: TabLeader;
};

/**
 * Line spacing rule
 */
export type LineSpacingRule = "auto" | "exact" | "atLeast";

/** Paragraph alignment/justification (`w:jc/@w:val`), generated from `ST_Jc`. */
export type { ParagraphAlignment };

/**
 * Complete paragraph formatting properties (w:pPr)
 */
export type SpacingExplicit = { before?: boolean; after?: boolean };

export type ParagraphFormatting = {
  // Alignment
  /** Paragraph alignment (w:jc) */
  alignment?: ParagraphAlignment;
  /** Text direction (w:bidi) */
  bidi?: boolean;
  /** Apply East Asian first/last-character line-breaking rules (w:kinsoku). */
  kinsoku?: boolean;
  /** Allow punctuation to hang beyond the text margin (w:overflowPunct). */
  overflowPunctuation?: boolean;

  // Spacing
  /** Spacing before in twips (w:spacing/@w:before) */
  spaceBefore?: number;
  /** Spacing after in twips (w:spacing/@w:after) */
  spaceAfter?: number;
  /** Line spacing value (w:spacing/@w:line) */
  lineSpacing?: number;
  /** Line spacing rule (w:spacing/@w:lineRule) */
  lineSpacingRule?: LineSpacingRule;
  /** Whether the paragraph participates in the section document grid (w:snapToGrid). */
  snapToGrid?: boolean;
  /** Auto space before (w:spacing/@w:beforeAutospacing) */
  beforeAutospacing?: boolean;
  /** Auto space after (w:spacing/@w:afterAutospacing) */
  afterAutospacing?: boolean;
  /** Which spacing sides came from this paragraph's own pPr. */
  spacingExplicit?: SpacingExplicit;

  // Indentation
  /** Left indent in twips (w:ind/@w:left) */
  indentLeft?: number;
  /** Right indent in twips (w:ind/@w:right) */
  indentRight?: number;
  /** First line indent in twips - positive for indent, negative for hanging (w:ind/@w:firstLine or @w:hanging) */
  indentFirstLine?: number;
  /** Whether first line is hanging indent */
  hangingIndent?: boolean;

  // Borders
  /** Paragraph borders (w:pBdr) */
  borders?: {
    top?: BorderSpec;
    bottom?: BorderSpec;
    left?: BorderSpec;
    right?: BorderSpec;
    between?: BorderSpec;
    bar?: BorderSpec;
  };

  // Background
  /** Paragraph shading (w:shd) */
  shading?: ShadingProperties;

  // Tab stops
  /** Custom tab stops (w:tabs) */
  tabs?: TabStop[];

  // Page break control
  /** Keep with next paragraph (w:keepNext) */
  keepNext?: boolean;
  /** Keep lines together (w:keepLines) */
  keepLines?: boolean;
  /** Widow/orphan control (w:widowControl) */
  widowControl?: boolean;
  /** Page break before (w:pageBreakBefore) */
  pageBreakBefore?: boolean;
  /** Contextual spacing — suppress space between paragraphs of the same style (w:contextualSpacing) */
  contextualSpacing?: boolean;

  // Numbering/List
  /**
   * The stated `w:numPr` (17.3.1.19). Absent means the tier states nothing and
   * inherits whatever the tier below it states; {@link ParagraphNumberingOverride}
   * carries the rest, so neither the reserved `w:numId w:val="0"` nor a level
   * stated without an id has a spelling any consumer has to recognise.
   */
  numPr?: ParagraphNumberingOverride;
  /**
   * When `numPr` was resolved from the paragraph STYLE's pPr rather than the
   * paragraph's own `<w:numPr>`, this records the style-sourced value. The
   * serializer omits `numPr` while it still equals this value — writing it as
   * direct formatting would flip Word's indent precedence (a directly
   * referenced level's indents beat the style's; a style-referenced level's
   * do not) and break the document on save/reload. Cleared the moment the
   * user changes the numbering (values diverge).
   */
  numPrFromStyle?: ParagraphNumberingOverride;
  /**
   * The `w:numberingChange` inside the paragraph's `w:numPr`, verbatim.
   *
   * It records the numbering a paragraph carried before a reviewer changed it,
   * with the author and date of that change. Nothing in the current model
   * derives it, and a rebuilt `w:numPr` that omits it discards a tracked
   * revision without telling anyone. The element is a historical snapshot, so
   * it travels as markup rather than as a parsed shape.
   */
  numberingChangeXml?: string;

  // Outline level (for TOC)
  /**
   * The stated `w:outlineLvl` (17.3.1.20). Absent means the paragraph states
   * none and inherits one; {@link OutlineLevel} carries the rest, so the
   * reserved `w:val="9"` cannot be mistaken for a tenth heading level.
   */
  outlineLevel?: OutlineLevel;

  // Style reference
  /** Paragraph style ID (w:pStyle) */
  styleId?: string;

  // Frame properties
  /** Text frame properties (w:framePr) */
  frame?: {
    dropCap?: "none" | "drop" | "margin";
    lines?: number;
    width?: number;
    height?: number;
    hSpace?: number;
    vSpace?: number;
    hAnchor?: "text" | "margin" | "page";
    vAnchor?: "text" | "margin" | "page";
    x?: number;
    y?: number;
    xAlign?: "left" | "center" | "right" | "inside" | "outside";
    yAlign?: "top" | "center" | "bottom" | "inside" | "outside" | "inline";
    wrap?: "around" | "auto" | "none" | "notBeside" | "through" | "tight";
  };

  // Suppress
  /** Suppress line numbers (w:suppressLineNumbers) */
  suppressLineNumbers?: boolean;
  /** Suppress auto hyphens (w:suppressAutoHyphens) */
  suppressAutoHyphens?: boolean;

  // Default run properties for this paragraph
  /** Run properties to apply to all runs (w:rPr) */
  runProperties?: TextFormatting;

  /**
   * Run-in heading: this paragraph's mark carries `<w:specVanish/>`
   * and the next paragraph should render inline on the same line.
   */
  runInWithNext?: boolean;
};

// ============================================================================
// TABLE FORMATTING (w:tblPr, w:trPr, w:tcPr)
// ============================================================================

/**
 * Table width type
 */
export type TableWidthType = "auto" | "dxa" | "nil" | "pct";

/**
 * Table measurement (width or height)
 */
export type TableMeasurement = {
  /** Value in twips (for dxa) or fifths of a percent (for pct) */
  value: number;
  /** Measurement type */
  type: TableWidthType;
};

/**
 * Table borders
 */
export type TableBorders = {
  top?: BorderSpec;
  bottom?: BorderSpec;
  left?: BorderSpec;
  right?: BorderSpec;
  insideH?: BorderSpec;
  insideV?: BorderSpec;
};

/**
 * Table cell borders, including w:tl2br (top-left to bottom-right) and
 * w:tr2bl (top-right to bottom-left).
 */
export type TableCellBorders = TableBorders & {
  topLeftToBottomRight?: BorderSpec;
  topRightToBottomLeft?: BorderSpec;
};

/**
 * Cell margins
 */
export type CellMargins = {
  top?: TableMeasurement;
  bottom?: TableMeasurement;
  left?: TableMeasurement;
  right?: TableMeasurement;
};

/**
 * `w:tblLook`: which of a table style's conditional formats the table asks for.
 *
 * Every flag is tri-state, because `CT_TblLook` states the same six facts twice
 * and the two spellings are not interchangeable. Absent means the author wrote
 * no such attribute, so the flag falls back to the matching bit of {@link
 * TableLook.val} and then to off; an explicit `false` states the region off and
 * overrides the bit. Writing only the true flags back turns the second and
 * third into the first, which is a different document.
 *
 * Read one with `resolveTableLook`, never on its own.
 */
export type TableLook = {
  /**
   * `w:val`, the legacy `ST_ShortHexNumber` bitmask, kept as the author spelled
   * it. Producers older than the attribute form write only this.
   */
  val?: string;
  firstColumn?: boolean;
  firstRow?: boolean;
  lastColumn?: boolean;
  lastRow?: boolean;
  noHBand?: boolean;
  noVBand?: boolean;
};

/**
 * Floating table properties
 */
export type FloatingTableProperties = {
  /** Horizontal anchor */
  horzAnchor?: "margin" | "page" | "text";
  /** Vertical anchor */
  vertAnchor?: "margin" | "page" | "text";
  /** Horizontal position */
  tblpX?: number;
  tblpXSpec?: "left" | "center" | "right" | "inside" | "outside";
  /** Vertical position */
  tblpY?: number;
  tblpYSpec?: "top" | "center" | "bottom" | "inside" | "outside" | "inline";
  /** Distance from surrounding text */
  topFromText?: number;
  bottomFromText?: number;
  leftFromText?: number;
  rightFromText?: number;
};

/**
 * The grid a reviewer replaced, as `w:tblGridChange` records it.
 *
 * `CT_TblGridChange` holds a `w:tblGrid` of its own, and that grid holds its
 * own `w:gridCol` children: a container nested in one of its own kind. It is a
 * snapshot rather than a description — nothing in the live column widths
 * derives it — so it is modelled as what it is, a count of columns each with
 * the width it stated.
 *
 * `w:w` is optional on a `w:gridCol`, and a column that stated no width is not
 * a column of width zero: `undefined` is what says the snapshot recorded a
 * column and no measure for it.
 */
export type TableGridChange = {
  /** `@w:id`: a physical revision id, re-minted on every save. */
  id: number;
  /** One entry per `w:gridCol`, in source order. */
  columnWidths: readonly (number | undefined)[];
};

/**
 * Table formatting properties (w:tblPr)
 */
export type TableFormatting = {
  /** Table width */
  width?: TableMeasurement;
  /** Table placement (`w:tblPr/w:jc`) */
  justification?: TableAlignment;
  /** Cell spacing */
  cellSpacing?: TableMeasurement;
  /** Table indent from left margin */
  indent?: TableMeasurement;
  /** Table borders */
  borders?: TableBorders;
  /** Default cell margins */
  cellMargins?: CellMargins;
  /** Table layout */
  layout?: "fixed" | "autofit";
  /** Table style ID */
  styleId?: string;
  /** Table look (conditional formatting flags) */
  look?: TableLook;
  /** Shading/background */
  shading?: ShadingProperties;
  /** Overlap for floating tables */
  overlap?: "never" | "overlap";
  /** Floating table properties */
  floating?: FloatingTableProperties;
  /** Right to left table */
  bidi?: boolean;
  /** Accessibility caption (`w:tblCaption`), authored in Word's alt-text dialog. */
  caption?: string;
  /** Long description (`w:tblDescription`), the caption's companion. */
  description?: string;
  /** Rows per band of the table style's row banding (`w:tblStyleRowBandSize`). */
  rowBandSize?: number;
  /** Columns per band (`w:tblStyleColBandSize`). */
  columnBandSize?: number;
  /** The `w:tblPr` children no reader took a typed value from. */
  preserved?: PreservedMarkup;
  /**
   * The table's `w:tblGrid`, verbatim.
   *
   * The grid is a sibling of `w:tblPr` rather than a child of it, but it is
   * the table's own property set that travels through the editable model, so
   * it rides along here. It is replayed only while the column widths still
   * agree with it; {@link gridChange} carries the part of it that a rebuild
   * must not drop.
   */
  gridSourceXml?: string;
  /**
   * The grid's `w:tblGridChange`.
   *
   * It records the grid as it stood before a reviewer resized a column, so it
   * is history: nothing in the current model derives it, and rebuilding the
   * grid from the column widths would accept the revision silently.
   */
  gridChange?: TableGridChange;
  /**
   * The element this formatting was parsed from, verbatim.
   *
   * A typed model covers what the editor understands, and a document carries
   * more than that: conditional-format flags, properties a later revision of
   * the format added, properties a producer wrote that nothing here reads.
   * Rebuilding the element from the model alone drops every one of them, so a
   * save that rewrites an untouched table would change it.
   *
   * The source therefore travels with the parse, and the serializer writes it
   * back unchanged while nothing in the model has moved. It is dropped the
   * moment an edit changes a modelled value, because the two would then
   * disagree and the source is the stale one.
   */
  sourceXml?: string;
};

/** Table properties a row may override in `w:tblPrEx`. */
export type TablePropertyExceptionFormatting = Pick<
  TableFormatting,
  | "width"
  | "justification"
  | "cellSpacing"
  | "indent"
  | "borders"
  | "shading"
  | "layout"
  | "cellMargins"
  | "look"
  | "preserved"
  | "sourceXml"
>;

/**
 * Table row formatting properties (w:trPr)
 */
export type TableRowFormatting = {
  /** Number of table grid columns omitted before the first cell */
  gridBefore?: number;
  /** Preferred width of the omitted leading grid columns */
  widthBefore?: TableMeasurement;
  /** Number of table grid columns omitted after the last cell */
  gridAfter?: number;
  /** Preferred width of the omitted trailing grid columns */
  widthAfter?: TableMeasurement;
  /** Row height */
  height?: TableMeasurement;
  /** Height rule */
  heightRule?: "auto" | "atLeast" | "exact";
  /** Header row (repeats on each page) */
  header?: boolean;
  /** Allow row to break across pages */
  cantSplit?: boolean;
  /** Row placement (`w:trPr/w:jc`) */
  justification?: TableAlignment;
  /** Hidden row */
  hidden?: boolean;
  /** Conditional format style */
  conditionalFormat?: ConditionalFormatStyle;
  /** The `w:trPr` children no reader took a typed value from. */
  preserved?: PreservedMarkup;
  /**
   * The element this formatting was parsed from, verbatim.
   *
   * A typed model covers what the editor understands, and a document carries
   * more than that: conditional-format flags, properties a later revision of
   * the format added, properties a producer wrote that nothing here reads.
   * Rebuilding the element from the model alone drops every one of them, so a
   * save that rewrites an untouched table would change it.
   *
   * The source therefore travels with the parse, and the serializer writes it
   * back unchanged while nothing in the model has moved. It is dropped the
   * moment an edit changes a modelled value, because the two would then
   * disagree and the source is the stale one.
   */
  sourceXml?: string;
};

/**
 * Conditional format style
 */
export type ConditionalFormatStyle = {
  /** First row */
  firstRow?: boolean;
  /** Last row */
  lastRow?: boolean;
  /** First column */
  firstColumn?: boolean;
  /** Last column */
  lastColumn?: boolean;
  /** Odd horizontal band */
  oddHBand?: boolean;
  /** Even horizontal band */
  evenHBand?: boolean;
  /** Odd vertical band */
  oddVBand?: boolean;
  /** Even vertical band */
  evenVBand?: boolean;
  /** Northwest corner */
  nwCell?: boolean;
  /** Northeast corner */
  neCell?: boolean;
  /** Southwest corner */
  swCell?: boolean;
  /** Southeast corner */
  seCell?: boolean;
};

/**
 * Table cell formatting properties (w:tcPr)
 */
export type TableCellFormatting = {
  /** Cell width */
  width?: TableMeasurement;
  /** Cell borders */
  borders?: TableCellBorders;
  /** Cell margins (override table default) */
  margins?: CellMargins;
  /** Cell shading/background */
  shading?: ShadingProperties;
  /** Vertical alignment */
  verticalAlign?: "top" | "center" | "bottom";
  /** Text direction (`w:textDirection`) */
  textDirection?: TextDirection;
  /** Grid span (horizontal merge) */
  gridSpan?: number;
  /** Vertical merge */
  vMerge?: "restart" | "continue";
  /** Fit text to cell width */
  fitText?: boolean;
  /** Wrap text */
  noWrap?: boolean;
  /** Hide cell marker */
  hideMark?: boolean;
  /** Conditional format style */
  conditionalFormat?: ConditionalFormatStyle;
  /** The `w:tcPr` children no reader took a typed value from. */
  preserved?: PreservedMarkup;
  /**
   * The element this formatting was parsed from, verbatim.
   *
   * A typed model covers what the editor understands, and a document carries
   * more than that: conditional-format flags, properties a later revision of
   * the format added, properties a producer wrote that nothing here reads.
   * Rebuilding the element from the model alone drops every one of them, so a
   * save that rewrites an untouched table would change it.
   *
   * The source therefore travels with the parse, and the serializer writes it
   * back unchanged while nothing in the model has moved. It is dropped the
   * moment an edit changes a modelled value, because the two would then
   * disagree and the source is the stale one.
   */
  sourceXml?: string;
};
