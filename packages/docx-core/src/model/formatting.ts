/**
 * Text, Paragraph, and Table Formatting Types
 *
 * Properties that control how text, paragraphs, and table structures
 * are formatted in OOXML (w:rPr, w:pPr, w:tblPr, etc.).
 */

import type { ColorValue, BorderSpec, ShadingProperties } from "./colors";

// ============================================================================
// TEXT FORMATTING (Run Properties - rPr)
// ============================================================================

/**
 * Underline style options
 */
export const UNDERLINE_STYLES = Object.freeze([
  "none",
  "single",
  "words",
  "double",
  "thick",
  "dotted",
  "dottedHeavy",
  "dash",
  "dashedHeavy",
  "dashLong",
  "dashLongHeavy",
  "dotDash",
  "dashDotHeavy",
  "dotDotDash",
  "dashDotDotHeavy",
  "wave",
  "wavyHeavy",
  "wavyDouble",
] as const);

export type UnderlineStyle = (typeof UNDERLINE_STYLES)[number];

/**
 * Text effect animations
 */
export const TEXT_EFFECTS = Object.freeze([
  "none",
  "blinkBackground",
  "lights",
  "antsBlack",
  "antsRed",
  "shimmer",
  "sparkle",
] as const);

export type TextEffect = (typeof TEXT_EFFECTS)[number];

/**
 * Emphasis mark type
 */
export const EMPHASIS_MARKS = Object.freeze([
  "none",
  "dot",
  "comma",
  "circle",
  "underDot",
] as const);

export type EmphasisMark = (typeof EMPHASIS_MARKS)[number];

export const VERTICAL_ALIGNMENTS = Object.freeze(["baseline", "superscript", "subscript"] as const);

export const HIGHLIGHT_COLORS = Object.freeze([
  "black",
  "blue",
  "cyan",
  "darkBlue",
  "darkCyan",
  "darkGray",
  "darkGreen",
  "darkMagenta",
  "darkRed",
  "darkYellow",
  "green",
  "lightGray",
  "magenta",
  "none",
  "red",
  "white",
  "yellow",
] as const);

export const FONT_HINTS = Object.freeze(["default", "eastAsia", "cs"] as const);

export const ASCII_THEME_FONTS = Object.freeze([
  "majorAscii",
  "majorHAnsi",
  "majorEastAsia",
  "majorBidi",
  "minorAscii",
  "minorHAnsi",
  "minorEastAsia",
  "minorBidi",
] as const);

export const TEXT_FORMATTING_VISUAL_GROUPS = Object.freeze([
  "allCaps",
  "bold",
  "characterSpacing",
  "color",
  "effect",
  "emboss",
  "emphasisMark",
  "fontFamily",
  "fontSize",
  "hidden",
  "highlight",
  "imprint",
  "italic",
  "language",
  "outline",
  "rtl",
  "shading",
  "shadow",
  "smallCaps",
  "strike",
  "underline",
  "vertAlign",
] as const);

export type TextFormattingVisualGroup = (typeof TEXT_FORMATTING_VISUAL_GROUPS)[number];

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
  vertAlign?: (typeof VERTICAL_ALIGNMENTS)[number];

  // Capitalization
  /** Small caps (w:smallCaps) */
  smallCaps?: boolean;
  /** All caps (w:caps) */
  allCaps?: boolean;

  // Visibility
  /** Hidden text (w:vanish) */
  hidden?: boolean;

  // Colors and highlighting
  /** Text color (w:color) */
  color?: ColorValue;
  /** Highlight/background color (w:highlight) */
  highlight?: (typeof HIGHLIGHT_COLORS)[number];
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
    hint?: (typeof FONT_HINTS)[number];
    /** Theme font reference */
    asciiTheme?: (typeof ASCII_THEME_FONTS)[number];
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
};

type SelfDescribingFieldMap<Value, Descriptor> = {
  [Field in keyof Value]-?: Descriptor & { field: Field };
};

type TextFormattingPropertyDescriptor = {
  comparison: "exact";
  validation:
    | "boolean"
    | "color"
    | "effect"
    | "emphasis"
    | "finite-number"
    | "font-family"
    | "highlight"
    | "language"
    | "nonnegative-number"
    | "shading"
    | "string"
    | "underline"
    | "vertical-alignment";
  visualGroup: TextFormattingVisualGroup | null;
  fastPath: "character-style" | "structural" | "visual";
};

/** Total semantic descriptor for every modeled run-formatting property. */
export const TEXT_FORMATTING_PROPERTY_DESCRIPTORS = {
  bold: {
    field: "bold",
    comparison: "exact",
    validation: "boolean",
    visualGroup: "bold",
    fastPath: "visual",
  },
  boldCs: {
    field: "boldCs",
    comparison: "exact",
    validation: "boolean",
    visualGroup: null,
    fastPath: "structural",
  },
  italic: {
    field: "italic",
    comparison: "exact",
    validation: "boolean",
    visualGroup: "italic",
    fastPath: "visual",
  },
  italicCs: {
    field: "italicCs",
    comparison: "exact",
    validation: "boolean",
    visualGroup: null,
    fastPath: "structural",
  },
  underline: {
    field: "underline",
    comparison: "exact",
    validation: "underline",
    visualGroup: "underline",
    fastPath: "visual",
  },
  strike: {
    field: "strike",
    comparison: "exact",
    validation: "boolean",
    visualGroup: "strike",
    fastPath: "visual",
  },
  doubleStrike: {
    field: "doubleStrike",
    comparison: "exact",
    validation: "boolean",
    visualGroup: "strike",
    fastPath: "visual",
  },
  vertAlign: {
    field: "vertAlign",
    comparison: "exact",
    validation: "vertical-alignment",
    visualGroup: "vertAlign",
    fastPath: "visual",
  },
  smallCaps: {
    field: "smallCaps",
    comparison: "exact",
    validation: "boolean",
    visualGroup: "smallCaps",
    fastPath: "visual",
  },
  allCaps: {
    field: "allCaps",
    comparison: "exact",
    validation: "boolean",
    visualGroup: "allCaps",
    fastPath: "visual",
  },
  hidden: {
    field: "hidden",
    comparison: "exact",
    validation: "boolean",
    visualGroup: "hidden",
    fastPath: "visual",
  },
  color: {
    field: "color",
    comparison: "exact",
    validation: "color",
    visualGroup: "color",
    fastPath: "visual",
  },
  highlight: {
    field: "highlight",
    comparison: "exact",
    validation: "highlight",
    visualGroup: "highlight",
    fastPath: "visual",
  },
  shading: {
    field: "shading",
    comparison: "exact",
    validation: "shading",
    visualGroup: "shading",
    fastPath: "visual",
  },
  fontSize: {
    field: "fontSize",
    comparison: "exact",
    validation: "nonnegative-number",
    visualGroup: "fontSize",
    fastPath: "visual",
  },
  fontSizeCs: {
    field: "fontSizeCs",
    comparison: "exact",
    validation: "nonnegative-number",
    visualGroup: null,
    fastPath: "structural",
  },
  fontFamily: {
    field: "fontFamily",
    comparison: "exact",
    validation: "font-family",
    visualGroup: "fontFamily",
    fastPath: "visual",
  },
  language: {
    field: "language",
    comparison: "exact",
    validation: "language",
    visualGroup: "language",
    fastPath: "visual",
  },
  spacing: {
    field: "spacing",
    comparison: "exact",
    validation: "finite-number",
    visualGroup: "characterSpacing",
    fastPath: "visual",
  },
  position: {
    field: "position",
    comparison: "exact",
    validation: "finite-number",
    visualGroup: "characterSpacing",
    fastPath: "visual",
  },
  scale: {
    field: "scale",
    comparison: "exact",
    validation: "finite-number",
    visualGroup: "characterSpacing",
    fastPath: "visual",
  },
  kerning: {
    field: "kerning",
    comparison: "exact",
    validation: "nonnegative-number",
    visualGroup: "characterSpacing",
    fastPath: "visual",
  },
  effect: {
    field: "effect",
    comparison: "exact",
    validation: "effect",
    visualGroup: "effect",
    fastPath: "visual",
  },
  emphasisMark: {
    field: "emphasisMark",
    comparison: "exact",
    validation: "emphasis",
    visualGroup: "emphasisMark",
    fastPath: "visual",
  },
  emboss: {
    field: "emboss",
    comparison: "exact",
    validation: "boolean",
    visualGroup: "emboss",
    fastPath: "visual",
  },
  imprint: {
    field: "imprint",
    comparison: "exact",
    validation: "boolean",
    visualGroup: "imprint",
    fastPath: "visual",
  },
  outline: {
    field: "outline",
    comparison: "exact",
    validation: "boolean",
    visualGroup: "outline",
    fastPath: "visual",
  },
  shadow: {
    field: "shadow",
    comparison: "exact",
    validation: "boolean",
    visualGroup: "shadow",
    fastPath: "visual",
  },
  rtl: {
    field: "rtl",
    comparison: "exact",
    validation: "boolean",
    visualGroup: "rtl",
    fastPath: "visual",
  },
  cs: {
    field: "cs",
    comparison: "exact",
    validation: "boolean",
    visualGroup: null,
    fastPath: "structural",
  },
  styleId: {
    field: "styleId",
    comparison: "exact",
    validation: "string",
    visualGroup: null,
    fastPath: "character-style",
  },
} as const satisfies SelfDescribingFieldMap<TextFormatting, TextFormattingPropertyDescriptor>;

export const TEXT_FORMATTING_FONT_FAMILY_FIELD_DESCRIPTORS = {
  ascii: { field: "ascii", validation: "string" },
  hAnsi: { field: "hAnsi", validation: "string" },
  eastAsia: { field: "eastAsia", validation: "string" },
  cs: { field: "cs", validation: "string" },
  hint: { field: "hint", validation: "font-hint" },
  asciiTheme: { field: "asciiTheme", validation: "ascii-theme" },
  hAnsiTheme: { field: "hAnsiTheme", validation: "string" },
  eastAsiaTheme: { field: "eastAsiaTheme", validation: "string" },
  csTheme: { field: "csTheme", validation: "string" },
} as const satisfies SelfDescribingFieldMap<
  NonNullable<TextFormatting["fontFamily"]>,
  { validation: "ascii-theme" | "font-hint" | "string" }
>;

export const TEXT_FORMATTING_LANGUAGE_FIELD_DESCRIPTORS = {
  val: { field: "val", validation: "string" },
  eastAsia: { field: "eastAsia", validation: "string" },
  bidi: { field: "bidi", validation: "string" },
} as const satisfies SelfDescribingFieldMap<
  NonNullable<TextFormatting["language"]>,
  { validation: "string" }
>;

export const TEXT_FORMATTING_UNDERLINE_FIELD_DESCRIPTORS = {
  style: { field: "style", validation: "underline-style" },
  color: { field: "color", validation: "color" },
} as const satisfies SelfDescribingFieldMap<
  NonNullable<TextFormatting["underline"]>,
  { validation: "color" | "underline-style" }
>;

/** Structural equality for descriptor-owned text-formatting values. */
export const sameTextFormattingValue = (left: unknown, right: unknown): boolean => {
  if (left === right) return true;
  if (typeof left !== "object" || left === null || typeof right !== "object" || right === null) {
    return false;
  }
  if (Array.isArray(left) || Array.isArray(right)) return false;
  const leftEntries = Object.entries(left);
  const rightEntries = Object.entries(right);
  if (leftEntries.length !== rightEntries.length) return false;
  return leftEntries.every(
    ([field, value]) =>
      Object.prototype.hasOwnProperty.call(right, field) &&
      sameTextFormattingValue(value, Reflect.get(right, field)),
  );
};

/** Whether two complete modeled run-property records are structurally equal. */
export const sameTextFormatting = (
  left: TextFormatting | undefined,
  right: TextFormatting | undefined,
): boolean =>
  Object.values(TEXT_FORMATTING_PROPERTY_DESCRIPTORS).every(({ field }) =>
    sameTextFormattingValue(left?.[field], right?.[field]),
  );

// ============================================================================
// PARAGRAPH FORMATTING (Paragraph Properties - pPr)
// ============================================================================

/**
 * Tab stop alignment
 */
export type TabStopAlignment = "left" | "center" | "right" | "decimal" | "bar" | "clear" | "num";

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

/**
 * Paragraph alignment/justification
 */
export type ParagraphAlignment =
  | "left"
  | "center"
  | "right"
  | "both"
  | "distribute"
  | "mediumKashida"
  | "highKashida"
  | "lowKashida"
  | "thaiDistribute";

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
  /** Numbering properties (w:numPr) */
  numPr?: {
    /** Numbering definition ID (w:numId) */
    numId?: number;
    /** List level (0-8) (w:ilvl) */
    ilvl?: number;
  };
  /**
   * When `numPr` was resolved from the paragraph STYLE's pPr rather than the
   * paragraph's own `<w:numPr>`, this records the style-sourced value. The
   * serializer omits `numPr` while it still equals this value — writing it as
   * direct formatting would flip Word's indent precedence (a directly
   * referenced level's indents beat the style's; a style-referenced level's
   * do not) and break the document on save/reload. Cleared the moment the
   * user changes the numbering (values diverge).
   */
  numPrFromStyle?: {
    numId?: number;
    ilvl?: number;
  };

  // Outline level (for TOC)
  /** Outline level 0-9 (w:outlineLvl) */
  outlineLevel?: number;

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
 * Table look flags (for table styles)
 */
export type TableLook = {
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
 * Table formatting properties (w:tblPr)
 */
export type TableFormatting = {
  /** Table width */
  width?: TableMeasurement;
  /** Table justification */
  justification?: "left" | "center" | "right";
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
  /**
   * The table's `w:tblGrid`, verbatim.
   *
   * The grid is a sibling of `w:tblPr` rather than a child of it, but it is
   * the table's own property set that travels through the editable model, so
   * it rides along here. Rebuilding the grid from the column widths alone
   * drops `w:tblGridChange` — the tracked record of a grid a reviewer resized
   * — which nothing else in the model carries.
   */
  gridSourceXml?: string;
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
  /** Row justification */
  justification?: "left" | "center" | "right";
  /** Hidden row */
  hidden?: boolean;
  /** Conditional format style */
  conditionalFormat?: ConditionalFormatStyle;
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
  /** Text direction */
  textDirection?: "lr" | "lrV" | "rl" | "rlV" | "tb" | "tbV" | "tbRl" | "tbRlV" | "btLr";
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
