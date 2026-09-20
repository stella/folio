/**
 * Document Content Model
 *
 * All content-bearing types: runs, hyperlinks, bookmarks, fields,
 * images, shapes, tables, lists, paragraphs, headers/footers,
 * footnotes/endnotes, and sections.
 *
 * These types form a deeply interrelated tree (Paragraph ↔ Table ↔ ShapeTextBody)
 * and are kept together to avoid circular import issues.
 */

import type { ColorValue, ThemeColorSlot, BorderSpec } from "./colors";
import type {
  TextFormatting,
  ParagraphFormatting,
  TableFormatting,
  TableRowFormatting,
  TableCellFormatting,
} from "./formatting";
import type { NumberFormat, ListRendering } from "./lists";
import type { PreservedAttribute, PreservedMarkup } from "./preservedMarkup";

// ============================================================================
// RUN CONTENT TYPES
// ============================================================================

/**
 * Plain text content
 */
export type TextContent = {
  type: "text";
  /** The text string */
  text: string;
};

/**
 * Tab character
 */
export type PositionalTab = {
  relativeTo?: "margin" | "indent";
  alignment?: "left" | "center" | "right";
  leader?: "none" | "dot" | "hyphen" | "underscore" | "middleDot";
};

export type TabContent = {
  type: "tab";
  /** Word positional-tab properties (`w:ptab`); absent for a regular `w:tab`. */
  positional?: PositionalTab;
};

/**
 * Line break
 */
export type BreakContent = {
  type: "break";
  /** Break type */
  breakType?: "page" | "column" | "textWrapping";
  /** Clear type for text wrapping break */
  clear?: "none" | "left" | "right" | "all";
};

/**
 * Symbol character (special font character)
 */
export type SymbolContent = {
  type: "symbol";
  /** Font name */
  font: string;
  /** Character code */
  char: string;
};

const OOXML_SYMBOL_CHARACTER_PATTERN = /^[\dA-Fa-f]{4}$/u;

/** Whether a symbol character is exactly four hexadecimal digits. */
export const isOoxmlSymbolCharacter = (value: string): boolean =>
  OOXML_SYMBOL_CHARACTER_PATTERN.test(value);

/**
 * Footnote or endnote reference
 */
export type NoteReferenceContent = {
  type: "footnoteRef" | "endnoteRef";
  /** Note ID */
  id: number;
};

/**
 * Field character (begin/separate/end)
 */
export type FieldCharContent = {
  type: "fieldChar";
  /** Field character type */
  charType: "begin" | "separate" | "end";
  /** `@w:fldLock`: absent states nothing, `false` is an explicit unlock. */
  fldLock?: boolean;
  /** `@w:dirty`: absent states nothing, `false` explicitly forbids a recompute. */
  dirty?: boolean;
  /**
   * Cached display value from a child `<w:numberingChange w:original="…"/>`.
   * Word writes this on the end fldChar of self-numbering fields (LISTNUM,
   * AUTONUM, …) so the static "(a)" / "1." text is recoverable without
   * re-evaluating the field. Used as a fallback `fieldResult` when the
   * field has no `separate` run.
   */
  originalValue?: string;
};

/**
 * Field instruction text
 */
export type InstrTextContent = {
  type: "instrText";
  /** Field instruction */
  text: string;
};

/**
 * Soft hyphen
 */
export type SoftHyphenContent = {
  type: "softHyphen";
};

/**
 * Non-breaking hyphen
 */
export type NoBreakHyphenContent = {
  type: "noBreakHyphen";
};

/** Cached pagination boundary emitted by a previous layout pass. */
export type RenderedPageBreakContent = {
  type: "renderedPageBreak";
};

/**
 * A run child folio does not model, kept byte-for-byte at its source position.
 *
 * Dropping one is worse than a loss. The parser's keep rule asks the source
 * element whether a run carried a payload while the serializer writes the
 * model, so an unmodelled child makes the two disagree for exactly one save:
 * save 1 writes a run with nothing in it, the next parse drops that run, and
 * save 2 differs from save 1. Holding the markup in the model is what makes
 * parse∘serialize a fixed point for the whole class at once.
 */
export type PreservedXmlContent = {
  type: "preservedXml";
  /** Replayable markup for one run child, as `captureVerbatimXml` wrote it. */
  xml: string;
  /**
   * The visible text the markup contributes, empty when it shows nothing.
   * `w:ruby` is the case that matters: its `w:rubyBase` is the text a reader
   * sees, so text extraction, markdown and layout would otherwise lose a word.
   */
  text: string;
};

/** Raw XML handling modes for drawings that Folio cannot model completely. */
export const DRAWING_RAW_XML_MODES = {
  PRESERVE_ONLY: "preserveOnly",
  /** The modeled image is a rendered preview of richer markup, so regeneration cannot reproduce the drawing. */
  PREVIEW_ONLY: "previewOnly",
} as const;

/** Raw XML handling mode for a drawing. */
export type DrawingRawXmlMode = (typeof DRAWING_RAW_XML_MODES)[keyof typeof DRAWING_RAW_XML_MODES];

/** Drawing/image reference with replayable XML required for preservation-only content. */
export type DrawingContent =
  | {
      type: "drawing";
      /** Image data */
      image: Image;
      /** Original OOXML for package-preserving round-trips while the editable image is unchanged. */
      rawXml?: string;
      /** Editable image projection fingerprint when `rawXml` was captured, used to invalidate replay after model edits. */
      rawImageFingerprint?: string;
      rawXmlMode?: never;
    }
  | {
      type: "drawing";
      /** Image data */
      image: Image;
      /** Original OOXML required to replay a drawing without an editable projection. */
      rawXml: string;
      /** Explicitly classifies raw XML that has no editable projected representation. */
      rawXmlMode: typeof DRAWING_RAW_XML_MODES.PRESERVE_ONLY;
    }
  | {
      type: "drawing";
      /** A render of `rawXml`, not a projection of it: editing it cannot describe the drawing. */
      image: Image;
      /** Original OOXML, the only faithful representation of the drawing. */
      rawXml: string;
      /** Required: a stale fingerprint is what marks the preview unsaveable rather than regenerable. */
      rawImageFingerprint: string;
      rawXmlMode: typeof DRAWING_RAW_XML_MODES.PREVIEW_ONLY;
    };

/**
 * Shape reference
 */
export type ShapeContent = {
  type: "shape";
  /** Shape data */
  shape: Shape;
};

/**
 * All possible run content types
 */
export type RunContent =
  | TextContent
  | TabContent
  | BreakContent
  | SymbolContent
  | NoteReferenceContent
  | FieldCharContent
  | InstrTextContent
  | SoftHyphenContent
  | NoBreakHyphenContent
  | RenderedPageBreakContent
  | PreservedXmlContent
  | DrawingContent
  | ShapeContent;

// ============================================================================
// RUN (w:r)
// ============================================================================

/**
 * A run is a contiguous region of text with the same formatting
 */
export type Run = {
  type: "run";
  /** Text formatting properties */
  formatting?: TextFormatting;
  /** Run-level tracked property changes (w:rPrChange) */
  propertyChanges?: RunPropertyChange[];
  /** Run content (text, tabs, breaks, etc.) */
  content: RunContent[];
  /**
   * Attributes `w:r` carried that this record has no field for.
   *
   * `CT_R` declares `w:rsidR`, `w:rsidDel` and `w:rsidRPr`, and folio rebuilt
   * every run without them. The remainder survives a save; the editor has no
   * run record to carry it on, which the container contract states and the
   * census records as `editorProjection`.
   */
  preservedAttributes?: PreservedAttribute[];
};

// ============================================================================
// HYPERLINKS & BOOKMARKS
// ============================================================================

/**
 * Hyperlink (w:hyperlink)
 */
export type Hyperlink = {
  type: "hyperlink";
  /** Relationship ID for external link */
  rId?: string;
  /** Resolved URL (from relationships) */
  href?: string;
  /** Internal bookmark anchor */
  anchor?: string;
  /** Tooltip text */
  tooltip?: string;
  /** Target frame */
  target?: string;
  /** Link history tracking */
  history?: boolean;
  /** Document location */
  docLocation?: string;
  /**
   * The link's content: runs, bookmark boundaries, and markup folio does not
   * model kept where the source put it.
   *
   * `CT_Hyperlink` is `EG_PContent`, so a link may hold a permission range, a
   * proofing error, a smart tag or a custom-XML revision range between its
   * runs. The capture is a member of this union rather than a sink beside it
   * for the reason `PreservedInline` gives: position inside the link is what
   * decides whether the markup goes with the link when the link moves.
   */
  children: (Run | BookmarkStart | BookmarkEnd | PreservedInline)[];
};

/**
 * Which side of a `w:customXml` element a range marker was displaced to
 * (`w:displacedByCustomXml`, ST_DisplacedByCustomXml).
 */
export type DisplacedByCustomXml = "next" | "prev";

/**
 * The attributes of `CT_MarkupRange`, the base type of every range marker:
 * bookmarks, comment ranges, move ranges and their ends.
 *
 * The three types below mirror the schema's own derivation chain
 * (`CT_MarkupRange` → `CT_Bookmark` → `CT_MoveBookmark`) so a marker's model
 * is the element's attribute set rather than a hand-picked subset of it. Every
 * marker that lost attributes on save lost them by declaring its own fields.
 */
export type MarkupRangeMarker = {
  /** Pairs a start with its end (`w:id`). */
  id: number;
  displacedByCustomXml?: DisplacedByCustomXml;
};

/** `CT_Bookmark`: a named range, optionally scoped to table columns. */
export type BookmarkRangeMarker = MarkupRangeMarker & {
  name: string;
  /** Column index for table bookmarks */
  colFirst?: number;
  colLast?: number;
};

/**
 * `CT_MoveBookmark`: a bookmark range attributed to the move that created it.
 *
 * The schema makes `w:author` required, so the field is not optional: a marker
 * that cannot name an author cannot be written at all, and the parser supplies
 * the same `Unknown` fallback a tracked change gets. `w:date` is required too,
 * but a date folio does not have is not a date it may invent, so a document
 * that arrived without one keeps arriving without one.
 */
export type MoveBookmarkMarker = BookmarkRangeMarker & {
  author: string;
  date?: string;
};

/**
 * Bookmark start marker (w:bookmarkStart)
 */
export type BookmarkStart = { type: "bookmarkStart" } & BookmarkRangeMarker;

/**
 * Bookmark end marker (w:bookmarkEnd)
 */
export type BookmarkEnd = { type: "bookmarkEnd" } & MarkupRangeMarker;

// ============================================================================
// FIELDS
// ============================================================================

/**
 * Known field types
 */
export type FieldType =
  | "PAGE"
  | "NUMPAGES"
  | "NUMWORDS"
  | "NUMCHARS"
  | "DATE"
  | "TIME"
  | "CREATEDATE"
  | "SAVEDATE"
  | "PRINTDATE"
  | "AUTHOR"
  | "TITLE"
  | "SUBJECT"
  | "KEYWORDS"
  | "COMMENTS"
  | "FILENAME"
  | "FILESIZE"
  | "TEMPLATE"
  | "DOCPROPERTY"
  | "DOCVARIABLE"
  | "REF"
  | "PAGEREF"
  | "NOTEREF"
  | "HYPERLINK"
  | "TOC"
  | "TOA"
  | "INDEX"
  | "SEQ"
  | "STYLEREF"
  | "AUTONUM"
  | "AUTONUMLGL"
  | "AUTONUMOUT"
  | "LISTNUM"
  | "IF"
  | "MERGEFIELD"
  | "NEXT"
  | "NEXTIF"
  | "ASK"
  | "SET"
  | "QUOTE"
  | "INCLUDETEXT"
  | "INCLUDEPICTURE"
  | "SYMBOL"
  | "ADVANCE"
  | "EDITTIME"
  | "REVNUM"
  | "SECTION"
  | "SECTIONPAGES"
  | "USERADDRESS"
  | "USERNAME"
  | "USERINITIALS"
  | "UNKNOWN";

/**
 * Simple field (w:fldSimple)
 */
export type SimpleField = {
  type: "simpleField";
  /** Field instruction (e.g., "PAGE \\* MERGEFORMAT") */
  instruction: string;
  /** Parsed field type */
  fieldType: FieldType;
  /**
   * The field's cached display, and any markup folio does not model between
   * the runs that carry it. `CT_SimpleField` is `EG_PContent` plus
   * `w:fldData`, so everything `EG_PContent` admits can sit here.
   */
  content: (Run | Hyperlink | PreservedInline)[];
  /** `@w:fldLock`: absent states nothing, `false` is an explicit unlock. */
  fldLock?: boolean;
  /** `@w:dirty`: absent states nothing, `false` explicitly forbids a recompute. */
  dirty?: boolean;
};

/**
 * Complex field (w:fldChar begin/separate/end with w:instrText)
 */
export type ComplexField = {
  type: "complexField";
  /** Field instruction */
  instruction: string;
  /** Parsed field type */
  fieldType: FieldType;
  /** Field code runs */
  fieldCode: Run[];
  /** Display result runs */
  fieldResult: Run[];
  /**
   * Run formatting captured from the field's structural run(s) (`w:rPr` on the
   * `begin`/code run). Used as the run-formatting fallback when the field has
   * no separate result run to read it from, e.g. a footer `PAGE` number whose
   * formatting lives on the field run (eigenpal/docx-editor#909).
   */
  formatting?: TextFormatting;
  /** `@w:fldLock`: absent states nothing, `false` is an explicit unlock. */
  fldLock?: boolean;
  /** `@w:dirty`: absent states nothing, `false` explicitly forbids a recompute. */
  dirty?: boolean;
};

export type Field = SimpleField | ComplexField;

// ============================================================================
// IMAGES
// ============================================================================

/**
 * Image size specification
 */
export type ImageSize = {
  /** Width in EMUs (English Metric Units) */
  width: number;
  /** Height in EMUs */
  height: number;
};

/**
 * Image wrap type for floating images
 */
export type ImageWrap = {
  type: "inline" | "square" | "tight" | "through" | "topAndBottom" | "behind" | "inFront";
  /** Wrap text direction */
  wrapText?: "bothSides" | "left" | "right" | "largest";
  /** Distance from text */
  distT?: number;
  distB?: number;
  distL?: number;
  distR?: number;
};

/**
 * Position for floating images
 */
export type ImagePosition = {
  /** Horizontal positioning */
  horizontal: {
    relativeTo:
      | "character"
      | "column"
      | "insideMargin"
      | "leftMargin"
      | "margin"
      | "outsideMargin"
      | "page"
      | "rightMargin";
    alignment?: "left" | "right" | "center" | "inside" | "outside";
    posOffset?: number;
  };
  /** Vertical positioning */
  vertical: {
    relativeTo:
      | "insideMargin"
      | "line"
      | "margin"
      | "outsideMargin"
      | "page"
      | "paragraph"
      | "topMargin"
      | "bottomMargin";
    alignment?: "top" | "bottom" | "center" | "inside" | "outside";
    posOffset?: number;
  };
};

/**
 * Image transformation
 */
export type ImageTransform = {
  /** Rotation in degrees */
  rotation?: number;
  /** Flip horizontal */
  flipH?: boolean;
  /** Flip vertical */
  flipV?: boolean;
};

/**
 * Image padding/margins
 */
export type ImagePadding = {
  top?: number;
  bottom?: number;
  left?: number;
  right?: number;
};

/**
 * Image crop fractions in [0, 1] applied to each side of the source bitmap.
 * Mirrors the four `<a:srcRect>` attributes (`l`, `t`, `r`, `b`) defined in
 * ECMA-376 §20.1.8.55, stored in 1/100000 units on the wire.
 *
 * eigenpal #424 (image-crop subset).
 */
export type ImageCrop = {
  left?: number;
  top?: number;
  right?: number;
  bottom?: number;
};

/**
 * `a:graphicFrameLocks` manipulation locks (ECMA-376 §20.1.2.2.19), in schema
 * attribute order. A field is undefined when the authored XML omits the
 * attribute, so a round-trip never materializes a lock the author never wrote.
 */
export type ImageFrameLocks = {
  noGrp?: boolean;
  noDrilldown?: boolean;
  noSelect?: boolean;
  noChangeAspect?: boolean;
  noMove?: boolean;
  noResize?: boolean;
};

/**
 * Embedded image (w:drawing)
 */
export type Image = {
  type: "image";
  /** Unique ID */
  id?: string;
  /**
   * Relationship id for the image data, absent when the drawing carries none.
   *
   * A `w:drawing` whose graphic is not a picture — a chart, a diagram, an OLE
   * frame — and one with no `a:graphic` at all have no `a:blip`, so there is no
   * id to record. Absence is spelled `undefined` rather than `""` so it can
   * never reach a relationship lookup as a key.
   */
  rId?: string;
  /** Resolved image data (base64 or blob URL) */
  src?: string;
  /** Image MIME type */
  mimeType?: string;
  /** Original filename */
  filename?: string;
  /** Authored non-visual drawing name (`wp:docPr@name`) */
  docPrName?: string;
  /** Alt text for accessibility */
  alt?: string;
  /** Authored non-visual drawing title (`wp:docPr@title`) */
  title?: string;
  /** Image size */
  size: ImageSize;
  /** Original size before any transforms */
  originalSize?: ImageSize;
  /** Wrap settings */
  wrap: ImageWrap;
  /** Position for floating images */
  position?: ImagePosition;
  /** Image transformations */
  transform?: ImageTransform;
  /** Padding around image */
  padding?: ImagePadding;
  /** Source-bitmap crop (wp:srcRect), eigenpal #424 */
  crop?: ImageCrop;
  /** `wp:cNvGraphicFramePr > a:graphicFrameLocks`: authored manipulation locks. */
  frameLocks?: ImageFrameLocks;
  /**
   * Opacity in [0, 1] (OOXML `a:alphaModFix amt`). Undefined or `1` means
   * fully opaque. Mirrors eigenpal docx-editor #424.
   */
  opacity?: number;
  /**
   * `wp:anchor layoutInCell` — when true (OOXML default), an anchored image
   * inside a table cell is constrained to the cell. When false, the image
   * escapes the cell into the page area. Round-tripped on save so the
   * author's intent survives; undefined means "use the spec default".
   */
  layoutInCell?: boolean;
  /**
   * `wp:anchor allowOverlap` — when true (OOXML default), anchored objects
   * may overlap; when false, Word repositions them to avoid collisions. We
   * don't currently reposition, but we round-trip the flag so saving
   * preserves the author's intent; undefined means "use the spec default".
   */
  allowOverlap?: boolean;
  /**
   * The image carries no information a reader needs, so assistive technology
   * skips it. Word writes this as an extension on `wp:docPr`
   * (`{C183D7F6-B498-43B3-948B-1728B52AA6E4}` holding
   * `adec:decorative val="1"`), not as an attribute: `CT_NonVisualDrawingProps`
   * has no `@decorative`.
   *
   * Not `hidden`. A decorative image is displayed and skipped by a screen
   * reader; a hidden one is not displayed at all.
   */
  decorative?: boolean;
  /** `wp:docPr @hidden`: the drawing is not displayed. */
  hidden?: boolean;
  /**
   * The `a:ext` entries of the `wp:docPr` extension list that folio does not
   * model — a creation id, a local-DPI hint — captured as XML in source order.
   * Written back beside the decorative extension so an extension folio has no
   * opinion about still survives a save.
   */
  docPrExtensions?: string[];
  /** Hyperlink URL for clickable image */
  hlinkHref?: string;
  /** Relationship ID for the clickable image hyperlink */
  hlinkRId?: string;
  /** Image outline/border */
  outline?: ShapeOutline;
  /** Image effects */
  effects?: {
    brightness?: number;
    contrast?: number;
    saturation?: number;
  };
};

// ============================================================================
// SHAPES & TEXT BOXES
// ============================================================================

/**
 * Shape types
 */
export type ShapeType =
  // Basic shapes
  | "rect"
  | "roundRect"
  | "ellipse"
  | "triangle"
  | "rtTriangle"
  | "parallelogram"
  | "trapezoid"
  | "pentagon"
  | "hexagon"
  | "heptagon"
  | "octagon"
  | "decagon"
  | "dodecagon"
  | "star4"
  | "star5"
  | "star6"
  | "star7"
  | "star8"
  | "star10"
  | "star12"
  | "star16"
  | "star24"
  | "star32"
  // Lines and connectors
  | "line"
  | "straightConnector1"
  | "bentConnector2"
  | "bentConnector3"
  | "bentConnector4"
  | "bentConnector5"
  | "curvedConnector2"
  | "curvedConnector3"
  | "curvedConnector4"
  | "curvedConnector5"
  // Arrows
  | "rightArrow"
  | "leftArrow"
  | "upArrow"
  | "downArrow"
  | "leftRightArrow"
  | "upDownArrow"
  | "quadArrow"
  | "leftRightUpArrow"
  | "bentArrow"
  | "uturnArrow"
  | "leftUpArrow"
  | "bentUpArrow"
  | "curvedRightArrow"
  | "curvedLeftArrow"
  | "curvedUpArrow"
  | "curvedDownArrow"
  | "stripedRightArrow"
  | "notchedRightArrow"
  | "homePlate"
  | "chevron"
  | "rightArrowCallout"
  | "downArrowCallout"
  | "leftArrowCallout"
  | "upArrowCallout"
  | "leftRightArrowCallout"
  | "quadArrowCallout"
  | "circularArrow"
  // Flowchart
  | "flowChartProcess"
  | "flowChartAlternateProcess"
  | "flowChartDecision"
  | "flowChartInputOutput"
  | "flowChartPredefinedProcess"
  | "flowChartInternalStorage"
  | "flowChartDocument"
  | "flowChartMultidocument"
  | "flowChartTerminator"
  | "flowChartPreparation"
  | "flowChartManualInput"
  | "flowChartManualOperation"
  | "flowChartConnector"
  | "flowChartOffpageConnector"
  | "flowChartPunchedCard"
  | "flowChartPunchedTape"
  | "flowChartSummingJunction"
  | "flowChartOr"
  | "flowChartCollate"
  | "flowChartSort"
  | "flowChartExtract"
  | "flowChartMerge"
  | "flowChartOnlineStorage"
  | "flowChartDelay"
  | "flowChartMagneticTape"
  | "flowChartMagneticDisk"
  | "flowChartMagneticDrum"
  | "flowChartDisplay"
  // Callouts
  | "wedgeRectCallout"
  | "wedgeRoundRectCallout"
  | "wedgeEllipseCallout"
  | "cloudCallout"
  | "borderCallout1"
  | "borderCallout2"
  | "borderCallout3"
  | "accentCallout1"
  | "accentCallout2"
  | "accentCallout3"
  | "callout1"
  | "callout2"
  | "callout3"
  | "accentBorderCallout1"
  | "accentBorderCallout2"
  | "accentBorderCallout3"
  // Other
  | "actionButtonBlank"
  | "actionButtonHome"
  | "actionButtonHelp"
  | "actionButtonInformation"
  | "actionButtonBackPrevious"
  | "actionButtonForwardNext"
  | "actionButtonBeginning"
  | "actionButtonEnd"
  | "actionButtonReturn"
  | "actionButtonDocument"
  | "actionButtonSound"
  | "actionButtonMovie"
  | "irregularSeal1"
  | "irregularSeal2"
  | "frame"
  | "halfFrame"
  | "corner"
  | "diagStripe"
  | "chord"
  | "arc"
  | "bracketPair"
  | "bracePair"
  | "leftBracket"
  | "rightBracket"
  | "leftBrace"
  | "rightBrace"
  | "can"
  | "cube"
  | "bevel"
  | "donut"
  | "noSmoking"
  | "blockArc"
  | "foldedCorner"
  | "smileyFace"
  | "heart"
  | "lightningBolt"
  | "sun"
  | "moon"
  | "cloud"
  | "snip1Rect"
  | "snip2SameRect"
  | "snip2DiagRect"
  | "snipRoundRect"
  | "round1Rect"
  | "round2SameRect"
  | "round2DiagRect"
  | "plaque"
  | "teardrop"
  | "mathPlus"
  | "mathMinus"
  | "mathMultiply"
  | "mathDivide"
  | "mathEqual"
  | "mathNotEqual"
  | "gear6"
  | "gear9"
  | "funnel"
  | "pieWedge"
  | "pie"
  | "leftCircularArrow"
  | "leftRightCircularArrow"
  | "swooshArrow"
  | "textBox";

/** Authored adjustment formula for a preset DrawingML geometry. */
export type ShapeGeometryAdjustment = {
  name: string;
  formula: string;
};

/**
 * Shape fill type
 */
export type ShapeFill = {
  type: "none" | "solid" | "gradient" | "pattern" | "picture";
  /**
   * Authored DrawingML retained when the normalized projection is incomplete.
   * Serializers prefer this value; omit it when modifying the fill.
   */
  rawXml?: string;
  /** Solid fill color */
  color?: ColorValue;
  /** Gradient stops for gradient fill */
  gradient?: {
    type: "linear" | "radial" | "rectangular" | "path";
    angle?: number;
    stops: {
      position: number; // 0-100000
      color: ColorValue;
    }[];
  };
};

/**
 * Shape outline/stroke
 */
export type ShapeOutline = {
  /**
   * Authored DrawingML retained for lossless round-trips. Serializers prefer
   * this value; callers modifying structured outline fields must omit it.
   */
  rawXml?: string;
  /** Line width in EMUs */
  width?: number;
  /** Line color */
  color?: ColorValue;
  /** Line style */
  style?:
    | "solid"
    | "dot"
    | "dash"
    | "lgDash"
    | "dashDot"
    | "lgDashDot"
    | "lgDashDotDot"
    | "sysDot"
    | "sysDash"
    | "sysDashDot"
    | "sysDashDotDot";
  /** Line cap */
  cap?: "flat" | "round" | "square";
  /** Line join */
  join?: "bevel" | "miter" | "round";
  /** Head arrow */
  headEnd?: {
    type: "none" | "triangle" | "stealth" | "diamond" | "oval" | "arrow";
    width?: "sm" | "med" | "lg";
    length?: "sm" | "med" | "lg";
  };
  /** Tail arrow */
  tailEnd?: {
    type: "none" | "triangle" | "stealth" | "diamond" | "oval" | "arrow";
    width?: "sm" | "med" | "lg";
    length?: "sm" | "med" | "lg";
  };
};

/**
 * Text body inside a shape
 */
export type ShapeTextBody = {
  /** Authored DrawingML WordArt metadata. */
  wordArt?: {
    /** Authored `wps:bodyPr/@fromWordArt`; omitted when the source omitted it. */
    fromWordArt?: boolean;
    preset?: string;
    adjustments?: ShapeGeometryAdjustment[];
  };
  /** Text direction */
  vertical?: boolean;
  /** Rotation */
  rotation?: number;
  /** Anchor/vertical alignment */
  anchor?: "top" | "middle" | "bottom" | "distributed" | "justified";
  /** Anchor center */
  anchorCenter?: boolean;
  /** Auto fit */
  autoFit?: "none" | "normal" | "shape";
  /** Horizontal text wrapping inside the shape */
  textWrap?: "square" | "none";
  /** Text margins */
  margins?: {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
  };
  /** Block content inside the shape */
  content: (Paragraph | Table)[];
};

/**
 * Shape/drawing object (wps:wsp)
 */
export type Shape = {
  type: "shape";
  /** Shape type preset */
  shapeType: ShapeType;
  /** Authored preset-geometry adjustments, in document order. */
  geometryAdjustments?: ShapeGeometryAdjustment[];
  /** Unique ID */
  id?: string;
  /** Authored non-visual drawing name (`wp:docPr@name` / `wps:cNvPr@name`) */
  name?: string;
  /** Alt text for accessibility (`wp:docPr@descr` / `wps:cNvPr@descr`) */
  alt?: string;
  /** Authored non-visual drawing title (`wp:docPr@title` / `wps:cNvPr@title`) */
  title?: string;
  /** Size in EMUs */
  size: ImageSize;
  /** Position for floating shapes */
  position?: ImagePosition;
  /** Wrap settings */
  wrap?: ImageWrap;
  /** Fill */
  fill?: ShapeFill;
  /** Outline/stroke */
  outline?: ShapeOutline;
  /** Transform */
  transform?: ImageTransform;
  /** Text content inside the shape */
  textBody?: ShapeTextBody;
  /** Custom geometry points */
  customGeometry?: string;
};

/**
 * Text box (floating text container)
 */
export type TextBox = {
  type: "textBox";
  /** Unique ID */
  id?: string;
  /** Authored non-visual drawing name (`wp:docPr@name` / `wps:cNvPr@name`) */
  name?: string;
  /** Alt text for accessibility (`wp:docPr@descr` / `wps:cNvPr@descr`) */
  alt?: string;
  /** Authored non-visual drawing title (`wp:docPr@title` / `wps:cNvPr@title`) */
  title?: string;
  /** Size */
  size: ImageSize;
  /** Position */
  position?: ImagePosition;
  /** Wrap settings */
  wrap?: ImageWrap;
  /** Fill */
  fill?: ShapeFill;
  /** Outline */
  outline?: ShapeOutline;
  /** DrawingML transform applied to the text-box frame and content. */
  transform?: ImageTransform;
  /** Text and table content */
  content: (Paragraph | Table)[];
  /** Authored DrawingML WordArt metadata. */
  wordArt?: ShapeTextBody["wordArt"];
  /** Text fitting behavior */
  autoFit?: ShapeTextBody["autoFit"];
  /** Horizontal text wrapping inside the box */
  textWrap?: ShapeTextBody["textWrap"];
  /** Vertical text alignment inside the box */
  verticalAlign?: ShapeTextBody["anchor"];
  /** Internal margins */
  margins?: {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
  };
};

// ============================================================================
// TABLES
// ============================================================================

/**
 * Table cell
 */
export type TableCell = {
  type: "tableCell";
  /** Cell formatting */
  formatting?: TableCellFormatting;
  /** Cell-level tracked property changes (w:tcPrChange) */
  propertyChanges?: TableCellPropertyChange[];
  /** Tracked structural changes (cell insert/delete/merge) */
  structuralChange?: TableStructuralChangeInfo;
  /**
   * Cell content.
   *
   * Derived from {@link BlockContent} rather than listed, so a new block kind
   * cannot land without a decision here; `BlockSdt` is excluded because folio
   * unwraps a `w:sdt` inside a cell into its `sdtContent` children rather
   * than modelling the wrapper, and keeping the branch out of the cell keeps
   * the model's recursion out of every table.
   */
  content: TableCellBlock[];
};

/**
 * Table row
 */
export type TableRow = {
  type: "tableRow";
  /** Row formatting */
  formatting?: TableRowFormatting;
  /**
   * The table properties this row overrides (`w:tblPrEx`).
   *
   * `CT_TblPrEx` is the middle of `CT_TblPrBase` — width, justification, cell
   * spacing, indent, borders, shading, layout, cell margins and the look — so
   * it is the table's own property set narrowed to what a row may restate, and
   * it travels in the same shape. Word writes one when a table is built by
   * merging two, and a consumer reads it in place of the table's for this row
   * alone, so losing it silently restyles the row.
   *
   * It is not `formatting`: that is `w:trPr`, the row's own geometry, and the
   * two are different elements in different places in `CT_Row`.
   */
  tablePropertyExceptions?: TableFormatting;
  /** Row-level tracked property changes (w:trPrChange) */
  propertyChanges?: TableRowPropertyChange[];
  /** Tracked changes to the property exceptions (w:tblPrExChange) */
  tablePropertyExceptionChanges?: TablePropertyExceptionChange[];
  /** Tracked structural changes (row insert/delete) */
  structuralChange?: TableStructuralChangeInfo;
  /** Cells in this row */
  cells: TableCell[];
  /**
   * Row markup `cells` cannot hold, with its position among them.
   *
   * `CT_Row` declares a permission range, a proofing error, a custom-XML
   * revision range and the row-level move and comment ranges beside its
   * cells, and none of them is a cell. This is the sink case rather than the
   * union case the inline levels use: the row models one kind of child, so
   * there is no member to be, and `index` counts the cells that preceded the
   * capture.
   */
  preserved?: PreservedMarkup;
  /**
   * Attributes `w:tr` carried that this record has no field for.
   *
   * `CT_Row` declares `w:rsidR`, `w:rsidDel`, `w:rsidTr` and `w:rsidRPr`. The
   * remainder follows the record the same way a paragraph's does: a row the
   * editor creates has none, and a row split off another does not inherit one.
   */
  preservedAttributes?: PreservedAttribute[];
};

/**
 * Table (w:tbl)
 */
export type Table = {
  type: "table";
  /** Table formatting */
  formatting?: TableFormatting;
  /** Table-level tracked property changes (w:tblPrChange) */
  propertyChanges?: TablePropertyChange[];
  /** Column widths in twips */
  columnWidths?: number[];
  /** Table rows */
  rows: TableRow[];
  /**
   * Table markup `rows` cannot hold, with its position among them.
   *
   * `CT_Tbl` declares a permission range, a proofing error, the table-level
   * comment and move ranges and the eight custom-XML revision ranges beside
   * `w:tr`, and none of them is a row. The row sink one level down has the
   * same shape for the same reason: the table models one kind of child, so
   * there is no union member to be, and `index` counts the rows that preceded
   * the capture. `w:tblPr` and `w:tblGrid` precede every row in the content
   * model and are read and written elsewhere, so they are not in the sink and
   * the index counts rows only.
   */
  preserved?: PreservedMarkup;
};

// ============================================================================
// COMMENTS
// ============================================================================

/**
 * A comment (w:comment) from comments.xml
 */
export type Comment = {
  /** Comment ID (matches commentRangeStart/End) */
  id: number;
  /** Author name */
  author: string;
  /** Author initials */
  initials?: string;
  /** Date */
  date?: string;
  /** Comment content (paragraphs) */
  content: Paragraph[];
  /** Formatting of the structural annotation-reference run in the first paragraph */
  annotationReferenceFormatting?: TextFormatting;
  /** Parent comment ID (for replies) */
  parentId?: number;
  /** Whether the comment is resolved/done */
  done?: boolean;
  /**
   * Body markup `content` cannot hold: a table, an equation, a content
   * control, a bookmark or range marker, a tracked-change wrapper. The
   * schema lets a comment body hold everything a document body can, and a
   * reviewer's words disappearing on save is not an acceptable simplification
   * of that.
   */
  preserved?: PreservedMarkup;
};

/**
 * Comment range start marker in paragraph content
 */
export type CommentRangeStart = { type: "commentRangeStart" } & MarkupRangeMarker;

/**
 * Comment range end marker in paragraph content
 */
export type CommentRangeEnd = { type: "commentRangeEnd" } & MarkupRangeMarker;

/**
 * Point comment reference (w:commentReference without an explicit range).
 * Word sometimes stores comments this way; we anchor them to nearby text for display.
 */
export type CommentReference = {
  type: "commentReference";
  id: number;
};

// ============================================================================
// MATH EQUATIONS
// ============================================================================

/**
 * Math equation content (m:oMath or m:oMathPara)
 */
export type MathEquation = {
  type: "mathEquation";
  /** Whether this is a block (oMathPara) or inline (oMath) equation */
  display: "inline" | "block";
  /** Raw OMML XML for round-trip preservation */
  ommlXml: string;
  /** Plain text representation for accessibility/fallback */
  plainText?: string;
};

// ============================================================================
// TRACKED CHANGES
// ============================================================================

/**
 * Largest value a revision id (`w:id`) may carry: 2^31 - 1.
 *
 * `w:id` on `<w:ins>`/`<w:del>` comes from `CT_Markup`, typed
 * `ST_DecimalNumber` — which ECMA-376 defines as an unbounded integer. The
 * bound below is an implementation limit: conforming consumers read
 * `ST_DecimalNumber` into a signed 32-bit int, and a value past this bound
 * overflows on load and surfaces as an unreadable document. Port of
 * eigenpal/docx-editor#1093.
 */
export const MAX_REVISION_ID = 2_147_483_647;

/**
 * Coerce any revision id — including one parsed from an untrusted DOCX — into
 * the range serialized `w:id` attributes may occupy.
 *
 * - Malformed (negative, fractional, `NaN`, `Infinity`): collapse to `0`.
 * - Well-formed but out of range: fold modulo the range (not clamp) so a
 *   contiguous run of overflowing ids stays distinguishable.
 */
export function normalizeRevisionId(id: number): number {
  if (!Number.isInteger(id) || id < 0) {
    return 0;
  }
  if (id > MAX_REVISION_ID) {
    return id % (MAX_REVISION_ID + 1);
  }
  return id;
}

/**
 * Tracked change metadata (w:ins, w:del attributes)
 */
export type TrackedChangeInfo = {
  /** Revision ID */
  id: number;
  /** Author who made the change */
  author: string;
  /** Date of the change */
  date?: string;
  /**
   * Author initials (w:initials). Optional attribution used by the review UI
   * and carried through the round-trip when present on the source document.
   */
  initials?: string;
  /**
   * The UTC companion to `w:date` a recent producer writes alongside it, kept
   * with the prefix the document bound it under: the attribute names a
   * namespace the part declares on its root, and re-emitting it under a
   * prefix of our choosing would name one that is not there.
   */
  utcDate?: { attribute: string; value: string };
};

/**
 * Generic tracked property-change wrapper metadata (w:*PrChange)
 */
export type PropertyChangeInfo = {
  /** Optional revision session ID */
  rsid?: string;
} & TrackedChangeInfo;

/**
 * Inline content that may sit inside a run-level tracked-change wrapper.
 *
 * The transparent wrappers belong here. A `w:bdo` / `w:dir` states how its
 * content is laid out and a `w:sdt` states what its content is bound to;
 * neither says anything about the revision, so a revision that holds one
 * holds it as content. Modelling it otherwise forces the parser to lift the
 * wrapper out to a sibling, which takes the wrapped text out of the revision
 * with it: the text is then neither inserted nor deleted, and accepting and
 * rejecting the revision both keep it.
 */
export type TrackedRunContent =
  | Run
  | Hyperlink
  | BookmarkStart
  | BookmarkEnd
  | SimpleField
  | ComplexField
  | InlineSdt
  | InlineWrapper
  // CT_RunTrackChange permits both m:oMath and m:oMathPara.
  | MathEquation
  | PreservedInline
  | TrackedRunChange;

/**
 * Insertion wrapper (w:ins) — runs inserted by tracked changes
 */
export type Insertion = {
  type: "insertion";
  /** Tracked change metadata */
  info: TrackedChangeInfo;
  /** Inserted content */
  content: TrackedRunContent[];
};

/**
 * Deletion wrapper (w:del) — runs deleted by tracked changes
 */
export type Deletion = {
  type: "deletion";
  /** Tracked change metadata */
  info: TrackedChangeInfo;
  /** Deleted content */
  content: TrackedRunContent[];
};

/**
 * Move-from wrapper (w:moveFrom) â€” content moved away from this position
 */
export type MoveFrom = {
  type: "moveFrom";
  /** Tracked change metadata */
  info: TrackedChangeInfo;
  /** Moved content */
  content: TrackedRunContent[];
};

/**
 * Move-to wrapper (w:moveTo) â€” content moved into this position
 */
export type MoveTo = {
  type: "moveTo";
  /** Tracked change metadata */
  info: TrackedChangeInfo;
  /** Moved content */
  content: TrackedRunContent[];
};

/**
 * Which Unicode bidirectional control a `w:bdo`/`w:dir` wrapper is.
 *
 * `w:dir` is an embedding: the run of text inside it is laid out in the given
 * direction and the bidirectional algorithm still resolves the characters
 * within it. `w:bdo` is an override: the algorithm is switched off inside and
 * every character is laid out in the given direction, which is what makes a
 * Latin word inside an `rtl` override read backwards. CSS spells the pair
 * `unicode-bidi: embed` and `unicode-bidi: bidi-override`, and HTML gives the
 * second an element of its own, `<bdo>`.
 */
export const BIDI_CONTROLS = { embedding: "embedding", override: "override" } as const;

export type BidiControl = (typeof BIDI_CONTROLS)[keyof typeof BIDI_CONTROLS];

/**
 * A transparent inline wrapper, discriminated by what kind of wrapper it is.
 *
 * Transparent means it says something about its content without constraining
 * it: it nests, it may hold anything paragraph content may hold, and dropping
 * it changes what the reader sees or what the markup states rather than what
 * the text is. `bidi` is the one kind folio parses today; a smart tag and a
 * custom-XML wrapper are the same shape and land with their parser.
 *
 * `bidi` is `w:dir` / `w:bdo` — ECMA-376 §17.3.2.8, §17.3.2.3.
 */
export type InlineWrapper = {
  type: "inlineWrapper";
  kind: "bidi";
  control: BidiControl;
  /** `w:val`; absent in the source means the wrapper states no direction. */
  direction?: "ltr" | "rtl";
  content: ParagraphContent[];
};

/**
 * Move-from range start marker (w:moveFromRangeStart) — ECMA-376 §17.13.5.22
 * Pairs with moveFromRangeEnd to delimit the source of a move in the document.
 */
export type MoveFromRangeStart = { type: "moveFromRangeStart" } & MoveBookmarkMarker;

/**
 * Move-from range end marker (w:moveFromRangeEnd)
 */
export type MoveFromRangeEnd = { type: "moveFromRangeEnd" } & MarkupRangeMarker;

/**
 * Move-to range start marker (w:moveToRangeStart) — ECMA-376 §17.13.5.24
 * Pairs with moveToRangeEnd to delimit the destination of a move.
 */
export type MoveToRangeStart = { type: "moveToRangeStart" } & MoveBookmarkMarker;

/**
 * Move-to range end marker (w:moveToRangeEnd)
 */
export type MoveToRangeEnd = { type: "moveToRangeEnd" } & MarkupRangeMarker;

/**
 * Run-level tracked wrappers represented in WordprocessingML.
 */
export type TrackedRunChange = Insertion | Deletion | MoveFrom | MoveTo;

/**
 * Run property change (w:rPrChange)
 */
export type RunPropertyChange = {
  type: "runPropertyChange";
  /** Tracked change metadata */
  info: PropertyChangeInfo;
  /** Run properties before the tracked change */
  previousFormatting?: TextFormatting;
  /** Run properties after the tracked change (editor model convenience) */
  currentFormatting?: TextFormatting;
};

/**
 * Paragraph property change (w:pPrChange)
 */
export type ParagraphPropertyChange = {
  type: "paragraphPropertyChange";
  /** Tracked change metadata */
  info: PropertyChangeInfo;
  /** Paragraph properties before the tracked change */
  previousFormatting?: ParagraphFormatting;
  /** Paragraph properties after the tracked change (editor model convenience) */
  currentFormatting?: ParagraphFormatting;
};

/**
 * Table property change (w:tblPrChange)
 */
export type TablePropertyChange = {
  type: "tablePropertyChange";
  /** Tracked change metadata */
  info: PropertyChangeInfo;
  /** Table properties before the tracked change */
  previousFormatting?: TableFormatting;
  /** Table properties after the tracked change (editor model convenience) */
  currentFormatting?: TableFormatting;
};

/**
 * Table property exception change (w:tblPrExChange)
 *
 * Its own type rather than a `TablePropertyChange`: the two carry the same
 * shape and are written as different elements in different containers, and a
 * shared discriminator would let one reach the other's serializer.
 */
export type TablePropertyExceptionChange = {
  type: "tablePropertyExceptionChange";
  /** Tracked change metadata */
  info: PropertyChangeInfo;
  /** Property exceptions before the tracked change */
  previousFormatting?: TableFormatting;
  /** Property exceptions after the tracked change (editor model convenience) */
  currentFormatting?: TableFormatting;
};

/**
 * Table row property change (w:trPrChange)
 */
export type TableRowPropertyChange = {
  type: "tableRowPropertyChange";
  /** Tracked change metadata */
  info: PropertyChangeInfo;
  /** Row properties before the tracked change */
  previousFormatting?: TableRowFormatting;
  /** Row properties after the tracked change (editor model convenience) */
  currentFormatting?: TableRowFormatting;
};

/**
 * Table cell property change (w:tcPrChange)
 */
export type TableCellPropertyChange = {
  type: "tableCellPropertyChange";
  /** Tracked change metadata */
  info: PropertyChangeInfo;
  /** Cell properties before the tracked change */
  previousFormatting?: TableCellFormatting;
  /** Cell properties after the tracked change (editor model convenience) */
  currentFormatting?: TableCellFormatting;
  /**
   * The cell's own insertion, deletion or merge as the snapshot recorded it.
   *
   * `CT_TcPrInner` is `CT_TcPrBase` plus `EG_CellMarkupElements`, so the
   * `w:tcPr` inside a `w:tcPrChange` may carry `w:cellIns`, `w:cellDel` or
   * `w:cellMerge` — "before this property change, the cell stood inserted".
   * It is not the cell's current revision, which is
   * {@link TableCell.structuralChange}, so it rides the change that recorded it.
   */
  previousStructuralChange?: TableStructuralChangeInfo;
};

/**
 * Section property change (w:sectPrChange)
 */
export type SectionPropertyChange = {
  type: "sectionPropertyChange";
  /** Tracked change metadata */
  info: PropertyChangeInfo;
  /** Section properties before the tracked change */
  previousProperties?: SectionProperties;
  /** Complete prior header/footer selection, stored in Folio's ignorable MCE extension. */
  previousReferences?: {
    headerReferences?: HeaderReference[];
    footerReferences?: FooterReference[];
  };
  /** Section properties after the tracked change (editor model convenience) */
  currentProperties?: SectionProperties;
};

/**
 * Table structural tracked change metadata (row/cell insert/delete/merge)
 */
export type TableStructuralChangeInfo =
  | {
      type: "tableRowInsertion" | "tableRowDeletion" | "tableCellInsertion" | "tableCellDeletion";
      /** Tracked change metadata */
      info: TrackedChangeInfo;
    }
  | {
      type: "tableCellMerge";
      /** Tracked change metadata */
      info: TrackedChangeInfo;
      /** Vertical merge state applied by the revision. */
      verticalMerge?: "continue" | "rest";
      /** Vertical merge state that existed before the revision. */
      verticalMergeOriginal?: "continue" | "rest";
    };

// ============================================================================
// STRUCTURED DOCUMENT TAGS (SDT / Content Controls)
// ============================================================================

/**
 * SDT type (content control type)
 */
export type SdtType =
  | "richText"
  | "plainText"
  | "date"
  | "dropdown"
  | "comboBox"
  | "checkbox"
  | "picture"
  | "buildingBlockGallery"
  | "group"
  | "unknown";

/**
 * SDT properties (`w:sdtPr`).
 *
 * Modeled fields are a read-only projection for downstream tooling
 * (tag/alias addressing, template extraction). They are NOT the
 * serialization source: the original `<w:sdtPr>` is captured verbatim in
 * `rawPropertiesXml` and replayed on save, which preserves element order
 * (`CT_SdtPr` is an `xsd:sequence`), avoids double-emission, and keeps
 * unmodeled features (`w:dataBinding`, `w15:repeatingSection`, `@lastValue`,
 * `w:sdtEndPr`) lossless.
 */
export type SdtProperties = {
  /** SDT type (projection; round-trip uses `rawPropertiesXml`). */
  sdtType: SdtType;
  /** Numeric id (`w:id/@w:val`). */
  id?: number;
  /** Alias (friendly name, `w:alias`). */
  alias?: string;
  /** Tag (developer identifier, `w:tag`). */
  tag?: string;
  /** Lock setting (`w:lock`). */
  lock?: "sdtLocked" | "contentLocked" | "sdtContentLocked" | "unlocked";
  /**
   * Placeholder building-block name (`w:placeholder/w:docPart@w:val`) — a
   * reference to a glossary docPart, not the literal placeholder text.
   */
  placeholder?: string;
  /** Whether the placeholder is currently shown (`w:showingPlcHdr`). */
  showingPlaceholder?: boolean;
  /** Date display format (`w:date/w:dateFormat@w:val`). */
  dateFormat?: string;
  /**
   * Bound date value (`w:date/@w:fullDate`), ISO 8601. Independent of
   * `dateFormat` (which controls display): the body text may show a
   * formatted version like "2 June 2026" while this stays as
   * `2026-06-02T00:00:00Z` so Word's date binding round-trips losslessly.
   */
  dateValueISO?: string;
  /** Dropdown/combobox list items. */
  listItems?: { displayText: string; value: string }[];
  /**
   * Selected dropdown / comboBox value (`w:dropDownList@w:lastValue`).
   * Persisted as the OOXML value, independent of the body display text.
   * Without this, the serializer had to recover the saved value by
   * matching the body's display text against `listItems`, which picked
   * the wrong entry when two items shared a `displayText`.
   */
  dropdownLastValue?: string;
  /** Checkbox checked state (`w14:checkbox/w14:checked`). */
  checked?: boolean;
  /**
   * Verbatim `<w:sdtPr>…</w:sdtPr>` captured at parse time. Replayed on
   * serialize so unmodeled OOXML features (data binding, repeating sections,
   * `@lastValue`, custom XML mappings) survive round-trip.
   */
  rawPropertiesXml?: string;
  /** Verbatim `<w:sdtEndPr>…</w:sdtEndPr>` captured at parse time. */
  rawEndPropertiesXml?: string;
  /**
   * Verbatim XML for any non-content direct children of `<w:sdt>` that
   * appear BEFORE `<w:sdtContent>` — MS-OE376 §2.5.2.30 documents 16
   * range-marker elements Word emits as direct sdt siblings (bookmark,
   * comment range, custom XML range, tracked-change range). Captured at
   * parse time and replayed on serialize so comment threads or tracked
   * changes that span an SDT boundary round-trip without losing a
   * delimiter.
   */
  rawSdtChildrenBeforeContent?: string;
  /** Verbatim XML for non-content sdt children that appear AFTER `<w:sdtContent>`. */
  rawSdtChildrenAfterContent?: string;
};

/**
 * Inline SDT (content control within a paragraph).
 *
 * OOXML allows runs, hyperlinks, simple/complex fields, nested SDTs,
 * tracked insertions/deletions/moves, the bidirectional controls, and math at
 * this level. All of them must survive parse → edit → save so docProps-bound
 * fields and reviewed template content do not lose their wrapper on round-trip.
 */
export type InlineSdt = {
  type: "inlineSdt";
  /** SDT properties */
  properties: SdtProperties;
  /** Inline content held inside the control */
  content: (
    | Run
    | Hyperlink
    | SimpleField
    | ComplexField
    | InlineSdt
    | InlineWrapper
    | Insertion
    | Deletion
    | MoveFrom
    | MoveTo
    | MathEquation
    | PreservedInline
  )[];
};

/**
 * Block-level SDT (content control wrapping paragraphs/tables).
 *
 * Content is `BlockContent[]` (not just `(Paragraph | Table)[]`) because
 * OOXML allows block SDTs to nest — e.g. a repeating-section control
 * whose row is itself a content control.
 */
export type BlockSdt = {
  type: "blockSdt";
  /** SDT properties (raw XML in `properties.rawPropertiesXml` round-trips losslessly). */
  properties: SdtProperties;
  /** Block content inside the control. */
  content: BlockContent[];
};

// ============================================================================
// PARAGRAPH
// ============================================================================

/**
 * Paragraph content types
 */
/**
 * Inline-level markup folio does not model, kept at its source position.
 *
 * The run-level twin is `PreservedXmlContent` and the block-level one is
 * `PreservedBlock`; this is the level between them. A paragraph, a run-level
 * tracked-change wrapper, a bidirectional wrapper and an inline content
 * control all admit `w:permStart`, `w:proofErr`, `w:customXml` and the eight
 * custom-XML revision ranges, none of which is a run and none of which folio
 * models. Inside a tracked-change wrapper the position is the point: markup
 * lifted out of a `w:ins` is markup the reviewer no longer accepts or rejects
 * along with the change, so the capture is a member of the wrapper's own
 * content union rather than a sibling beside it.
 *
 * Opaque, so it holds no fields and no comment anchors; `text` is what the
 * markup puts on the line, which is empty for everything except a transparent
 * wrapper such as `w:customXml`.
 */
export type PreservedInline = {
  type: "preservedInline";
  /** Replayable markup for one child, as `captureVerbatimXml` wrote it. */
  xml: string;
  /** The visible text the markup contributes, empty when it shows nothing. */
  text: string;
};

/**
 * Paragraph content types
 */
export type ParagraphContent =
  | Run
  | Hyperlink
  | BookmarkStart
  | BookmarkEnd
  | SimpleField
  | ComplexField
  | InlineSdt
  | CommentRangeStart
  | CommentRangeEnd
  | CommentReference
  | Insertion
  | Deletion
  | MoveFrom
  | MoveTo
  | MoveFromRangeStart
  | MoveFromRangeEnd
  | MoveToRangeStart
  | MoveToRangeEnd
  | InlineWrapper
  | MathEquation
  | PreservedInline;

/**
 * The kinds a paragraph-mark tracked change can be (ECMA-376 §17.13.5).
 *
 * Written as a child of `<w:pPr><w:rPr>` — `<w:ins/>` when the paragraph
 * break itself was inserted in track-changes mode (the user pressed Enter
 * mid-paragraph), `<w:del/>` when the paragraph break is pending deletion
 * (Backspace at paragraph start or Delete at paragraph end). The mark is
 * independent of the inline runs the paragraph carries.
 *
 * A relocated paragraph's break is `<w:moveFrom/>` at the source and
 * `<w:moveTo/>` at the destination. They resolve exactly as `del` and `ins`
 * do — a move is a deletion and an insertion that a reader is told belong
 * together — but they are not those kinds: writing `w:del` on a moved
 * paragraph's mark reports the relocation as a deletion as well.
 */
export const PARAGRAPH_MARK_CHANGE_KINDS = Object.freeze([
  "moveFrom",
  "moveTo",
  "ins",
  "del",
] as const);

/** One of {@link PARAGRAPH_MARK_CHANGE_KINDS}. */
export type ParagraphMarkChangeKind = (typeof PARAGRAPH_MARK_CHANGE_KINDS)[number];

/** A paragraph-mark tracked change, written under `<w:pPr><w:rPr>`. */
export type ParagraphMarkChange = {
  kind: ParagraphMarkChangeKind;
  info: TrackedChangeInfo;
};

export const REVIEW_CARRIERS = {
  TERMINAL_TABLE: "terminal-table",
} as const;

/** A Folio-private untracked paragraph used to resolve a terminal table deletion. */
export type ReviewCarrier = (typeof REVIEW_CARRIERS)[keyof typeof REVIEW_CARRIERS];

/** Paragraph (w:p) */
export type Paragraph = {
  type: "paragraph";
  /** Unique paragraph ID */
  paraId?: string;
  /** Text ID */
  textId?: string;
  /** Paragraph formatting */
  formatting?: ParagraphFormatting;
  /** Paragraph-level tracked property changes (w:pPrChange) */
  propertyChanges?: ParagraphPropertyChange[];
  /** Paragraph-mark insertion / deletion (w:pPr / w:rPr / w:ins | w:del) */
  pPrMark?: ParagraphMarkChange;
  /** Folio-private review-resolution carrier; ignored by standard OOXML consumers. */
  reviewCarrier?: ReviewCarrier;
  /** Paragraph content */
  content: ParagraphContent[];
  /** Computed list rendering (if this is a list item) */
  listRendering?: ListRendering;
  /** Word's cached layout says this paragraph started on a new rendered page. */
  renderedPageBreakBefore?: boolean;
  /** Section properties (if this paragraph ends a section) */
  sectionProperties?: SectionProperties;
  /**
   * Attributes `w:p` carried that this record has no field for.
   *
   * Word writes a revision-session id (`w:rsidR` and its family) on nearly
   * every paragraph, and a rebuild that dropped it rewrote the document's
   * revision history. The remainder follows the record: a paragraph the
   * editor creates from scratch has none, and when an edit splits one in two
   * the half that keeps the authored identity keeps the remainder, because a
   * revision id inherited from a neighbour is a claim about history nobody
   * made.
   */
  preservedAttributes?: PreservedAttribute[];
};

// ============================================================================
// HEADERS & FOOTERS
// ============================================================================

/**
 * Header/footer type
 */
export type HeaderFooterType = "default" | "first" | "even";

/**
 * Header or footer reference
 */
export type HeaderReference = {
  type: HeaderFooterType;
  rId: string;
};

export type FooterReference = {
  type: HeaderFooterType;
  rId: string;
};

/**
 * Header or footer content
 */
export type HeaderFooter = {
  type: "header" | "footer";
  /** Header/footer type */
  hdrFtrType: HeaderFooterType;
  /** Content (paragraphs, tables, block-level content controls). */
  content: BlockContent[];
  /**
   * Document watermark detected in this header part. Word emits
   * watermarks as VML or DrawingML behind-content shapes inside header
   * parts; the body paragraph that contains them is empty otherwise.
   * The modeled `Watermark` is exposed alongside `content` so callers
   * can render and edit it without walking raw runs.
   */
  watermark?: Watermark;
  /**
   * Verbatim XML of the paragraph(s) containing the source watermark
   * shape. Captured at parse time so an untouched DOCX serializes the
   * watermark byte-exact even though `runParser` does not surface VML /
   * DrawingML at the run level. Cleared (or rewritten) when callers
   * mutate the modeled watermark via the headless API.
   */
  rawWatermarkXml?: string;
  /**
   * Verbatim XML of the whole part, captured at parse time so an unedited
   * header or footer re-emits byte-identically on save (VML OLE wrappers,
   * smart tags, and other constructs the model cannot fully represent).
   * Replayed only while `verbatimFingerprint` still matches the modeled
   * fields; cleared on first edit.
   */
  verbatimXml?: string;
  /** Fingerprint of the modeled fields when `verbatimXml` was captured. */
  verbatimFingerprint?: string;
  /**
   * Index where the watermark paragraph sat among block-level siblings
   * in the source header. The serializer inserts the watermark (raw or
   * synthesized) at this position so a header that originally placed
   * the watermark after visible text round-trips with the same flow.
   * Undefined when no watermark was parsed or when callers built the
   * watermark programmatically — in that case the serializer emits it
   * at the top of the header (the same position Word's own UI uses).
   */
  watermarkBlockIndex?: number;
};

/**
 * Document watermark (MS Word's behind-content page decoration).
 */
export type Watermark = TextWatermark | PictureWatermark;

export type TextWatermark = {
  kind: "text";
  /** Visible string. Required. */
  text: string;
  /** Font family. Word's default is Calibri. */
  font?: string;
  /**
   * Hex color (`"C0C0C0"`), `"auto"`, or `undefined` for the producer
   * default. Word emits `#C0C0C0` (light gray) for text watermarks.
   */
  color?: string;
  /**
   * `true` = diagonal (Word default, -45°), `false` = horizontal.
   * Stored as a boolean since the only Word-supported rotations are
   * -45 and 0.
   */
  diagonal?: boolean;
  /**
   * Opacity 0..1. Word's interactive UI exposes a "transparency"
   * percentage; folio stores it as an opacity scalar for renderer
   * convenience. Default ~0.5.
   */
  opacity?: number;
};

export type PictureWatermark = {
  kind: "picture";
  /** Relationship id of the image part in `word/_rels/header*.xml.rels`. */
  imageRId: string;
  /**
   * Stable identity of the image the `imageRId` resolved to in the header it
   * was parsed from — an absolute package path for embedded media (e.g.
   * `word/media/image1.png`) or the URL for a linked image (see
   * {@link imageTargetExternal}). Relationship ids are scoped per header part
   * and commonly repeat, so propagating a watermark across headers rebinds
   * against this anchored target rather than the (ambiguous) source rId; each
   * target header's relationship is written relative to its own part location.
   */
  imageTarget?: string;
  /**
   * When true, {@link imageTarget} is an external (linked) URL written back
   * with `TargetMode="External"`, not an embedded package path.
   */
  imageTargetExternal?: boolean;
  /** Optional scale factor (1.0 = native, 0.5 = half-size). */
  scale?: number;
  /**
   * Display width in points from the VML shape, captured at parse so the
   * source aspect ratio survives a save that re-synthesizes the watermark
   * (Word stretches the image to the shape box, so a non-2:1 box distorts a
   * non-2:1 image). Absent for synthesized watermarks, which fall back to
   * Word's default box scaled by `scale`.
   */
  widthPt?: number;
  /** Display height in points from the VML shape. See {@link widthPt}. */
  heightPt?: number;
  /**
   * Whether Word's "washout" effect was applied (low contrast).
   * Default true — Word emits washout=true on every picture
   * watermark inserted via Insert → Watermark.
   */
  washout?: boolean;
};

// ============================================================================
// FOOTNOTES & ENDNOTES
// ============================================================================

/**
 * Footnote position
 */
export type FootnotePosition = "pageBottom" | "beneathText" | "sectEnd" | "docEnd";

/**
 * Endnote position
 */
export type EndnotePosition = "sectEnd" | "docEnd";

/**
 * Number restart type
 */
export type NoteNumberRestart = "continuous" | "eachSect" | "eachPage";

/**
 * Footnote properties
 */
export type FootnoteProperties = {
  position?: FootnotePosition;
  numFmt?: NumberFormat;
  numStart?: number;
  numRestart?: NoteNumberRestart;
};

/**
 * Endnote properties
 */
export type EndnoteProperties = {
  position?: EndnotePosition;
  numFmt?: NumberFormat;
  numStart?: number;
  numRestart?: NoteNumberRestart;
};

/**
 * Footnote (w:footnote)
 */
export type Footnote = {
  type: "footnote";
  /** Footnote ID */
  id: number;
  /** Special footnote type */
  noteType?: "normal" | "separator" | "continuationSeparator" | "continuationNotice";
  /**
   * Content. Note bodies may carry block-level `<w:sdt>` content
   * controls (citation slots, bound metadata fields) — preserved as
   * `BlockSdt` so the rest of folio's SDT round-trip + mutate APIs
   * work in notes the same as they do in the main body. Mirrors the
   * shape upstream eigenpal/docx-editor#678 fixed for the same case.
   */
  content: BlockContent[];
};

/**
 * Endnote (w:endnote)
 */
export type Endnote = {
  type: "endnote";
  /** Endnote ID */
  id: number;
  /** Special endnote type */
  noteType?: "normal" | "separator" | "continuationSeparator" | "continuationNotice";
  /**
   * Content. Like `Footnote.content`, may carry block-level `<w:sdt>`
   * preserved as `BlockSdt` so SDT round-trip works inside endnotes.
   */
  content: BlockContent[];
};

// ============================================================================
// SECTION PROPERTIES
// ============================================================================

/**
 * Page orientation
 */
export type PageOrientation = "portrait" | "landscape";

/**
 * Section start type
 */
export type SectionStart = "continuous" | "nextPage" | "oddPage" | "evenPage" | "nextColumn";

/**
 * Vertical alignment
 */
export type VerticalAlign = "top" | "center" | "both" | "bottom";

/**
 * Line number restart type
 */
export type LineNumberRestart = "continuous" | "newPage" | "newSection";

/**
 * Column definition
 */
export type Column = {
  /** Column width in twips */
  width?: number;
  /** Space after column in twips */
  space?: number;
};

/**
 * Section properties (w:sectPr)
 */
export type SectionTextDirection =
  | "lrTb"
  | "tbRl"
  | "btLr"
  | "lrTbV"
  | "tbRlV"
  | "tbLrV"
  | "tb"
  | "rl"
  | "lr"
  | "tbV"
  | "rlV"
  | "lrV";

export type SectionProperties = {
  // Page size
  /** Page width in twips */
  pageWidth?: number;
  /** Page height in twips */
  pageHeight?: number;
  /** Page orientation */
  orientation?: PageOrientation;

  // Margins
  /** Top margin in twips */
  marginTop?: number;
  /** Bottom margin in twips */
  marginBottom?: number;
  /** Left margin in twips */
  marginLeft?: number;
  /** Right margin in twips */
  marginRight?: number;
  /** Header distance from top in twips */
  headerDistance?: number;
  /** Footer distance from bottom in twips */
  footerDistance?: number;
  /** Gutter margin in twips */
  gutter?: number;

  // Columns
  /** Number of columns */
  columnCount?: number;
  /** Space between columns in twips */
  columnSpace?: number;
  /** Equal width columns */
  equalWidth?: boolean;
  /** Separator line between columns */
  separator?: boolean;
  /** Individual column definitions */
  columns?: Column[];

  // Section behavior
  /** Section start type */
  sectionStart?: SectionStart;
  /** Vertical alignment of text */
  verticalAlign?: VerticalAlign;
  /** Section text direction */
  textDirection?: SectionTextDirection;
  /** Right-to-left section */
  bidi?: boolean;

  // Headers and footers
  /** Header references */
  headerReferences?: HeaderReference[];
  /** Footer references */
  footerReferences?: FooterReference[];
  /** Different first page header/footer */
  titlePg?: boolean;
  /** Different odd/even page headers/footers */
  evenAndOddHeaders?: boolean;

  // Line numbers
  /** Line numbering settings */
  lineNumbers?: {
    start?: number;
    countBy?: number;
    distance?: number;
    restart?: LineNumberRestart;
  };

  // Page numbers
  /** Page numbering settings */
  pageNumbering?: {
    format?: NumberFormat;
    start?: number;
    chapterStyle?: number;
    chapterSeparator?: string;
  };

  // Page borders
  /** Page borders */
  pageBorders?: {
    top?: BorderSpec;
    bottom?: BorderSpec;
    left?: BorderSpec;
    right?: BorderSpec;
    /** Display setting */
    display?: "allPages" | "firstPage" | "notFirstPage";
    /** Offset from */
    offsetFrom?: "page" | "text";
    /** Z-order */
    zOrder?: "front" | "back";
  };

  // Background
  /** Page background */
  background?: {
    color?: ColorValue;
    themeColor?: ThemeColorSlot;
    themeTint?: string;
    themeShade?: string;
  };

  // Footnote/Endnote properties
  /** Footnote properties for this section */
  footnotePr?: FootnoteProperties;
  /** Number of footnote columns in this section (`w15:footnoteColumns`) */
  footnoteColumns?: number;
  /** Endnote properties for this section */
  endnotePr?: EndnoteProperties;

  // Document grid
  /** Document grid */
  docGrid?: {
    type?: "default" | "lines" | "linesAndChars" | "snapToChars";
    linePitch?: number;
    charSpace?: number;
  };

  // Paper source
  /** First page paper source */
  paperSrcFirst?: number;
  /** Other pages paper source */
  paperSrcOther?: number;

  // Section-level flags and relationships
  /** Protected forms in this section */
  formProtection?: boolean;
  /** Suppress endnotes in this section */
  noEndnote?: boolean;
  /** Use right-to-left gutter in this section */
  rtlGutter?: boolean;
  /** Relationship id for printer settings */
  printerSettingsRelationshipId?: string;

  /** Section-level tracked property changes (w:sectPrChange) */
  propertyChanges?: SectionPropertyChange[];
  /** The `w:sectPr` children no reader took a typed value from. */
  preserved?: PreservedMarkup;
  /**
   * Attributes `w:sectPr` carried that this record has no field for.
   *
   * `AG_SectPrAttributes` declares `w:rsidR`, `w:rsidDel`, `w:rsidRPr` and
   * `w:rsidSect`, on the body's section properties and on a paragraph's alike.
   * The record travels through the editor whole, so the remainder does too.
   */
  preservedAttributes?: PreservedAttribute[];
};

// ============================================================================
// SECTION & DOCUMENT BODY
// ============================================================================

/**
 * A block-level child folio does not model, kept where it stood.
 *
 * `w:body`, `w:tc`, `w:hdr`, `w:ftr`, an SDT's content and a footnote all
 * admit more than paragraphs, tables and content controls: `w:permStart` is
 * the whole of a document-protection range, `w:altChunk` is an entire imported
 * document, `m:oMathPara` is a display equation, and a comment or move range
 * may open between two blocks. Position is their meaning, so the capture is a
 * block in its own right rather than a field riding on a neighbour: it sits
 * between the same two siblings in the model, in the editor and in the saved
 * part, and nothing has to keep an index honest as the blocks around it move.
 *
 * Being opaque, it holds no text, no fields and no comment anchors; a walker
 * looking for any of those may skip it, and every walker that rebuilds block
 * content must write it back.
 */
export type PreservedBlock = {
  type: "preservedBlock";
  /** Replayable markup for one child, as `captureVerbatimXml` wrote it. */
  xml: string;
};

/**
 * Block-level content types
 */
export type BlockContent = Paragraph | Table | BlockSdt | PreservedBlock;

/** {@link BlockContent} minus the branch folio does not model inside a cell. */
export type TableCellBlock = Exclude<BlockContent, BlockSdt>;

/**
 * Section (implicit or explicit based on sectPr)
 */
export type Section = {
  /** Section properties */
  properties: SectionProperties;
  /** Content in this section */
  content: BlockContent[];
  /** Headers for this section */
  headers?: Map<HeaderFooterType, HeaderFooter>;
  /** Footers for this section */
  footers?: Map<HeaderFooterType, HeaderFooter>;
};

/**
 * Document body (w:body)
 */
export type DocumentBody = {
  /** All content (paragraphs, tables) */
  content: BlockContent[];
  /** Sections (derived from sectPr in paragraphs and final sectPr) */
  sections?: Section[];
  /** Final section properties (from body's sectPr) */
  finalSectionProperties?: SectionProperties;
  /** Comments from comments.xml */
  comments?: Comment[];
};
