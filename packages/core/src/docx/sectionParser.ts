/**
 * Section Properties Parser - Parse section properties (w:sectPr)
 *
 * Section properties define page layout and settings for a section of the document.
 * They appear in two places:
 * 1. Within a paragraph's properties (w:p/w:pPr/w:sectPr) - marks end of a section
 * 2. At the end of the document body (w:body/w:sectPr) - final section properties
 *
 * OOXML Reference:
 * - w:pgSz: Page size (width, height, orientation)
 * - w:pgMar: Page margins (top, bottom, left, right, header, footer, gutter)
 * - w:cols: Column definitions
 * - w:type: Section start type
 * - w:vAlign: Vertical alignment
 * - w:headerReference, w:footerReference: Header/footer references
 * - w:titlePg: Different first page
 * - w:lnNumType: Line numbering
 * - w:pgBorders: Page borders
 * - w:docGrid: Document grid
 * - w:footnotePr, w:endnotePr: Footnote/endnote properties
 */

import type {
  FooterReference,
  HeaderReference,
  SectionProperties,
  SectionPropertyChange,
  PageOrientation,
  SectionStart,
  VerticalAlign,
  LineNumberRestart,
  Column,
} from "../types/document";

import { attributeRemainder, NO_MODELLED_ATTRIBUTES } from "./attributeRemainder";
import {
  CAPTURE,
  type ChildHandlers,
  dispatchChildren,
  keptUnless,
  sequencePositions,
} from "./containerChildren";
import { parseHeaderReference, parseFooterReference } from "./headerFooterRefParser";
import type { ParseContext } from "./parseContext";
import { parseFootnoteProperties, parseEndnoteProperties } from "./notePropertiesParser";
import { NumberFormatSchema, ThemeColorSlotSchema, narrowEnum } from "./parserEnums";
import { parseBorderSpec } from "./borderParser";
import {
  findChild,
  findChildren,
  getAttribute,
  parseNumericAttribute,
  parseBooleanElement,
  parseOnOffAttribute,
} from "./xmlParser";
import type { XmlElement } from "./xmlParser";
import { parsePropertyChangeInfo } from "./trackedChangeInfo";
import { parseSectionReferenceHistory } from "./sectionReferenceHistory";

/**
 * Sanity cap on `w:cols/@w:num`. Word's column picker tops out well below
 * this; a hostile/corrupt value here would force the layout engine to
 * generate a proportional band per column.
 */
const MAX_SECTION_COLUMNS = 45;

/**
 * Sanity cap on `w:pgSz` `@w:w`/`@w:h`, in twips (~22in). Matches the
 * existing `w:defaultTabStop` cap (`settingsParser.ts`) and bounds the
 * ruler tick generators, which are sized off the page dimensions.
 */
const MAX_PAGE_DIMENSION_TWIPS = 31_680;

// ============================================================================
// HELPER PARSERS
// ============================================================================

/**
 * Parse page orientation
 */
function parseOrientation(orient: string | null): PageOrientation | undefined {
  switch (orient) {
    case "landscape":
      return "landscape";
    case "portrait":
      return "portrait";
    default:
      return undefined;
  }
}

/**
 * Parse section start type
 */
function parseSectionStart(type: string | null): SectionStart | undefined {
  switch (type) {
    case "continuous":
      return "continuous";
    case "nextPage":
      return "nextPage";
    case "oddPage":
      return "oddPage";
    case "evenPage":
      return "evenPage";
    case "nextColumn":
      return "nextColumn";
    default:
      return undefined;
  }
}

/**
 * Parse vertical alignment
 */
function parseVerticalAlign(align: string | null): VerticalAlign | undefined {
  switch (align) {
    case "top":
      return "top";
    case "center":
      return "center";
    case "both":
      return "both";
    case "bottom":
      return "bottom";
    default:
      return undefined;
  }
}

function parseTextDirection(val: string | null): SectionProperties["textDirection"] | undefined {
  switch (val) {
    case "lrTb":
    case "tbRl":
    case "btLr":
    case "lrTbV":
    case "tbRlV":
    case "tbLrV":
    case "tb":
    case "rl":
    case "lr":
    case "tbV":
    case "rlV":
    case "lrV":
      return val;
    default:
      return undefined;
  }
}

/**
 * Parse line number restart type
 */
function parseLineNumberRestart(restart: string | null): LineNumberRestart | undefined {
  switch (restart) {
    case "continuous":
      return "continuous";
    case "newPage":
      return "newPage";
    case "newSection":
      return "newSection";
    default:
      return undefined;
  }
}

// ============================================================================
// MAIN PARSER
// ============================================================================

/**
 * Parse section properties (w:sectPr)
 *
 * @param sectPr - The w:sectPr element
 * @returns SectionProperties object
 */
export function parseSectionProperties(
  sectPr: XmlElement | null,
  context?: ParseContext,
): SectionProperties {
  const props: SectionProperties = {};

  if (!sectPr) {
    return props;
  }

  const headerRefs: HeaderReference[] = [];
  const footerRefs: FooterReference[] = [];
  const propertyChanges: SectionPropertyChange[] = [];

  const handlers: ChildHandlers<"section-properties"> = {
    headerReference: (child) => {
      const ref = parseHeaderReference(child, context);
      if (ref) {
        headerRefs.push(ref);
      }
      return keptUnless(ref !== null);
    },
    footerReference: (child) => {
      const ref = parseFooterReference(child, context);
      if (ref) {
        footerRefs.push(ref);
      }
      return keptUnless(ref !== null);
    },
    footnotePr: (child) => {
      const footnotePr = parseFootnoteProperties(child);
      if (Object.keys(footnotePr).length > 0) {
        props.footnotePr = footnotePr;
        return undefined;
      }
      return CAPTURE;
    },
    endnotePr: (child) => {
      const endnotePr = parseEndnoteProperties(child);
      if (Object.keys(endnotePr).length > 0) {
        props.endnotePr = endnotePr;
        return undefined;
      }
      return CAPTURE;
    },
    type: (child) => {
      const sectionStart = parseSectionStart(getAttribute(child, "w", "val"));
      if (sectionStart) {
        props.sectionStart = sectionStart;
      }
      return keptUnless(sectionStart !== undefined);
    },
    pgSz: (child) => {
      const width = parseNumericAttribute(child, "w", "w");
      if (width !== undefined) {
        props.pageWidth = Math.min(width, MAX_PAGE_DIMENSION_TWIPS);
      }
      const height = parseNumericAttribute(child, "w", "h");
      if (height !== undefined) {
        props.pageHeight = Math.min(height, MAX_PAGE_DIMENSION_TWIPS);
      }
      const orientation = parseOrientation(getAttribute(child, "w", "orient"));
      if (orientation) {
        props.orientation = orientation;
      }
      return keptUnless(width !== undefined || height !== undefined || orientation !== undefined);
    },
    pgMar: (child) => {
      const margins = [
        ["top", "marginTop"],
        ["bottom", "marginBottom"],
        ["left", "marginLeft"],
        ["right", "marginRight"],
        ["header", "headerDistance"],
        ["footer", "footerDistance"],
        ["gutter", "gutter"],
      ] as const;
      let taken = false;
      for (const [attribute, field] of margins) {
        const value = parseNumericAttribute(child, "w", attribute);
        if (value !== undefined) {
          props[field] = value;
          taken = true;
        }
      }
      return keptUnless(taken);
    },
    paperSrc: (child) => {
      const first = parseNumericAttribute(child, "w", "first");
      if (first !== undefined) {
        props.paperSrcFirst = first;
      }
      const other = parseNumericAttribute(child, "w", "other");
      if (other !== undefined) {
        props.paperSrcOther = other;
      }
      return keptUnless(first !== undefined || other !== undefined);
    },
    pgBorders: (child) => keptUnless(readPageBorders(props, child, context)),
    lnNumType: (child) => {
      const lineNumbers: NonNullable<SectionProperties["lineNumbers"]> = {};
      for (const [attribute, field] of [
        ["start", "start"],
        ["countBy", "countBy"],
        ["distance", "distance"],
      ] as const) {
        const value = parseNumericAttribute(child, "w", attribute);
        if (value !== undefined) {
          lineNumbers[field] = value;
        }
      }
      const restart = parseLineNumberRestart(getAttribute(child, "w", "restart"));
      if (restart) {
        lineNumbers.restart = restart;
      }
      if (Object.keys(lineNumbers).length === 0) {
        return CAPTURE;
      }
      props.lineNumbers = lineNumbers;
      return undefined;
    },
    pgNumType: (child) => {
      const pageNumbering: NonNullable<SectionProperties["pageNumbering"]> = {};
      const format = narrowEnum(getAttribute(child, "w", "fmt"), NumberFormatSchema);
      if (format) {
        pageNumbering.format = format;
      }
      const start = parseNumericAttribute(child, "w", "start");
      if (start !== undefined) {
        pageNumbering.start = start;
      }
      const chapterStyle = parseNumericAttribute(child, "w", "chapStyle");
      if (chapterStyle !== undefined) {
        pageNumbering.chapterStyle = chapterStyle;
      }
      const chapterSeparator = getAttribute(child, "w", "chapSep");
      if (chapterSeparator) {
        pageNumbering.chapterSeparator = chapterSeparator;
      }
      if (Object.keys(pageNumbering).length === 0) {
        return CAPTURE;
      }
      props.pageNumbering = pageNumbering;
      return undefined;
    },
    cols: (child) => keptUnless(readColumns(props, child)),
    formProt: (child) => {
      props.formProtection = parseBooleanElement(child, "w", context);
    },
    vAlign: (child) => {
      const verticalAlign = parseVerticalAlign(getAttribute(child, "w", "val"));
      if (verticalAlign) {
        props.verticalAlign = verticalAlign;
      }
      return keptUnless(verticalAlign !== undefined);
    },
    noEndnote: (child) => {
      props.noEndnote = parseBooleanElement(child, "w", context);
    },
    titlePg: (child) => {
      props.titlePg = parseBooleanElement(child, "w", context);
    },
    textDirection: (child) => {
      const textDirection = parseTextDirection(getAttribute(child, "w", "val"));
      if (textDirection) {
        props.textDirection = textDirection;
      }
      return keptUnless(textDirection !== undefined);
    },
    bidi: (child) => {
      props.bidi = parseBooleanElement(child, "w", context);
    },
    rtlGutter: (child) => {
      props.rtlGutter = parseBooleanElement(child, "w", context);
    },
    docGrid: (child) => {
      const docGrid: NonNullable<SectionProperties["docGrid"]> = {};
      const gridType = getAttribute(child, "w", "type");
      if (
        gridType === "default" ||
        gridType === "lines" ||
        gridType === "linesAndChars" ||
        gridType === "snapToChars"
      ) {
        docGrid.type = gridType;
      }
      const linePitch = parseNumericAttribute(child, "w", "linePitch");
      if (linePitch !== undefined) {
        docGrid.linePitch = linePitch;
      }
      const charSpace = parseNumericAttribute(child, "w", "charSpace");
      if (charSpace !== undefined) {
        docGrid.charSpace = charSpace;
      }
      // Every `w:docGrid` attribute is optional, and the serializer writes an
      // attribute-less element back, so the empty record is the state.
      props.docGrid = docGrid;
    },
    printerSettings: (child) => {
      const relationshipId = getAttribute(child, "r", "id");
      if (relationshipId) {
        props.printerSettingsRelationshipId = relationshipId;
      }
      return keptUnless(Boolean(relationshipId));
    },
    // A revision, not a property: it is read into `propertyChanges` and the
    // serializer writes it back from there.
    sectPrChange: (child) => {
      const previousSectPr = findChild(child, "w", "sectPr");
      const change: SectionPropertyChange = {
        type: "sectionPropertyChange",
        info: parsePropertyChangeInfo(child),
      };
      if (previousSectPr) {
        change.previousProperties = parseSectionProperties(previousSectPr);
      }
      const previousReferences = parseSectionReferenceHistory(child);
      if (previousReferences !== undefined) {
        change.previousReferences = previousReferences;
      }
      propertyChanges.push(change);
    },
  };

  const preserved = dispatchChildren({
    element: sectPr,
    container: "section-properties",
    handlers,
    capturePosition: sequencePositions("section-properties", sectPr),
    // Three names the Transitional content model does not declare for a
    // section and folio reads anyway: `w:background` and `w:evenAndOddHeaders`
    // belong to the document and the settings part and are written here by
    // producers, and `w15:footnoteColumns` is a later revision's extension.
    // Naming them is a claim that folio reads them here; anything else in this
    // element still goes to the sink.
    undeclared: {
      background: (child) => keptUnless(readBackground(props, child)),
      evenAndOddHeaders: (child) => {
        props.evenAndOddHeaders = parseBooleanElement(child, "w", context);
      },
      footnoteColumns: (child) => {
        const columns = parseNumericAttribute(child, "w", "val");
        if (columns !== undefined) {
          props.footnoteColumns = columns;
        }
        return keptUnless(columns !== undefined);
      },
    },
  });
  if (preserved) {
    props.preserved = preserved;
  }

  if (headerRefs.length > 0) {
    props.headerReferences = headerRefs;
  }
  if (footerRefs.length > 0) {
    props.footerReferences = footerRefs;
  }
  if (propertyChanges.length > 0) {
    props.propertyChanges = propertyChanges;
  }

  // `AG_SectPrAttributes`'s four `w:rsid*` attributes, and anything else the
  // source put on the element: `serializeSectionProperties` writes a bare
  // `<w:sectPr>` start tag, so the record models none of them.
  const remainder = attributeRemainder({ element: sectPr, modelled: NO_MODELLED_ATTRIBUTES });
  if (remainder) {
    props.preservedAttributes = remainder;
  }

  return props;
}

/** `w:cols`, into the four scalars and the per-column list. Did it state anything? */
function readColumns(props: SectionProperties, cols: XmlElement): boolean {
  const num = parseNumericAttribute(cols, "w", "num");
  if (num !== undefined) {
    props.columnCount = Math.min(num, MAX_SECTION_COLUMNS);
  }

  const space = parseNumericAttribute(cols, "w", "space");
  if (space !== undefined) {
    props.columnSpace = space;
  }

  const equalWidth = parseOnOffAttribute(cols, "w", "equalWidth");
  if (equalWidth !== undefined) {
    props.equalWidth = equalWidth;
  }

  const separator = parseOnOffAttribute(cols, "w", "sep");
  if (separator !== undefined) {
    props.separator = separator;
  }

  const colElements = findChildren(cols, "w", "col");
  if (colElements.length > 0) {
    props.columns = colElements.map((colEl): Column => {
      const column: Column = {};
      const colWidth = parseNumericAttribute(colEl, "w", "w");
      if (colWidth !== undefined) {
        column.width = colWidth;
      }
      const colSpace = parseNumericAttribute(colEl, "w", "space");
      if (colSpace !== undefined) {
        column.space = colSpace;
      }
      return column;
    });

    // Infer column count from w:col entries when w:num is absent
    if (props.columnCount === undefined) {
      props.columnCount = colElements.length;
    }
  }

  return (
    num !== undefined ||
    space !== undefined ||
    equalWidth !== undefined ||
    separator !== undefined ||
    colElements.length > 0
  );
}

/** `w:pgBorders`, into the four sides and the three placement attributes. */
function readPageBorders(
  props: SectionProperties,
  pgBorders: XmlElement,
  context: ParseContext | undefined,
): boolean {
  const pageBorders: NonNullable<SectionProperties["pageBorders"]> = {};

  for (const side of ["top", "bottom", "left", "right"] as const) {
    const border = parseBorderSpec(findChild(pgBorders, "w", side), context);
    if (border) {
      pageBorders[side] = border;
    }
  }

  const display = getAttribute(pgBorders, "w", "display");
  if (display === "allPages" || display === "firstPage" || display === "notFirstPage") {
    pageBorders.display = display;
  }

  const offsetFrom = getAttribute(pgBorders, "w", "offsetFrom");
  if (offsetFrom === "page" || offsetFrom === "text") {
    pageBorders.offsetFrom = offsetFrom;
  }

  const zOrder = getAttribute(pgBorders, "w", "zOrder");
  if (zOrder === "front" || zOrder === "back") {
    pageBorders.zOrder = zOrder;
  }

  if (Object.keys(pageBorders).length === 0) {
    return false;
  }
  props.pageBorders = pageBorders;
  return true;
}

/** `w:background`, the page colour a producer wrote on the section. */
function readBackground(props: SectionProperties, background: XmlElement): boolean {
  const pageBackground: NonNullable<SectionProperties["background"]> = {};

  const colorVal = getAttribute(background, "w", "color");
  if (colorVal && colorVal !== "auto") {
    pageBackground.color = { rgb: colorVal };
  }

  const backgroundThemeColor = narrowEnum(
    getAttribute(background, "w", "themeColor"),
    ThemeColorSlotSchema,
  );
  if (backgroundThemeColor) {
    pageBackground.themeColor = backgroundThemeColor;
  }

  const themeTint = getAttribute(background, "w", "themeTint");
  if (themeTint) {
    pageBackground.themeTint = themeTint;
  }

  const themeShade = getAttribute(background, "w", "themeShade");
  if (themeShade) {
    pageBackground.themeShade = themeShade;
  }

  if (Object.keys(pageBackground).length === 0) {
    return false;
  }
  props.background = pageBackground;
  return true;
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Get page width in pixels (96 DPI)
 *
 * @param props - Section properties
 * @param defaultWidth - Default width in twips (default: 12240 = 8.5 inches)
 * @returns Width in pixels
 */
export function getPageWidthPixels(
  props: SectionProperties,
  defaultWidth: number = 12_240,
): number {
  const twips = props.pageWidth ?? defaultWidth;
  // 1 inch = 1440 twips, 1 inch = 96 pixels at 96 DPI
  return Math.round((twips / 1440) * 96);
}

/**
 * Get page height in pixels (96 DPI)
 *
 * @param props - Section properties
 * @param defaultHeight - Default height in twips (default: 15840 = 11 inches)
 * @returns Height in pixels
 */
export function getPageHeightPixels(
  props: SectionProperties,
  defaultHeight: number = 15_840,
): number {
  const twips = props.pageHeight ?? defaultHeight;
  return Math.round((twips / 1440) * 96);
}

/**
 * Get content width (page width minus margins) in pixels
 *
 * @param props - Section properties
 * @returns Content width in pixels
 */
export function getContentWidthPixels(props: SectionProperties): number {
  const pageWidth = props.pageWidth ?? 12_240;
  const marginLeft = props.marginLeft ?? 1440; // 1 inch default
  const marginRight = props.marginRight ?? 1440;
  const twips = pageWidth - marginLeft - marginRight;
  return Math.round((twips / 1440) * 96);
}

/**
 * Get content height (page height minus margins) in pixels
 *
 * @param props - Section properties
 * @returns Content height in pixels
 */
export function getContentHeightPixels(props: SectionProperties): number {
  const pageHeight = props.pageHeight ?? 15_840;
  const marginTop = props.marginTop ?? 1440;
  const marginBottom = props.marginBottom ?? 1440;
  const twips = pageHeight - marginTop - marginBottom;
  return Math.round((twips / 1440) * 96);
}

/**
 * Get margins in pixels
 *
 * @param props - Section properties
 * @returns Object with all margins in pixels
 */
export function getMarginsPixels(props: SectionProperties): {
  top: number;
  bottom: number;
  left: number;
  right: number;
  header: number;
  footer: number;
  gutter: number;
} {
  const twipsToPixels = (twips: number | undefined, defaultTwips: number) =>
    Math.round(((twips ?? defaultTwips) / 1440) * 96);

  return {
    top: twipsToPixels(props.marginTop, 1440),
    bottom: twipsToPixels(props.marginBottom, 1440),
    left: twipsToPixels(props.marginLeft, 1440),
    right: twipsToPixels(props.marginRight, 1440),
    header: twipsToPixels(props.headerDistance, 720), // 0.5 inch default
    footer: twipsToPixels(props.footerDistance, 720),
    gutter: twipsToPixels(props.gutter, 0),
  };
}

/**
 * Check if section has different first page header/footer
 */
export function hasDifferentFirstPage(props: SectionProperties): boolean {
  return props.titlePg === true;
}

/**
 * Check if section has different odd/even page headers/footers
 */
export function hasDifferentOddEven(props: SectionProperties): boolean {
  return props.evenAndOddHeaders === true;
}

/**
 * Get effective column count (minimum 1)
 */
export function getColumnCount(props: SectionProperties): number {
  return Math.max(1, props.columnCount ?? 1);
}

/**
 * Check if section is landscape
 */
export function isLandscape(props: SectionProperties): boolean {
  return props.orientation === "landscape";
}

/**
 * Check if section has page borders
 */
export function hasPageBorders(props: SectionProperties): boolean {
  if (!props.pageBorders) {
    return false;
  }
  return !!(
    props.pageBorders.top ||
    props.pageBorders.bottom ||
    props.pageBorders.left ||
    props.pageBorders.right
  );
}

/**
 * Check if section has line numbers
 */
export function hasLineNumbers(props: SectionProperties): boolean {
  return !!props.lineNumbers;
}

/**
 * Get default section properties (US Letter size, 1 inch margins)
 */
export function getDefaultSectionProperties(): SectionProperties {
  return {
    pageWidth: 12_240, // 8.5 inches
    pageHeight: 15_840, // 11 inches
    orientation: "portrait",
    marginTop: 1440, // 1 inch
    marginBottom: 1440,
    marginLeft: 1440,
    marginRight: 1440,
    headerDistance: 720, // 0.5 inch
    footerDistance: 720,
    gutter: 0,
    columnCount: 1,
    columnSpace: 720, // 0.5 inch
    equalWidth: true,
    sectionStart: "nextPage",
    verticalAlign: "top",
  };
}

/**
 * Merge section properties (later values override earlier)
 *
 * @param base - Base properties
 * @param override - Override properties
 * @returns Merged properties
 */
export function mergeSectionProperties(
  base: SectionProperties,
  override: SectionProperties,
): SectionProperties {
  const result: SectionProperties = { ...base };

  // Simple properties - override if present
  if (override.pageWidth !== undefined) {
    result.pageWidth = override.pageWidth;
  }
  if (override.pageHeight !== undefined) {
    result.pageHeight = override.pageHeight;
  }
  if (override.orientation !== undefined) {
    result.orientation = override.orientation;
  }
  if (override.marginTop !== undefined) {
    result.marginTop = override.marginTop;
  }
  if (override.marginBottom !== undefined) {
    result.marginBottom = override.marginBottom;
  }
  if (override.marginLeft !== undefined) {
    result.marginLeft = override.marginLeft;
  }
  if (override.marginRight !== undefined) {
    result.marginRight = override.marginRight;
  }
  if (override.headerDistance !== undefined) {
    result.headerDistance = override.headerDistance;
  }
  if (override.footerDistance !== undefined) {
    result.footerDistance = override.footerDistance;
  }
  if (override.gutter !== undefined) {
    result.gutter = override.gutter;
  }
  if (override.columnCount !== undefined) {
    result.columnCount = override.columnCount;
  }
  if (override.columnSpace !== undefined) {
    result.columnSpace = override.columnSpace;
  }
  if (override.equalWidth !== undefined) {
    result.equalWidth = override.equalWidth;
  }
  if (override.separator !== undefined) {
    result.separator = override.separator;
  }
  if (override.columns !== undefined) {
    result.columns = override.columns;
  }
  if (override.sectionStart !== undefined) {
    result.sectionStart = override.sectionStart;
  }
  if (override.verticalAlign !== undefined) {
    result.verticalAlign = override.verticalAlign;
  }
  if (override.textDirection !== undefined) {
    result.textDirection = override.textDirection;
  }
  if (override.bidi !== undefined) {
    result.bidi = override.bidi;
  }
  if (override.headerReferences !== undefined) {
    result.headerReferences = override.headerReferences;
  }
  if (override.footerReferences !== undefined) {
    result.footerReferences = override.footerReferences;
  }
  if (override.titlePg !== undefined) {
    result.titlePg = override.titlePg;
  }
  if (override.evenAndOddHeaders !== undefined) {
    result.evenAndOddHeaders = override.evenAndOddHeaders;
  }
  if (override.lineNumbers !== undefined) {
    result.lineNumbers = override.lineNumbers;
  }
  if (override.pageNumbering !== undefined) {
    result.pageNumbering = override.pageNumbering;
  }
  if (override.pageBorders !== undefined) {
    result.pageBorders = override.pageBorders;
  }
  if (override.background !== undefined) {
    result.background = override.background;
  }
  if (override.footnotePr !== undefined) {
    result.footnotePr = override.footnotePr;
  }
  if (override.endnotePr !== undefined) {
    result.endnotePr = override.endnotePr;
  }
  if (override.docGrid !== undefined) {
    result.docGrid = override.docGrid;
  }
  if (override.paperSrcFirst !== undefined) {
    result.paperSrcFirst = override.paperSrcFirst;
  }
  if (override.paperSrcOther !== undefined) {
    result.paperSrcOther = override.paperSrcOther;
  }
  if (override.formProtection !== undefined) {
    result.formProtection = override.formProtection;
  }
  if (override.noEndnote !== undefined) {
    result.noEndnote = override.noEndnote;
  }
  if (override.rtlGutter !== undefined) {
    result.rtlGutter = override.rtlGutter;
  }
  if (override.printerSettingsRelationshipId !== undefined) {
    result.printerSettingsRelationshipId = override.printerSettingsRelationshipId;
  }

  return result;
}
