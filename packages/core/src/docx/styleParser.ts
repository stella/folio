/**
 * Style Parser - Parse styles.xml with full inheritance resolution
 *
 * Parses all style types (paragraph, character, table, list) with
 * complete basedOn inheritance chain resolution.
 *
 * OOXML Reference:
 * - Style file is at: word/styles.xml
 * - Uses WordprocessingML namespace (w:)
 *
 * Style Cascade (lowest to highest priority):
 * 1. Document defaults (w:docDefaults)
 * 2. Parent style properties (w:basedOn chain)
 * 3. Current style properties
 * 4. Direct formatting in document
 */

import type {
  Theme,
  Style,
  StyleType,
  StyleDefinitions,
  DocDefaults,
  TextFormatting,
  TableFormatting,
  TableRowFormatting,
  TableCellFormatting,
  ColorValue,
  TableBorders,
  TableCellBorders,
  CellMargins,
  TableMeasurement,
} from "../types/document";
import { resolveDefaultParagraphStyle } from "./defaultParagraphStyle";
// The one `w:pPr` reader. A style's property set is the same set a paragraph
// carries, and the private copy that used to live here read fourteen of its
// children and let the rest fall off the end — including `w:framePr`, the
// Strict `w:ind` spellings and every child the sink now keeps.
import { parseParagraphProperties } from "./paragraphParser";
import { parseTableLook } from "./tableParser";
import { mergeParagraphFormatting } from "../utils/paragraphFormattingMerge";
import { mergeStyleTextFormatting } from "../utils/textFormattingMerge";
import { parseHorizontalScalePercent } from "../utils/horizontalScale";
import {
  ConditionalStyleTypeSchema,
  EmphasisMarkSchema,
  FontHintSchema,
  FontThemeSchema,
  HighlightColorSchema,
  StyleTypeSchema,
  TableCellTextDirectionSchema,
  TableRowHeightRuleSchema,
  TableWidthTypeSchema,
  TextEffectSchema,
  ThemeColorSlotSchema,
  UnderlineStyleSchema,
  narrowEnum,
} from "./parserEnums";
import { resolveThemeFontRef } from "./themeParser";
import { parseShading } from "./shadingParser";
import { parseBorderSpec } from "./borderParser";
import {
  parseXmlDocument,
  findChild,
  findChildren,
  getAttribute,
  getLocalName,
  parseBooleanElement,
  parseNumericAttribute,
  parseOnOffValue,
  parseTableMeasurementValue,
} from "./xmlParser";
import type { XmlElement } from "./xmlParser";

/**
 * Style map keyed by styleId
 */
export type StyleMap = Map<string, Style>;

export type ParsedStylesPackage = {
  styleDefinitions: StyleDefinitions;
  styles: StyleMap;
};

const findLastRunToggle = (rPr: XmlElement, localName: string): XmlElement | null =>
  findChildren(rPr, "w", localName).at(-1) ?? null;

/**
 * Parse text formatting properties (w:rPr)
 */
function parseRunProperties(
  rPr: XmlElement | null,
  theme: Theme | null,
): TextFormatting | undefined {
  if (!rPr) {
    return undefined;
  }

  const formatting: TextFormatting = {};

  // Bold
  const b = findLastRunToggle(rPr, "b");
  if (b) {
    formatting.bold = parseBooleanElement(b);
  }

  const bCs = findLastRunToggle(rPr, "bCs");
  if (bCs) {
    formatting.boldCs = parseBooleanElement(bCs);
  }

  // Italic
  const i = findLastRunToggle(rPr, "i");
  if (i) {
    formatting.italic = parseBooleanElement(i);
  }

  const iCs = findLastRunToggle(rPr, "iCs");
  if (iCs) {
    formatting.italicCs = parseBooleanElement(iCs);
  }

  // Underline
  const u = findChild(rPr, "w", "u");
  if (u) {
    const style = narrowEnum(getAttribute(u, "w", "val"), UnderlineStyleSchema);
    if (style) {
      formatting.underline = { style };
      const colorVal = getAttribute(u, "w", "color");
      const themeColor = getAttribute(u, "w", "themeColor");
      if (colorVal || themeColor) {
        formatting.underline.color = parseColorValue(
          colorVal,
          themeColor,
          getAttribute(u, "w", "themeTint"),
          getAttribute(u, "w", "themeShade"),
        );
      }
    }
  }

  // Strikethrough
  const strike = findLastRunToggle(rPr, "strike");
  if (strike) {
    formatting.strike = parseBooleanElement(strike);
  }

  const dstrike = findChild(rPr, "w", "dstrike");
  if (dstrike) {
    formatting.doubleStrike = parseBooleanElement(dstrike);
  }

  // Vertical alignment (superscript/subscript)
  const vertAlign = findChild(rPr, "w", "vertAlign");
  if (vertAlign) {
    const val = getAttribute(vertAlign, "w", "val");
    if (val === "superscript" || val === "subscript" || val === "baseline") {
      formatting.vertAlign = val;
    }
  }

  // Capitalization
  const smallCaps = findLastRunToggle(rPr, "smallCaps");
  if (smallCaps) {
    formatting.smallCaps = parseBooleanElement(smallCaps);
  }

  const caps = findLastRunToggle(rPr, "caps");
  if (caps) {
    formatting.allCaps = parseBooleanElement(caps);
  }

  // Hidden
  const vanish = findLastRunToggle(rPr, "vanish");
  if (vanish) {
    formatting.hidden = parseBooleanElement(vanish);
  }

  // Color
  const color = findChild(rPr, "w", "color");
  if (color) {
    formatting.color = parseColorValue(
      getAttribute(color, "w", "val"),
      getAttribute(color, "w", "themeColor"),
      getAttribute(color, "w", "themeTint"),
      getAttribute(color, "w", "themeShade"),
    );
  }

  // Highlight
  const highlight = findChild(rPr, "w", "highlight");
  if (highlight) {
    const val = narrowEnum(getAttribute(highlight, "w", "val"), HighlightColorSchema);
    if (val) {
      formatting.highlight = val;
    }
  }

  // Character shading
  const shd = findChild(rPr, "w", "shd");
  if (shd) {
    const shadingResult = parseShading(shd);
    if (shadingResult) {
      formatting.shading = shadingResult;
    }
  }

  // Font size (in half-points)
  const sz = findChild(rPr, "w", "sz");
  if (sz) {
    const val = parseNumericAttribute(sz, "w", "val");
    if (val !== undefined) {
      formatting.fontSize = val;
    }
  }

  const szCs = findChild(rPr, "w", "szCs");
  if (szCs) {
    const val = parseNumericAttribute(szCs, "w", "val");
    if (val !== undefined) {
      formatting.fontSizeCs = val;
    }
  }

  // Font family
  const rFonts = findChild(rPr, "w", "rFonts");
  if (rFonts) {
    const fontFamily: NonNullable<TextFormatting["fontFamily"]> = {};
    const ascii = getAttribute(rFonts, "w", "ascii");
    if (ascii) {
      fontFamily.ascii = ascii;
    }
    const hAnsi = getAttribute(rFonts, "w", "hAnsi");
    if (hAnsi) {
      fontFamily.hAnsi = hAnsi;
    }
    const eastAsia = getAttribute(rFonts, "w", "eastAsia");
    if (eastAsia) {
      fontFamily.eastAsia = eastAsia;
    }
    const csFont = getAttribute(rFonts, "w", "cs");
    if (csFont) {
      fontFamily.cs = csFont;
    }
    const hint = narrowEnum(getAttribute(rFonts, "w", "hint"), FontHintSchema);
    if (hint) {
      fontFamily.hint = hint;
    }

    // Theme font references - resolve to actual font names
    const asciiThemeRaw = getAttribute(rFonts, "w", "asciiTheme");
    const asciiTheme = narrowEnum(asciiThemeRaw, FontThemeSchema);
    if (asciiTheme) {
      fontFamily.asciiTheme = asciiTheme;
      // Also resolve the actual font name for convenience
      if (theme && !fontFamily.ascii) {
        const resolved = resolveThemeFontRef(theme, asciiTheme);
        if (resolved) {
          fontFamily.ascii = resolved;
        }
      }
    }
    const hAnsiTheme = narrowEnum(getAttribute(rFonts, "w", "hAnsiTheme"), FontThemeSchema);
    if (hAnsiTheme) {
      fontFamily.hAnsiTheme = hAnsiTheme;
      if (theme && !fontFamily.hAnsi) {
        const resolved = resolveThemeFontRef(theme, hAnsiTheme);
        if (resolved) {
          fontFamily.hAnsi = resolved;
        }
      }
    }
    const eastAsiaTheme = narrowEnum(getAttribute(rFonts, "w", "eastAsiaTheme"), FontThemeSchema);
    if (eastAsiaTheme) {
      fontFamily.eastAsiaTheme = eastAsiaTheme;
      if (theme && !fontFamily.eastAsia) {
        const resolved = resolveThemeFontRef(theme, eastAsiaTheme);
        if (resolved) {
          fontFamily.eastAsia = resolved;
        }
      }
    }
    const csTheme = narrowEnum(getAttribute(rFonts, "w", "cstheme"), FontThemeSchema);
    if (csTheme) {
      fontFamily.csTheme = csTheme;
      if (theme && !fontFamily.cs) {
        const resolved = resolveThemeFontRef(theme, csTheme);
        if (resolved) {
          fontFamily.cs = resolved;
        }
      }
    }

    formatting.fontFamily = fontFamily;
  }

  const lang = findChild(rPr, "w", "lang");
  if (lang) {
    const val = getAttribute(lang, "w", "val") || undefined;
    const eastAsia = getAttribute(lang, "w", "eastAsia") || undefined;
    const bidi = getAttribute(lang, "w", "bidi") || undefined;
    if (val || eastAsia || bidi) {
      formatting.language = {
        ...(val ? { val } : {}),
        ...(eastAsia ? { eastAsia } : {}),
        ...(bidi ? { bidi } : {}),
      };
    }
  }

  // Character spacing (in twips)
  const spacing = findChild(rPr, "w", "spacing");
  if (spacing) {
    const val = parseNumericAttribute(spacing, "w", "val");
    if (val !== undefined) {
      formatting.spacing = val;
    }
  }

  // Position (raised/lowered in half-points)
  const position = findChild(rPr, "w", "position");
  if (position) {
    const val = parseNumericAttribute(position, "w", "val");
    if (val !== undefined) {
      formatting.position = val;
    }
  }

  // Scale (horizontal text scale percentage)
  const w = findChild(rPr, "w", "w");
  if (w) {
    const val = parseHorizontalScalePercent(getAttribute(w, "w", "val"));
    if (val !== undefined) {
      formatting.scale = val;
    }
  }

  // Kerning
  const kern = findChild(rPr, "w", "kern");
  if (kern) {
    const val = parseNumericAttribute(kern, "w", "val");
    if (val !== undefined) {
      formatting.kerning = val;
    }
  }

  // Text effects
  const effect = findChild(rPr, "w", "effect");
  if (effect) {
    const val = narrowEnum(getAttribute(effect, "w", "val"), TextEffectSchema);
    if (val) {
      formatting.effect = val;
    }
  }

  // Emphasis mark
  const em = findChild(rPr, "w", "em");
  if (em) {
    const val = narrowEnum(getAttribute(em, "w", "val"), EmphasisMarkSchema);
    if (val) {
      formatting.emphasisMark = val;
    }
  }

  // Other effects
  const emboss = findLastRunToggle(rPr, "emboss");
  if (emboss) {
    formatting.emboss = parseBooleanElement(emboss);
  }

  const imprint = findLastRunToggle(rPr, "imprint");
  if (imprint) {
    formatting.imprint = parseBooleanElement(imprint);
  }

  const outline = findLastRunToggle(rPr, "outline");
  if (outline) {
    formatting.outline = parseBooleanElement(outline);
  }

  const shadow = findLastRunToggle(rPr, "shadow");
  if (shadow) {
    formatting.shadow = parseBooleanElement(shadow);
  }

  // RTL and complex script
  const rtl = findChild(rPr, "w", "rtl");
  if (rtl) {
    formatting.rtl = parseBooleanElement(rtl);
  }

  const cs = findChild(rPr, "w", "cs");
  if (cs) {
    formatting.cs = parseBooleanElement(cs);
  }

  // Character style reference
  const rStyle = findChild(rPr, "w", "rStyle");
  if (rStyle) {
    const val = getAttribute(rStyle, "w", "val");
    if (val) {
      formatting.styleId = val;
    }
  }

  return Object.keys(formatting).length > 0 ? formatting : undefined;
}

/**
 * Parse color value from attributes
 */
function parseColorValue(
  rgb: string | null,
  themeColor: string | null,
  themeTint: string | null,
  themeShade: string | null,
): ColorValue {
  const color: ColorValue = {};

  if (rgb && rgb !== "auto") {
    color.rgb = rgb;
  } else if (rgb === "auto") {
    color.auto = true;
  }

  const validatedThemeColor = narrowEnum(themeColor, ThemeColorSlotSchema);
  if (validatedThemeColor) {
    color.themeColor = validatedThemeColor;
  }

  if (themeTint) {
    color.themeTint = themeTint;
  }

  if (themeShade) {
    color.themeShade = themeShade;
  }

  return color;
}

/**
 * Parse table measurement (width/height with type)
 */
function parseTableMeasurement(element: XmlElement | null): TableMeasurement | undefined {
  if (!element) {
    return undefined;
  }

  const rawType = getAttribute(element, "w", "type");
  const type = rawType === null ? "dxa" : narrowEnum(rawType, TableWidthTypeSchema);
  const w = type ? parseTableMeasurementValue(element, type) : undefined;

  if (w !== undefined && type) {
    return { value: w, type };
  }

  return undefined;
}

/**
 * Parse table borders
 */
function parseTableBorders(tblBorders: XmlElement | null): TableBorders | undefined {
  if (!tblBorders) {
    return undefined;
  }

  const borders: TableBorders = {};

  const top = parseBorderSpec(findChild(tblBorders, "w", "top"));
  if (top) {
    borders.top = top;
  }

  const bottom = parseBorderSpec(findChild(tblBorders, "w", "bottom"));
  if (bottom) {
    borders.bottom = bottom;
  }

  const left = parseBorderSpec(
    findChild(tblBorders, "w", "left") ?? findChild(tblBorders, "w", "start"),
  );
  if (left) {
    borders.left = left;
  }

  const right = parseBorderSpec(
    findChild(tblBorders, "w", "right") ?? findChild(tblBorders, "w", "end"),
  );
  if (right) {
    borders.right = right;
  }

  const insideH = parseBorderSpec(findChild(tblBorders, "w", "insideH"));
  if (insideH) {
    borders.insideH = insideH;
  }

  const insideV = parseBorderSpec(findChild(tblBorders, "w", "insideV"));
  if (insideV) {
    borders.insideV = insideV;
  }

  return Object.keys(borders).length > 0 ? borders : undefined;
}

/** Parse cell-only diagonal borders in addition to the shared table sides. */
function parseTableCellBorders(tcBorders: XmlElement | null): TableCellBorders | undefined {
  if (!tcBorders) {
    return undefined;
  }

  const borders: TableCellBorders = { ...parseTableBorders(tcBorders) };
  const topLeftToBottomRight = parseBorderSpec(findChild(tcBorders, "w", "tl2br"));
  if (topLeftToBottomRight) {
    borders.topLeftToBottomRight = topLeftToBottomRight;
  }

  const topRightToBottomLeft = parseBorderSpec(findChild(tcBorders, "w", "tr2bl"));
  if (topRightToBottomLeft) {
    borders.topRightToBottomLeft = topRightToBottomLeft;
  }

  return Object.keys(borders).length > 0 ? borders : undefined;
}

/**
 * Parse cell margins
 */
function parseCellMargins(tblCellMar: XmlElement | null): CellMargins | undefined {
  if (!tblCellMar) {
    return undefined;
  }

  const margins: CellMargins = {};

  const top = parseTableMeasurement(findChild(tblCellMar, "w", "top"));
  if (top) {
    margins.top = top;
  }

  const bottom = parseTableMeasurement(findChild(tblCellMar, "w", "bottom"));
  if (bottom) {
    margins.bottom = bottom;
  }

  const left = parseTableMeasurement(findChild(tblCellMar, "w", "left"));
  if (left) {
    margins.left = left;
  }

  const right = parseTableMeasurement(findChild(tblCellMar, "w", "right"));
  if (right) {
    margins.right = right;
  }

  return Object.keys(margins).length > 0 ? margins : undefined;
}

/**
 * Parse table formatting properties (w:tblPr)
 */
function parseTableProperties(
  tblPr: XmlElement | null,
  _theme: Theme | null,
): TableFormatting | undefined {
  if (!tblPr) {
    return undefined;
  }

  const formatting: TableFormatting = {};

  // Table width
  const tblW = findChild(tblPr, "w", "tblW");
  if (tblW) {
    const widthResult = parseTableMeasurement(tblW);
    if (widthResult) {
      formatting.width = widthResult;
    }
  }

  // Table alignment/justification
  const jc = findChild(tblPr, "w", "jc");
  if (jc) {
    const val = getAttribute(jc, "w", "val");
    if (val === "left" || val === "center" || val === "right") {
      formatting.justification = val;
    }
  }

  // Cell spacing
  const tblCellSpacing = findChild(tblPr, "w", "tblCellSpacing");
  if (tblCellSpacing) {
    const cellSpacingResult = parseTableMeasurement(tblCellSpacing);
    if (cellSpacingResult) {
      formatting.cellSpacing = cellSpacingResult;
    }
  }

  // Table indent
  const tblInd = findChild(tblPr, "w", "tblInd");
  if (tblInd) {
    const indentResult = parseTableMeasurement(tblInd);
    if (indentResult) {
      formatting.indent = indentResult;
    }
  }

  // Table borders
  const tblBorders = findChild(tblPr, "w", "tblBorders");
  if (tblBorders) {
    const bordersResult = parseTableBorders(tblBorders);
    if (bordersResult) {
      formatting.borders = bordersResult;
    }
  }

  // Cell margins
  const tblCellMar = findChild(tblPr, "w", "tblCellMar");
  if (tblCellMar) {
    const marginsResult = parseCellMargins(tblCellMar);
    if (marginsResult) {
      formatting.cellMargins = marginsResult;
    }
  }

  // Table layout
  const tblLayout = findChild(tblPr, "w", "tblLayout");
  if (tblLayout) {
    const val = getAttribute(tblLayout, "w", "type");
    if (val === "fixed" || val === "autofit") {
      formatting.layout = val;
    }
  }

  // Table style
  const tblStyle = findChild(tblPr, "w", "tblStyle");
  if (tblStyle) {
    const val = getAttribute(tblStyle, "w", "val");
    if (val) {
      formatting.styleId = val;
    }
  }

  // Table look
  const tblLook = findChild(tblPr, "w", "tblLook");
  if (tblLook) {
    const lookResult = parseTableLook(tblLook);
    if (lookResult) {
      formatting.look = lookResult;
    }
  }

  // Shading
  const shd = findChild(tblPr, "w", "shd");
  if (shd) {
    const shadingResult = parseShading(shd);
    if (shadingResult) {
      formatting.shading = shadingResult;
    }
  }

  // Bidi
  const bidiVisual = findChild(tblPr, "w", "bidiVisual");
  if (bidiVisual) {
    formatting.bidi = parseBooleanElement(bidiVisual);
  }

  return Object.keys(formatting).length > 0 ? formatting : undefined;
}

/**
 * Parse table row formatting properties (w:trPr)
 */
function parseTableRowProperties(trPr: XmlElement | null): TableRowFormatting | undefined {
  if (!trPr) {
    return undefined;
  }

  const formatting: TableRowFormatting = {};

  // Row height
  const trHeight = findChild(trPr, "w", "trHeight");
  if (trHeight) {
    const heightResult = parseTableMeasurement(trHeight);
    if (heightResult) {
      formatting.height = heightResult;
    }
    const hRule = narrowEnum(getAttribute(trHeight, "w", "hRule"), TableRowHeightRuleSchema);
    if (hRule) {
      formatting.heightRule = hRule;
    }
  }

  // Header row
  const tblHeader = findChild(trPr, "w", "tblHeader");
  if (tblHeader) {
    formatting.header = parseBooleanElement(tblHeader);
  }

  // Can't split
  const cantSplit = findChild(trPr, "w", "cantSplit");
  if (cantSplit) {
    formatting.cantSplit = parseBooleanElement(cantSplit);
  }

  // Row justification
  const jc = findChild(trPr, "w", "jc");
  if (jc) {
    const val = getAttribute(jc, "w", "val");
    if (val === "left" || val === "center" || val === "right") {
      formatting.justification = val;
    }
  }

  // Hidden
  const hidden = findChild(trPr, "w", "hidden");
  if (hidden) {
    formatting.hidden = parseBooleanElement(hidden);
  }

  return Object.keys(formatting).length > 0 ? formatting : undefined;
}

/**
 * Parse table cell formatting properties (w:tcPr)
 */
function parseTableCellProperties(
  tcPr: XmlElement | null,
  _theme: Theme | null,
): TableCellFormatting | undefined {
  if (!tcPr) {
    return undefined;
  }

  const formatting: TableCellFormatting = {};

  // Cell width
  const tcW = findChild(tcPr, "w", "tcW");
  if (tcW) {
    const widthResult = parseTableMeasurement(tcW);
    if (widthResult) {
      formatting.width = widthResult;
    }
  }

  // Cell borders
  const tcBorders = findChild(tcPr, "w", "tcBorders");
  if (tcBorders) {
    const bordersResult = parseTableCellBorders(tcBorders);
    if (bordersResult) {
      formatting.borders = bordersResult;
    }
  }

  // Cell margins
  const tcMar = findChild(tcPr, "w", "tcMar");
  if (tcMar) {
    const marginsResult = parseCellMargins(tcMar);
    if (marginsResult) {
      formatting.margins = marginsResult;
    }
  }

  // Shading
  const shd = findChild(tcPr, "w", "shd");
  if (shd) {
    const shadingResult = parseShading(shd);
    if (shadingResult) {
      formatting.shading = shadingResult;
    }
  }

  // Vertical alignment
  const vAlign = findChild(tcPr, "w", "vAlign");
  if (vAlign) {
    const val = getAttribute(vAlign, "w", "val");
    if (val === "top" || val === "center" || val === "bottom") {
      formatting.verticalAlign = val;
    }
  }

  // Text direction
  const textDirection = findChild(tcPr, "w", "textDirection");
  if (textDirection) {
    const val = narrowEnum(getAttribute(textDirection, "w", "val"), TableCellTextDirectionSchema);
    if (val) {
      formatting.textDirection = val;
    }
  }

  // Grid span (horizontal merge)
  const gridSpan = findChild(tcPr, "w", "gridSpan");
  if (gridSpan) {
    const val = parseNumericAttribute(gridSpan, "w", "val");
    if (val !== undefined) {
      formatting.gridSpan = val;
    }
  }

  // Vertical merge
  const vMerge = findChild(tcPr, "w", "vMerge");
  if (vMerge) {
    const val = getAttribute(vMerge, "w", "val");
    formatting.vMerge = val === "restart" ? "restart" : "continue";
  }

  // Fit text
  const tcFitText = findChild(tcPr, "w", "tcFitText");
  if (tcFitText) {
    formatting.fitText = parseBooleanElement(tcFitText);
  }

  // No wrap
  const noWrap = findChild(tcPr, "w", "noWrap");
  if (noWrap) {
    formatting.noWrap = parseBooleanElement(noWrap);
  }

  // Hide mark
  const hideMark = findChild(tcPr, "w", "hideMark");
  if (hideMark) {
    formatting.hideMark = parseBooleanElement(hideMark);
  }

  return Object.keys(formatting).length > 0 ? formatting : undefined;
}

type StyleChildren = {
  basedOn?: XmlElement;
  hidden?: XmlElement;
  link?: XmlElement;
  name?: XmlElement;
  next?: XmlElement;
  personal?: XmlElement;
  pPr?: XmlElement;
  qFormat?: XmlElement;
  rPr?: XmlElement;
  semiHidden?: XmlElement;
  tblPr?: XmlElement;
  tblStylePrs: XmlElement[];
  tcPr?: XmlElement;
  trPr?: XmlElement;
  uiPriority?: XmlElement;
  unhideWhenUsed?: XmlElement;
};

function collectStyleChildren(styleEl: XmlElement): StyleChildren {
  const children: StyleChildren = { tblStylePrs: [] };

  for (const child of styleEl.elements ?? []) {
    if (child.type !== "element") {
      continue;
    }

    switch (getLocalName(child.name || "")) {
      case "basedOn":
        children.basedOn ??= child;
        break;
      case "hidden":
        children.hidden ??= child;
        break;
      case "link":
        children.link ??= child;
        break;
      case "name":
        children.name ??= child;
        break;
      case "next":
        children.next ??= child;
        break;
      case "personal":
        children.personal ??= child;
        break;
      case "pPr":
        children.pPr ??= child;
        break;
      case "qFormat":
        children.qFormat ??= child;
        break;
      case "rPr":
        children.rPr ??= child;
        break;
      case "semiHidden":
        children.semiHidden ??= child;
        break;
      case "tblPr":
        children.tblPr ??= child;
        break;
      case "tblStylePr":
        children.tblStylePrs.push(child);
        break;
      case "tcPr":
        children.tcPr ??= child;
        break;
      case "trPr":
        children.trPr ??= child;
        break;
      case "uiPriority":
        children.uiPriority ??= child;
        break;
      case "unhideWhenUsed":
        children.unhideWhenUsed ??= child;
        break;
    }
  }

  return children;
}

/**
 * Parse a single style element (w:style)
 */
function parseStyle(styleEl: XmlElement, theme: Theme | null): Style {
  const rawType = getAttribute(styleEl, "w", "type");
  const style: Style = {
    styleId: getAttribute(styleEl, "w", "styleId") ?? "",
    type: narrowEnum(rawType, StyleTypeSchema) ?? "paragraph",
  };

  // Default flag
  const defaultAttr = getAttribute(styleEl, "w", "default");
  if (defaultAttr) {
    style.default = parseOnOffValue(defaultAttr) ?? false;
  }

  const children = collectStyleChildren(styleEl);

  // Name
  const nameEl = children.name;
  if (nameEl) {
    const nameVal = getAttribute(nameEl, "w", "val");
    if (nameVal) {
      style.name = nameVal;
    }
  }

  // Based on (inheritance)
  const basedOn = children.basedOn;
  if (basedOn) {
    const basedOnVal = getAttribute(basedOn, "w", "val");
    if (basedOnVal) {
      style.basedOn = basedOnVal;
    }
  }

  // Next style
  const next = children.next;
  if (next) {
    const nextVal = getAttribute(next, "w", "val");
    if (nextVal) {
      style.next = nextVal;
    }
  }

  // Linked style
  const link = children.link;
  if (link) {
    const linkVal = getAttribute(link, "w", "val");
    if (linkVal) {
      style.link = linkVal;
    }
  }

  // UI Priority
  const uiPriority = children.uiPriority;
  if (uiPriority) {
    const val = parseNumericAttribute(uiPriority, "w", "val");
    if (val !== undefined) {
      style.uiPriority = val;
    }
  }

  // Hidden/Semi-hidden
  const hidden = children.hidden;
  if (hidden) {
    style.hidden = parseBooleanElement(hidden);
  }

  const semiHidden = children.semiHidden;
  if (semiHidden) {
    style.semiHidden = parseBooleanElement(semiHidden);
  }

  // Unhide when used
  const unhideWhenUsed = children.unhideWhenUsed;
  if (unhideWhenUsed) {
    style.unhideWhenUsed = parseBooleanElement(unhideWhenUsed);
  }

  // Quick format
  const qFormat = children.qFormat;
  if (qFormat) {
    style.qFormat = parseBooleanElement(qFormat);
  }

  // Personal/custom style
  const personal = children.personal;
  if (personal) {
    style.personal = parseBooleanElement(personal);
  }

  // Paragraph properties
  const pPr = children.pPr;
  if (pPr) {
    const pPrResult = parseParagraphProperties(pPr, theme);
    if (pPrResult) {
      style.pPr = pPrResult;
    }
  }

  // Run properties
  const rPr = children.rPr;
  if (rPr) {
    const rPrResult = parseRunProperties(rPr, theme);
    if (rPrResult) {
      style.rPr = rPrResult;
    }
  }

  // Table properties (for table styles)
  const tblPr = children.tblPr;
  if (tblPr) {
    const tblPrResult = parseTableProperties(tblPr, theme);
    if (tblPrResult) {
      style.tblPr = tblPrResult;
    }
  }

  // Table row properties
  const trPr = children.trPr;
  if (trPr) {
    const trPrResult = parseTableRowProperties(trPr);
    if (trPrResult) {
      style.trPr = trPrResult;
    }
  }

  // Table cell properties
  const tcPr = children.tcPr;
  if (tcPr) {
    const tcPrResult = parseTableCellProperties(tcPr, theme);
    if (tcPrResult) {
      style.tcPr = tcPrResult;
    }
  }

  // Table style conditional formatting (tblStylePr)
  const tblStylePrs = children.tblStylePrs;
  if (tblStylePrs.length > 0) {
    style.tblStylePr = [];

    for (const tblStylePr of tblStylePrs) {
      const conditionalType = narrowEnum(
        getAttribute(tblStylePr, "w", "type"),
        ConditionalStyleTypeSchema,
      );
      if (conditionalType) {
        const conditionalStyle: NonNullable<Style["tblStylePr"]>[number] = {
          type: conditionalType,
        };

        const condPPr = findChild(tblStylePr, "w", "pPr");
        if (condPPr) {
          const condPPrResult = parseParagraphProperties(condPPr, theme);
          if (condPPrResult) {
            conditionalStyle.pPr = condPPrResult;
          }
        }

        const condRPr = findChild(tblStylePr, "w", "rPr");
        if (condRPr) {
          const condRPrResult = parseRunProperties(condRPr, theme);
          if (condRPrResult) {
            conditionalStyle.rPr = condRPrResult;
          }
        }

        const condTblPr = findChild(tblStylePr, "w", "tblPr");
        if (condTblPr) {
          const condTblPrResult = parseTableProperties(condTblPr, theme);
          if (condTblPrResult) {
            conditionalStyle.tblPr = condTblPrResult;
          }
        }

        const condTrPr = findChild(tblStylePr, "w", "trPr");
        if (condTrPr) {
          const condTrPrResult = parseTableRowProperties(condTrPr);
          if (condTrPrResult) {
            conditionalStyle.trPr = condTrPrResult;
          }
        }

        const condTcPr = findChild(tblStylePr, "w", "tcPr");
        if (condTcPr) {
          const condTcPrResult = parseTableCellProperties(condTcPr, theme);
          if (condTcPrResult) {
            conditionalStyle.tcPr = condTcPrResult;
          }
        }

        style.tblStylePr.push(conditionalStyle);
      }
    }
  }

  return style;
}

/**
 * Parse document defaults (w:docDefaults)
 */
function parseDocDefaults(
  docDefaults: XmlElement | null,
  theme: Theme | null,
): DocDefaults | undefined {
  if (!docDefaults) {
    return undefined;
  }

  const result: DocDefaults = {};

  // Default run properties
  const rPrDefault = findChild(docDefaults, "w", "rPrDefault");
  if (rPrDefault) {
    const rPr = findChild(rPrDefault, "w", "rPr");
    if (rPr) {
      const rPrResult = parseRunProperties(rPr, theme);
      if (rPrResult) {
        result.rPr = rPrResult;
      }
    }
  }

  // Default paragraph properties
  const pPrDefault = findChild(docDefaults, "w", "pPrDefault");
  if (pPrDefault) {
    const pPr = findChild(pPrDefault, "w", "pPr");
    if (pPr) {
      const pPrResult = parseParagraphProperties(pPr, theme);
      if (pPrResult) {
        result.pPr = pPrResult;
      }
    }
  }

  // Return the (possibly empty) defaults whenever the `<w:docDefaults>` element
  // is present, so callers can distinguish "document declares empty defaults"
  // (zero spacing / single line) from "no docDefaults at all". The early guard
  // above already returns undefined for an absent element
  // (eigenpal/docx-editor#909).
  return result;
}

/**
 * Resolve style inheritance chain
 */
function resolveStyleInheritance(
  style: Style,
  styleMap: StyleMap,
  visited = new Set<string>(),
): Style {
  // Prevent circular inheritance
  if (visited.has(style.styleId)) {
    return style;
  }
  visited.add(style.styleId);

  // If no basedOn, return as-is
  if (!style.basedOn) {
    return style;
  }

  // Get parent style
  const parentStyle = styleMap.get(style.basedOn);
  if (!parentStyle) {
    return style;
  }

  // Recursively resolve parent
  const resolvedParent = resolveStyleInheritance(parentStyle, styleMap, visited);

  // Merge parent into this style (this style overrides parent)
  const resolved: Style = { ...style };

  const mergedPPr = mergeParagraphFormatting(resolvedParent.pPr, style.pPr);
  if (mergedPPr) {
    resolved.pPr = mergedPPr;
  }

  const mergedRPr = mergeStyleTextFormatting(resolvedParent.rPr, style.rPr);
  if (mergedRPr) {
    resolved.rPr = mergedRPr;
  }

  // Merge table properties if this is a table style
  if (style.type === "table") {
    if (resolvedParent.tblPr || style.tblPr) {
      resolved.tblPr = {
        ...resolvedParent.tblPr,
        ...style.tblPr,
      };
    }
    if (resolvedParent.trPr || style.trPr) {
      resolved.trPr = { ...resolvedParent.trPr, ...style.trPr };
    }
    if (resolvedParent.tcPr || style.tcPr) {
      resolved.tcPr = { ...resolvedParent.tcPr, ...style.tcPr };
    }
  }

  return resolved;
}

/**
 * Parse styles.xml content
 *
 * @param stylesXml - XML content of styles.xml
 * @param theme - Parsed theme for resolving theme references
 * @returns StyleMap with resolved inheritance
 */
export function parseStyles(stylesXml: string, theme: Theme | null): StyleMap {
  const doc = parseXmlDocument(stylesXml);
  if (!doc) {
    return new Map();
  }

  return parseStylesFromDocument(doc, theme);
}

function parseStylesFromDocument(doc: XmlElement, theme: Theme | null): StyleMap {
  const styleMap: StyleMap = new Map();

  try {
    // First pass: parse all styles without inheritance resolution
    const styleElements = findChildren(doc, "w", "style");
    for (const styleEl of styleElements) {
      const style = parseStyle(styleEl, theme);
      if (style.styleId) {
        styleMap.set(style.styleId, style);
      }
    }

    // Second pass: resolve inheritance
    for (const [styleId, style] of styleMap) {
      const resolved = resolveStyleInheritance(style, styleMap);
      styleMap.set(styleId, resolved);
    }
  } catch {
    // Malformed style inheritance leaves unresolved styles in place.
  }

  return styleMap;
}

/**
 * Parse complete style definitions including docDefaults
 *
 * @param stylesXml - XML content of styles.xml
 * @param theme - Parsed theme for resolving theme references
 * @returns StyleDefinitions with docDefaults and resolved styles
 */
export function parseStyleDefinitions(
  stylesXml: string,
  theme: Theme | null,
  resolvedStyles?: StyleMap,
): StyleDefinitions {
  const doc = parseXmlDocument(stylesXml);
  if (!doc) {
    return { styles: [] };
  }

  return parseStyleDefinitionsFromDocument(doc, theme, resolvedStyles);
}

function parseStyleDefinitionsFromDocument(
  doc: XmlElement,
  theme: Theme | null,
  resolvedStyles?: StyleMap,
): StyleDefinitions {
  const result: StyleDefinitions = {
    styles: [],
  };

  try {
    // Parse document defaults
    const docDefaultsEl = findChild(doc, "w", "docDefaults");
    const parsedDocDefaults = parseDocDefaults(docDefaultsEl, theme);
    if (parsedDocDefaults) {
      result.docDefaults = parsedDocDefaults;
    }

    // Parse latent styles
    const latentStylesEl = findChild(doc, "w", "latentStyles");
    if (latentStylesEl) {
      const latentStyles: NonNullable<StyleDefinitions["latentStyles"]> = {
        defLockedState:
          parseOnOffValue(getAttribute(latentStylesEl, "w", "defLockedState")) ?? false,
        defSemiHidden: parseOnOffValue(getAttribute(latentStylesEl, "w", "defSemiHidden")) ?? false,
        defUnhideWhenUsed:
          parseOnOffValue(getAttribute(latentStylesEl, "w", "defUnhideWhenUsed")) ?? false,
        defQFormat: parseOnOffValue(getAttribute(latentStylesEl, "w", "defQFormat")) ?? false,
      };
      const defUIPriority = parseNumericAttribute(latentStylesEl, "w", "defUIPriority");
      if (defUIPriority !== undefined) {
        latentStyles.defUIPriority = defUIPriority;
      }
      const count = parseNumericAttribute(latentStylesEl, "w", "count");
      if (count !== undefined) {
        latentStyles.count = count;
      }
      result.latentStyles = latentStyles;
    }

    // Parse styles with full inheritance resolution
    const styleMap = resolvedStyles ?? parseStylesFromDocument(doc, theme);
    result.styles = Array.from(styleMap.values());
  } catch {
    // Malformed styles return the partial definitions parsed so far.
  }

  return result;
}

export function parseStylesPackage(stylesXml: string, theme: Theme | null): ParsedStylesPackage {
  const doc = parseXmlDocument(stylesXml);
  if (!doc) {
    return {
      styleDefinitions: { styles: [] },
      styles: new Map(),
    };
  }

  const styles = parseStylesFromDocument(doc, theme);
  return {
    styleDefinitions: parseStyleDefinitionsFromDocument(doc, theme, styles),
    styles,
  };
}

/**
 * Get the resolved properties for a style
 *
 * @param styleId - Style ID to look up
 * @param styleMap - Style map from parseStyles
 * @returns Resolved style or undefined
 */
export function getResolvedStyle(styleId: string, styleMap: StyleMap): Style | undefined {
  return styleMap.get(styleId);
}

/**
 * Get the default paragraph style
 */
export function getDefaultParagraphStyle(styleMap: StyleMap): Style | undefined {
  return resolveDefaultParagraphStyle(styleMap.values());
}

/**
 * Get the default character style
 */
export function getDefaultCharacterStyle(styleMap: StyleMap): Style | undefined {
  for (const style of styleMap.values()) {
    if (style.type === "character" && style.default) {
      return style;
    }
  }
  return undefined;
}

/**
 * Get all styles of a specific type
 */
export function getStylesByType(styleMap: StyleMap, type: StyleType): Style[] {
  const result: Style[] = [];
  for (const style of styleMap.values()) {
    if (style.type === type) {
      result.push(style);
    }
  }
  return result;
}
