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
  ParagraphFormatting,
  TableFormatting,
  TableRowFormatting,
  TableCellFormatting,
  TabStop,
  TableBorders,
  TableCellBorders,
  CellMargins,
  TableMeasurement,
} from "../types/document";
import { resolveDefaultParagraphStyle } from "./defaultParagraphStyle";
import { parseTableLook } from "./tableParser";
import { mergeParagraphFormatting } from "../utils/paragraphFormattingMerge";
import { mergeStyleTextFormatting } from "../utils/textFormattingMerge";
import {
  ConditionalStyleTypeSchema,
  LineSpacingRuleSchema,
  ParagraphAlignmentSchema,
  StyleTypeSchema,
  TableAlignmentSchema,
  TableRowHeightRuleSchema,
  TableWidthTypeSchema,
  TabLeaderSchema,
  TabStopAlignmentSchema,
  TextDirectionSchema,
  narrowEnum,
} from "./parserEnums";
import { parseRunProperties, RUN_PROPERTY_OWNERS } from "./runParser";
import { parseShading } from "./shadingParser";
import { parseBorderSpec } from "./borderParser";
import {
  parseXmlDocument,
  findChild,
  findChildren,
  getAttribute,
  getLocalName,
  parseBooleanElement,
  parseNumberingLevelAttribute,
  parseNumericAttribute,
  parseOnOffValue,
  parseTableMeasurementValue,
  parseOnOffAttribute,
} from "./xmlParser";
import type { XmlElement } from "./xmlParser";
import { findChildAnySpelling, numericAttributeAnySpelling } from "./strictNames";

/**
 * Style map keyed by styleId
 */
export type StyleMap = Map<string, Style>;

export type ParsedStylesPackage = {
  styleDefinitions: StyleDefinitions;
  styles: StyleMap;
};

/**
 * Parse tab stops (w:tabs)
 */
function parseTabStops(tabs: XmlElement | null): TabStop[] | undefined {
  if (!tabs) {
    return undefined;
  }

  const tabElements = findChildren(tabs, "w", "tab");
  if (tabElements.length === 0) {
    return undefined;
  }

  const result: TabStop[] = [];

  for (const tab of tabElements) {
    const pos = parseNumericAttribute(tab, "w", "pos");
    const alignment = narrowEnum(getAttribute(tab, "w", "val"), TabStopAlignmentSchema);

    if (pos !== undefined && alignment) {
      const tabStop: TabStop = {
        position: pos,
        alignment,
      };

      const leader = narrowEnum(getAttribute(tab, "w", "leader"), TabLeaderSchema);
      if (leader && leader !== "none") {
        tabStop.leader = leader;
      }

      result.push(tabStop);
    }
  }

  return result.length > 0 ? result : undefined;
}

/**
 * Parse paragraph formatting properties (w:pPr)
 */
function parseParagraphProperties(
  pPr: XmlElement | null,
  theme: Theme | null,
): ParagraphFormatting | undefined {
  if (!pPr) {
    return undefined;
  }

  const formatting: ParagraphFormatting = {};

  // Alignment
  const jc = findChild(pPr, "w", "jc");
  if (jc) {
    const val = narrowEnum(getAttribute(jc, "w", "val"), ParagraphAlignmentSchema);
    if (val) {
      formatting.alignment = val;
    }
  }

  // Bidi
  const bidi = findChild(pPr, "w", "bidi");
  if (bidi) {
    formatting.bidi = parseBooleanElement(bidi);
  }

  const snapToGrid = findChild(pPr, "w", "snapToGrid");
  if (snapToGrid) {
    formatting.snapToGrid = parseBooleanElement(snapToGrid);
  }

  // Spacing
  const spacing = findChild(pPr, "w", "spacing");
  if (spacing) {
    const before = parseNumericAttribute(spacing, "w", "before");
    if (before !== undefined) {
      formatting.spaceBefore = before;
    }

    const after = parseNumericAttribute(spacing, "w", "after");
    if (after !== undefined) {
      formatting.spaceAfter = after;
    }

    const line = parseNumericAttribute(spacing, "w", "line");
    if (line !== undefined) {
      formatting.lineSpacing = line;
    }

    const lineRule = narrowEnum(getAttribute(spacing, "w", "lineRule"), LineSpacingRuleSchema);
    if (lineRule) {
      formatting.lineSpacingRule = lineRule;
    }

    const beforeAutospacing = parseOnOffAttribute(spacing, "w", "beforeAutospacing");
    if (beforeAutospacing !== undefined) {
      formatting.beforeAutospacing = beforeAutospacing;
    }

    const afterAutospacing = parseOnOffAttribute(spacing, "w", "afterAutospacing");
    if (afterAutospacing !== undefined) {
      formatting.afterAutospacing = afterAutospacing;
    }
  }

  // Indentation
  const ind = findChild(pPr, "w", "ind");
  if (ind) {
    const left = numericAttributeAnySpelling(ind, "CT_Ind @left");
    if (left !== undefined) {
      formatting.indentLeft = left;
    }

    const right = numericAttributeAnySpelling(ind, "CT_Ind @right");
    if (right !== undefined) {
      formatting.indentRight = right;
    }

    const firstLine = parseNumericAttribute(ind, "w", "firstLine");
    if (firstLine !== undefined) {
      formatting.indentFirstLine = firstLine;
    }

    const hanging = parseNumericAttribute(ind, "w", "hanging");
    if (hanging !== undefined) {
      formatting.indentFirstLine = -hanging;
      formatting.hangingIndent = true;
    }
  }

  // Borders
  const pBdr = findChild(pPr, "w", "pBdr");
  if (pBdr) {
    const borders: ParagraphFormatting["borders"] = {};
    const top = parseBorderSpec(findChild(pBdr, "w", "top"));
    if (top) {
      borders.top = top;
    }
    const bottom = parseBorderSpec(findChild(pBdr, "w", "bottom"));
    if (bottom) {
      borders.bottom = bottom;
    }
    const left = parseBorderSpec(findChild(pBdr, "w", "left"));
    if (left) {
      borders.left = left;
    }
    const right = parseBorderSpec(findChild(pBdr, "w", "right"));
    if (right) {
      borders.right = right;
    }
    const between = parseBorderSpec(findChild(pBdr, "w", "between"));
    if (between) {
      borders.between = between;
    }
    const bar = parseBorderSpec(findChild(pBdr, "w", "bar"));
    if (bar) {
      borders.bar = bar;
    }

    if (Object.keys(borders).length > 0) {
      formatting.borders = borders;
    }
  }

  // Shading
  const shd = findChild(pPr, "w", "shd");
  if (shd) {
    const shadingResult = parseShading(shd);
    if (shadingResult) {
      formatting.shading = shadingResult;
    }
  }

  // Tab stops
  const tabs = findChild(pPr, "w", "tabs");
  if (tabs) {
    const tabStops = parseTabStops(tabs);
    if (tabStops) {
      formatting.tabs = tabStops;
    }
  }

  // Page break control
  const keepNext = findChild(pPr, "w", "keepNext");
  if (keepNext) {
    formatting.keepNext = parseBooleanElement(keepNext);
  }

  const keepLines = findChild(pPr, "w", "keepLines");
  if (keepLines) {
    formatting.keepLines = parseBooleanElement(keepLines);
  }

  const widowControl = findChild(pPr, "w", "widowControl");
  if (widowControl) {
    formatting.widowControl = parseBooleanElement(widowControl);
  }

  const pageBreakBefore = findChild(pPr, "w", "pageBreakBefore");
  if (pageBreakBefore) {
    formatting.pageBreakBefore = parseBooleanElement(pageBreakBefore);
  }

  const contextualSpacing = findChild(pPr, "w", "contextualSpacing");
  if (contextualSpacing) {
    formatting.contextualSpacing = parseBooleanElement(contextualSpacing);
  }

  // Numbering properties
  const numPr = findChild(pPr, "w", "numPr");
  if (numPr) {
    const numId = findChild(numPr, "w", "numId");
    const ilvl = findChild(numPr, "w", "ilvl");

    if (numId || ilvl) {
      formatting.numPr = {};
      if (numId) {
        const val = parseNumericAttribute(numId, "w", "val");
        if (val !== undefined) {
          formatting.numPr.numId = val;
        }
      }
      if (ilvl) {
        const val = parseNumberingLevelAttribute(ilvl);
        if (val !== undefined) {
          formatting.numPr.ilvl = val;
        }
      }
    }
  }

  // Outline level
  const outlineLvl = findChild(pPr, "w", "outlineLvl");
  if (outlineLvl) {
    const val = parseNumericAttribute(outlineLvl, "w", "val");
    if (val !== undefined) {
      formatting.outlineLevel = val;
    }
  }

  // Style reference
  const pStyle = findChild(pPr, "w", "pStyle");
  if (pStyle) {
    const val = getAttribute(pStyle, "w", "val");
    if (val) {
      formatting.styleId = val;
    }
  }

  // Suppress line numbers
  const suppressLineNumbers = findChild(pPr, "w", "suppressLineNumbers");
  if (suppressLineNumbers) {
    formatting.suppressLineNumbers = parseBooleanElement(suppressLineNumbers);
  }

  // Suppress auto hyphens
  const suppressAutoHyphens = findChild(pPr, "w", "suppressAutoHyphens");
  if (suppressAutoHyphens) {
    formatting.suppressAutoHyphens = parseBooleanElement(suppressAutoHyphens);
  }

  const kinsoku = findChild(pPr, "w", "kinsoku");
  if (kinsoku) {
    formatting.kinsoku = parseBooleanElement(kinsoku);
  }

  const overflowPunct = findChild(pPr, "w", "overflowPunct");
  if (overflowPunct) {
    formatting.overflowPunctuation = parseBooleanElement(overflowPunct);
  }

  // Run properties for this paragraph (default run formatting)
  const rPr = findChild(pPr, "w", "rPr");
  if (rPr) {
    const runProps = parseRunProperties(rPr, theme, RUN_PROPERTY_OWNERS.standalone);
    if (runProps) {
      formatting.runProperties = runProps;
    }
  }

  return Object.keys(formatting).length > 0 ? formatting : undefined;
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

  const left = parseBorderSpec(findChildAnySpelling(tblBorders, "CT_Border left"));
  if (left) {
    borders.left = left;
  }

  const right = parseBorderSpec(findChildAnySpelling(tblBorders, "CT_Border right"));
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

  const left = parseTableMeasurement(findChildAnySpelling(tblCellMar, "CT_TblWidth left"));
  if (left) {
    margins.left = left;
  }

  const right = parseTableMeasurement(findChildAnySpelling(tblCellMar, "CT_TblWidth right"));
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

  // Table placement (w:jc), narrowed against `ST_JcTable`.
  const justification = narrowEnum(
    getAttribute(findChild(tblPr, "w", "jc"), "w", "val"),
    TableAlignmentSchema,
  );
  if (justification) {
    formatting.justification = justification;
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

  // Row placement (w:jc), the same `ST_JcTable` the table's own carries.
  const justification = narrowEnum(
    getAttribute(findChild(trPr, "w", "jc"), "w", "val"),
    TableAlignmentSchema,
  );
  if (justification) {
    formatting.justification = justification;
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
    const val = narrowEnum(getAttribute(textDirection, "w", "val"), TextDirectionSchema);
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
    const rPrResult = parseRunProperties(rPr, theme, RUN_PROPERTY_OWNERS.standalone);
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
          const condRPrResult = parseRunProperties(condRPr, theme, RUN_PROPERTY_OWNERS.standalone);
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
      const rPrResult = parseRunProperties(rPr, theme, RUN_PROPERTY_OWNERS.standalone);
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
