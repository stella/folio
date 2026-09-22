/**
 * Table Parser - Parse tables with full OOXML structure
 *
 * OOXML tables consist of:
 * - w:tbl - Table element
 * - w:tblPr - Table properties (width, borders, style)
 * - w:tblGrid - Column width definitions
 * - w:tr - Table rows
 * - w:trPr - Row properties (height, header)
 * - w:tc - Table cells
 * - w:tcPr - Cell properties (width, borders, merge)
 *
 * Cell merging:
 * - Horizontal: w:gridSpan (how many grid columns this cell spans)
 * - Vertical: w:vMerge (restart = start of merge, continue = continuation)
 *
 * OOXML Reference:
 * - w:tbl contains w:tblPr, w:tblGrid, and w:tr elements
 * - w:tr contains w:trPr and w:tc elements
 * - w:tc contains w:tcPr and content (paragraphs, tables)
 */

import type {
  Table,
  TablePreservedMarkup,
  TableRow,
  TableCell,
  TableFormatting,
  TablePropertyExceptionFormatting,
  TableGridChange,
  TableRowFormatting,
  TableCellFormatting,
  TablePropertyChange,
  TablePropertyExceptionChange,
  TableRowPropertyChange,
  TableCellPropertyChange,
  TableStructuralChangeInfo,
  TableMeasurement,
  TableWidthType,
  TableBorders,
  TableCellBorders,
  TableLook,
  CellMargins,
  FloatingTableProperties,
  ConditionalFormatStyle,
  Paragraph,
  PositionedBookmarkMarker,
  TableCellBlock,
  PreservedBlock,
  PreservedChild,
  SdtProperties,
  Theme,
  RelationshipMap,
  MediaFile,
} from "../types/document";
import { attributeRemainder, NO_MODELLED_ATTRIBUTES } from "./attributeRemainder";
import { parseBookmarkEnd, parseBookmarkStart } from "./bookmarkParser";
import { TABLE_LOOK_FLAGS } from "./tableLook";
import {
  CAPTURE,
  type ChildHandlers,
  dispatchChildren,
  keptUnless,
  ownedElsewhere,
  sequencePositions,
  withPreservedChildren,
} from "./containerChildren";
import type { NumberingMap } from "./numberingParser";
import { parseParagraph } from "./paragraphParser";
import { captureSdtSiblingMarkers, parseSdtProperties } from "./sdtProperties";
import { enrichParagraphTextBoxes } from "./paragraphTextBoxEnrichment";
import {
  FloatingTableXSpecSchema,
  FloatingTableYSpecSchema,
  TableAlignmentSchema,
  TextDirectionSchema,
  narrowEnum,
} from "./parserEnums";
import type { StyleMap } from "./styleParser";
import { captureVerbatimXml } from "./verbatimCapture";
import { parseBorderSpec } from "./borderParser";
import { parseShading } from "./shadingParser";
import {
  cloneElement,
  findChild,
  findChildren,
  findWordprocessingChild,
  getAttribute,
  getAttributeByNamespaceUri,
  getLocalName,
  mergeXmlnsDeclarations,
  parseNumericAttribute,
  parseTableMeasurementValue,
  parseBooleanElement,
  selectAlternateContentBranch,
  parseOnOffAttribute,
  WORDPROCESSINGML_NAMESPACE_URIS,
} from "./xmlParser";
import type { XmlElement } from "./xmlParser";
import { findChildAnySpelling } from "./strictNames";
import { parsePropertyChangeInfo, parseTrackedChangeInfo } from "./trackedChangeInfo";
import { percentageSpelling, transitionalSlotEncoding } from "./transitionalSpelling";

/**
 * Sanity cap on `w:gridSpan` (and the derived table column count). Word's
 * practical column limit is 63; a hostile/corrupt value here would blow up
 * every downstream structure sized by column count (TableMap, border grids,
 * cell-grid arrays).
 */
export const MAX_TABLE_COLUMNS = 63;

// ============================================================================
// TABLE MEASUREMENT PARSING
// ============================================================================

/**
 * Parse a table measurement (width, height, etc.)
 *
 * @param element - Element with w:w and w:type attributes
 * @returns Parsed measurement or undefined
 */
export function parseTableMeasurement(element: XmlElement | null): TableMeasurement | undefined {
  if (!element) {
    return undefined;
  }

  const declared = getAttribute(element, "w", "type");
  const declaredType: TableWidthType | undefined =
    declared === "auto" || declared === "dxa" || declared === "nil" || declared === "pct"
      ? declared
      : undefined;

  const type = tableWidthType(element, declaredType);

  return { value: parseTableMeasurementValue(element, type) ?? 0, type };
}

/**
 * What unit `w:w` counts in.
 *
 * `w:type` is optional on `CT_TblWidth` and the schema gives it no default, so
 * reading an absent one as `dxa` turned `w:w="50%"` into 50 twips: a
 * full-width table became a hairline. `w:w` is `ST_MeasurementOrPercent`, so a
 * value spelled with a `%` is a percentage whatever `w:type` says — the
 * spelling carries its own unit, and no number of twips is ever written that
 * way. `auto` and `nil` are left alone: neither reads `w:w` as a width at all.
 */
const tableWidthType = (
  element: XmlElement,
  declared: TableWidthType | undefined,
): TableWidthType => {
  if (declared === "auto" || declared === "nil") {
    return declared;
  }
  const raw = getAttribute(element, "w", "w");
  const spelledAsPercent = raw !== null && percentageSpelling(raw) !== undefined;
  const slotTakesPercent =
    transitionalSlotEncoding(element.namespaceUri, getLocalName(element.name), "w")?.percent !==
    undefined;
  if (spelledAsPercent && slotTakesPercent) {
    return "pct";
  }
  return declared ?? "dxa";
};

/**
 * Parse width from an element (shorthand for common case)
 */
function parseWidth(element: XmlElement | null): TableMeasurement | undefined {
  return parseTableMeasurement(element);
}

// ============================================================================
// BORDER PARSING
// ============================================================================

/**
 * Parse a single border specification
 *
 * @param element - Border element (w:top, w:bottom, etc.)
 * @returns Parsed border or undefined
 */

/**
 * Parse table borders (w:tblBorders or w:tcBorders)
 *
 * @param bordersElement - The borders container element
 * @returns Parsed borders or undefined
 */
export function parseTableBorders(bordersElement: XmlElement | null): TableBorders | undefined {
  if (!bordersElement) {
    return undefined;
  }

  const borders: TableBorders = {};

  const top = parseBorderSpec(findChild(bordersElement, "w", "top"));
  if (top) {
    borders.top = top;
  }

  const bottom = parseBorderSpec(findChild(bordersElement, "w", "bottom"));
  if (bottom) {
    borders.bottom = bottom;
  }

  const left = parseBorderSpec(findChildAnySpelling(bordersElement, "CT_Border left"));
  if (left) {
    borders.left = left;
  }

  const right = parseBorderSpec(findChildAnySpelling(bordersElement, "CT_Border right"));
  if (right) {
    borders.right = right;
  }

  const insideH = parseBorderSpec(findChild(bordersElement, "w", "insideH"));
  if (insideH) {
    borders.insideH = insideH;
  }

  const insideV = parseBorderSpec(findChild(bordersElement, "w", "insideV"));
  if (insideV) {
    borders.insideV = insideV;
  }

  // Return undefined if no borders were parsed
  if (Object.keys(borders).length === 0) {
    return undefined;
  }

  return borders;
}

/** Parse cell-only diagonal borders in addition to the shared table sides. */
function parseTableCellBorders(bordersElement: XmlElement | null): TableCellBorders | undefined {
  if (!bordersElement) {
    return undefined;
  }

  const borders: TableCellBorders = { ...parseTableBorders(bordersElement) };
  const topLeftToBottomRight = parseBorderSpec(findChild(bordersElement, "w", "tl2br"));
  if (topLeftToBottomRight) {
    borders.topLeftToBottomRight = topLeftToBottomRight;
  }

  const topRightToBottomLeft = parseBorderSpec(findChild(bordersElement, "w", "tr2bl"));
  if (topRightToBottomLeft) {
    borders.topRightToBottomLeft = topRightToBottomLeft;
  }

  return Object.keys(borders).length > 0 ? borders : undefined;
}

// ============================================================================
// CELL MARGINS PARSING
// ============================================================================

/**
 * Parse cell margins (w:tblCellMar or w:tcMar)
 *
 * @param marginsElement - The margins container element
 * @returns Parsed margins or undefined
 */
export function parseCellMargins(marginsElement: XmlElement | null): CellMargins | undefined {
  if (!marginsElement) {
    return undefined;
  }

  const margins: CellMargins = {};

  const top = parseWidth(findChild(marginsElement, "w", "top"));
  if (top) {
    margins.top = top;
  }

  const bottom = parseWidth(findChild(marginsElement, "w", "bottom"));
  if (bottom) {
    margins.bottom = bottom;
  }

  const left = parseWidth(findChildAnySpelling(marginsElement, "CT_TblWidth left"));
  if (left) {
    margins.left = left;
  }

  const right = parseWidth(findChildAnySpelling(marginsElement, "CT_TblWidth right"));
  if (right) {
    margins.right = right;
  }

  if (Object.keys(margins).length === 0) {
    return undefined;
  }

  return margins;
}

// ============================================================================
// SHADING PARSING
// ============================================================================

// ============================================================================
// TABLE LOOK PARSING
// ============================================================================

/**
 * Read a `w:tblLook` (the only reader; `styleParser` calls this one).
 *
 * What the author wrote, and nothing more: `w:val` verbatim and each flag as
 * stated, absent, `false` or `true`. Folding `w:val`'s bits into the flags here
 * would forget which of the two the document said, and writing the result back
 * would invent attributes the author never had. `resolveTableLook` owns the
 * other direction.
 *
 * @param lookElement - The w:tblLook element
 * @returns Parsed table look or undefined
 */
export function parseTableLook(lookElement: XmlElement | null): TableLook | undefined {
  if (!lookElement) {
    return undefined;
  }

  const look: TableLook = {};

  const val = getAttribute(lookElement, "w", "val");
  if (val !== null && val !== "") {
    look.val = val;
  }

  for (const flag of TABLE_LOOK_FLAGS) {
    const stated = parseOnOffAttribute(lookElement, "w", flag);
    if (stated !== undefined) {
      look[flag] = stated;
    }
  }

  if (Object.keys(look).length === 0) {
    return undefined;
  }

  return look;
}

// ============================================================================
// FLOATING TABLE PROPERTIES
// ============================================================================

/**
 * Parse floating table properties (w:tblpPr)
 *
 * @param tblpPrElement - The w:tblpPr element
 * @returns Parsed floating properties or undefined
 */
export function parseFloatingTableProperties(
  tblpPrElement: XmlElement | null,
): FloatingTableProperties | undefined {
  if (!tblpPrElement) {
    return undefined;
  }

  const floating: FloatingTableProperties = {};

  // Horizontal anchor
  const horzAnchor = getAttribute(tblpPrElement, "w", "horzAnchor");
  if (horzAnchor === "margin" || horzAnchor === "page" || horzAnchor === "text") {
    floating.horzAnchor = horzAnchor;
  }

  // Vertical anchor
  const vertAnchor = getAttribute(tblpPrElement, "w", "vertAnchor");
  if (vertAnchor === "margin" || vertAnchor === "page" || vertAnchor === "text") {
    floating.vertAnchor = vertAnchor;
  }

  // Horizontal position
  const tblpX = parseNumericAttribute(tblpPrElement, "w", "tblpX");
  if (tblpX !== undefined) {
    floating.tblpX = tblpX;
  }

  const tblpXSpec = narrowEnum(
    getAttribute(tblpPrElement, "w", "tblpXSpec"),
    FloatingTableXSpecSchema,
  );
  if (tblpXSpec) {
    floating.tblpXSpec = tblpXSpec;
  }

  // Vertical position
  const tblpY = parseNumericAttribute(tblpPrElement, "w", "tblpY");
  if (tblpY !== undefined) {
    floating.tblpY = tblpY;
  }

  const tblpYSpec = narrowEnum(
    getAttribute(tblpPrElement, "w", "tblpYSpec"),
    FloatingTableYSpecSchema,
  );
  if (tblpYSpec) {
    floating.tblpYSpec = tblpYSpec;
  }

  // Distance from text
  const topFromText = parseNumericAttribute(tblpPrElement, "w", "topFromText");
  if (topFromText !== undefined) {
    floating.topFromText = topFromText;
  }

  const bottomFromText = parseNumericAttribute(tblpPrElement, "w", "bottomFromText");
  if (bottomFromText !== undefined) {
    floating.bottomFromText = bottomFromText;
  }

  const leftFromText = parseNumericAttribute(tblpPrElement, "w", "leftFromText");
  if (leftFromText !== undefined) {
    floating.leftFromText = leftFromText;
  }

  const rightFromText = parseNumericAttribute(tblpPrElement, "w", "rightFromText");
  if (rightFromText !== undefined) {
    floating.rightFromText = rightFromText;
  }

  if (Object.keys(floating).length === 0) {
    return undefined;
  }

  return floating;
}

// ============================================================================
// TABLE PROPERTIES PARSING (w:tblPr)
// ============================================================================

/**
 * Parse table properties (w:tblPr)
 *
 * @param tblPrElement - The w:tblPr element
 * @returns Parsed table formatting
 */
/**
 * Children of a property element that record a revision rather than a
 * property. They are parsed into records of their own and written back from
 * those, so the capture leaves them out: a capture that kept them would put a
 * resolved revision back into the document after it was accepted.
 */
const REVISION_PROPERTY_CHILDREN: ReadonlySet<string> = new Set([
  "tblPrChange",
  "tblPrExChange",
  "trPrChange",
  "tcPrChange",
  "ins",
  "del",
  "cellIns",
  "cellDel",
  "cellMerge",
]);

/**
 * The property element a formatting object was parsed from, kept beside the
 * typed values so a save that did not touch them writes the element back
 * exactly as it arrived.
 */
const withSourceXml = <TFormatting extends { sourceXml?: string }>(
  formatting: TFormatting,
  element: XmlElement,
): TFormatting => ({
  ...formatting,
  sourceXml: captureVerbatimXml(
    cloneElement(element, {
      elements: (element.elements ?? []).filter(
        (child) => !REVISION_PROPERTY_CHILDREN.has(getLocalName(child.name)),
      ),
    }),
  ),
});

/**
 * The nine children `w:tblPr` and `w:tblPrEx` both declare, read once.
 *
 * `CT_TblPrEx` is the middle of `CT_TblPrBase`: the same elements, declared in
 * the same order, without the ones that describe the table as a whole. Two
 * handler maps over the same names would be two answers to the same question,
 * and the one folio would notice is the day they stop agreeing.
 */
const sharedTablePropertyHandlers = (formatting: TablePropertyExceptionFormatting) => ({
  tblW: (child: XmlElement) => {
    const width = parseWidth(child);
    if (width) {
      formatting.width = width;
    }
    return keptUnless(width !== undefined);
  },
  jc: (child: XmlElement) => {
    const justification = narrowEnum(getAttribute(child, "w", "val"), TableAlignmentSchema);
    if (justification !== undefined) {
      formatting.justification = justification;
    }
    return keptUnless(justification !== undefined);
  },
  tblCellSpacing: (child: XmlElement) => {
    const cellSpacing = parseWidth(child);
    if (cellSpacing) {
      formatting.cellSpacing = cellSpacing;
    }
    return keptUnless(cellSpacing !== undefined);
  },
  tblInd: (child: XmlElement) => {
    const indent = parseWidth(child);
    if (indent) {
      formatting.indent = indent;
    }
    return keptUnless(indent !== undefined);
  },
  tblBorders: (child: XmlElement) => {
    const borders = parseTableBorders(child);
    if (borders) {
      formatting.borders = borders;
    }
    return keptUnless(borders !== undefined);
  },
  shd: (child: XmlElement) => {
    const shading = parseShading(child);
    if (shading) {
      formatting.shading = shading;
    }
    return keptUnless(shading !== undefined);
  },
  tblLayout: (child: XmlElement) => {
    const layout = getAttribute(child, "w", "type");
    if (layout === "fixed" || layout === "autofit") {
      formatting.layout = layout;
      return undefined;
    }
    return CAPTURE;
  },
  tblCellMar: (child: XmlElement) => {
    const cellMargins = parseCellMargins(child);
    if (cellMargins) {
      formatting.cellMargins = cellMargins;
    }
    return keptUnless(cellMargins !== undefined);
  },
  tblLook: (child: XmlElement) => {
    const look = parseTableLook(child);
    if (look) {
      formatting.look = look;
    }
    return keptUnless(look !== undefined);
  },
});

/**
 * The children the table walks skip, and the reader that takes each instead.
 *
 * Every one of them is read off its container element before the child walk
 * runs, so capturing it here as well would write it twice. The claims are at
 * module scope, not inside the parsers, so they are registered when the module
 * loads rather than the first time a table is parsed; `ownedElsewhere` is the
 * only way to make one, which is what keeps the set complete.
 */
const TABLE_PROPERTY_CHANGE_OWNER = ownedElsewhere({
  container: "table-properties",
  child: "tblPrChange",
  // A revision, not a property: `parseTablePropertyChanges` reads it into
  // `Table.propertyChanges` and the serializer writes it back from there.
  reader: "tableParser#parseTable",
});

const TABLE_PROPERTY_EXCEPTION_CHANGE_OWNER = ownedElsewhere({
  container: "table-property-exceptions",
  child: "tblPrExChange",
  reader: "tableParser#parseTableRow",
});

const CELL_PROPERTIES_OWNER = ownedElsewhere({
  container: "block-content",
  child: "tcPr",
  reader: "tableParser#parseTableCell",
});

/**
 * The revisions a row's and a cell's property set carry.
 *
 * Each is read off the property element before its child walk runs, into
 * `TableRow.structuralChange` / `propertyChanges` and the cell's equivalents,
 * and written back from there. Capturing one here as well would write it twice,
 * and worse: a capture replays a resolved revision into the document after a
 * reviewer accepted it.
 */
const ROW_INSERTION_OWNER = ownedElsewhere({
  container: "row-properties",
  child: "ins",
  reader: "tableParser#parseTableRow",
});

const ROW_DELETION_OWNER = ownedElsewhere({
  container: "row-properties",
  child: "del",
  reader: "tableParser#parseTableRow",
});

const ROW_PROPERTY_CHANGE_OWNER = ownedElsewhere({
  container: "row-properties",
  child: "trPrChange",
  reader: "tableParser#parseTableRow",
});

const CELL_INSERTION_OWNER = ownedElsewhere({
  container: "cell-properties",
  child: "cellIns",
  reader: "tableParser#parseTableCell",
});

const CELL_DELETION_OWNER = ownedElsewhere({
  container: "cell-properties",
  child: "cellDel",
  reader: "tableParser#parseTableCell",
});

const CELL_MERGE_OWNER = ownedElsewhere({
  container: "cell-properties",
  child: "cellMerge",
  reader: "tableParser#parseTableCell",
});

const CELL_PROPERTY_CHANGE_OWNER = ownedElsewhere({
  container: "cell-properties",
  child: "tcPrChange",
  reader: "tableParser#parseTableCell",
});

const ROW_CHILD_OWNERS = {
  trPr: ownedElsewhere({
    container: "row-content",
    child: "trPr",
    reader: "tableParser#parseTableRowProperties",
  }),
  tblPrEx: ownedElsewhere({
    container: "row-content",
    child: "tblPrEx",
    reader: "tableParser#parseTablePropertyExceptions",
  }),
};

const TABLE_CHILD_OWNERS = {
  tblPr: ownedElsewhere({
    container: "table-content",
    child: "tblPr",
    reader: "tableParser#parseTableProperties",
  }),
  tblGrid: ownedElsewhere({
    container: "table-content",
    child: "tblGrid",
    reader: "tableParser#parseTableGrid",
  }),
};

export function parseTableProperties(tblPrElement: XmlElement | null): TableFormatting | undefined {
  if (!tblPrElement) {
    return undefined;
  }

  const formatting: TableFormatting = {};

  const handlers: ChildHandlers<"table-properties"> = {
    ...sharedTablePropertyHandlers(formatting),
    tblStyle: (child) => {
      const styleId = getAttribute(child, "w", "val");
      if (styleId) {
        formatting.styleId = styleId;
      }
      return keptUnless(Boolean(styleId));
    },
    tblpPr: (child) => {
      const floating = parseFloatingTableProperties(child);
      if (floating) {
        formatting.floating = floating;
      }
      return keptUnless(floating !== undefined);
    },
    tblOverlap: (child) => {
      const overlap = getAttribute(child, "w", "val");
      if (overlap === "never" || overlap === "overlap") {
        formatting.overlap = overlap;
        return undefined;
      }
      return CAPTURE;
    },
    // `CT_OnOff` with no `w:val` is the value `on`, so an empty element states
    // something and the tri-state reader keeps it.
    bidiVisual: (child) => {
      formatting.bidi = parseBooleanElement(child);
    },
    tblStyleRowBandSize: (child) => {
      const size = parseNumericAttribute(child, "w", "val");
      if (size !== undefined) {
        formatting.rowBandSize = size;
      }
      return keptUnless(size !== undefined);
    },
    tblStyleColBandSize: (child) => {
      const size = parseNumericAttribute(child, "w", "val");
      if (size !== undefined) {
        formatting.columnBandSize = size;
      }
      return keptUnless(size !== undefined);
    },
    tblCaption: (child) => {
      const caption = getAttribute(child, "w", "val");
      if (caption !== null) {
        formatting.caption = caption;
      }
      return keptUnless(caption !== null);
    },
    tblDescription: (child) => {
      const description = getAttribute(child, "w", "val");
      if (description !== null) {
        formatting.description = description;
      }
      return keptUnless(description !== null);
    },
    tblPrChange: TABLE_PROPERTY_CHANGE_OWNER,
  };

  const preserved = dispatchChildren({
    element: tblPrElement,
    container: "table-properties",
    handlers,
    capturePosition: sequencePositions("table-properties", tblPrElement),
  });
  if (preserved) {
    formatting.preserved = preserved;
  }

  if (Object.keys(formatting).length === 0) {
    return undefined;
  }

  return withSourceXml(formatting, tblPrElement);
}

/**
 * A revision is the author, the date and the id; the snapshot may be empty.
 *
 * `<w:tblPrChange …><w:tblPr/></w:tblPrChange>` on a table that states no
 * properties of its own records that a reviewer changed the table's
 * properties from nothing to nothing — an accepted style change, most often.
 * Dropping it because neither snapshot carried a typed value threw the
 * revision away, and with it the reviewer's ability to accept or reject.
 * `w:pPrChange` and `w:rPrChange` never did this.
 */
function parseTablePropertyChanges(
  tblPrElement: XmlElement | null,
  currentFormatting: TableFormatting | undefined,
): TablePropertyChange[] | undefined {
  if (!tblPrElement) {
    return undefined;
  }

  const changes = findChildren(tblPrElement, "w", "tblPrChange").map(
    (changeElement): TablePropertyChange => {
      const previousTblPr = findChild(changeElement, "w", "tblPr");
      const change: TablePropertyChange = {
        type: "tablePropertyChange",
        info: parsePropertyChangeInfo(changeElement),
      };
      const prev = parseTableProperties(previousTblPr);
      if (prev !== undefined) {
        change.previousFormatting = prev;
      }
      if (currentFormatting !== undefined) {
        change.currentFormatting = currentFormatting;
      }
      return change;
    },
  );

  return changes.length > 0 ? changes : undefined;
}

/**
 * Parse a row's table property exceptions (`w:tblPrEx`).
 *
 * The nine children are the table's own, read by the same handlers, so an
 * exception and the property it overrides can never be read into two different
 * shapes. `w:tblPrExChange` is a revision rather than a property and is read by
 * {@link parseTablePropertyExceptionChanges}.
 */
export function parseTablePropertyExceptions(
  tblPrExElement: XmlElement | null,
): TablePropertyExceptionFormatting | undefined {
  if (!tblPrExElement) {
    return undefined;
  }

  const formatting: TablePropertyExceptionFormatting = {};

  const handlers: ChildHandlers<"table-property-exceptions"> = {
    ...sharedTablePropertyHandlers(formatting),
    tblPrExChange: TABLE_PROPERTY_EXCEPTION_CHANGE_OWNER,
  };

  const preserved = dispatchChildren({
    element: tblPrExElement,
    container: "table-property-exceptions",
    handlers,
    capturePosition: sequencePositions("table-property-exceptions", tblPrExElement),
  });
  if (preserved) {
    formatting.preserved = preserved;
  }

  // No empty-record guard, unlike the table's own properties: `w:tblPr` is
  // required on a `w:tbl` and written back whatever the model holds, while
  // `w:tblPrEx` is optional, so its presence is itself the value. Returning
  // `undefined` for `<w:tblPrEx/>` deleted the element, and a row that
  // overrides the table's properties with nothing is not a row that does not
  // override them.
  return withSourceXml(formatting, tblPrExElement);
}

/** `w:tblPrExChange`, read the way `w:tblPrChange` is. */
function parseTablePropertyExceptionChanges(
  tblPrExElement: XmlElement | null,
  currentFormatting: TablePropertyExceptionFormatting | undefined,
): TablePropertyExceptionChange[] | undefined {
  if (!tblPrExElement) {
    return undefined;
  }

  const changes = findChildren(tblPrExElement, "w", "tblPrExChange").map(
    (changeElement): TablePropertyExceptionChange => {
      const change: TablePropertyExceptionChange = {
        type: "tablePropertyExceptionChange",
        info: parsePropertyChangeInfo(changeElement),
      };
      const prev = parseTablePropertyExceptions(findChild(changeElement, "w", "tblPrEx"));
      if (prev !== undefined) {
        change.previousFormatting = prev;
      }
      if (currentFormatting !== undefined) {
        change.currentFormatting = currentFormatting;
      }
      return change;
    },
  );

  return changes.length > 0 ? changes : undefined;
}

function parseTableRowPropertyChanges(
  trPrElement: XmlElement | null,
  currentFormatting: TableRowFormatting | undefined,
): TableRowPropertyChange[] | undefined {
  if (!trPrElement) {
    return undefined;
  }

  const changes = findChildren(trPrElement, "w", "trPrChange").map(
    (changeElement): TableRowPropertyChange => {
      const previousTrPr = findChild(changeElement, "w", "trPr");
      const change: TableRowPropertyChange = {
        type: "tableRowPropertyChange",
        info: parsePropertyChangeInfo(changeElement),
      };
      const prev = parseTableRowProperties(previousTrPr);
      if (prev !== undefined) {
        change.previousFormatting = prev;
      }
      if (currentFormatting !== undefined) {
        change.currentFormatting = currentFormatting;
      }
      return change;
    },
  );

  return changes.length > 0 ? changes : undefined;
}

function parseTableCellPropertyChanges(
  tcPrElement: XmlElement | null,
  currentFormatting: TableCellFormatting | undefined,
): TableCellPropertyChange[] | undefined {
  if (!tcPrElement) {
    return undefined;
  }

  const changes = findChildren(tcPrElement, "w", "tcPrChange").map(
    (changeElement): TableCellPropertyChange => {
      const previousTcPr = findChild(changeElement, "w", "tcPr");
      const change: TableCellPropertyChange = {
        type: "tableCellPropertyChange",
        info: parsePropertyChangeInfo(changeElement),
      };
      const prev = parseTableCellProperties(previousTcPr);
      if (prev !== undefined) {
        change.previousFormatting = prev;
      }
      // `CT_TcPrInner` declares the cell's structural revision, so a snapshot
      // may state one; it belongs to the snapshot rather than to the cell.
      const previousStructuralChange = parseTableCellStructuralChange(previousTcPr);
      if (previousStructuralChange !== undefined) {
        change.previousStructuralChange = previousStructuralChange;
      }
      if (currentFormatting !== undefined) {
        change.currentFormatting = currentFormatting;
      }
      return change;
    },
  );

  return changes.length > 0 ? changes : undefined;
}

function parseTableRowStructuralChange(
  trPrElement: XmlElement | null,
): TableStructuralChangeInfo | undefined {
  if (!trPrElement) {
    return undefined;
  }

  const insertion = findChild(trPrElement, "w", "ins");
  if (insertion) {
    return {
      type: "tableRowInsertion",
      info: parseTrackedChangeInfo(insertion),
    };
  }

  const deletion = findChild(trPrElement, "w", "del");
  if (deletion) {
    return {
      type: "tableRowDeletion",
      info: parseTrackedChangeInfo(deletion),
    };
  }

  return undefined;
}

function parseTableCellStructuralChange(
  tcPrElement: XmlElement | null,
): TableStructuralChangeInfo | undefined {
  if (!tcPrElement) {
    return undefined;
  }

  const insertion = findChild(tcPrElement, "w", "cellIns");
  if (insertion) {
    return {
      type: "tableCellInsertion",
      info: parseTrackedChangeInfo(insertion),
    };
  }

  const deletion = findChild(tcPrElement, "w", "cellDel");
  if (deletion) {
    return {
      type: "tableCellDeletion",
      info: parseTrackedChangeInfo(deletion),
    };
  }

  const merge = findChild(tcPrElement, "w", "cellMerge");
  if (merge) {
    const verticalMerge = parseTableCellVerticalMergeRevisionValue(
      getAttribute(merge, "w", "vMerge"),
    );
    const verticalMergeOriginal = parseTableCellVerticalMergeRevisionValue(
      getAttribute(merge, "w", "vMergeOrig"),
    );
    return {
      type: "tableCellMerge",
      info: parseTrackedChangeInfo(merge),
      ...(verticalMerge !== undefined ? { verticalMerge } : {}),
      ...(verticalMergeOriginal !== undefined ? { verticalMergeOriginal } : {}),
    };
  }

  return undefined;
}

function parseTableCellVerticalMergeRevisionValue(
  value: string | null | undefined,
): "continue" | "rest" | undefined {
  if (value === "cont") {
    return "continue";
  }
  if (value === "rest") {
    return "rest";
  }
  return undefined;
}

// ============================================================================
// TABLE ROW PROPERTIES PARSING (w:trPr)
// ============================================================================

/**
 * Parse table row properties (w:trPr)
 *
 * @param trPrElement - The w:trPr element
 * @returns Parsed row formatting
 */
export function parseTableRowProperties(
  trPrElement: XmlElement | null,
): TableRowFormatting | undefined {
  if (!trPrElement) {
    return undefined;
  }

  const formatting: TableRowFormatting = {};

  const handlers: ChildHandlers<"row-properties"> = {
    cnfStyle: (child) => {
      const conditionalFormat = parseConditionalFormatStyle(child);
      if (conditionalFormat) {
        formatting.conditionalFormat = conditionalFormat;
      }
      return keptUnless(conditionalFormat !== undefined);
    },
    // `w:divId` names an HTML `div` the row belonged to in a web page Word
    // round-tripped. Nothing in the editor has a place for it, so it travels
    // as markup.
    divId: CAPTURE,
    gridBefore: (child) => {
      const gridBefore = parseNumericAttribute(child, "w", "val");
      if (gridBefore !== undefined && gridBefore > 0) {
        formatting.gridBefore = gridBefore;
      }
      return keptUnless(gridBefore !== undefined && gridBefore > 0);
    },
    gridAfter: (child) => {
      const gridAfter = parseNumericAttribute(child, "w", "val");
      if (gridAfter !== undefined && gridAfter > 0) {
        formatting.gridAfter = gridAfter;
      }
      return keptUnless(gridAfter !== undefined && gridAfter > 0);
    },
    wBefore: (child) => {
      const widthBefore = parseTableMeasurement(child);
      if (widthBefore) {
        formatting.widthBefore = widthBefore;
      }
      return keptUnless(widthBefore !== undefined);
    },
    wAfter: (child) => {
      const widthAfter = parseTableMeasurement(child);
      if (widthAfter) {
        formatting.widthAfter = widthAfter;
      }
      return keptUnless(widthAfter !== undefined);
    },
    // `CT_OnOff` with no `w:val` is the value `on`, so an empty element states
    // something and the tri-state reader keeps an explicit off apart from an
    // absent element.
    cantSplit: (child) => {
      formatting.cantSplit = parseBooleanElement(child);
    },
    // `w:trHeight` carries the height on `w:val`, not on `w:w`.
    trHeight: (child) => {
      const heightVal = parseNumericAttribute(child, "w", "val");
      if (heightVal === undefined || heightVal <= 0) {
        return CAPTURE;
      }
      formatting.height = { value: heightVal, type: "dxa" };
      const hRule = getAttribute(child, "w", "hRule");
      if (hRule === "auto" || hRule === "atLeast" || hRule === "exact") {
        formatting.heightRule = hRule;
      }
      return undefined;
    },
    tblHeader: (child) => {
      formatting.header = parseBooleanElement(child);
    },
    // The cell spacing a row overrides. `TableRowFormatting` has no field for
    // it and `w:tblPrEx` is where a row states table geometry, so the element
    // travels as markup rather than being read into a shape nothing writes.
    tblCellSpacing: CAPTURE,
    jc: (child) => {
      const justification = narrowEnum(getAttribute(child, "w", "val"), TableAlignmentSchema);
      if (justification !== undefined) {
        formatting.justification = justification;
      }
      return keptUnless(justification !== undefined);
    },
    hidden: (child) => {
      formatting.hidden = parseBooleanElement(child);
    },
    ins: ROW_INSERTION_OWNER,
    del: ROW_DELETION_OWNER,
    trPrChange: ROW_PROPERTY_CHANGE_OWNER,
  };

  const preserved = dispatchChildren({
    element: trPrElement,
    container: "row-properties",
    handlers,
    capturePosition: sequencePositions("row-properties", trPrElement),
  });
  if (preserved) {
    formatting.preserved = preserved;
  }

  // No empty-record guard: see `parseTableCellProperties` for why the element's
  // presence is the value.
  return withSourceXml(formatting, trPrElement);
}

// ============================================================================
// TABLE CELL PROPERTIES PARSING (w:tcPr)
// ============================================================================

/**
 * Parse conditional format style (for table style conditional formatting)
 *
 * @param cnfElement - The w:cnfStyle element
 * @returns Parsed conditional format or undefined
 */
export function parseConditionalFormatStyle(
  cnfElement: XmlElement | null,
): ConditionalFormatStyle | undefined {
  if (!cnfElement) {
    return undefined;
  }

  const style: ConditionalFormatStyle = {};

  // Parse individual flags
  if (parseOnOffAttribute(cnfElement, "w", "firstRow") === true) {
    style.firstRow = true;
  }

  if (parseOnOffAttribute(cnfElement, "w", "lastRow") === true) {
    style.lastRow = true;
  }

  if (parseOnOffAttribute(cnfElement, "w", "firstColumn") === true) {
    style.firstColumn = true;
  }

  if (parseOnOffAttribute(cnfElement, "w", "lastColumn") === true) {
    style.lastColumn = true;
  }

  if (parseOnOffAttribute(cnfElement, "w", "oddHBand") === true) {
    style.oddHBand = true;
  }

  if (parseOnOffAttribute(cnfElement, "w", "evenHBand") === true) {
    style.evenHBand = true;
  }

  if (parseOnOffAttribute(cnfElement, "w", "oddVBand") === true) {
    style.oddVBand = true;
  }

  if (parseOnOffAttribute(cnfElement, "w", "evenVBand") === true) {
    style.evenVBand = true;
  }

  // Corner cells
  if (parseOnOffAttribute(cnfElement, "w", "firstRowFirstColumn") === true) {
    style.nwCell = true;
  }

  if (parseOnOffAttribute(cnfElement, "w", "firstRowLastColumn") === true) {
    style.neCell = true;
  }

  if (parseOnOffAttribute(cnfElement, "w", "lastRowFirstColumn") === true) {
    style.swCell = true;
  }

  if (parseOnOffAttribute(cnfElement, "w", "lastRowLastColumn") === true) {
    style.seCell = true;
  }

  // Also check for the val attribute (binary flags string)
  const val = getAttribute(cnfElement, "w", "val");
  if (val && val.length === 12) {
    // Binary string format: XXXXXXXXXXXXXX
    // Position meanings from left to right
    if (val[0] === "1") {
      style.firstRow = true;
    }
    if (val[1] === "1") {
      style.lastRow = true;
    }
    if (val[2] === "1") {
      style.firstColumn = true;
    }
    if (val[3] === "1") {
      style.lastColumn = true;
    }
    if (val[4] === "1") {
      style.oddVBand = true;
    }
    if (val[5] === "1") {
      style.evenVBand = true;
    }
    if (val[6] === "1") {
      style.oddHBand = true;
    }
    if (val[7] === "1") {
      style.evenHBand = true;
    }
    if (val[8] === "1") {
      style.nwCell = true;
    }
    if (val[9] === "1") {
      style.neCell = true;
    }
    if (val[10] === "1") {
      style.swCell = true;
    }
    if (val[11] === "1") {
      style.seCell = true;
    }
  }

  if (Object.keys(style).length === 0) {
    return undefined;
  }

  return style;
}

/**
 * Parse table cell properties (w:tcPr)
 *
 * @param tcPrElement - The w:tcPr element
 * @returns Parsed cell formatting
 */
export function parseTableCellProperties(
  tcPrElement: XmlElement | null,
): TableCellFormatting | undefined {
  if (!tcPrElement) {
    return undefined;
  }

  const formatting: TableCellFormatting = {};

  const handlers: ChildHandlers<"cell-properties"> = {
    cnfStyle: (child) => {
      const conditionalFormat = parseConditionalFormatStyle(child);
      if (conditionalFormat) {
        formatting.conditionalFormat = conditionalFormat;
      }
      return keptUnless(conditionalFormat !== undefined);
    },
    tcW: (child) => {
      const width = parseWidth(child);
      if (width) {
        formatting.width = width;
      }
      return keptUnless(width !== undefined);
    },
    gridSpan: (child) => {
      const gridSpan = parseNumericAttribute(child, "w", "val");
      if (gridSpan !== undefined && gridSpan > 1) {
        formatting.gridSpan = Math.min(gridSpan, MAX_TABLE_COLUMNS);
        return undefined;
      }
      return CAPTURE;
    },
    // The legacy horizontal merge `w:gridSpan` replaced. folio models the span
    // and nothing reads `w:hMerge`, so it travels as markup: rewriting it as a
    // span would change how a consumer that still honours it lays the row out.
    hMerge: CAPTURE,
    vMerge: (child) => {
      // No `w:val`, or `w:val="continue"`, is a continuation.
      formatting.vMerge = getAttribute(child, "w", "val") === "restart" ? "restart" : "continue";
    },
    tcBorders: (child) => {
      const borders = parseTableCellBorders(child);
      if (borders) {
        formatting.borders = borders;
      }
      return keptUnless(borders !== undefined);
    },
    shd: (child) => {
      const shading = parseShading(child);
      if (shading) {
        formatting.shading = shading;
      }
      return keptUnless(shading !== undefined);
    },
    // `CT_OnOff` with no `w:val` is the value `on`; the tri-state reader keeps
    // an explicit off apart from an absent element.
    noWrap: (child) => {
      formatting.noWrap = parseBooleanElement(child);
    },
    tcMar: (child) => {
      const margins = parseCellMargins(child);
      if (margins) {
        formatting.margins = margins;
      }
      return keptUnless(margins !== undefined);
    },
    textDirection: (child) => {
      const textDir = narrowEnum(getAttribute(child, "w", "val"), TextDirectionSchema);
      if (textDir) {
        formatting.textDirection = textDir;
      }
      return keptUnless(textDir !== undefined);
    },
    tcFitText: (child) => {
      formatting.fitText = parseBooleanElement(child);
    },
    vAlign: (child) => {
      const vAlign = getAttribute(child, "w", "val");
      if (vAlign === "top" || vAlign === "center" || vAlign === "bottom") {
        formatting.verticalAlign = vAlign;
        return undefined;
      }
      return CAPTURE;
    },
    hideMark: (child) => {
      formatting.hideMark = parseBooleanElement(child);
    },
    // `w:headers` names the header cells this one is described by, as
    // accessibility metadata keyed on bookmark names. folio models neither the
    // list nor the bookmarks it points at, so the element travels whole rather
    // than being rebuilt from names it would have to invent.
    headers: CAPTURE,
    cellIns: CELL_INSERTION_OWNER,
    cellDel: CELL_DELETION_OWNER,
    cellMerge: CELL_MERGE_OWNER,
    tcPrChange: CELL_PROPERTY_CHANGE_OWNER,
  };

  const preserved = dispatchChildren({
    element: tcPrElement,
    container: "cell-properties",
    handlers,
    capturePosition: sequencePositions("cell-properties", tcPrElement),
  });
  if (preserved) {
    formatting.preserved = preserved;
  }

  // No empty-record guard, the decision `w:tblPrEx` already made: the element
  // is optional, so its presence is the value, and returning "no properties"
  // for `<w:tcPr/>` deleted the element a producer wrote. The record is
  // present-and-empty rather than absent, and the carrier is the element
  // rather than the properties it yielded — which is what kept an empty
  // property set out of the model and out of the save.
  return withSourceXml(formatting, tcPrElement);
}

// ============================================================================
// ROW- AND CELL-LEVEL CONTENT CONTROLS
// ============================================================================

/**
 * Record a `CT_SdtRow` / `CT_SdtCell` on the rows or cells it wrapped.
 *
 * Both are transparent: what they hold is ordinary rows and cells, and the
 * control adds properties and an end mark. folio keeps the children modelled
 * and puts the control on each of them rather than modelling a wrapper
 * between them, because a table's children are rows and a row's are cells,
 * and neither has a node to spare for something that is not one.
 *
 * Outermost first, so a control inside a control — a bound row inside a
 * repeating section — comes back nested the way it was written. The
 * recursion reads the inner one first, so the outer prepends.
 */
type TablePreservedChild = NonNullable<TablePreservedMarkup["children"]>[number];

const recordContentControl = <Wrapped extends { contentControls?: SdtProperties[] }>(
  wrapped: readonly Wrapped[],
  carriers: readonly { contentControls?: SdtProperties[] }[],
  sdtElement: XmlElement,
): void => {
  const properties = parseSdtProperties(
    findWordprocessingChild(sdtElement, "sdtPr"),
    findWordprocessingChild(sdtElement, "sdtEndPr"),
  );
  const siblings = captureSdtSiblingMarkers(sdtElement);
  if (siblings.before.length > 0) {
    properties.rawSdtChildrenBeforeContent = siblings.before;
  }
  if (siblings.after.length > 0) {
    properties.rawSdtChildrenAfterContent = siblings.after;
  }
  for (const item of wrapped) {
    item.contentControls = [properties, ...(item.contentControls ?? [])];
  }
  for (const item of carriers) {
    item.contentControls = [properties, ...(item.contentControls ?? [])];
  }
};

// ============================================================================
// CELL CONTENT PARSING
// ============================================================================

/**
 * Parse table cell content, including nested block controls.
 *
 * @param tcElement - The w:tc element
 * @param styles - Style definitions
 * @param theme - Theme for color/font resolution
 * @param numbering - Numbering definitions for lists
 * @param rels - Relationships for hyperlinks
 * @param media - Media files for images
 * @returns Array of content blocks
 */
type TableParseOptions = { inHeaderFooter?: boolean; rootXmlns?: Record<string, string> };

/**
 * Accumulate a table container's own `xmlns:*` onto the inherited in-scope set,
 * so a `w:pict` nested inside a cell still resolves prefixes scoped on the
 * `w:tbl` / `w:tr` / `w:tc` wrapper.
 */
function withContainerXmlns(
  options: TableParseOptions | undefined,
  element: XmlElement,
): TableParseOptions {
  return { ...options, rootXmlns: mergeXmlnsDeclarations(options?.rootXmlns ?? {}, element) };
}

function parseCellContent(
  tcElement: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  options?: { inHeaderFooter?: boolean; rootXmlns?: Record<string, string> },
): TableCellBlock[] {
  const findLastFlowBlock = (blocks: readonly TableCellBlock[]): Paragraph | Table | undefined => {
    for (let index = blocks.length - 1; index >= 0; index -= 1) {
      const block = blocks[index];
      if (block?.type === "paragraph" || block?.type === "table") {
        return block;
      }
      if (block?.type === "blockSdt") {
        const nested = findLastFlowBlock(block.content);
        if (nested) {
          return nested;
        }
      }
    }
    return undefined;
  };

  const parseCellChildren = (
    element: XmlElement,
    childOptions: TableParseOptions | undefined,
    requireTrailingParagraph: boolean,
  ): TableCellBlock[] => {
    const modelled: TableCellBlock[] = [];
    const captured: PreservedChild[] = [];

    // AlternateContent is transparent to the cell's block sequence. An SDT is
    // not: it owns a nested block sequence whose preservation positions must
    // be counted independently from the surrounding cell.
    const dispatchCellChildren = (
      childContainer: XmlElement,
      nestedOptions: TableParseOptions | undefined,
    ): void => {
      const preserved = dispatchChildren({
        element: childContainer,
        container: "block-content",
        capturePosition: () => modelled.length,
        undeclared: {
          AlternateContent: (child) => {
            const selectedBranch = selectAlternateContentBranch(child);
            if (!selectedBranch) {
              return;
            }
            dispatchCellChildren(
              selectedBranch,
              withContainerXmlns(withContainerXmlns(nestedOptions, child), selectedBranch),
            );
          },
        },
        handlers: {
          p: (child) => {
            const para = parseParagraph(child, styles, theme, numbering, rels, media, {
              ...nestedOptions,
              runConsolidation: "deferred",
            });
            enrichParagraphTextBoxes(
              para,
              child,
              styles,
              theme,
              numbering,
              rels,
              media,
              parseTable,
            );
            modelled.push(para);
          },
          tbl: (child) => {
            const table = parseTable(child, styles, theme, numbering, rels, media, nestedOptions);
            if (!table) {
              return;
            }
            modelled.push(table);
          },
          sdt: (child) => {
            const properties = parseSdtProperties(
              findWordprocessingChild(child, "sdtPr"),
              findWordprocessingChild(child, "sdtEndPr"),
            );
            const siblings = captureSdtSiblingMarkers(child);
            if (siblings.before.length > 0) {
              properties.rawSdtChildrenBeforeContent = siblings.before;
            }
            if (siblings.after.length > 0) {
              properties.rawSdtChildrenAfterContent = siblings.after;
            }
            const sdtContent = findWordprocessingChild(child, "sdtContent");
            if (!sdtContent) {
              return CAPTURE;
            }
            modelled.push({
              type: "blockSdt",
              properties,
              content: parseCellChildren(
                sdtContent,
                withContainerXmlns(withContainerXmlns(nestedOptions, child), sdtContent),
                false,
              ),
            });
            return undefined;
          },
          // `CT_Tc` declares the marker beside its blocks, so the cell keeps it
          // there rather than folding it into a neighbouring paragraph.
          bookmarkStart: (child) => {
            modelled.push(parseBookmarkStart(child));
          },
          bookmarkEnd: (child) => {
            modelled.push(parseBookmarkEnd(child));
          },
          tcPr: CELL_PROPERTIES_OWNER,
          // Declared for `w:body`, not for a cell; the handler map is total over
          // the union every block container shares.
          sectPr: CAPTURE,
          altChunk: CAPTURE,
          commentRangeEnd: CAPTURE,
          commentRangeStart: CAPTURE,
          customXml: CAPTURE,
          customXmlDelRangeEnd: CAPTURE,
          customXmlDelRangeStart: CAPTURE,
          customXmlInsRangeEnd: CAPTURE,
          customXmlInsRangeStart: CAPTURE,
          customXmlMoveFromRangeEnd: CAPTURE,
          customXmlMoveFromRangeStart: CAPTURE,
          customXmlMoveToRangeEnd: CAPTURE,
          customXmlMoveToRangeStart: CAPTURE,
          del: CAPTURE,
          ins: CAPTURE,
          moveFrom: CAPTURE,
          moveFromRangeEnd: CAPTURE,
          moveFromRangeStart: CAPTURE,
          moveTo: CAPTURE,
          moveToRangeEnd: CAPTURE,
          moveToRangeStart: CAPTURE,
          permEnd: CAPTURE,
          permStart: CAPTURE,
          proofErr: CAPTURE,
        },
      });
      if (preserved?.children) {
        captured.push(...preserved.children);
      }
    };

    dispatchCellChildren(element, childOptions);

    // `CT_Tc` ends in a paragraph: a cell holds at least one, and a nested table
    // is never its last block. A package that ends a cell with a table, or with
    // nothing, states a cell no consumer can render as written, and every one of
    // them reads the implied empty paragraph there instead. Parsing it as a fact
    // keeps the model's cells the shape the format allows, so a comparison
    // against such a package is not asked to delete a paragraph mark that has to
    // stay. Neither opaque markup nor a bookmark marker is a paragraph, so the
    // block that decides this is the last paragraph or table the cell holds.
    if (requireTrailingParagraph) {
      const lastBlock = findLastFlowBlock(modelled);
      if (lastBlock?.type !== "paragraph") {
        modelled.push({ type: "paragraph", content: [] });
      }
    }

    return withPreservedChildren(
      modelled,
      { children: captured },
      (xml): PreservedBlock => ({ type: "preservedBlock", xml }),
    );
  };

  return parseCellChildren(tcElement, options, true);
}

// ============================================================================
// TABLE CELL PARSING
// ============================================================================

/**
 * Parse a table cell (w:tc)
 *
 * @param tcElement - The w:tc element
 * @param styles - Style definitions
 * @param theme - Theme for color/font resolution
 * @param numbering - Numbering definitions for lists
 * @param rels - Relationships for hyperlinks
 * @param media - Media files for images
 * @returns Parsed table cell
 */
export function parseTableCell(
  tcElement: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  options?: { inHeaderFooter?: boolean; rootXmlns?: Record<string, string> },
): TableCell {
  const cell: TableCell = {
    type: "tableCell",
    content: [],
  };
  const id = getAttributeByNamespaceUri(tcElement, WORDPROCESSINGML_NAMESPACE_URIS, "id");
  if (id !== null) {
    cell.id = id;
  }

  // Parse cell properties (w:tcPr)
  const tcPrElement = findChild(tcElement, "w", "tcPr");
  const formatting = parseTableCellProperties(tcPrElement);
  if (formatting) {
    cell.formatting = formatting;
  }
  const cellPropChanges = parseTableCellPropertyChanges(tcPrElement, formatting);
  if (cellPropChanges !== undefined) {
    cell.propertyChanges = cellPropChanges;
  }
  const cellStructChange = parseTableCellStructuralChange(tcPrElement);
  if (cellStructChange !== undefined) {
    cell.structuralChange = cellStructChange;
  }

  // Parse content, threading the cell's own xmlns down the in-scope set.
  cell.content = parseCellContent(
    tcElement,
    styles,
    theme,
    numbering,
    rels,
    media,
    withContainerXmlns(options, tcElement),
  );

  return cell;
}

// ============================================================================
// TABLE ROW PARSING
// ============================================================================

/**
 * Parse a table row (w:tr)
 *
 * @param trElement - The w:tr element
 * @param styles - Style definitions
 * @param theme - Theme for color/font resolution
 * @param numbering - Numbering definitions for lists
 * @param rels - Relationships for hyperlinks
 * @param media - Media files for images
 * @returns Parsed table row
 */
export function parseTableRow(
  trElement: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  options?: { inHeaderFooter?: boolean; rootXmlns?: Record<string, string> },
): TableRow {
  const row: TableRow = {
    type: "tableRow",
    cells: [],
  };

  // The table properties this row overrides (w:tblPrEx). `CT_Row` declares it
  // before `w:trPr`, and the two are read off the element rather than by the
  // child walk below, which is why {@link ROW_CHILD_OWNERS} names this reader.
  const tblPrExElement = findChild(trElement, "w", "tblPrEx");
  const exceptions = parseTablePropertyExceptions(tblPrExElement);
  if (exceptions) {
    row.tablePropertyExceptions = exceptions;
  }
  const exceptionChanges = parseTablePropertyExceptionChanges(tblPrExElement, exceptions);
  if (exceptionChanges !== undefined) {
    row.tablePropertyExceptionChanges = exceptionChanges;
  }

  // Parse row properties (w:trPr)
  const trPrElement = findChild(trElement, "w", "trPr");
  const formatting = parseTableRowProperties(trPrElement);
  if (formatting) {
    row.formatting = formatting;
  }
  const rowPropChanges = parseTableRowPropertyChanges(trPrElement, formatting);
  if (rowPropChanges !== undefined) {
    row.propertyChanges = rowPropChanges;
  }
  const rowStructChange = parseTableRowStructuralChange(trPrElement);
  if (rowStructChange !== undefined) {
    row.structuralChange = rowStructChange;
  }

  // Parse cells, threading the row's own xmlns down the in-scope set.
  const rowOptions = withContainerXmlns(options, trElement);
  const bookmarks: PositionedBookmarkMarker[] = [];
  const preservedChildren: TablePreservedChild[] = [];

  /**
   * One row's children, or a cell-level content control's.
   *
   * folio keeps the cells `w:sdt` holds as the row's own and records the
   * control on each of them, so the recursion walks the control's content
   * with the same map; the sink is the row's either way, because the control
   * holds no capture of its own.
   */
  const dispatchRowChildren = (
    element: XmlElement,
    childOptions: TableParseOptions | undefined,
  ): void => {
    const captured = dispatchChildren({
      element,
      container: "row-content",
      capturePosition: () => row.cells.length,
      handlers: {
        tc: (child) => {
          row.cells.push(
            parseTableCell(child, styles, theme, numbering, rels, media, childOptions),
          );
        },

        sdt: (child) => {
          const sdtContent = findWordprocessingChild(child, "sdtContent");
          if (!sdtContent) {
            return CAPTURE;
          }
          const firstWrapped = row.cells.length;
          const firstCaptured = preservedChildren.length;
          const firstBookmark = bookmarks.length;
          const sdtOptions = withContainerXmlns(childOptions, child);
          dispatchRowChildren(sdtContent, withContainerXmlns(sdtOptions, sdtContent));
          const wrapped = row.cells.slice(firstWrapped);
          if (wrapped.length === 0) {
            preservedChildren.splice(firstCaptured);
            bookmarks.splice(firstBookmark);
            return CAPTURE;
          }
          recordContentControl(
            wrapped,
            [...preservedChildren.slice(firstCaptured), ...bookmarks.slice(firstBookmark)],
            child,
          );
          return undefined;
        },

        // A bookmark that selects whole rows opens and closes here, between
        // two cells. The row models one kind of child, so the marker keeps its
        // place as an index among the cells rather than as a member — and it
        // stays a typed marker rather than joining the verbatim sink, because
        // a bookmark the model cannot see is a bookmark whose partner the
        // editor deletes.
        bookmarkStart: (child) => {
          bookmarks.push({ index: row.cells.length, marker: parseBookmarkStart(child) });
        },
        bookmarkEnd: (child) => {
          bookmarks.push({ index: row.cells.length, marker: parseBookmarkEnd(child) });
        },

        ...ROW_CHILD_OWNERS,

        commentRangeEnd: CAPTURE,
        commentRangeStart: CAPTURE,
        customXml: CAPTURE,
        customXmlDelRangeEnd: CAPTURE,
        customXmlDelRangeStart: CAPTURE,
        customXmlInsRangeEnd: CAPTURE,
        customXmlInsRangeStart: CAPTURE,
        customXmlMoveFromRangeEnd: CAPTURE,
        customXmlMoveFromRangeStart: CAPTURE,
        customXmlMoveToRangeEnd: CAPTURE,
        customXmlMoveToRangeStart: CAPTURE,
        del: CAPTURE,
        ins: CAPTURE,
        moveFrom: CAPTURE,
        moveFromRangeEnd: CAPTURE,
        moveFromRangeStart: CAPTURE,
        moveTo: CAPTURE,
        moveToRangeEnd: CAPTURE,
        moveToRangeStart: CAPTURE,
        permEnd: CAPTURE,
        permStart: CAPTURE,
        proofErr: CAPTURE,
        // A row nested directly in a row is legal markup folio has no model
        // for; captured whole rather than flattened into this row's cells,
        // which would move its content into a row the author did not write.
        tr: CAPTURE,
      },
    });
    if (captured?.children) {
      preservedChildren.push(...captured.children);
    }
  };

  dispatchRowChildren(trElement, rowOptions);
  if (preservedChildren.length > 0) {
    row.preserved = { children: preservedChildren };
  }
  if (bookmarks.length > 0) {
    row.bookmarks = bookmarks;
  }

  // `CT_Row`'s four `w:rsid*` attributes, and anything else the source put on
  // the element: `serializeTableRow` writes a bare `<w:tr>` start tag, so the
  // row models none of them.
  const remainingAttributes = attributeRemainder({
    element: trElement,
    modelled: NO_MODELLED_ATTRIBUTES,
  });
  if (remainingAttributes) {
    row.preservedAttributes = remainingAttributes;
  }

  return row;
}

// ============================================================================
// TABLE GRID PARSING
// ============================================================================

/**
 * Parse table grid (w:tblGrid) for column widths
 *
 * @param tblGridElement - The w:tblGrid element
 * @returns Array of column widths in twips
 */
export function parseTableGrid(tblGridElement: XmlElement | null): number[] | undefined {
  if (!tblGridElement) {
    return undefined;
  }

  const widths: number[] = [];
  let hasExplicitZero = false;

  const gridCols = findChildren(tblGridElement, "w", "gridCol");
  for (const col of gridCols) {
    const statedWidth = parseNumericAttribute(col, "w", "w");
    hasExplicitZero ||= statedWidth === 0;
    const width = statedWidth ?? 0;
    widths.push(width);
  }

  // An omitted width and an explicit zero both feed zero into the layout
  // fallback, but only the latter is authored data the editor must retain.
  if (widths.length > 0 && widths.every((width) => width <= 0) && !hasExplicitZero) {
    return undefined;
  }

  return widths.length > 0 ? widths : undefined;
}

/**
 * The grid a `w:tblGridChange` snapshots, from the change's own `w:tblGrid`.
 *
 * {@link parseTableGrid} reads the *live* grid, where a column with no `w:w`
 * is a column folio has no width for and a grid of nothing but those is no
 * grid at all. A snapshot is a record of what stood, so every `w:gridCol`
 * survives whether or not it stated a width, and `undefined` is what says the
 * column stated none.
 */
function parseTableGridChange(gridElement: XmlElement): TableGridChange | undefined {
  const changeElement = findChild(gridElement, "w", "tblGridChange");
  if (!changeElement) {
    return undefined;
  }
  const snapshot = findChild(changeElement, "w", "tblGrid");
  return {
    id: parseTrackedChangeInfo(changeElement).id,
    columnWidths: (snapshot === null ? [] : findChildren(snapshot, "w", "gridCol")).map((column) =>
      parseNumericAttribute(column, "w", "w"),
    ),
  };
}

function hasRowGridOffsets(rowElement: XmlElement): boolean {
  const trPrElement = findChild(rowElement, "w", "trPr");
  if (!trPrElement) {
    return false;
  }

  const gridBefore =
    parseNumericAttribute(findChild(trPrElement, "w", "gridBefore"), "w", "val") ?? 0;
  const gridAfter =
    parseNumericAttribute(findChild(trPrElement, "w", "gridAfter"), "w", "val") ?? 0;

  return gridBefore > 0 || gridAfter > 0;
}

function getTableGridWidth(table: Table): number | null {
  const totalWidth = table.columnWidths?.reduce((sum, width) => sum + width, 0);
  if (!totalWidth || totalWidth <= 0) {
    return null;
  }
  return totalWidth;
}

function cellWidthCoversTableGrid(cell: TableCell, tableGridWidth: number | null): boolean {
  const width = cell.formatting?.width;
  if (!width || width.type !== "dxa" || tableGridWidth === null) {
    return false;
  }
  return width.value >= tableGridWidth;
}

function inferImplicitSingleCellRowSpans(table: Table, rowsWithGridOffsets: Set<number>): void {
  const gridColumnCount = table.columnWidths?.length ?? 0;
  if (gridColumnCount <= 1) {
    return;
  }

  const tableGridWidth = getTableGridWidth(table);
  if (tableGridWidth === null) {
    return;
  }

  for (const [rowIndex, row] of table.rows.entries()) {
    if (row.cells.length !== 1) {
      continue;
    }
    if (rowsWithGridOffsets.has(rowIndex)) {
      continue;
    }

    const cell = row.cells.at(0);
    if (!cell) {
      continue;
    }

    const currentSpan = cell.formatting?.gridSpan ?? 1;
    if (currentSpan >= gridColumnCount) {
      continue;
    }

    if (cell.formatting?.vMerge) {
      continue;
    }
    if (cell.formatting?.gridSpan != null) {
      continue;
    }
    if (!cellWidthCoversTableGrid(cell, tableGridWidth)) {
      continue;
    }

    cell.formatting = {
      ...cell.formatting,
      gridSpan: gridColumnCount,
    };
  }
}

// ============================================================================
// MAIN TABLE PARSING
// ============================================================================

/**
 * Parse a table element (w:tbl)
 *
 * @param tblElement - The w:tbl element
 * @param styles - Style definitions
 * @param theme - Theme for color/font resolution
 * @param numbering - Numbering definitions for lists
 * @param rels - Relationships for hyperlinks
 * @param media - Media files for images
 * @returns Parsed table
 */
export function parseTable(
  tblElement: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  options?: { inHeaderFooter?: boolean; rootXmlns?: Record<string, string> },
): Table | undefined {
  const table: Table = {
    type: "table",
    rows: [],
  };

  // Parse table properties (w:tblPr)
  const tblPrElement = findChild(tblElement, "w", "tblPr");
  const formatting = parseTableProperties(tblPrElement);
  if (formatting) {
    table.formatting = formatting;
  }
  const tblPropChanges = parseTablePropertyChanges(tblPrElement, formatting);
  if (tblPropChanges !== undefined) {
    table.propertyChanges = tblPropChanges;
  }

  // Parse table grid (w:tblGrid)
  const gridElement = findChild(tblElement, "w", "tblGrid");
  const columnWidths = parseTableGrid(gridElement);
  if (columnWidths) {
    table.columnWidths = columnWidths;
  }
  // The grid element travels with the table's formatting, so a save that did
  // not resize a column writes it back with whatever it carried. A save that
  // did resize one rebuilds the grid, and `w:tblGridChange` — the tracked
  // record of the grid a reviewer replaced — has to travel on its own to
  // survive that rebuild.
  if (gridElement) {
    const gridChange = parseTableGridChange(gridElement);
    table.formatting = {
      ...table.formatting,
      gridSourceXml: captureVerbatimXml(gridElement),
      ...(gridChange === undefined ? {} : { gridChange }),
    };
  }

  // Parse rows, threading the table's own xmlns down the in-scope set.
  const tableOptions = withContainerXmlns(options, tblElement);
  const rowsWithGridOffsets = new Set<number>();
  const preservedChildren: TablePreservedChild[] = [];
  const bookmarks: PositionedBookmarkMarker[] = [];

  /**
   * One table's children, or a row-level content control's.
   *
   * folio keeps the rows `w:sdt` holds as the table's own and records the
   * control on each of them, so the recursion walks the control's content
   * with the same map; the sink is the table's either way, because the
   * control holds no capture of its own.
   */
  const dispatchTableChildren = (
    element: XmlElement,
    childOptions: TableParseOptions | undefined,
  ): void => {
    const captured = dispatchChildren({
      element,
      container: "table-content",
      capturePosition: () => table.rows.length,
      handlers: {
        tr: (child) => {
          const rowIndex = table.rows.length;
          table.rows.push(
            parseTableRow(child, styles, theme, numbering, rels, media, childOptions),
          );
          if (hasRowGridOffsets(child)) {
            rowsWithGridOffsets.add(rowIndex);
          }
        },

        sdt: (child) => {
          const sdtContent = findWordprocessingChild(child, "sdtContent");
          if (!sdtContent) {
            return CAPTURE;
          }
          const firstWrapped = table.rows.length;
          const firstCaptured = preservedChildren.length;
          const firstBookmark = bookmarks.length;
          const sdtOptions = withContainerXmlns(childOptions, child);
          dispatchTableChildren(sdtContent, withContainerXmlns(sdtOptions, sdtContent));
          const wrapped = table.rows.slice(firstWrapped);
          if (wrapped.length === 0) {
            preservedChildren.splice(firstCaptured);
            bookmarks.splice(firstBookmark);
            return CAPTURE;
          }
          recordContentControl(
            wrapped,
            [...preservedChildren.slice(firstCaptured), ...bookmarks.slice(firstBookmark)],
            child,
          );
          return undefined;
        },

        ...TABLE_CHILD_OWNERS,

        // A bookmark that selects a whole table opens and closes here. The
        // sink would keep the bytes and hide the marker from the pass that
        // pairs it with a `w:bookmarkStart` inside a cell, so it is typed and
        // positioned by the rows that preceded it.
        bookmarkStart: (child) => {
          bookmarks.push({ index: table.rows.length, marker: parseBookmarkStart(child) });
        },
        bookmarkEnd: (child) => {
          bookmarks.push({ index: table.rows.length, marker: parseBookmarkEnd(child) });
        },
        commentRangeEnd: CAPTURE,
        commentRangeStart: CAPTURE,
        // Kept whole rather than unwrapped, so everything the wrapper holds
        // survives the save as the bytes the source wrote.
        customXml: CAPTURE,
        // Only `w:customXml` declares it, and that wrapper is captured whole,
        // so this walk never meets one. Capture is still the right answer for
        // a source that writes it where the schema does not admit it.
        customXmlPr: CAPTURE,
        customXmlDelRangeEnd: CAPTURE,
        customXmlDelRangeStart: CAPTURE,
        customXmlInsRangeEnd: CAPTURE,
        customXmlInsRangeStart: CAPTURE,
        customXmlMoveFromRangeEnd: CAPTURE,
        customXmlMoveFromRangeStart: CAPTURE,
        customXmlMoveToRangeEnd: CAPTURE,
        customXmlMoveToRangeStart: CAPTURE,
        del: CAPTURE,
        ins: CAPTURE,
        moveFrom: CAPTURE,
        moveFromRangeEnd: CAPTURE,
        moveFromRangeStart: CAPTURE,
        moveTo: CAPTURE,
        moveToRangeEnd: CAPTURE,
        moveToRangeStart: CAPTURE,
        permEnd: CAPTURE,
        permStart: CAPTURE,
        proofErr: CAPTURE,
      },
    });
    if (captured?.children) {
      preservedChildren.push(...captured.children);
    }
  };

  // A `w:tbl` nested directly in a `w:tbl` is markup the content model does
  // not declare, so the sink's default keeps it whole; flattening it would
  // move its rows into a table the author did not write.
  dispatchTableChildren(tblElement, tableOptions);

  // OOXML encountered in the wild can contain placeholder w:tbl elements
  // without rows. They have no visible content, while the canonical model
  // intentionally requires every table to contain a row. Omit the placeholder
  // instead of inventing a visible row or rejecting the entire document.
  if (table.rows.length === 0) {
    return undefined;
  }

  if (preservedChildren.length > 0) {
    table.preserved = { children: preservedChildren };
  }
  if (bookmarks.length > 0) {
    table.bookmarks = bookmarks;
  }

  inferImplicitSingleCellRowSpans(table, rowsWithGridOffsets);

  return table;
}

// ============================================================================
// TABLE UTILITIES
// ============================================================================

/**
 * Get the number of columns in a table
 *
 * Uses the table grid if available, otherwise counts cells in first row.
 *
 * @param table - The table to measure
 * @returns Number of columns
 */
export function getTableColumnCount(table: Table): number {
  if (table.columnWidths && table.columnWidths.length > 0) {
    return table.columnWidths.length;
  }

  if (table.rows.length === 0) {
    return 0;
  }

  // SAFETY: rows.length > 0 verified above
  // Count cells in first row, accounting for grid span
  return table.rows[0]!.cells.reduce((count, cell) => count + (cell.formatting?.gridSpan ?? 1), 0);
}

/**
 * Get the number of rows in a table
 *
 * @param table - The table to measure
 * @returns Number of rows
 */
export function getTableRowCount(table: Table): number {
  return table.rows.length;
}

/**
 * Check if a cell is part of a vertical merge
 *
 * @param cell - The cell to check
 * @returns true if cell continues a vertical merge
 */
export function isCellMergeContinuation(cell: TableCell): boolean {
  return cell.formatting?.vMerge === "continue";
}

/**
 * Check if a cell starts a vertical merge
 *
 * @param cell - The cell to check
 * @returns true if cell starts a vertical merge
 */
export function isCellMergeStart(cell: TableCell): boolean {
  return cell.formatting?.vMerge === "restart";
}

/**
 * Check if a cell spans multiple columns
 *
 * @param cell - The cell to check
 * @returns true if cell spans multiple columns
 */
export function isCellHorizontallyMerged(cell: TableCell): boolean {
  return (cell.formatting?.gridSpan ?? 1) > 1;
}

/**
 * Get the plain text content of a table
 *
 * @param table - The table to extract text from
 * @returns Plain text content
 */
export function getTableText(table: Table): string {
  const rows: string[] = [];

  for (const row of table.rows) {
    const cells: string[] = [];

    for (const cell of row.cells) {
      const cellText = cell.content
        .filter((c): c is Paragraph => c.type === "paragraph")
        .map((p) => getParagraphText(p))
        .join("\n");
      cells.push(cellText);
    }

    rows.push(cells.join("\t"));
  }

  return rows.join("\n");
}

/**
 * Helper to get paragraph text (simplified)
 */
function getParagraphText(para: Paragraph): string {
  return para.content
    .filter((c) => "content" in c)
    .flatMap((run) => {
      if (!("content" in run) || !Array.isArray(run.content)) {
        return [];
      }
      return run.content
        .filter(
          (c: unknown): c is { type: "text"; text: string } =>
            typeof c === "object" && c !== null && "type" in c && c.type === "text" && "text" in c,
        )
        .map((c) => c.text);
    })
    .join("");
}

/**
 * Check if table has header row
 *
 * @param table - The table to check
 * @returns true if first row is marked as header
 */
export function hasHeaderRow(table: Table): boolean {
  if (table.rows.length === 0) {
    return false;
  }
  // SAFETY: rows.length > 0 verified above
  return table.rows[0]!.formatting?.header === true;
}

/**
 * Get all header rows from a table
 *
 * @param table - The table to search
 * @returns Array of header rows
 */
export function getHeaderRows(table: Table): TableRow[] {
  return table.rows.filter((row) => row.formatting?.header === true);
}

/**
 * Check if table is a floating table
 *
 * @param table - The table to check
 * @returns true if table has floating properties
 */
export function isFloatingTable(table: Table): boolean {
  return table.formatting?.floating !== undefined;
}
