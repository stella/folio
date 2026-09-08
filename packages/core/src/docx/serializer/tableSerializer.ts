/**
 * Table Serializer - Serialize tables to OOXML XML
 *
 * Converts Table objects back to <w:tbl> XML format for DOCX files.
 * Handles all table, row, and cell properties including merged cells.
 *
 * OOXML Reference:
 * - Table: w:tbl
 * - Table properties: w:tblPr
 * - Table grid: w:tblGrid
 * - Table row: w:tr
 * - Row properties: w:trPr
 * - Table cell: w:tc
 * - Cell properties: w:tcPr
 */

import type {
  Table,
  TableRow,
  TableCell,
  TableFormatting,
  TableRowFormatting,
  TableCellFormatting,
  TablePropertyChange,
  TableRowPropertyChange,
  TableCellPropertyChange,
  TableStructuralChangeInfo,
  TableMeasurement,
  TableBorders,
  TableCellBorders,
  BorderSpec,
  TableLook,
  CellMargins,
  FloatingTableProperties,
  ConditionalFormatStyle,
  ShadingProperties,
  Paragraph,
} from "../../types/document";
import { canonicalJson } from "../../utils/canonicalJson";
import { isValidHexColor } from "../../utils/colorResolver";
import {
  parseTableCellProperties,
  parseTableGrid,
  parseTableProperties,
  parseTableRowProperties,
} from "../tableParser";
import { OOXML_NAMESPACE_SCOPE, parseXml, type XmlElement } from "../xmlParser";
import { serializeBorder } from "./borderSerializer";
import { serializeTrackedChangeAttributes } from "./trackedChangeAttributes";
import { escapeXml, intAttr } from "./xmlUtils";

type ParagraphSerializer = (paragraph: Paragraph) => string;

/**
 * The captured element with the revision children written back into it.
 *
 * The capture holds properties only, so a revision the model carries — a row
 * marked inserted, a `w:tcPrChange` — is appended here rather than forcing the
 * whole element to be rebuilt and losing everything the model does not cover.
 */
const withRevisionChildren = (sourceXml: string, revisions: readonly string[]): string => {
  const written = revisions.join("");
  if (written.length === 0) {
    return sourceXml;
  }
  if (sourceXml.endsWith("/>")) {
    const name = sourceXml.slice(1, sourceXml.length - 2).split(/[\s/]/u)[0] ?? "";
    return `${sourceXml.slice(0, -2)}>${written}</${name}>`;
  }
  const close = sourceXml.lastIndexOf("</");
  return close === -1
    ? sourceXml
    : `${sourceXml.slice(0, close)}${written}${sourceXml.slice(close)}`;
};

/** `w:trPr/w:ins` | `w:trPr/w:del`, the row's own structural revision. */
const rowStructuralChangeXml = (change: TableStructuralChangeInfo | undefined): string[] => {
  if (change?.type === "tableRowInsertion") {
    return [`<w:ins ${serializeTrackedChangeAttributes(change.info)}/>`];
  }
  if (change?.type === "tableRowDeletion") {
    return [`<w:del ${serializeTrackedChangeAttributes(change.info)}/>`];
  }
  return [];
};

/** `w:tcPr/w:cellIns` | `w:cellDel` | `w:cellMerge`, the cell's own. */
const cellStructuralChangeXml = (change: TableStructuralChangeInfo | undefined): string[] => {
  if (change?.type === "tableCellInsertion") {
    return [`<w:cellIns ${serializeTrackedChangeAttributes(change.info)}/>`];
  }
  if (change?.type === "tableCellDeletion") {
    return [`<w:cellDel ${serializeTrackedChangeAttributes(change.info)}/>`];
  }
  if (change?.type !== "tableCellMerge") {
    return [];
  }
  const attrs = [serializeTrackedChangeAttributes(change.info)];
  if (change.verticalMerge) {
    attrs.push(`w:vMerge="${change.verticalMerge === "continue" ? "cont" : "rest"}"`);
  }
  if (change.verticalMergeOriginal) {
    attrs.push(`w:vMergeOrig="${change.verticalMergeOriginal === "continue" ? "cont" : "rest"}"`);
  }
  return [`<w:cellMerge ${attrs.join(" ")}/>`];
};

/**
 * A property set without the elements kept beside it. Only the typed values
 * decide whether a capture still describes the node; the captures themselves
 * are what is being decided about.
 */
const withoutCaptures = (
  formatting: { sourceXml?: string; gridSourceXml?: string } | undefined,
): Record<string, unknown> => {
  const { sourceXml: _source, gridSourceXml: _grid, ...rest } = formatting ?? {};
  return rest;
};

/**
 * The element a formatting object was parsed from, when it still describes it.
 *
 * `sourceXml` is only usable while the typed values around it are the ones it
 * was parsed into. Anything may edit a `Document` in place — a host building
 * one by hand, a migration, a test — and a capture written back over a changed
 * model would silently discard the change. So the capture is re-parsed and
 * checked rather than trusted: it is written only when parsing it reproduces
 * the formatting it sits on.
 *
 * The re-parse runs under {@link OOXML_NAMESPACE_SCOPE} because a captured
 * fragment carries no `xmlns` of its own. A document that binds the
 * WordprocessingML prefix differently fails the check and is rebuilt from the
 * model, which is what happened to every document before the capture existed.
 */
const verifiedSourceXml = <TFormatting extends { sourceXml?: string }>(
  formatting: TFormatting | undefined,
  parse: (element: XmlElement | null) => TFormatting | undefined,
): string | null => {
  if (formatting === undefined || formatting.sourceXml === undefined) {
    return null;
  }
  const { sourceXml } = formatting;
  const reparsed = parse(parseXml(sourceXml, OOXML_NAMESPACE_SCOPE).elements?.[0] ?? null);
  return canonicalJson(withoutCaptures(reparsed)) === canonicalJson(withoutCaptures(formatting))
    ? sourceXml
    : null;
};

// ============================================================================
// MEASUREMENT SERIALIZATION
// ============================================================================

/**
 * Serialize a table measurement (width, height)
 */
function serializeMeasurement(
  measurement: TableMeasurement | undefined,
  elementName: string,
): string {
  if (!measurement) {
    return "";
  }

  const attrs: string[] = [`w:w="${intAttr(measurement.value)}"`, `w:type="${measurement.type}"`];

  return `<w:${elementName} ${attrs.join(" ")}/>`;
}

// ============================================================================
// BORDER SERIALIZATION
// ============================================================================

/**
 * Serialize table borders (w:tblBorders or w:tcBorders)
 */
function serializeTableBorderParts(borders: TableBorders): string[] {
  const parts: string[] = [];

  const appendBorder = (border: BorderSpec | undefined, name: string): void => {
    const xml = serializeBorder(border, name);
    if (xml) {
      parts.push(xml);
    }
  };

  appendBorder(borders.top, "top");
  appendBorder(borders.left, "left");
  appendBorder(borders.bottom, "bottom");
  appendBorder(borders.right, "right");
  appendBorder(borders.insideH, "insideH");
  appendBorder(borders.insideV, "insideV");

  return parts;
}

function serializeTableBorders(borders: TableBorders | undefined, elementName: string): string {
  if (!borders) {
    return "";
  }

  const parts = serializeTableBorderParts(borders);

  if (parts.length === 0) {
    return "";
  }

  return `<w:${elementName}>${parts.join("")}</w:${elementName}>`;
}

function serializeTableCellBorders(borders: TableCellBorders | undefined): string {
  if (!borders) {
    return "";
  }

  const parts = serializeTableBorderParts(borders);
  const topLeftToBottomRight = serializeBorder(borders.topLeftToBottomRight, "tl2br");
  if (topLeftToBottomRight) {
    parts.push(topLeftToBottomRight);
  }
  const topRightToBottomLeft = serializeBorder(borders.topRightToBottomLeft, "tr2bl");
  if (topRightToBottomLeft) {
    parts.push(topRightToBottomLeft);
  }

  return parts.length > 0 ? `<w:tcBorders>${parts.join("")}</w:tcBorders>` : "";
}

// ============================================================================
// CELL MARGINS SERIALIZATION
// ============================================================================

/**
 * Serialize cell margins (w:tblCellMar or w:tcMar)
 */
function serializeCellMargins(margins: CellMargins | undefined, elementName: string): string {
  if (!margins) {
    return "";
  }

  const parts: string[] = [];

  if (margins.top) {
    parts.push(serializeMeasurement(margins.top, "top"));
  }

  if (margins.left) {
    parts.push(serializeMeasurement(margins.left, "left"));
  }

  if (margins.bottom) {
    parts.push(serializeMeasurement(margins.bottom, "bottom"));
  }

  if (margins.right) {
    parts.push(serializeMeasurement(margins.right, "right"));
  }

  if (parts.length === 0) {
    return "";
  }

  return `<w:${elementName}>${parts.join("")}</w:${elementName}>`;
}

// ============================================================================
// SHADING SERIALIZATION
// ============================================================================

/**
 * Serialize shading properties (w:shd)
 */
function serializeShading(shading: ShadingProperties | undefined): string {
  if (!shading) {
    return "";
  }

  const attrs: string[] = [];

  // Pattern/val
  if (shading.pattern) {
    attrs.push(`w:val="${escapeXml(shading.pattern)}"`);
  } else {
    attrs.push('w:val="clear"');
  }

  // Color (pattern color)
  if (shading.color?.rgb && isValidHexColor(shading.color.rgb)) {
    attrs.push(`w:color="${escapeXml(shading.color.rgb)}"`);
  } else if (shading.color?.auto) {
    attrs.push('w:color="auto"');
  }

  // Fill (background color)
  if (shading.fill?.rgb && isValidHexColor(shading.fill.rgb)) {
    attrs.push(`w:fill="${escapeXml(shading.fill.rgb)}"`);
  } else if (shading.fill?.auto) {
    attrs.push('w:fill="auto"');
  }

  // Theme fill
  if (shading.fill?.themeColor) {
    attrs.push(`w:themeFill="${escapeXml(shading.fill.themeColor)}"`);
  }

  if (shading.fill?.themeTint) {
    attrs.push(`w:themeFillTint="${escapeXml(shading.fill.themeTint)}"`);
  }

  if (shading.fill?.themeShade) {
    attrs.push(`w:themeFillShade="${escapeXml(shading.fill.themeShade)}"`);
  }

  if (attrs.length === 0) {
    return "";
  }

  return `<w:shd ${attrs.join(" ")}/>`;
}

// ============================================================================
// TABLE LOOK SERIALIZATION
// ============================================================================

/**
 * Serialize table look flags (w:tblLook)
 */
function serializeTableLook(look: TableLook | undefined): string {
  if (!look) {
    return "";
  }

  const attrs: string[] = [];

  if (look.firstRow) {
    attrs.push('w:firstRow="1"');
  }

  if (look.lastRow) {
    attrs.push('w:lastRow="1"');
  }

  if (look.firstColumn) {
    attrs.push('w:firstColumn="1"');
  }

  if (look.lastColumn) {
    attrs.push('w:lastColumn="1"');
  }

  if (look.noHBand) {
    attrs.push('w:noHBand="1"');
  }

  if (look.noVBand) {
    attrs.push('w:noVBand="1"');
  }

  if (attrs.length === 0) {
    return "";
  }

  return `<w:tblLook ${attrs.join(" ")}/>`;
}

// ============================================================================
// FLOATING TABLE PROPERTIES SERIALIZATION
// ============================================================================

/**
 * Serialize floating table properties (w:tblpPr)
 */
function serializeFloatingTableProperties(floating: FloatingTableProperties | undefined): string {
  if (!floating) {
    return "";
  }

  const attrs: string[] = [];

  if (floating.horzAnchor) {
    attrs.push(`w:horzAnchor="${floating.horzAnchor}"`);
  }

  if (floating.vertAnchor) {
    attrs.push(`w:vertAnchor="${floating.vertAnchor}"`);
  }

  if (floating.tblpX !== undefined) {
    attrs.push(`w:tblpX="${intAttr(floating.tblpX)}"`);
  }

  if (floating.tblpXSpec) {
    attrs.push(`w:tblpXSpec="${floating.tblpXSpec}"`);
  }

  if (floating.tblpY !== undefined) {
    attrs.push(`w:tblpY="${intAttr(floating.tblpY)}"`);
  }

  if (floating.tblpYSpec) {
    attrs.push(`w:tblpYSpec="${floating.tblpYSpec}"`);
  }

  if (floating.topFromText !== undefined) {
    attrs.push(`w:topFromText="${intAttr(floating.topFromText)}"`);
  }

  if (floating.bottomFromText !== undefined) {
    attrs.push(`w:bottomFromText="${intAttr(floating.bottomFromText)}"`);
  }

  if (floating.leftFromText !== undefined) {
    attrs.push(`w:leftFromText="${intAttr(floating.leftFromText)}"`);
  }

  if (floating.rightFromText !== undefined) {
    attrs.push(`w:rightFromText="${intAttr(floating.rightFromText)}"`);
  }

  if (attrs.length === 0) {
    return "";
  }

  return `<w:tblpPr ${attrs.join(" ")}/>`;
}

// ============================================================================
// TABLE PROPERTIES SERIALIZATION (w:tblPr)
// ============================================================================

/**
 * Serialize table formatting properties (w:tblPr)
 */
export function serializeTableFormatting(
  formatting: TableFormatting | undefined,
  propertyChanges?: TablePropertyChange[],
): string {
  // Nothing in the model moved, so the element goes back as it arrived —
  // conditional-format flags, producer-specific children and all — with the
  // revisions the model carries written into it. See `TableFormatting.sourceXml`.
  const tableSource = verifiedSourceXml(formatting, parseTableProperties);
  if (tableSource !== null) {
    return withRevisionChildren(
      tableSource,
      (propertyChanges ?? []).map((change) => serializeTablePropertyChange(change)),
    );
  }

  const parts: string[] = [];

  // CT_TblPrBase is a SEQUENCE (ECMA-376 §17.4.60), so the children are
  // written in the order it declares: tblStyle, tblpPr, tblOverlap,
  // bidiVisual, tblW, jc, tblCellSpacing, tblInd, tblBorders, shd, tblLayout,
  // tblCellMar, tblLook. A consumer validating the part refuses one written in
  // any other order.
  if (formatting) {
    if (formatting.styleId) {
      parts.push(`<w:tblStyle w:val="${escapeXml(formatting.styleId)}"/>`);
    }

    const floatingXml = serializeFloatingTableProperties(formatting.floating);
    if (floatingXml) {
      parts.push(floatingXml);
    }

    if (formatting.overlap) {
      parts.push(`<w:tblOverlap w:val="${formatting.overlap}"/>`);
    }

    if (formatting.bidi !== undefined) {
      parts.push(formatting.bidi ? "<w:bidiVisual/>" : '<w:bidiVisual w:val="0"/>');
    }

    const widthXml = serializeMeasurement(formatting.width, "tblW");
    if (widthXml) {
      parts.push(widthXml);
    }

    if (formatting.justification) {
      parts.push(`<w:jc w:val="${formatting.justification}"/>`);
    }

    const cellSpacingXml = serializeMeasurement(formatting.cellSpacing, "tblCellSpacing");
    if (cellSpacingXml) {
      parts.push(cellSpacingXml);
    }

    const indentXml = serializeMeasurement(formatting.indent, "tblInd");
    if (indentXml) {
      parts.push(indentXml);
    }

    const bordersXml = serializeTableBorders(formatting.borders, "tblBorders");
    if (bordersXml) {
      parts.push(bordersXml);
    }

    const shadingXml = serializeShading(formatting.shading);
    if (shadingXml) {
      parts.push(shadingXml);
    }

    if (formatting.layout) {
      parts.push(`<w:tblLayout w:type="${formatting.layout}"/>`);
    }

    const marginsXml = serializeCellMargins(formatting.cellMargins, "tblCellMar");
    if (marginsXml) {
      parts.push(marginsXml);
    }

    const lookXml = serializeTableLook(formatting.look);
    if (lookXml) {
      parts.push(lookXml);
    }
  }

  if (propertyChanges && propertyChanges.length > 0) {
    parts.push(...propertyChanges.map((change) => serializeTablePropertyChange(change)));
  }

  if (parts.length === 0) {
    return "";
  }

  return `<w:tblPr>${parts.join("")}</w:tblPr>`;
}

function extractTblPrInner(tblPrXml: string): string {
  if (!tblPrXml.startsWith("<w:tblPr>") || !tblPrXml.endsWith("</w:tblPr>")) {
    return "";
  }
  return tblPrXml.slice("<w:tblPr>".length, -"</w:tblPr>".length);
}

function serializeTablePropertyChange(change: TablePropertyChange): string {
  const attrs = serializeTrackedChangeAttributes(change.info);
  const previousTblPrXml = serializeTableFormatting(change.previousFormatting) || "<w:tblPr/>";
  const previousTblPrInner = extractTblPrInner(previousTblPrXml);
  const normalizedPreviousTblPr =
    previousTblPrInner.length > 0 ? `<w:tblPr>${previousTblPrInner}</w:tblPr>` : "<w:tblPr/>";

  return `<w:tblPrChange ${attrs}>${normalizedPreviousTblPr}</w:tblPrChange>`;
}

// ============================================================================
// TABLE ROW PROPERTIES SERIALIZATION (w:trPr)
// ============================================================================

/**
 * Serialize table row formatting properties (w:trPr)
 */
export function serializeTableRowFormatting(
  formatting: TableRowFormatting | undefined,
  propertyChanges?: TableRowPropertyChange[],
  structuralChange?: TableStructuralChangeInfo,
): string {
  // See `serializeTableFormatting`: the source element is written back while
  // the model holds what it was parsed into, with the revisions spliced in.
  const rowSource = verifiedSourceXml(formatting, parseTableRowProperties);
  if (rowSource !== null) {
    return withRevisionChildren(rowSource, [
      ...rowStructuralChangeXml(structuralChange),
      ...(propertyChanges ?? []).map((change) => serializeTableRowPropertyChange(change)),
    ]);
  }

  const parts: string[] = [];

  if (formatting) {
    const cnfStyleXml = serializeConditionalFormatStyle(formatting.conditionalFormat);
    if (cnfStyleXml) {
      parts.push(cnfStyleXml);
    }

    if (formatting.gridBefore) {
      parts.push(`<w:gridBefore w:val="${intAttr(formatting.gridBefore)}"/>`);
    }

    if (formatting.widthBefore) {
      parts.push(
        `<w:wBefore w:w="${intAttr(formatting.widthBefore.value)}" w:type="${formatting.widthBefore.type}"/>`,
      );
    }

    if (formatting.gridAfter) {
      parts.push(`<w:gridAfter w:val="${intAttr(formatting.gridAfter)}"/>`);
    }

    if (formatting.widthAfter) {
      parts.push(
        `<w:wAfter w:w="${intAttr(formatting.widthAfter.value)}" w:type="${formatting.widthAfter.type}"/>`,
      );
    }

    // Can't split
    if (formatting.cantSplit) {
      parts.push("<w:cantSplit/>");
    }

    // Header row
    if (formatting.header) {
      parts.push("<w:tblHeader/>");
    }

    // Row height
    if (formatting.height) {
      const attrs: string[] = [`w:val="${intAttr(formatting.height.value)}"`];

      if (formatting.heightRule) {
        attrs.push(`w:hRule="${formatting.heightRule}"`);
      }

      parts.push(`<w:trHeight ${attrs.join(" ")}/>`);
    }

    // Row justification
    if (formatting.justification) {
      parts.push(`<w:jc w:val="${formatting.justification}"/>`);
    }

    // Hidden
    if (formatting.hidden) {
      parts.push("<w:hidden/>");
    }
  }

  parts.push(...rowStructuralChangeXml(structuralChange));

  if (propertyChanges && propertyChanges.length > 0) {
    parts.push(...propertyChanges.map((change) => serializeTableRowPropertyChange(change)));
  }

  if (parts.length === 0) {
    return "";
  }

  return `<w:trPr>${parts.join("")}</w:trPr>`;
}

function extractTrPrInner(trPrXml: string): string {
  if (!trPrXml.startsWith("<w:trPr>") || !trPrXml.endsWith("</w:trPr>")) {
    return "";
  }
  return trPrXml.slice("<w:trPr>".length, -"</w:trPr>".length);
}

function serializeTableRowPropertyChange(change: TableRowPropertyChange): string {
  const attrs = serializeTrackedChangeAttributes(change.info);
  const previousTrPrXml = serializeTableRowFormatting(change.previousFormatting) || "<w:trPr/>";
  const previousTrPrInner = extractTrPrInner(previousTrPrXml);
  const normalizedPreviousTrPr =
    previousTrPrInner.length > 0 ? `<w:trPr>${previousTrPrInner}</w:trPr>` : "<w:trPr/>";

  return `<w:trPrChange ${attrs}>${normalizedPreviousTrPr}</w:trPrChange>`;
}

// ============================================================================
// CONDITIONAL FORMAT STYLE SERIALIZATION
// ============================================================================

/**
 * Serialize conditional format style (w:cnfStyle)
 */
function serializeConditionalFormatStyle(style: ConditionalFormatStyle | undefined): string {
  if (!style) {
    return "";
  }

  // Build the 12-character binary string
  const bits = [
    style.firstRow ? "1" : "0",
    style.lastRow ? "1" : "0",
    style.firstColumn ? "1" : "0",
    style.lastColumn ? "1" : "0",
    style.oddVBand ? "1" : "0",
    style.evenVBand ? "1" : "0",
    style.oddHBand ? "1" : "0",
    style.evenHBand ? "1" : "0",
    style.nwCell ? "1" : "0",
    style.neCell ? "1" : "0",
    style.swCell ? "1" : "0",
    style.seCell ? "1" : "0",
  ];

  const val = bits.join("");

  // Only serialize if any bits are set
  if (val === "000000000000") {
    return "";
  }

  return `<w:cnfStyle w:val="${val}"/>`;
}

// ============================================================================
// TABLE CELL PROPERTIES SERIALIZATION (w:tcPr)
// ============================================================================

/**
 * Serialize table cell formatting properties (w:tcPr)
 */
export function serializeTableCellFormatting(
  formatting: TableCellFormatting | undefined,
  propertyChanges?: TableCellPropertyChange[],
  structuralChange?: TableStructuralChangeInfo,
): string {
  // See `serializeTableFormatting`: the source element is written back while
  // the model holds what it was parsed into, with the revisions spliced in.
  const cellSource = verifiedSourceXml(formatting, parseTableCellProperties);
  if (cellSource !== null) {
    return withRevisionChildren(cellSource, [
      ...cellStructuralChangeXml(structuralChange),
      ...(propertyChanges ?? []).map((change) => serializeTableCellPropertyChange(change)),
    ]);
  }

  const parts: string[] = [];

  if (formatting) {
    // Conditional format style
    const cnfStyleXml = serializeConditionalFormatStyle(formatting.conditionalFormat);
    if (cnfStyleXml) {
      parts.push(cnfStyleXml);
    }

    // Cell width
    const widthXml = serializeMeasurement(formatting.width, "tcW");
    if (widthXml) {
      parts.push(widthXml);
    }

    // Grid span (horizontal merge)
    if (formatting.gridSpan && formatting.gridSpan > 1) {
      parts.push(`<w:gridSpan w:val="${intAttr(formatting.gridSpan)}"/>`);
    }

    // Vertical merge
    if (formatting.vMerge) {
      if (formatting.vMerge === "restart") {
        parts.push('<w:vMerge w:val="restart"/>');
      } else {
        // continue is the default when w:vMerge has no value
        parts.push("<w:vMerge/>");
      }
    }

    // Cell borders
    const bordersXml = serializeTableCellBorders(formatting.borders);
    if (bordersXml) {
      parts.push(bordersXml);
    }

    // Shading
    const shadingXml = serializeShading(formatting.shading);
    if (shadingXml) {
      parts.push(shadingXml);
    }

    // No wrap
    if (formatting.noWrap) {
      parts.push("<w:noWrap/>");
    }

    // Cell margins
    const marginsXml = serializeCellMargins(formatting.margins, "tcMar");
    if (marginsXml) {
      parts.push(marginsXml);
    }

    // Text direction
    if (formatting.textDirection) {
      parts.push(`<w:textDirection w:val="${formatting.textDirection}"/>`);
    }

    // Fit text
    if (formatting.fitText) {
      parts.push("<w:tcFitText/>");
    }

    // Vertical alignment
    if (formatting.verticalAlign) {
      parts.push(`<w:vAlign w:val="${formatting.verticalAlign}"/>`);
    }

    // Hide mark
    if (formatting.hideMark === true) {
      parts.push("<w:hideMark/>");
    } else if (formatting.hideMark === false) {
      parts.push('<w:hideMark w:val="off"/>');
    }
  }

  parts.push(...cellStructuralChangeXml(structuralChange));

  if (propertyChanges && propertyChanges.length > 0) {
    parts.push(...propertyChanges.map((change) => serializeTableCellPropertyChange(change)));
  }

  if (parts.length === 0) {
    return "";
  }

  return `<w:tcPr>${parts.join("")}</w:tcPr>`;
}

function extractTcPrInner(tcPrXml: string): string {
  if (!tcPrXml.startsWith("<w:tcPr>") || !tcPrXml.endsWith("</w:tcPr>")) {
    return "";
  }
  return tcPrXml.slice("<w:tcPr>".length, -"</w:tcPr>".length);
}

function serializeTableCellPropertyChange(change: TableCellPropertyChange): string {
  const attrs = serializeTrackedChangeAttributes(change.info);
  const previousTcPrXml = serializeTableCellFormatting(change.previousFormatting) || "<w:tcPr/>";
  const previousTcPrInner = extractTcPrInner(previousTcPrXml);
  const normalizedPreviousTcPr =
    previousTcPrInner.length > 0 ? `<w:tcPr>${previousTcPrInner}</w:tcPr>` : "<w:tcPr/>";

  return `<w:tcPrChange ${attrs}>${normalizedPreviousTcPr}</w:tcPrChange>`;
}

// ============================================================================
// TABLE GRID SERIALIZATION
// ============================================================================

/**
 * Columns the grid has to declare: the widest row's span total, because a grid
 * narrower than a row leaves cells with no column to sit in.
 */
function gridColumnCount(table: Table): number {
  let widest = 0;
  for (const row of table.rows) {
    let columns = 0;
    for (const cell of row.cells) {
      columns += cell.formatting?.gridSpan ?? 1;
    }
    widest = Math.max(widest, columns);
  }
  return widest;
}

/**
 * Serialize table grid (w:tblGrid).
 *
 * `w:tblGrid` is required on every `w:tbl` and `w:w` is optional on a
 * `w:gridCol`, so a table whose column widths were never measured — one the
 * comparison creates, say — still declares a grid, just without widths.
 */
function serializeTableGrid(table: Table): string {
  const columnWidths = table.columnWidths;
  // The grid as it arrived, while it still states the widths the model holds.
  // A `w:tblGridChange` sits inside it and nothing else carries one.
  const gridSourceXml = table.formatting?.gridSourceXml;
  if (gridSourceXml !== undefined) {
    const parsed = parseTableGrid(
      parseXml(gridSourceXml, OOXML_NAMESPACE_SCOPE).elements?.[0] ?? null,
    );
    if (canonicalJson(parsed) === canonicalJson(columnWidths)) {
      return gridSourceXml;
    }
  }
  if (columnWidths && columnWidths.length > 0) {
    return `<w:tblGrid>${columnWidths.map((w) => `<w:gridCol w:w="${intAttr(w)}"/>`).join("")}</w:tblGrid>`;
  }

  const columns = gridColumnCount(table);
  return columns === 0
    ? "<w:tblGrid/>"
    : `<w:tblGrid>${"<w:gridCol/>".repeat(columns)}</w:tblGrid>`;
}

// ============================================================================
// CELL CONTENT SERIALIZATION
// ============================================================================

/**
 * Serialize cell content (paragraphs, nested tables)
 */
function serializeCellContent(
  content: (Paragraph | Table)[],
  serializeParagraph: ParagraphSerializer,
): string {
  const parts: string[] = [];

  for (const item of content) {
    if (item.type === "paragraph") {
      parts.push(serializeParagraph(item));
    } else {
      parts.push(serializeTable(item, serializeParagraph));
    }
  }

  // Ensure at least one empty paragraph (Word requires this)
  if (parts.length === 0) {
    parts.push("<w:p/>");
  }

  return parts.join("");
}

// ============================================================================
// TABLE CELL SERIALIZATION
// ============================================================================

/**
 * Serialize a table cell (w:tc)
 */
export function serializeTableCell(
  cell: TableCell,
  serializeParagraph: ParagraphSerializer,
): string {
  const parts: string[] = [];

  // Cell properties
  const tcPrXml = serializeTableCellFormatting(
    cell.formatting,
    cell.propertyChanges,
    cell.structuralChange,
  );
  if (tcPrXml) {
    parts.push(tcPrXml);
  }

  // Cell content
  parts.push(serializeCellContent(cell.content, serializeParagraph));

  return `<w:tc>${parts.join("")}</w:tc>`;
}

// ============================================================================
// TABLE ROW SERIALIZATION
// ============================================================================

/**
 * Serialize a table row (w:tr)
 */
export function serializeTableRow(row: TableRow, serializeParagraph: ParagraphSerializer): string {
  const parts: string[] = [];

  // Row properties
  const trPrXml = serializeTableRowFormatting(
    row.formatting,
    row.propertyChanges,
    row.structuralChange,
  );
  if (trPrXml) {
    parts.push(trPrXml);
  }

  // Cells
  for (const cell of row.cells) {
    parts.push(serializeTableCell(cell, serializeParagraph));
  }

  return `<w:tr>${parts.join("")}</w:tr>`;
}

// ============================================================================
// MAIN TABLE SERIALIZATION
// ============================================================================

/**
 * Serialize a table to OOXML XML (w:tbl)
 *
 * @param table - The table to serialize
 * @returns XML string for the table
 */
export function serializeTable(table: Table, serializeParagraph: ParagraphSerializer): string {
  // `w:tbl` is `w:tblPr, w:tblGrid, (rows)*`: both properties and grid are
  // required and precede every row. Emitting them only when the model carried
  // something to put in them made a table with neither — the shape a tracked
  // table insertion produces — open with a `w:tr` the content model has no
  // place for. Empty elements are the valid way to say "nothing here".
  const tblPrXml =
    serializeTableFormatting(table.formatting, table.propertyChanges) || "<w:tblPr/>";
  const parts: string[] = [tblPrXml, serializeTableGrid(table)];

  for (const row of table.rows) {
    parts.push(serializeTableRow(row, serializeParagraph));
  }

  return `<w:tbl>${parts.join("")}</w:tbl>`;
}

/**
 * Serialize multiple tables to OOXML XML
 *
 * @param tables - The tables to serialize
 * @returns XML string for all tables
 */
export function serializeTables(tables: Table[], serializeParagraph: ParagraphSerializer): string {
  return tables.map((table) => serializeTable(table, serializeParagraph)).join("");
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Check if a table has any rows
 */
export function hasTableRows(table: Table): boolean {
  return table.rows.length > 0;
}

/**
 * Check if a table has formatting
 */
export function hasTableFormatting(table: Table): boolean {
  return table.formatting !== undefined && Object.keys(table.formatting).length > 0;
}

/**
 * Check if a row has any cells
 */
export function hasRowCells(row: TableRow): boolean {
  return row.cells.length > 0;
}

/**
 * Check if a row has formatting
 */
export function hasRowFormatting(row: TableRow): boolean {
  return row.formatting !== undefined && Object.keys(row.formatting).length > 0;
}

/**
 * Check if a cell has any content
 */
export function hasCellContent(cell: TableCell): boolean {
  return cell.content.length > 0;
}

/**
 * Check if a cell has formatting
 */
export function hasCellFormatting(cell: TableCell): boolean {
  return cell.formatting !== undefined && Object.keys(cell.formatting).length > 0;
}

/**
 * Get the number of columns in a table
 */
export function getTableColumnCount(table: Table): number {
  if (table.columnWidths && table.columnWidths.length > 0) {
    return table.columnWidths.length;
  }

  if (table.rows.length === 0) {
    return 0;
  }

  // Count cells in first row, accounting for grid span
  // SAFETY: rows.length > 0 verified above
  return table.rows[0]!.cells.reduce((count, cell) => count + (cell.formatting?.gridSpan ?? 1), 0);
}

/**
 * Get the number of rows in a table
 */
export function getTableRowCount(table: Table): number {
  return table.rows.length;
}

/**
 * Create an empty table
 */
export function createEmptyTable(rows: number = 1, cols: number = 1): Table {
  const tableRows: TableRow[] = [];

  for (let r = 0; r < rows; r++) {
    const cells: TableCell[] = [];
    for (let c = 0; c < cols; c++) {
      cells.push({
        type: "tableCell",
        content: [{ type: "paragraph", content: [] }],
      });
    }
    tableRows.push({
      type: "tableRow",
      cells,
    });
  }

  return {
    type: "table",
    rows: tableRows,
  };
}

/**
 * Create a table cell with text content
 */
export function createTextCell(text: string, formatting?: TableCellFormatting): TableCell {
  return {
    type: "tableCell",
    ...(formatting !== undefined ? { formatting } : {}),
    content: [
      {
        type: "paragraph",
        content: [
          {
            type: "run",
            content: [{ type: "text", text }],
          },
        ],
      },
    ],
  };
}
