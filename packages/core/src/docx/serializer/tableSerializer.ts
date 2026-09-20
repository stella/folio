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
  TablePropertyExceptionChange,
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
  TableCellBlock,
} from "../../types/document";
import { canonicalJson } from "../../utils/canonicalJson";
import { isValidHexColor } from "../../utils/colorResolver";
import {
  parseTableCellProperties,
  parseTableGrid,
  parseTableProperties,
  parseTablePropertyExceptions,
  parseTableRowProperties,
} from "../tableParser";
import { serializePreservedAttributes } from "../attributeRemainder";
import { serializeSequenceChildren, serializeWithPreservedChildren } from "../containerChildren";
import { TABLE_LOOK_FLAGS } from "../tableLook";
import { sanitizeCapturedXmlElement } from "../verbatimCapture";
import { NAMESPACES, OOXML_NAMESPACE_SCOPE, parseXml, type XmlElement } from "../xmlParser";
import { serializeBorder } from "./borderSerializer";
import { serializeTrackedChangeAttributes } from "./trackedChangeAttributes";
import { intAttr } from "./xmlUtils";
import { escapeXmlAttribute } from "@stll/docx-core";

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

/**
 * A structural revision as its own declared child, or nothing.
 *
 * The element name is part of the answer rather than something the caller
 * restates: the property set is written by declared-child name, and a
 * `<w:del/>` filed under `ins` would land at the wrong ordinal.
 */
type StructuralChangeXml<Name extends string> = { name: Name; xml: string } | undefined;

/** `w:trPr/w:ins` | `w:trPr/w:del`, the row's own structural revision. */
const rowStructuralChangeXml = (
  change: TableStructuralChangeInfo | undefined,
): StructuralChangeXml<"ins" | "del"> => {
  if (change?.type === "tableRowInsertion") {
    return { name: "ins", xml: `<w:ins ${serializeTrackedChangeAttributes(change.info)}/>` };
  }
  if (change?.type === "tableRowDeletion") {
    return { name: "del", xml: `<w:del ${serializeTrackedChangeAttributes(change.info)}/>` };
  }
  return undefined;
};

/** `w:tcPr/w:cellIns` | `w:cellDel` | `w:cellMerge`, the cell's own. */
const cellStructuralChangeXml = (
  change: TableStructuralChangeInfo | undefined,
): StructuralChangeXml<"cellIns" | "cellDel" | "cellMerge"> => {
  if (change?.type === "tableCellInsertion") {
    return {
      name: "cellIns",
      xml: `<w:cellIns ${serializeTrackedChangeAttributes(change.info)}/>`,
    };
  }
  if (change?.type === "tableCellDeletion") {
    return {
      name: "cellDel",
      xml: `<w:cellDel ${serializeTrackedChangeAttributes(change.info)}/>`,
    };
  }
  if (change?.type !== "tableCellMerge") {
    return undefined;
  }
  const attrs = [serializeTrackedChangeAttributes(change.info)];
  if (change.verticalMerge) {
    attrs.push(`w:vMerge="${change.verticalMerge === "continue" ? "cont" : "rest"}"`);
  }
  if (change.verticalMergeOriginal) {
    attrs.push(`w:vMergeOrig="${change.verticalMergeOriginal === "continue" ? "cont" : "rest"}"`);
  }
  return { name: "cellMerge", xml: `<w:cellMerge ${attrs.join(" ")}/>` };
};

/** The one structural revision a property set carries, as a modelled entry. */
const structuralChangeEntry = <Name extends string>(
  name: Name,
  change: StructuralChangeXml<Name>,
): readonly [name: Name, xml: string] => [name, change?.name === name ? change.xml : ""];

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
    attrs.push(`w:val="${escapeXmlAttribute(shading.pattern)}"`);
  } else {
    attrs.push('w:val="clear"');
  }

  // Color (pattern color)
  if (shading.color?.rgb && isValidHexColor(shading.color.rgb)) {
    attrs.push(`w:color="${escapeXmlAttribute(shading.color.rgb)}"`);
  } else if (shading.color?.auto) {
    attrs.push('w:color="auto"');
  }

  // Fill (background color)
  if (shading.fill?.rgb && isValidHexColor(shading.fill.rgb)) {
    attrs.push(`w:fill="${escapeXmlAttribute(shading.fill.rgb)}"`);
  } else if (shading.fill?.auto) {
    attrs.push('w:fill="auto"');
  }

  // Theme fill
  if (shading.fill?.themeColor) {
    attrs.push(`w:themeFill="${escapeXmlAttribute(shading.fill.themeColor)}"`);
  }

  if (shading.fill?.themeTint) {
    attrs.push(`w:themeFillTint="${escapeXmlAttribute(shading.fill.themeTint)}"`);
  }

  if (shading.fill?.themeShade) {
    attrs.push(`w:themeFillShade="${escapeXmlAttribute(shading.fill.themeShade)}"`);
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
 * Serialize a `w:tblLook`: what the author wrote, in the order they wrote it.
 *
 * A flag is written when it was authored, true or false. Writing only the true
 * ones drops an explicit `w:lastRow="0"`, and an absent flag is not an off one:
 * it falls back to `w:val`'s bit, so the two spell different tables under the
 * same table style. `w:val` leads, then the flags in `CT_TblLook` order, which
 * is the order Word writes them.
 */
function serializeTableLook(look: TableLook | undefined): string {
  if (!look) {
    return "";
  }

  const attrs: string[] = [];

  if (look.val !== undefined) {
    attrs.push(`w:val="${escapeXmlAttribute(look.val)}"`);
  }

  for (const flag of TABLE_LOOK_FLAGS) {
    const stated = look[flag];
    if (stated !== undefined) {
      attrs.push(`w:${flag}="${stated ? "1" : "0"}"`);
    }
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

/** `<w:name w:val="…"/>`, or nothing when the model holds no value. */
const tagWithVal = (name: string, value: string | undefined): string =>
  value === undefined ? "" : `<w:${name} w:val="${escapeXmlAttribute(value)}"/>`;

const numberTag = (name: string, value: number | undefined): string =>
  value === undefined ? "" : `<w:${name} w:val="${intAttr(value)}"/>`;

/** `CT_OnOff`: present means on, and an explicit off is not an absent one. */
const serializeOnOffElement = (value: boolean | undefined, name: string): string => {
  if (value === undefined) {
    return "";
  }
  return value ? `<w:${name}/>` : `<w:${name} w:val="0"/>`;
};

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

  // `CT_TblPr` is a sequence (ECMA-376 §17.4.60) and a consumer refuses a
  // `w:tblPr` whose children are in any other order. The order is the
  // generated declared-child list rather than the order of the statements
  // below, so it cannot drift from the schema the census and the contract read.
  const parts = serializeSequenceChildren({
    container: "table-properties",
    modelled: [
      ["tblStyle", formatting?.styleId ? tagWithVal("tblStyle", formatting.styleId) : ""],
      ["tblpPr", serializeFloatingTableProperties(formatting?.floating)],
      ["tblOverlap", formatting?.overlap ? `<w:tblOverlap w:val="${formatting.overlap}"/>` : ""],
      ["bidiVisual", serializeOnOffElement(formatting?.bidi, "bidiVisual")],
      ["tblStyleRowBandSize", numberTag("tblStyleRowBandSize", formatting?.rowBandSize)],
      ["tblStyleColBandSize", numberTag("tblStyleColBandSize", formatting?.columnBandSize)],
      ["tblW", serializeMeasurement(formatting?.width, "tblW")],
      ["jc", formatting?.justification ? `<w:jc w:val="${formatting.justification}"/>` : ""],
      ["tblCellSpacing", serializeMeasurement(formatting?.cellSpacing, "tblCellSpacing")],
      ["tblInd", serializeMeasurement(formatting?.indent, "tblInd")],
      ["tblBorders", serializeTableBorders(formatting?.borders, "tblBorders")],
      ["shd", serializeShading(formatting?.shading)],
      ["tblLayout", formatting?.layout ? `<w:tblLayout w:type="${formatting.layout}"/>` : ""],
      ["tblCellMar", serializeCellMargins(formatting?.cellMargins, "tblCellMar")],
      ["tblLook", serializeTableLook(formatting?.look)],
      ["tblCaption", tagWithVal("tblCaption", formatting?.caption)],
      ["tblDescription", tagWithVal("tblDescription", formatting?.description)],
      [
        "tblPrChange",
        (propertyChanges ?? []).map((change) => serializeTablePropertyChange(change)).join(""),
      ],
    ],
    preserved: formatting?.preserved,
  });

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
// TABLE PROPERTY EXCEPTIONS SERIALIZATION (w:tblPrEx)
// ============================================================================

/**
 * Serialize a row's table property exceptions (w:tblPrEx).
 *
 * `CT_TblPrEx` is a sequence like `CT_TblPr`, and the nine children it shares
 * with it are written by the same statements: the modelled half is keyed by
 * element name and `serializeSequenceChildren` puts it in the generated order,
 * so the exception and the property it overrides cannot come out spelled two
 * different ways.
 */
function serializeTablePropertyExceptions(
  exceptions: TableFormatting | undefined,
  propertyChanges?: TablePropertyExceptionChange[],
): string {
  if (exceptions === undefined && (propertyChanges ?? []).length === 0) {
    return "";
  }

  // See `serializeTableFormatting`: the source element is written back while
  // the model holds what it was parsed into, with the revisions spliced in.
  const exceptionSource = verifiedSourceXml(exceptions, parseTablePropertyExceptions);
  if (exceptionSource !== null) {
    return withRevisionChildren(
      exceptionSource,
      (propertyChanges ?? []).map((change) => serializeTablePropertyExceptionChange(change)),
    );
  }

  const parts = serializeSequenceChildren({
    container: "table-property-exceptions",
    modelled: [
      ["tblW", serializeMeasurement(exceptions?.width, "tblW")],
      ["jc", exceptions?.justification ? `<w:jc w:val="${exceptions.justification}"/>` : ""],
      ["tblCellSpacing", serializeMeasurement(exceptions?.cellSpacing, "tblCellSpacing")],
      ["tblInd", serializeMeasurement(exceptions?.indent, "tblInd")],
      ["tblBorders", serializeTableBorders(exceptions?.borders, "tblBorders")],
      ["shd", serializeShading(exceptions?.shading)],
      ["tblLayout", exceptions?.layout ? `<w:tblLayout w:type="${exceptions.layout}"/>` : ""],
      ["tblCellMar", serializeCellMargins(exceptions?.cellMargins, "tblCellMar")],
      ["tblLook", serializeTableLook(exceptions?.look)],
      [
        "tblPrExChange",
        (propertyChanges ?? [])
          .map((change) => serializeTablePropertyExceptionChange(change))
          .join(""),
      ],
    ],
    preserved: exceptions?.preserved,
  });

  // `<w:tblPrEx/>` rather than nothing: the element is optional, so the row
  // that wrote an empty one said something the absent element does not.
  return parts.length === 0 ? "<w:tblPrEx/>" : `<w:tblPrEx>${parts.join("")}</w:tblPrEx>`;
}

function serializeTablePropertyExceptionChange(change: TablePropertyExceptionChange): string {
  const attrs = serializeTrackedChangeAttributes(change.info);
  const previous = serializeTablePropertyExceptions(change.previousFormatting) || "<w:tblPrEx/>";
  return `<w:tblPrExChange ${attrs}>${previous}</w:tblPrExChange>`;
}

// ============================================================================
// TABLE ROW PROPERTIES SERIALIZATION (w:trPr)
// ============================================================================

/** `w:trHeight`, whose value rides `w:val` rather than `w:w`. */
const serializeRowHeight = (formatting: TableRowFormatting | undefined): string => {
  if (!formatting?.height) {
    return "";
  }
  const attrs = [`w:val="${intAttr(formatting.height.value)}"`];
  if (formatting.heightRule) {
    attrs.push(`w:hRule="${formatting.heightRule}"`);
  }
  return `<w:trHeight ${attrs.join(" ")}/>`;
};

/**
 * Serialize table row formatting properties (w:trPr)
 */
export function serializeTableRowFormatting(
  formatting: TableRowFormatting | undefined,
  propertyChanges?: TableRowPropertyChange[],
  structuralChange?: TableStructuralChangeInfo,
): string {
  const rowStructuralChange = rowStructuralChangeXml(structuralChange);

  // See `serializeTableFormatting`: the source element is written back while
  // the model holds what it was parsed into, with the revisions spliced in.
  const rowSource = verifiedSourceXml(formatting, parseTableRowProperties);
  if (rowSource !== null) {
    return withRevisionChildren(rowSource, [
      ...(rowStructuralChange ? [rowStructuralChange.xml] : []),
      ...(propertyChanges ?? []).map((change) => serializeTableRowPropertyChange(change)),
    ]);
  }

  // `CT_TrPrBase` is a repeated choice, so the row's properties have no order
  // a consumer enforces; `CT_TrPr` closes a sequence over it, so `w:ins`,
  // `w:del` and `w:trPrChange` do have to come last. The generated
  // declared-child list is both, and writing by it puts the sink's captures
  // back between the same neighbours the parser read them between.
  const parts = serializeSequenceChildren({
    container: "row-properties",
    modelled: [
      ["cnfStyle", serializeConditionalFormatStyle(formatting?.conditionalFormat)],
      [
        "gridBefore",
        formatting?.gridBefore ? `<w:gridBefore w:val="${intAttr(formatting.gridBefore)}"/>` : "",
      ],
      [
        "gridAfter",
        formatting?.gridAfter ? `<w:gridAfter w:val="${intAttr(formatting.gridAfter)}"/>` : "",
      ],
      ["wBefore", serializeMeasurement(formatting?.widthBefore, "wBefore")],
      ["wAfter", serializeMeasurement(formatting?.widthAfter, "wAfter")],
      ["cantSplit", serializeOnOffElement(formatting?.cantSplit, "cantSplit")],
      ["trHeight", serializeRowHeight(formatting)],
      ["tblHeader", serializeOnOffElement(formatting?.header, "tblHeader")],
      ["jc", formatting?.justification ? `<w:jc w:val="${formatting.justification}"/>` : ""],
      ["hidden", serializeOnOffElement(formatting?.hidden, "hidden")],
      structuralChangeEntry("ins", rowStructuralChange),
      structuralChangeEntry("del", rowStructuralChange),
      [
        "trPrChange",
        (propertyChanges ?? []).map((change) => serializeTableRowPropertyChange(change)).join(""),
      ],
    ],
    preserved: formatting?.preserved,
  });

  // `<w:trPr/>` rather than nothing when the row carried one: the element is
  // optional on a `w:tr`, so a row that wrote an empty one said something an
  // absent element does not. A row that never had one has no formatting at
  // all, and still writes nothing.
  if (parts.length === 0) {
    return formatting === undefined ? "" : "<w:trPr/>";
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

/** `w:vMerge`: no `w:val` is the continuation, so only a restart states one. */
const serializeVerticalMerge = (vMerge: TableCellFormatting["vMerge"]): string => {
  if (vMerge === undefined) {
    return "";
  }
  return vMerge === "restart" ? '<w:vMerge w:val="restart"/>' : "<w:vMerge/>";
};

/**
 * `w:hideMark`, whose explicit off is spelled `off` rather than `0`.
 *
 * Both are `ST_OnOff` and the rest of this file writes `0`; the spelling is
 * pinned here by `tableParser.test.ts`, so the two live side by side until one
 * of them is chosen for the whole package.
 */
const serializeHideMark = (hideMark: boolean | undefined): string => {
  if (hideMark === undefined) {
    return "";
  }
  return hideMark ? "<w:hideMark/>" : '<w:hideMark w:val="off"/>';
};

/**
 * Serialize table cell formatting properties (w:tcPr)
 */
export function serializeTableCellFormatting(
  formatting: TableCellFormatting | undefined,
  propertyChanges?: TableCellPropertyChange[],
  structuralChange?: TableStructuralChangeInfo,
): string {
  const cellStructuralChange = cellStructuralChangeXml(structuralChange);

  // See `serializeTableFormatting`: the source element is written back while
  // the model holds what it was parsed into, with the revisions spliced in.
  const cellSource = verifiedSourceXml(formatting, parseTableCellProperties);
  if (cellSource !== null) {
    return withRevisionChildren(cellSource, [
      ...(cellStructuralChange ? [cellStructuralChange.xml] : []),
      ...(propertyChanges ?? []).map((change) => serializeTableCellPropertyChange(change)),
    ]);
  }

  // `CT_TcPrBase` is a sequence and so is every type extending it, so a
  // consumer refuses a `w:tcPr` whose children are in any other order. The
  // order is the generated declared-child list rather than the order of the
  // statements below, so it cannot drift from the schema.
  const parts = serializeSequenceChildren({
    container: "cell-properties",
    modelled: [
      ["cnfStyle", serializeConditionalFormatStyle(formatting?.conditionalFormat)],
      ["tcW", serializeMeasurement(formatting?.width, "tcW")],
      [
        "gridSpan",
        formatting?.gridSpan && formatting.gridSpan > 1
          ? `<w:gridSpan w:val="${intAttr(formatting.gridSpan)}"/>`
          : "",
      ],
      ["vMerge", serializeVerticalMerge(formatting?.vMerge)],
      ["tcBorders", serializeTableCellBorders(formatting?.borders)],
      ["shd", serializeShading(formatting?.shading)],
      ["noWrap", serializeOnOffElement(formatting?.noWrap, "noWrap")],
      ["tcMar", serializeCellMargins(formatting?.margins, "tcMar")],
      [
        "textDirection",
        formatting?.textDirection ? `<w:textDirection w:val="${formatting.textDirection}"/>` : "",
      ],
      ["tcFitText", serializeOnOffElement(formatting?.fitText, "tcFitText")],
      [
        "vAlign",
        formatting?.verticalAlign ? `<w:vAlign w:val="${formatting.verticalAlign}"/>` : "",
      ],
      ["hideMark", serializeHideMark(formatting?.hideMark)],
      structuralChangeEntry("cellIns", cellStructuralChange),
      structuralChangeEntry("cellDel", cellStructuralChange),
      structuralChangeEntry("cellMerge", cellStructuralChange),
      [
        "tcPrChange",
        (propertyChanges ?? []).map((change) => serializeTableCellPropertyChange(change)).join(""),
      ],
    ],
    preserved: formatting?.preserved,
  });

  // `<w:tcPr/>` for a cell that carried one, for the reason `w:trPr` does.
  if (parts.length === 0) {
    return formatting === undefined ? "" : "<w:tcPr/>";
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
  const previousTcPrXml =
    serializeTableCellFormatting(
      change.previousFormatting,
      undefined,
      change.previousStructuralChange,
    ) || "<w:tcPr/>";
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
  const gridSourceXml = table.formatting?.gridSourceXml;
  if (gridSourceXml !== undefined) {
    const parsed = parseTableGrid(
      parseXml(gridSourceXml, OOXML_NAMESPACE_SCOPE).elements?.[0] ?? null,
    );
    if (canonicalJson(parsed) === canonicalJson(columnWidths)) {
      return gridSourceXml;
    }
  }
  // `w:tblGridChange` records the grid a reviewer replaced. It is history, so
  // no width in the model derives it, and a rebuilt grid that leaves it out
  // accepts the reviewer's change without saying so. The schema declares it
  // after every `w:gridCol`, so it is appended.
  const gridChange = replayableGridChangeXml(table.formatting?.gridChangeXml) ?? "";
  if (columnWidths && columnWidths.length > 0) {
    const columns = columnWidths.map((w) => `<w:gridCol w:w="${intAttr(w)}"/>`).join("");
    return `<w:tblGrid>${columns}${gridChange}</w:tblGrid>`;
  }

  const columns = gridColumnCount(table);
  return columns === 0 && gridChange === ""
    ? "<w:tblGrid/>"
    : `<w:tblGrid>${"<w:gridCol/>".repeat(columns)}${gridChange}</w:tblGrid>`;
}

const GRID_CHANGE_ROOT_NAME: ReadonlySet<string> = new Set(["tblGridChange"]);
const WORDPROCESSINGML_NAMESPACE: ReadonlySet<string> = new Set([NAMESPACES.w]);

const replayableGridChangeXml = (gridChangeXml: string | undefined): string | null =>
  sanitizeCapturedXmlElement(gridChangeXml, {
    allowedLocalNames: GRID_CHANGE_ROOT_NAME,
    allowedNamespaceUris: WORDPROCESSINGML_NAMESPACE,
    inheritedNamespaceScope: OOXML_NAMESPACE_SCOPE,
  });

// ============================================================================
// CELL CONTENT SERIALIZATION
// ============================================================================

/**
 * Serialize cell content (paragraphs, nested tables)
 */
function serializeCellContent(
  content: readonly TableCellBlock[],
  serializeParagraph: ParagraphSerializer,
): string {
  const parts = content.map((block) => {
    switch (block.type) {
      case "paragraph":
        return serializeParagraph(block);
      case "table":
        return serializeTable(block, serializeParagraph);
      case "preservedBlock":
        return block.xml;
      default: {
        const unreachable: never = block;
        return unreachable;
      }
    }
  });

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

  // `CT_Row` opens with `w:tblPrEx` and only then `w:trPr`, so the exceptions
  // are written first: a row that states both in the other order is markup a
  // validating consumer refuses.
  const tblPrExXml = serializeTablePropertyExceptions(
    row.tablePropertyExceptions,
    row.tablePropertyExceptionChanges,
  );
  if (tblPrExXml) {
    parts.push(tblPrExXml);
  }

  // Row properties
  const trPrXml = serializeTableRowFormatting(
    row.formatting,
    row.propertyChanges,
    row.structuralChange,
  );
  if (trPrXml) {
    parts.push(trPrXml);
  }

  // Cells, with the row markup folio does not model back between the same
  // two of them. `w:trPr` and `w:tblPrEx` come first in the content model and
  // are written above, so the sink's index counts cells and nothing else.
  parts.push(
    serializeWithPreservedChildren(
      row.cells.map((cell) => serializeTableCell(cell, serializeParagraph)),
      row.preserved,
    ),
  );

  // The row models no attribute of its own, so every one it writes comes from
  // the remainder the parser kept.
  const attrs = serializePreservedAttributes([], row.preservedAttributes);
  return `<w:tr${attrs.length > 0 ? ` ${attrs.join(" ")}` : ""}>${parts.join("")}</w:tr>`;
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
  // Rows, with the table markup folio does not model back between the same
  // two of them. `w:tblPr` and `w:tblGrid` come first in the content model and
  // are written above, so the sink's index counts rows and nothing else.
  const parts: string[] = [
    tblPrXml,
    serializeTableGrid(table),
    serializeWithPreservedChildren(
      table.rows.map((row) => serializeTableRow(row, serializeParagraph)),
      table.preserved,
    ),
  ];

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
