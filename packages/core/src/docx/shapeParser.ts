/**
 * Shape Parser — parse DrawingML preset-geometry shapes from `<wps:wsp>`.
 *
 * Adapted from eigenpal docx-editor `shapeParser.ts`
 * (eigenpal/main:packages/core/src/docx/shapeParser.ts). Re-shaped to fit
 * folio's xml-parser helpers, picklist-based enum narrowing, and
 * exactOptionalPropertyTypes. The eigenpal version owned its own
 * fill/outline parser; folio already has `drawingUtils.parseFill` /
 * `parseOutline` so this module only extends them where gradient stops
 * and arrow-end metadata need more fidelity than the shared helpers
 * provide.
 *
 * OOXML structure (ECMA-376 §20.5.2 wordprocessingShape):
 *   w:drawing
 *     └── wp:inline | wp:anchor
 *         └── a:graphic
 *             └── a:graphicData
 *                 └── wps:wsp                  (the shape)
 *                     ├── wps:cNvPr            (id, name, descr, title)
 *                     ├── wps:spPr             (shape properties)
 *                     │   ├── a:xfrm           (size + rotation)
 *                     │   ├── a:prstGeom       (preset geometry)
 *                     │   ├── a:solidFill / a:gradFill / a:noFill
 *                     │   └── a:ln             (outline + arrowheads)
 *                     ├── wps:txbx             (optional text body)
 *                     └── wps:bodyPr           (text body properties)
 */

import type { ColorValue, ImageSize, ImageTransform, Shape, ShapeFill } from "../types/document";
import { parseDrawingAnchor } from "./drawingAnchor";
import {
  parseAnchorPosition,
  parseAnchorWrap,
  parseFill,
  parseInlineWrap,
  parseOutline,
} from "./drawingUtils";
import { parseNonVisualDrawingNames } from "./nonVisualDrawingProps";
import { narrowEnum, ShapeTypeSchema } from "./parserEnums";
import {
  findAllDeep,
  findChildByLocalName,
  findChildren,
  getAttribute,
  getChildElements,
  parseNumericAttribute,
  parseOnOffAttribute,
} from "./xmlParser";
import type { XmlElement } from "./xmlParser";

/** Convert OOXML rotation (1/60000ths of a degree) to degrees. */
function rotToDegrees(rot: string | null | undefined): number | undefined {
  if (rot === null || rot === undefined) {
    return undefined;
  }
  const val = Number.parseInt(rot, 10);
  return Number.isNaN(val) ? undefined : val / 60_000;
}

/**
 * Parse a fill with full gradient stop capture (eigenpal #21).
 * Falls through to `drawingUtils.parseFill` for solid / none cases.
 */
function parseShapeFill(spPr: XmlElement | null): ShapeFill | undefined {
  return parseFill(spPr);
}

// ---------------------------------------------------------------------------
// TRANSFORM (a:xfrm)
// ---------------------------------------------------------------------------

function parseTransform(xfrm: XmlElement | null): {
  size: ImageSize;
  transform?: ImageTransform;
} {
  if (!xfrm) {
    return { size: { width: 0, height: 0 } };
  }

  const ext = findChildByLocalName(xfrm, "ext");
  const cx = parseNumericAttribute(ext, null, "cx") ?? 0;
  const cy = parseNumericAttribute(ext, null, "cy") ?? 0;
  const size: ImageSize = { width: cx, height: cy };

  const rotation = rotToDegrees(getAttribute(xfrm, null, "rot"));
  // Authored, not truthy: `flipH="0"` is the default, and "the author said no
  // flip" has to stay distinguishable from "the author said nothing".
  const flipH = parseOnOffAttribute(xfrm, null, "flipH");
  const flipV = parseOnOffAttribute(xfrm, null, "flipV");

  if (rotation === undefined && flipH === undefined && flipV === undefined) {
    return { size };
  }
  const transform: ImageTransform = {};
  if (rotation !== undefined) {
    transform.rotation = rotation;
  }
  if (flipH !== undefined) {
    transform.flipH = flipH;
  }
  if (flipV !== undefined) {
    transform.flipV = flipV;
  }
  return { size, transform };
}

// ---------------------------------------------------------------------------
// SHAPE TYPE (a:prstGeom)
// ---------------------------------------------------------------------------

/**
 * Read `<a:prstGeom prst="…">` and narrow to the typed `ShapeType` union.
 * Unsupported geometry is not consumed as an editable shape because
 * `ShapeContent` currently serializes fresh preset geometry.
 */
function parseShapeType(spPr: XmlElement | null): Shape["shapeType"] {
  if (!spPr) {
    return "rect";
  }
  const prstGeom = findChildByLocalName(spPr, "prstGeom");
  if (prstGeom) {
    const prst = getAttribute(prstGeom, null, "prst");
    const narrowed = narrowEnum(prst, ShapeTypeSchema);
    if (narrowed) {
      return narrowed;
    }
  }
  return "rect";
}

function hasUnsupportedGeometry(spPr: XmlElement | null): boolean {
  if (!spPr) {
    return false;
  }
  const prstGeom = findChildByLocalName(spPr, "prstGeom");
  if (!prstGeom) {
    return findChildByLocalName(spPr, "custGeom") !== null;
  }
  const avLst = findChildByLocalName(prstGeom, "avLst");
  const prst = getAttribute(prstGeom, null, "prst");
  if (narrowEnum(prst, ShapeTypeSchema) === undefined) {
    return true;
  }
  const adjustmentChildren = avLst ? getChildElements(avLst) : [];
  if (adjustmentChildren.length === 0) {
    return false;
  }
  const adjustments = findChildren(avLst, "a", "gd");
  const adjustmentNames = new Set(
    adjustments.map((adjustment) => getAttribute(adjustment, null, "name")),
  );
  return !(
    prst === "rightBrace" &&
    adjustments.length === 2 &&
    adjustmentChildren.length === adjustments.length &&
    adjustmentNames.size === 2 &&
    adjustmentNames.has("adj1") &&
    adjustmentNames.has("adj2") &&
    adjustments.every((adjustment) =>
      /^val\s+-?\d+$/u.test(getAttribute(adjustment, null, "fmla") ?? ""),
    )
  );
}

function parseGeometryAdjustments(spPr: XmlElement | null): Shape["geometryAdjustments"] {
  const prstGeom = spPr ? findChildByLocalName(spPr, "prstGeom") : null;
  const avLst = prstGeom ? findChildByLocalName(prstGeom, "avLst") : null;
  if (!avLst) {
    return undefined;
  }
  const adjustments: NonNullable<Shape["geometryAdjustments"]> = [];
  for (const adjustment of findChildren(avLst, "a", "gd")) {
    const name = getAttribute(adjustment, null, "name");
    const formula = getAttribute(adjustment, null, "fmla");
    if (name !== null && formula !== null) {
      adjustments.push({ name, formula });
    }
  }
  return adjustments.length === 0 ? undefined : adjustments;
}

function hasUnsupportedRgbColorModifiers(spPr: XmlElement | null): boolean {
  for (const color of findAllDeep(spPr, "a", "srgbClr")) {
    if (color.elements?.some((child) => child.type === "element")) {
      return true;
    }
  }
  return false;
}

function hasUnmodeledFill(spPr: XmlElement | null): boolean {
  return (
    findChildByLocalName(spPr, "pattFill") !== null ||
    findChildByLocalName(spPr, "blipFill") !== null ||
    findChildByLocalName(spPr, "grpFill") !== null
  );
}

/**
 * Shape properties the serializer never emits: `Shape` models geometry, fill,
 * outline and transform only, so a shadow, glow, reflection or 3-D scene would
 * be dropped on save.
 */
const UNMODELED_EFFECT_ELEMENTS = ["effectLst", "effectDag", "scene3d", "sp3d"] as const;

/**
 * An empty `<a:effectLst/>` is Word's explicit "no effects" marker and models
 * fine; carried attributes (`<a:sp3d extrusionH="…"/>`) or children do not.
 */
function hasUnmodeledEffects(spPr: XmlElement | null): boolean {
  return UNMODELED_EFFECT_ELEMENTS.some((localName) => {
    const element = findChildByLocalName(spPr, localName);
    if (!element) {
      return false;
    }
    return getChildElements(element).length > 0 || Object.keys(element.attributes ?? {}).length > 0;
  });
}

function colorNeedsRawPreservation(color: ColorValue | undefined): boolean {
  return color !== undefined && color.rgb === undefined;
}

function fillNeedsRawPreservation(fill: ShapeFill | undefined): boolean {
  if (!fill) {
    return false;
  }
  if (fill.type === "solid") {
    return false;
  }
  if (fill.type === "gradient") {
    return fill.gradient?.stops.some((stop) => colorNeedsRawPreservation(stop.color)) ?? false;
  }
  return false;
}

/**
 * The single answer to "does this `wps:spPr` carry something the editable
 * `Shape` model would lose?". `parseShapeFromDrawing` refuses such a shape and
 * `shouldPreserveRawShapeDrawing` claims it for verbatim preservation, so both
 * must read the same predicate: a reason added to only one of them would
 * either drop the drawing or model it lossily.
 */
const shapeNeedsRawPreservation = (spPr: XmlElement | null): boolean =>
  hasUnsupportedGeometry(spPr) ||
  hasUnsupportedRgbColorModifiers(spPr) ||
  hasUnmodeledFill(spPr) ||
  hasUnmodeledEffects(spPr) ||
  fillNeedsRawPreservation(parseShapeFill(spPr));

// ---------------------------------------------------------------------------
// MAIN ENTRY POINTS
// ---------------------------------------------------------------------------

/**
 * Parse a `<wps:wsp>` element into a Shape model. Does NOT pick up the
 * extent / position / wrap fields — those live on the wrapping
 * `<wp:inline>` / `<wp:anchor>` and are handled by `parseShapeFromDrawing`.
 */
export function parseShape(node: XmlElement): Shape {
  const cNvPr = findChildByLocalName(node, "cNvPr");
  const spPr = findChildByLocalName(node, "spPr");

  const shapeType = parseShapeType(spPr);
  const xfrm = findChildByLocalName(spPr, "xfrm");
  const { size, transform } = parseTransform(xfrm);
  const fill = parseShapeFill(spPr);
  const outline = parseOutline(spPr);

  const id = cNvPr ? (getAttribute(cNvPr, null, "id") ?? undefined) : undefined;
  const names = parseNonVisualDrawingNames(cNvPr);

  const shape: Shape = {
    type: "shape",
    shapeType,
    size,
    ...names,
  };
  const geometryAdjustments = parseGeometryAdjustments(spPr);
  if (geometryAdjustments !== undefined) {
    shape.geometryAdjustments = geometryAdjustments;
  }
  if (id !== undefined) {
    shape.id = id;
  }
  if (fill !== undefined) {
    shape.fill = fill;
  }
  if (outline !== undefined) {
    shape.outline = outline;
  }
  if (transform !== undefined) {
    shape.transform = transform;
  }
  return shape;
}

/**
 * Parse a `<w:drawing>` element as a shape (i.e. the `<a:graphicData>` payload
 * is `<wps:wsp>`, not `<pic:pic>` or a text-box). Returns null when the
 * drawing is not a shape, or when it is a text-box (delegated to
 * `textBoxParser.parseTextBox`).
 */
export function parseShapeFromDrawing(drawingEl: XmlElement): Shape | null {
  const inline = findChildByLocalName(drawingEl, "inline");
  const anchor = findChildByLocalName(drawingEl, "anchor");
  const container = inline ?? anchor;
  if (!container) {
    return null;
  }

  const graphic = findChildByLocalName(container, "graphic");
  if (!graphic) {
    return null;
  }
  const graphicData = findChildByLocalName(graphic, "graphicData");
  if (!graphicData) {
    return null;
  }
  const wsp = findChildByLocalName(graphicData, "wsp");
  if (!wsp) {
    return null;
  }
  // Text boxes go through textBoxParser, not here; they carry their own
  // content tree that the block parser threads through paragraphs.
  if (findChildByLocalName(wsp, "txbx") !== null) {
    return null;
  }

  const spPr = findChildByLocalName(wsp, "spPr");
  if (shapeNeedsRawPreservation(spPr)) {
    return null;
  }

  const shape = parseShape(wsp);

  // The container's wp:extent supersedes spPr's a:ext when both exist.
  // Word can author straight connectors with a zero outer dimension while
  // keeping the actual line length in a:xfrm; retain that non-zero dimension.
  const extent = findChildByLocalName(container, "extent");
  if (extent) {
    const outerWidth = parseNumericAttribute(extent, null, "cx");
    const outerHeight = parseNumericAttribute(extent, null, "cy");
    const cx =
      outerWidth !== undefined && (outerWidth > 0 || shape.size.width === 0)
        ? outerWidth
        : shape.size.width;
    const cy =
      outerHeight !== undefined && (outerHeight > 0 || shape.size.height === 0)
        ? outerHeight
        : shape.size.height;
    shape.size = { width: cx, height: cy };
  }

  const isAnchor = container === anchor;
  if (isAnchor) {
    const position = parseAnchorPosition(container);
    if (position) {
      shape.position = position;
    }
    const wrap = parseAnchorWrap(container);
    if (wrap) {
      shape.wrap = wrap;
    }
    // The same anchor record a picture gets. The shape path used to read none
    // of these and the serializer wrote a constant for each, so a shape came
    // back unlocked, unstacked and free to leave its table cell.
    const drawingAnchor = parseDrawingAnchor(container);
    if (drawingAnchor) {
      shape.anchor = drawingAnchor;
    }
  } else {
    // `wp:inline` carries the same wrap insets `wp:anchor` does, and dropping
    // them here is what made the inline leg lose what the anchored one kept.
    shape.wrap = parseInlineWrap(container);
  }

  const docPr = findChildByLocalName(container, "docPr");
  if (docPr) {
    const id = getAttribute(docPr, null, "id");
    if (id !== null) {
      shape.id = id;
    }
    // `wp:docPr` is the drawing's own non-visual properties; `wps:cNvPr` names
    // the shape inside it. Where both carry a value the outer one wins.
    const names = parseNonVisualDrawingNames(docPr);
    if (names.name !== undefined) {
      shape.name = names.name;
    }
    if (names.alt !== undefined) {
      shape.alt = names.alt;
    }
    if (names.title !== undefined) {
      shape.title = names.title;
    }
  }

  return shape;
}

export function shouldPreserveRawShapeDrawing(drawingEl: XmlElement): boolean {
  const inline = findChildByLocalName(drawingEl, "inline");
  const anchor = findChildByLocalName(drawingEl, "anchor");
  const container = inline ?? anchor;
  if (!container) {
    return false;
  }
  const graphic = findChildByLocalName(container, "graphic");
  const graphicData = graphic ? findChildByLocalName(graphic, "graphicData") : null;
  const wsp = graphicData ? findChildByLocalName(graphicData, "wsp") : null;
  if (!wsp || findChildByLocalName(wsp, "txbx") !== null) {
    return false;
  }
  return shapeNeedsRawPreservation(findChildByLocalName(wsp, "spPr"));
}
