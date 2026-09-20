/**
 * Shared DrawingML Parsing Utilities
 *
 * Common functions used by imageParser and textBoxParser
 * for parsing DrawingML elements (positions, wrapping, colors, fills, outlines).
 */

import {
  isPresetLineDashVal,
  isSchemeColorValue,
  PARSE_WARNING_CODES,
  presetLineDashFrom,
  THEME_COLOR_BY_SCHEME_COLOR_VALUE,
  type ThemeColor,
  themeColorSlot,
} from "@stll/docx-core/model";

import type {
  EffectExtentSlots,
  ImagePadding,
  ImagePosition,
  ImageWrap,
  ImageWrapPolygon,
  SchemeColorSlot,
  ShapeFill,
  ShapeOutline,
  ColorValue,
  WrapDistances,
  WrapDistanceSlots,
  WrapPolygonPoint,
} from "../types/document";
import type { ParseContext } from "./parseContext";
import {
  ImageHorizontalAlignmentSchema,
  ImageHorizontalRelativeToSchema,
  ImageVerticalAlignmentSchema,
  ImageVerticalRelativeToSchema,
  ImageWrapTextSchema,
  narrowEnum,
} from "./parserEnums";
import { captureVerbatimXml } from "./verbatimCapture";
import { WORDPROCESSING_DRAWING_NAMESPACE_URIS } from "./drawingAnchor";
import {
  findChildByNamespaceUri,
  getChildElements,
  getAttribute,
  getLocalName,
  getNamespaceUri,
  getTextContent,
  parseNumericAttribute,
  findChildByLocalName,
  findChildrenByLocalName,
  parseOnOffValue,
} from "./xmlParser";
import type { XmlElement } from "./xmlParser";

// ============================================================================
// COLOR PARSING
// ============================================================================

/**
 * `a:schemeClr/@val` as a WordprocessingML theme colour, for the one DrawingML
 * reference a document part carries inline.
 *
 * Both directions live in `@stll/docx-core/model`, total over their schema
 * enumerations, so neither can omit a token.
 */
const schemeColorToThemeColor = (value: string): ThemeColor | undefined =>
  isSchemeColorValue(value) ? (THEME_COLOR_BY_SCHEME_COLOR_VALUE[value] ?? undefined) : undefined;

/**
 * sRGB hex per OOXML (ST_HexColorRGB): exactly six hex digits, case-insensitive.
 * Anything else is rejected so untrusted DOCX input cannot smuggle markup through
 * downstream renderers that interpolate the value into HTML/SVG.
 */
const HEX_COLOR_RE = /^[0-9A-Fa-f]{6}$/u;

function isHexColor(val: string | undefined | null): val is string {
  return typeof val === "string" && HEX_COLOR_RE.test(val);
}

/**
 * Common preset color names to RGB hex values.
 */
const PRESET_COLORS: Record<string, string> = {
  black: "000000",
  white: "FFFFFF",
  red: "FF0000",
  green: "00FF00",
  blue: "0000FF",
  yellow: "FFFF00",
  cyan: "00FFFF",
  magenta: "FF00FF",
};

/**
 * Apply color modifiers (shade, tint) from child elements of a color element.
 * Converts DrawingML 100000ths-scale values to hex (0-FF) for OOXML compatibility.
 */
function applyColorModifiers(color: ColorValue, element: XmlElement): ColorValue {
  const children = getChildElements(element);

  const shade = children.find((el) => el.name === "a:shade");
  if (shade) {
    const val = getAttribute(shade, null, "val");
    if (val) {
      color.themeShade = Math.round((Number.parseInt(val, 10) / 100_000) * 255)
        .toString(16)
        .padStart(2, "0")
        .toUpperCase();
    }
  }

  const tint = children.find((el) => el.name === "a:tint");
  if (tint) {
    const val = getAttribute(tint, null, "val");
    if (val) {
      color.themeTint = Math.round((Number.parseInt(val, 10) / 100_000) * 255)
        .toString(16)
        .padStart(2, "0")
        .toUpperCase();
    }
  }

  return color;
}

/**
 * Parse a color value from a DrawingML element.
 * Handles: a:srgbClr, a:schemeClr, a:sysClr, a:prstClr
 * Applies shade/tint modifiers when present.
 */
export function parseColorElement(element: XmlElement | null): ColorValue | undefined {
  if (!element) {
    return undefined;
  }

  const children = getChildElements(element);

  // sRGB color: a:srgbClr[@val] — must be 6 hex digits per OOXML
  const srgbClr = children.find((el) => el.name === "a:srgbClr");
  if (srgbClr) {
    const val = getAttribute(srgbClr, null, "val");
    if (isHexColor(val)) {
      return applyColorModifiers({ rgb: val.toUpperCase() }, srgbClr);
    }
  }

  // Scheme color (theme): a:schemeClr[@val]
  const schemeClr = children.find((el) => el.name === "a:schemeClr");
  if (schemeClr) {
    // `phClr` and anything outside `ST_SchemeColorVal` name no theme slot; the
    // reference falls through to the next colour kind rather than being read as
    // a colour it does not mean.
    const themeColor = schemeColorToThemeColor(getAttribute(schemeClr, null, "val") ?? "");
    if (themeColor !== undefined) {
      return applyColorModifiers({ themeColor }, schemeClr);
    }
  }

  // System color: a:sysClr[@lastClr] — fall back to black if missing/malformed
  const sysClr = children.find((el) => el.name === "a:sysClr");
  if (sysClr) {
    const lastClr = getAttribute(sysClr, null, "lastClr");
    return { rgb: isHexColor(lastClr) ? lastClr.toUpperCase() : "000000" };
  }

  // Preset color: a:prstClr[@val]
  const prstClr = children.find((el) => el.name === "a:prstClr");
  if (prstClr) {
    const val = getAttribute(prstClr, null, "val");
    if (val && PRESET_COLORS[val]) {
      return { rgb: PRESET_COLORS[val] };
    }
  }

  return undefined;
}

// ============================================================================
// FILL & OUTLINE PARSING
// ============================================================================

/**
 * Parse fill from shape properties (a:solidFill, a:noFill, a:gradFill).
 */
export function parseFill(spPr: XmlElement | null): ShapeFill | undefined {
  if (!spPr) {
    return undefined;
  }

  if (findChildByLocalName(spPr, "noFill")) {
    return { type: "none" };
  }

  const solidFill = findChildByLocalName(spPr, "solidFill");
  if (solidFill) {
    const color = parseColorElement(solidFill);
    return color !== undefined ? { type: "solid", color } : { type: "solid" };
  }

  const gradientFill = findChildByLocalName(spPr, "gradFill");
  if (gradientFill) {
    return parseGradientFill(gradientFill);
  }

  return undefined;
}

function parseGradientFill(gradientFill: XmlElement): ShapeFill {
  let type: "linear" | "radial" | "rectangular" | "path" = "linear";
  let angle: number | undefined;

  const linear = findChildByLocalName(gradientFill, "lin");
  if (linear) {
    const authoredAngle = getAttribute(linear, null, "ang");
    if (authoredAngle) {
      const parsedAngle = Number.parseInt(authoredAngle, 10);
      angle = Number.isNaN(parsedAngle) ? undefined : parsedAngle / 60_000;
    }
  }

  const path = findChildByLocalName(gradientFill, "path");
  if (path) {
    const pathType = getAttribute(path, null, "path");
    if (pathType === "circle") {
      type = "radial";
    } else if (pathType === "rect") {
      type = "rectangular";
    } else {
      type = "path";
    }
  }

  const stops: NonNullable<ShapeFill["gradient"]>["stops"] = [];
  const stopList = findChildByLocalName(gradientFill, "gsLst");
  if (stopList) {
    for (const stop of findChildrenByLocalName(stopList, "gs")) {
      const authoredPosition = getAttribute(stop, null, "pos");
      const parsedPosition = authoredPosition ? Number.parseInt(authoredPosition, 10) : 0;
      const color = parseColorElement(stop);
      if (color) {
        stops.push({
          position: Number.isNaN(parsedPosition) ? 0 : parsedPosition,
          color,
        });
      }
    }
  }

  return {
    type: "gradient",
    rawXml: captureVerbatimXml(gradientFill),
    gradient: {
      type,
      ...(angle !== undefined ? { angle } : {}),
      stops,
    },
  };
}

/**
 * Parse outline from shape properties (a:ln).
 */
export function parseOutline(
  spPr: XmlElement | null,
  context?: ParseContext,
): ShapeOutline | undefined {
  const ln = spPr ? findChildByLocalName(spPr, "ln") : null;
  if (!ln) {
    return undefined;
  }

  if (findChildByLocalName(ln, "noFill")) {
    return undefined;
  }

  const outline: ShapeOutline = {
    rawXml: captureVerbatimXml(ln),
  };

  const w = getAttribute(ln, null, "w");
  if (w) {
    const parsed = Number.parseInt(w, 10);
    if (!Number.isNaN(parsed)) {
      outline.width = parsed;
    }
  }

  const cap = getAttribute(ln, null, "cap");
  if (cap === "flat") {
    outline.cap = "flat";
  } else if (cap === "rnd") {
    outline.cap = "round";
  } else if (cap === "sq") {
    outline.cap = "square";
  }

  if (findChildByLocalName(ln, "bevel")) {
    outline.join = "bevel";
  } else if (findChildByLocalName(ln, "round")) {
    outline.join = "round";
  } else if (findChildByLocalName(ln, "miter")) {
    outline.join = "miter";
  }

  const solidFill = findChildByLocalName(ln, "solidFill");
  if (solidFill) {
    const color = parseColorElement(solidFill);
    if (color !== undefined) {
      outline.color = color;
    }
  }

  // `a:custDash` is a sibling element, not a `@val` of this one: an outline
  // that carries one has no preset dash to read, and its stop list replays
  // through `rawXml`.
  const prstDash = findChildByLocalName(ln, "prstDash");
  const rawDash = prstDash ? getAttribute(prstDash, null, "val") : null;
  if (rawDash) {
    if (!isPresetLineDashVal(rawDash)) {
      context?.warn({
        code: PARSE_WARNING_CODES.outlineDashOutsideEnum,
        element: prstDash?.name ?? "a:prstDash",
        value: rawDash,
      });
    }
    outline.dash = presetLineDashFrom(rawDash);
  }

  const headEnd = findChildByLocalName(ln, "headEnd");
  if (headEnd) {
    outline.headEnd = parseLineEnd(headEnd);
  }

  const tailEnd = findChildByLocalName(ln, "tailEnd");
  if (tailEnd) {
    outline.tailEnd = parseLineEnd(tailEnd);
  }

  return outline;
}

type LineEndType = NonNullable<ShapeOutline["headEnd"]>["type"];
type LineEndSize = NonNullable<ShapeOutline["headEnd"]>["width"];

const LINE_END_TYPES = [
  "none",
  "triangle",
  "stealth",
  "diamond",
  "oval",
  "arrow",
] as const satisfies readonly LineEndType[];

function parseLineEnd(element: XmlElement): NonNullable<ShapeOutline["headEnd"]> {
  const authoredType = getAttribute(element, null, "type");
  const type = LINE_END_TYPES.find((allowed) => allowed === authoredType) ?? "none";

  const authoredWidth = getAttribute(element, null, "w");
  const authoredLength = getAttribute(element, null, "len");
  const width: LineEndSize =
    authoredWidth === "sm" || authoredWidth === "med" || authoredWidth === "lg"
      ? authoredWidth
      : undefined;
  const length: LineEndSize =
    authoredLength === "sm" || authoredLength === "med" || authoredLength === "lg"
      ? authoredLength
      : undefined;

  return {
    type,
    ...(width !== undefined ? { width } : {}),
    ...(length !== undefined ? { length } : {}),
  };
}

// ============================================================================
// POSITION PARSING
// ============================================================================

/**
 * Parse horizontal position from wp:positionH element.
 */
/** A child of a WordprocessingDrawing element, resolved by namespace. */
const findDrawingChild = (parent: XmlElement | null, localName: string): XmlElement | null =>
  findChildByNamespaceUri(parent, WORDPROCESSING_DRAWING_NAMESPACE_URIS, localName);

export function parsePositionH(posH: XmlElement | null): ImagePosition["horizontal"] | undefined {
  if (!posH) {
    return undefined;
  }

  const relativeTo =
    narrowEnum(getAttribute(posH, null, "relativeFrom"), ImageHorizontalRelativeToSchema) ??
    "column";

  const alignEl = findDrawingChild(posH, "align");
  if (alignEl) {
    const text = getTextContent(alignEl);
    const alignment = narrowEnum(text, ImageHorizontalAlignmentSchema);
    if (alignment) {
      return { relativeTo, alignment };
    }
  }

  const posOffsetEl = findDrawingChild(posH, "posOffset");
  if (posOffsetEl) {
    const text = getTextContent(posOffsetEl);
    const posOffset = Number.parseInt(text, 10);
    return {
      relativeTo,
      posOffset: Number.isNaN(posOffset) ? 0 : posOffset,
    };
  }

  return { relativeTo, posOffset: 0 };
}

/**
 * Parse vertical position from wp:positionV element.
 */
export function parsePositionV(posV: XmlElement | null): ImagePosition["vertical"] | undefined {
  if (!posV) {
    return undefined;
  }

  const relativeTo =
    narrowEnum(getAttribute(posV, null, "relativeFrom"), ImageVerticalRelativeToSchema) ??
    "paragraph";

  const alignEl = findDrawingChild(posV, "align");
  if (alignEl) {
    const text = getTextContent(alignEl);
    const alignment = narrowEnum(text, ImageVerticalAlignmentSchema);
    if (alignment) {
      return { relativeTo, alignment };
    }
  }

  const posOffsetEl = findDrawingChild(posV, "posOffset");
  if (posOffsetEl) {
    const text = getTextContent(posOffsetEl);
    const posOffset = Number.parseInt(text, 10);
    return {
      relativeTo,
      posOffset: Number.isNaN(posOffset) ? 0 : posOffset,
    };
  }

  return { relativeTo, posOffset: 0 };
}

/**
 * Parse position for anchored drawings (combines positionH + positionV).
 */
export function parseAnchorPosition(anchor: XmlElement): ImagePosition {
  const positionH = findDrawingChild(anchor, "positionH");
  const positionV = findDrawingChild(anchor, "positionV");

  return {
    horizontal: parsePositionH(positionH) ?? { relativeTo: "column", posOffset: 0 },
    vertical: parsePositionV(positionV) ?? { relativeTo: "paragraph", posOffset: 0 },
  };
}

// ============================================================================
// WRAP PARSING
// ============================================================================

/**
 * `EG_WrapType`'s members, by local name: the `wp` prefix is the producer's
 * choice and a package free to bind the namespace elsewhere writes the same
 * wrap.
 */
const WRAP_ELEMENT_LOCAL_NAMES: Readonly<Record<string, ImageWrap["type"]>> = {
  wrapNone: "inFront",
  wrapSquare: "square",
  wrapTight: "tight",
  wrapThrough: "through",
  wrapTopAndBottom: "topAndBottom",
};

/** The `EG_WrapType` child of a `wp:anchor`, resolved by namespace. */
export function findWrapElement(anchor: XmlElement): XmlElement | null {
  for (const child of getChildElements(anchor)) {
    if (
      getLocalName(child.name) in WRAP_ELEMENT_LOCAL_NAMES &&
      WORDPROCESSING_DRAWING_NAMESPACE_URIS.has(getNamespaceUri(child) ?? "")
    ) {
      return child;
    }
  }
  return null;
}

/** The inset attributes `CT_Inline`, `CT_Anchor` and the wrap children share. */
const WRAP_DISTANCE_KEYS = ["distT", "distB", "distL", "distR"] as const;

/**
 * The insets one element states, or undefined when it states none.
 *
 * A wrap child's type declares a subset — `CT_WrapTight` and `CT_WrapThrough`
 * only `distL`/`distR`, `CT_WrapTopBottom` only `distT`/`distB` — and an
 * attribute the type does not declare cannot be present to read.
 */
const parseWrapDistances = (el: XmlElement | null): WrapDistances | undefined => {
  if (!el) {
    return undefined;
  }
  const distances: WrapDistances = {};
  let stated = false;
  for (const key of WRAP_DISTANCE_KEYS) {
    const value = parseNumericAttribute(el, null, key);
    if (value !== undefined) {
      distances[key] = value;
      stated = true;
    }
  }
  return stated ? distances : undefined;
};

/** The slot record for a drawing, or undefined when neither element stated an inset. */
const wrapDistanceSlots = (
  drawing: WrapDistances | undefined,
  wrapChild: WrapDistances | undefined,
): WrapDistanceSlots | undefined =>
  drawing === undefined && wrapChild === undefined
    ? undefined
    : {
        ...(drawing === undefined ? {} : { drawing }),
        ...(wrapChild === undefined ? {} : { wrapChild }),
      };

/**
 * `wp:effectExtent`'s four sides in EMU, exactly as the element states them.
 *
 * `CT_EffectExtent` requires all four, so a missing one is a malformed element
 * rather than an unstated side, and zero is its own value: an explicit all-zero
 * reservation on a wrap child says "this wrap reserves nothing", which is not
 * what an absent element says.
 */
export const parseEffectExtent = (el: XmlElement | null): ImagePadding | undefined =>
  el === null
    ? undefined
    : {
        left: parseNumericAttribute(el, null, "l") ?? 0,
        top: parseNumericAttribute(el, null, "t") ?? 0,
        right: parseNumericAttribute(el, null, "r") ?? 0,
        bottom: parseNumericAttribute(el, null, "b") ?? 0,
      };

/**
 * The drawing's own `wp:effectExtent`, which is also what `Image.padding` holds.
 *
 * An all-zero reservation is what Word writes on a drawing with no effect, and
 * a rebuild writes zeros for a drawing the record holds none for, so the two
 * are the same document and neither the padding nor the slot keeps it. The wrap
 * child's is read as authored instead: an absent one there inherits the
 * drawing's, so an explicit all-zero says something an absent one does not.
 */
export const parseDrawingEffectExtent = (el: XmlElement | null): ImagePadding | undefined => {
  const extent = parseEffectExtent(el);
  if (extent === undefined) {
    return undefined;
  }
  return extent.left === 0 && extent.top === 0 && extent.right === 0 && extent.bottom === 0
    ? undefined
    : extent;
};

/** The slot record, or undefined when neither element stated an effect extent. */
const effectExtentSlots = (
  drawing: ImagePadding | undefined,
  wrapChild: ImagePadding | undefined,
): EffectExtentSlots | undefined =>
  drawing === undefined && wrapChild === undefined
    ? undefined
    : {
        ...(drawing === undefined ? {} : { drawing }),
        ...(wrapChild === undefined ? {} : { wrapChild }),
      };

/** The wrap insets `CT_Inline` carries, read the same way for every graphic. */
export function parseInlineWrap(inlineEl: XmlElement): ImageWrap {
  const wrap: ImageWrap = { type: "inline" };
  const drawing = parseWrapDistances(inlineEl);
  for (const key of WRAP_DISTANCE_KEYS) {
    const value = drawing?.[key];
    if (value !== undefined) {
      wrap[key] = value;
    }
  }
  const slots = wrapDistanceSlots(drawing, undefined);
  if (slots !== undefined) {
    wrap.distanceSlots = slots;
  }
  // `CT_Inline` has no wrap child, so the drawing is the only carrier there is.
  const extents = effectExtentSlots(
    parseDrawingEffectExtent(findDrawingChild(inlineEl, "effectExtent")),
    undefined,
  );
  if (extents !== undefined) {
    wrap.effectExtentSlots = extents;
  }
  return wrap;
}

/** A `CT_Point2D`, which requires both coordinates. */
const parseWrapPolygonPoint = (el: XmlElement | null): WrapPolygonPoint | undefined => {
  const x = parseNumericAttribute(el, null, "x");
  const y = parseNumericAttribute(el, null, "y");
  return x === undefined || y === undefined ? undefined : { x, y };
};

/**
 * `wp:wrapPolygon`, the outline `wp:wrapTight` and `wp:wrapThrough` require.
 *
 * Undefined when the element is absent or states no `wp:start`, which
 * `CT_WrapPath` requires and a rebuild has nothing to write without. A path
 * with fewer `wp:lineTo` than the type admits is kept as it stands: reading it
 * as no polygon at all would hand the drawing the minted rectangle and move
 * text the source flows through the object.
 */
const parseWrapPolygon = (wrapEl: XmlElement): ImageWrapPolygon | undefined => {
  const polygonEl = findDrawingChild(wrapEl, "wrapPolygon");
  if (!polygonEl) {
    return undefined;
  }
  const startEl = findDrawingChild(polygonEl, "start");
  const start = parseWrapPolygonPoint(startEl);
  if (start === undefined) {
    return undefined;
  }
  const lineTo = getChildElements(polygonEl)
    .filter(
      (child) =>
        getLocalName(child.name) === "lineTo" &&
        WORDPROCESSING_DRAWING_NAMESPACE_URIS.has(getNamespaceUri(child) ?? ""),
    )
    .flatMap((child) => {
      const point = parseWrapPolygonPoint(child);
      return point === undefined ? [] : [point];
    });

  const edited = parseOnOffValue(getAttribute(polygonEl, null, "edited"));
  return {
    ...(edited === undefined ? {} : { edited }),
    start,
    lineTo,
  };
};

export type WrapElementOptions = {
  /** The `EG_WrapType` child of the anchor, or null when it declared none. */
  wrapEl: XmlElement | null;
  /** `wp:anchor/@behindDoc`, which decides which side of the text a `wrapNone` sits on. */
  behindDoc: boolean;
  /** The insets `wp:anchor` itself states. */
  anchorDistances?: WrapDistances;
  /** The `wp:effectExtent` `wp:anchor` itself states. */
  anchorEffectExtent?: ImagePadding;
};

/**
 * Parse wrap settings from a wrap element.
 *
 * Both the insets and the effect extent are declared twice over — on the
 * drawing and on the wrap child — and each is kept in the slot that stated it.
 * The insets differ in that the wrap child's take priority as the value in
 * force; the two effect extents are two values (the object's own reservation
 * and the one the text flow is computed against), so neither overrides the
 * other.
 */
export function parseWrapElement({
  wrapEl,
  behindDoc,
  anchorDistances,
  anchorEffectExtent,
}: WrapElementOptions): ImageWrap {
  const drawingDistances =
    anchorDistances !== undefined && WRAP_DISTANCE_KEYS.some((key) => anchorDistances[key] != null)
      ? anchorDistances
      : undefined;

  if (!wrapEl) {
    const wrap: ImageWrap = { type: behindDoc ? "behind" : "inFront" };
    for (const key of WRAP_DISTANCE_KEYS) {
      const value = drawingDistances?.[key];
      if (value !== undefined) {
        wrap[key] = value;
      }
    }
    const slots = wrapDistanceSlots(drawingDistances, undefined);
    if (slots !== undefined) {
      wrap.distanceSlots = slots;
    }
    // `wp:wrapNone` declares no effect extent, so the drawing is the only carrier.
    const extents = effectExtentSlots(anchorEffectExtent, undefined);
    if (extents !== undefined) {
      wrap.effectExtentSlots = extents;
    }
    return wrap;
  }

  const named = WRAP_ELEMENT_LOCAL_NAMES[getLocalName(wrapEl.name)] ?? "square";
  // `wp:wrapNone` says only that text does not flow around the object; which
  // side of the text it sits on is `@behindDoc`, on the anchor.
  const type = named === "inFront" && behindDoc ? "behind" : named;

  const wrap: ImageWrap = { type };

  const wrapText = narrowEnum(getAttribute(wrapEl, null, "wrapText"), ImageWrapTextSchema);
  if (wrapText) {
    wrap.wrapText = wrapText;
  }

  // Wrap child distances take priority, then anchor-level. Both are kept
  // beside the value in force, so a rebuild writes each back on the element
  // that stated it.
  const wrapChildDistances = parseWrapDistances(wrapEl);
  for (const key of WRAP_DISTANCE_KEYS) {
    const value = wrapChildDistances?.[key] ?? drawingDistances?.[key];
    if (value !== undefined) {
      wrap[key] = value;
    }
  }
  const slots = wrapDistanceSlots(drawingDistances, wrapChildDistances);
  if (slots !== undefined) {
    wrap.distanceSlots = slots;
  }

  // Only `CT_WrapSquare` and `CT_WrapTopBottom` declare one; the lookup on the
  // others finds nothing, so no kind test is needed to stay schema-honest.
  const extents = effectExtentSlots(
    anchorEffectExtent,
    parseEffectExtent(findDrawingChild(wrapEl, "effectExtent")),
  );
  if (extents !== undefined) {
    wrap.effectExtentSlots = extents;
  }

  const polygon = parseWrapPolygon(wrapEl);
  if (polygon !== undefined) {
    wrap.polygon = polygon;
  }

  return wrap;
}

/**
 * Parse wrap from an anchor element (finds wrap child internally).
 */
/**
 * Read `wp:anchor/@behindDoc`, the flag that puts an anchored object behind the
 * body text. The attribute is xsd:boolean, so `1`, `0`, `true` and `false` are
 * all legal spellings and producers differ: Word writes `1`, others `true`.
 * Absent means in front of the text.
 */
export function parseAnchorBehindDoc(anchor: XmlElement): boolean {
  return parseOnOffValue(getAttribute(anchor, null, "behindDoc")) ?? false;
}

export function parseAnchorWrap(anchor: XmlElement): ImageWrap | undefined {
  const anchorDistances = parseWrapDistances(anchor);
  const anchorEffectExtent = parseDrawingEffectExtent(findDrawingChild(anchor, "effectExtent"));
  return parseWrapElement({
    wrapEl: findWrapElement(anchor),
    behindDoc: parseAnchorBehindDoc(anchor),
    ...(anchorDistances === undefined ? {} : { anchorDistances }),
    ...(anchorEffectExtent === undefined ? {} : { anchorEffectExtent }),
  });
}

// ============================================================================
// COLOR RESOLUTION (for shapes/text boxes without theme context)
// ============================================================================

/**
 * Default theme color fallbacks (Office 2016 defaults).
 * Used when resolving theme colors without a Theme object.
 */
const DEFAULT_THEME_COLOR_HEX = {
  accent1: "5B9BD5",
  accent2: "ED7D31",
  accent3: "A5A5A5",
  accent4: "FFC000",
  accent5: "4472C4",
  accent6: "70AD47",
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "1F497D",
  lt2: "EEECE1",
  hlink: "0563C1",
  folHlink: "954F72",
} as const satisfies Record<SchemeColorSlot, string>;

/**
 * Resolve a ColorValue to a CSS hex string using default theme colors.
 * For use when no Theme object is available (e.g., shape/text box parsing).
 */
export function resolveColorValueToHex(color: ColorValue | undefined): string | undefined {
  if (!color) {
    return undefined;
  }

  if (color.rgb) {
    return `#${color.rgb}`;
  }

  if (color.themeColor) {
    const slot = themeColorSlot(color.themeColor);
    return `#${slot === undefined ? "000000" : DEFAULT_THEME_COLOR_HEX[slot]}`;
  }

  return undefined;
}
