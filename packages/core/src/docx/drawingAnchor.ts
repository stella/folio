/**
 * `wp:inline` and `wp:anchor` — the element a drawing hangs off, and the one
 * place that reads and writes its own attributes.
 *
 * A picture, a shape and a text box are the same `CT_Inline` or `CT_Anchor`
 * with a different graphic inside, and they had two writers. The picture path
 * wrote what the author stated; the shape path wrote `simplePos="0"
 * relativeHeight="251658240" locked="0" layoutInCell="1" allowOverlap="1"` and
 * `<wp:simplePos x="0" y="0"/>` on every shape it rebuilt, so a shape a
 * document stacked above a picture came back under it. One owner is what keeps
 * the two from drifting apart again.
 *
 * `CT_Anchor` requires five of these attributes plus `@behindDoc`, so a rebuilt
 * anchor states a value for each. The model holding nothing for one is the
 * author having stated nothing, and what is written then is OOXML's own default
 * rather than a value folio chose.
 */

import type { DrawingAnchor, ImagePadding, ImageWrap, WrapDistances } from "../types/document";
import { intAttr } from "./serializer/xmlUtils";
import {
  findChildByNamespaceUri,
  getAttribute,
  parseNumericAttribute,
  parseOnOffValue,
} from "./xmlParser";
import type { XmlElement } from "./xmlParser";

/** WordprocessingDrawing, Transitional and Strict (ECMA-376 Parts 1 and 4). */
export const WORDPROCESSING_DRAWING_NAMESPACE_URIS: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
]);

/** The anchor fields that are a single `ST_OnOff`-style attribute. */
type DrawingAnchorFlag = {
  [Key in keyof DrawingAnchor]-?: NonNullable<DrawingAnchor[Key]> extends boolean ? Key : never;
}[keyof DrawingAnchor];

type FlagAttribute = {
  /** The attribute's spelling on `wp:anchor`. */
  spelling: string;
  /**
   * What `CT_Anchor` says when the model holds nothing: a required attribute
   * still has to be written, an optional one is left off.
   */
  whenUnstated: boolean | "omit";
};

/**
 * Model key → attribute, total over the boolean half of {@link DrawingAnchor}.
 *
 * A flag added to the model fails this map until the rebuild path decides what
 * an anchor that does not state it is.
 */
const ANCHOR_FLAGS = {
  useSimplePosition: { spelling: "simplePos", whenUnstated: false },
  locked: { spelling: "locked", whenUnstated: false },
  layoutInCell: { spelling: "layoutInCell", whenUnstated: true },
  allowOverlap: { spelling: "allowOverlap", whenUnstated: true },
  hidden: { spelling: "hidden", whenUnstated: "omit" },
} as const satisfies Record<DrawingAnchorFlag, FlagAttribute>;

/** Every boolean the anchor models, from the total map above. */
export const DRAWING_ANCHOR_FLAG_KEYS = Object.keys(ANCHOR_FLAGS) as DrawingAnchorFlag[];

/**
 * `@relativeHeight` is required and unsigned; this is the value Word writes for
 * the first floating object it anchors, and the only one folio may invent.
 */
const DEFAULT_RELATIVE_HEIGHT = 251_658_240;

/** The wrap insets, which `CT_Inline` and `CT_Anchor` both carry. */
const WRAP_DISTANCE_ATTRIBUTES = {
  distT: "distT",
  distB: "distB",
  distL: "distL",
  distR: "distR",
} as const;

const WRAP_DISTANCE_KEYS = Object.keys(
  WRAP_DISTANCE_ATTRIBUTES,
) as (keyof typeof WRAP_DISTANCE_ATTRIBUTES)[];

/**
 * Read `CT_Anchor`'s own attributes and its `wp:simplePos` offsets.
 *
 * Undefined when the element states none of them: an anchor whose every
 * attribute is the spec default holds nothing the model needs to remember, and
 * recording an empty record would make a rebuild write values the source did
 * not.
 */
export const parseDrawingAnchor = (anchorEl: XmlElement): DrawingAnchor | undefined => {
  const anchor: DrawingAnchor = {};
  let stated = false;

  for (const key of DRAWING_ANCHOR_FLAG_KEYS) {
    const value = parseOnOffValue(getAttribute(anchorEl, null, ANCHOR_FLAGS[key].spelling));
    if (value !== undefined) {
      anchor[key] = value;
      stated = true;
    }
  }

  const relativeHeight = parseNumericAttribute(anchorEl, null, "relativeHeight");
  if (relativeHeight !== undefined) {
    anchor.relativeHeight = relativeHeight;
    stated = true;
  }

  const simplePosEl = findChildByNamespaceUri(
    anchorEl,
    WORDPROCESSING_DRAWING_NAMESPACE_URIS,
    "simplePos",
  );
  const x = parseNumericAttribute(simplePosEl, null, "x");
  const y = parseNumericAttribute(simplePosEl, null, "y");
  if (x !== undefined && y !== undefined) {
    anchor.simplePosition = { x, y };
    stated = true;
  }

  return stated ? anchor : undefined;
};

/**
 * The insets to write on each of the two elements that can carry them.
 *
 * A drawing whose slots say where its insets were authored gets each one back
 * on that element, which is the only way the wrap child's own set survives a
 * rebuild. A drawing with no slots, or one an editor command has moved an inset
 * on, states the value in force on the drawing: that is where OOXML reads an
 * inset from when the wrap child states none, so the effective value is right
 * either way, and inventing a wrap-child inset from a moved value would not be.
 */
export const resolveWrapDistances = (
  wrap: ImageWrap | undefined,
): { drawing: WrapDistances; wrapChild: WrapDistances } => {
  const slots = wrap?.distanceSlots;
  const unmoved =
    slots !== undefined &&
    WRAP_DISTANCE_KEYS.every(
      (key) => wrap?.[key] === (slots.wrapChild?.[key] ?? slots.drawing?.[key]),
    );
  if (unmoved) {
    return { drawing: slots.drawing ?? {}, wrapChild: slots.wrapChild ?? {} };
  }
  const drawing: WrapDistances = {};
  for (const key of WRAP_DISTANCE_KEYS) {
    const value = wrap?.[key];
    if (value !== undefined) {
      drawing[key] = value;
    }
  }
  return { drawing, wrapChild: {} };
};

/** `CT_EffectExtent`'s four sides, in the order the type declares them. */
const EFFECT_EXTENT_SIDES = [
  "left",
  "top",
  "right",
  "bottom",
] as const satisfies readonly (keyof ImagePadding)[];

/**
 * Two reservations are the same document when every side matches.
 *
 * An absent side and a zero one are one value: `CT_EffectExtent` requires all
 * four and the rebuild writes zero for a side the record holds nothing for, so
 * an all-zero record and no record at all say the same thing.
 */
const sameEffectExtent = (
  left: ImagePadding | undefined,
  right: ImagePadding | undefined,
): boolean => EFFECT_EXTENT_SIDES.every((side) => (left?.[side] ?? 0) === (right?.[side] ?? 0));

export type ResolvedEffectExtents = {
  /** `wp:inline` or `wp:anchor`; undefined writes the all-zero reservation. */
  drawing: ImagePadding | undefined;
  /** `wp:wrapSquare` or `wp:wrapTopAndBottom`, which write nothing when undefined. */
  wrapChild: ImagePadding | undefined;
};

/**
 * The `wp:effectExtent` to write on each of the two elements that carry one.
 *
 * {@link resolveWrapDistances}'s rule, one element over: a drawing whose slots
 * say where the reservations were authored gets each one back on that element,
 * which is the only way the wrap child's own survives a rebuild. Once the
 * drawing's reservation has been resized the two no longer describe the same
 * object, so the rebuild states the value in force on the drawing and lets
 * OOXML's own default apply to the wrap — keeping a wrap reservation computed
 * against the old size would flow text around a shape that is no longer there.
 *
 * `effective` is the reservation the model holds for the drawing itself:
 * `Image.padding`. A shape and a text box have no such field — the rebuild has
 * always written zeros — so they pass the drawing slot back and nothing can
 * have moved a value the model cannot hold.
 */
export const resolveEffectExtents = (
  wrap: ImageWrap | undefined,
  effective: ImagePadding | undefined,
): ResolvedEffectExtents => {
  const slots = wrap?.effectExtentSlots;
  if (slots === undefined || !sameEffectExtent(effective, slots.drawing)) {
    return { drawing: effective, wrapChild: undefined };
  }
  return { drawing: slots.drawing, wrapChild: slots.wrapChild };
};

/** `wp:effectExtent`, whose four sides `CT_EffectExtent` all requires. */
export const serializeEffectExtent = (extent: ImagePadding | undefined): string =>
  `<wp:effectExtent l="${intAttr(extent?.left ?? 0)}" t="${intAttr(extent?.top ?? 0)}" r="${intAttr(extent?.right ?? 0)}" b="${intAttr(extent?.bottom ?? 0)}"/>`;

/** The inset attributes an element states, in schema order, or nothing. */
export const serializeWrapDistances = (distances: WrapDistances): string => {
  const attrs = WRAP_DISTANCE_KEYS.flatMap((key) => {
    const value = distances[key];
    return value === undefined ? [] : [`${WRAP_DISTANCE_ATTRIBUTES[key]}="${intAttr(value)}"`];
  });
  return attrs.length > 0 ? ` ${attrs.join(" ")}` : "";
};

/**
 * Only the wrap insets `wp:inline` or `wp:anchor` itself carries. OOXML
 * defaults an omitted one to zero, so writing all four would put a value on
 * every drawing that never had one.
 */
export const serializeWrapDistanceAttributes = (wrap: ImageWrap | undefined): string =>
  serializeWrapDistances(resolveWrapDistances(wrap).drawing);

/** `wp:inline` carries the wrap insets and nothing else of `CT_Anchor`'s set. */
export const serializeInlineAttributes = (wrap: ImageWrap | undefined): string =>
  serializeWrapDistanceAttributes(wrap);

export type AnchorAttributesOptions = {
  anchor: DrawingAnchor | undefined;
  wrap: ImageWrap | undefined;
  /** `@behindDoc`, which the wrap type decides rather than the anchor record. */
  behindDoc: boolean;
};

/**
 * Every attribute `CT_Anchor` declares of its own, in schema order.
 *
 * The walk is over {@link DRAWING_ANCHOR_FLAG_KEYS} rather than a hand-written list, so
 * a flag added to the model cannot be written by the parser and forgotten here.
 */
export const serializeAnchorAttributes = ({
  anchor,
  wrap,
  behindDoc,
}: AnchorAttributesOptions): string => {
  const attrs: string[] = [];
  for (const key of DRAWING_ANCHOR_FLAG_KEYS) {
    const { spelling, whenUnstated } = ANCHOR_FLAGS[key];
    const stated = anchor?.[key] ?? whenUnstated;
    if (stated !== "omit") {
      attrs.push(`${spelling}="${stated ? "1" : "0"}"`);
    }
    // `@simplePos` opens the set, `@relativeHeight` and `@behindDoc` follow it,
    // and the remaining flags come after those.
    if (key === "useSimplePosition") {
      attrs.push(`relativeHeight="${intAttr(anchor?.relativeHeight ?? DEFAULT_RELATIVE_HEIGHT)}"`);
      attrs.push(`behindDoc="${behindDoc ? "1" : "0"}"`);
    }
  }
  return `${serializeWrapDistanceAttributes(wrap)} ${attrs.join(" ")}`;
};

/**
 * `wp:simplePos`, which `CT_Anchor` requires whether or not `@simplePos` says to
 * use it, so an anchor that authored no offsets still writes the origin.
 */
export const serializeSimplePos = (anchor: DrawingAnchor | undefined): string => {
  const { x, y } = anchor?.simplePosition ?? { x: 0, y: 0 };
  return `<wp:simplePos x="${intAttr(x)}" y="${intAttr(y)}"/>`;
};
