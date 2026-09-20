/**
 * `wp:wrapPolygon` — the outline a tight or through wrap flows text around, and
 * the one place that decides when folio may invent one.
 *
 * `CT_WrapTight` and `CT_WrapThrough` require the element and `CT_WrapPath`
 * requires a `wp:start` and at least two `wp:lineTo`, so a wrap of either kind
 * has to state a polygon whether or not the document authored one. The
 * serializer wrote a constant rectangle for every one of them and read nothing,
 * so an authored outline came back as a box. Minting and writing share this
 * module so the value folio invents cannot drift from the value it accepts.
 */

import type { ImageWrap, ImageWrapPolygon } from "../types/document";
import { intAttr } from "./serializer/xmlUtils";

/** `CT_WrapPath` coordinates are relative to a box of this many units. */
const WRAP_PATH_EXTENT = 21_600;

/**
 * The whole drawing, as a closed rectangle in the path's relative space: the
 * polygon that leaves text flowing exactly as the wrap's bounding box would.
 * The only wrap outline folio invents.
 */
const DEFAULT_WRAP_POLYGON: ImageWrapPolygon = {
  edited: false,
  start: { x: 0, y: 0 },
  lineTo: [
    { x: 0, y: WRAP_PATH_EXTENT },
    { x: WRAP_PATH_EXTENT, y: WRAP_PATH_EXTENT },
    { x: WRAP_PATH_EXTENT, y: 0 },
    { x: 0, y: 0 },
  ],
};

/**
 * The polygon a `wp:wrapTight` or `wp:wrapThrough` states: the authored one, or
 * the rectangle when the model holds none.
 *
 * An authored path is written back as it stands, including one shorter than
 * `CT_WrapPath`'s two `wp:lineTo`. folio is not the validator of its input, and
 * replacing a short outline with the full extent of the drawing would move text
 * that the source flows through the object — the same silent substitution this
 * module exists to end. The rectangle is minted only where the alternative is
 * omitting a child the type requires.
 */
export const requiredWrapPolygon = (authored: ImageWrapPolygon | undefined): ImageWrapPolygon =>
  authored ?? DEFAULT_WRAP_POLYGON;

/**
 * The polygon a wrap of `type` has to state, or undefined when its type
 * declares none.
 *
 * The editor command that changes a wrap type calls this, so a drawing that
 * becomes tight or through is minted its rectangle once — at the edit, as a
 * fact of the document the author can see and undo — and every save after that
 * writes back what the model then holds.
 */
export const wrapPolygonFor = (
  type: ImageWrap["type"],
  authored: ImageWrapPolygon | undefined,
): ImageWrapPolygon | undefined =>
  type === "tight" || type === "through" ? requiredWrapPolygon(authored) : undefined;

/** A copy, deep enough that a ProseMirror attr cannot share a point with the model. */
export const copiedWrapPolygon = (
  polygon: ImageWrapPolygon | undefined,
): ImageWrapPolygon | undefined =>
  polygon === undefined
    ? undefined
    : {
        ...polygon,
        start: { ...polygon.start },
        lineTo: polygon.lineTo.map((point) => ({ ...point })),
      };

/** `wp:wrapPolygon`, written from a polygon {@link wrapPolygonFor} has approved. */
export const serializeWrapPolygon = (polygon: ImageWrapPolygon): string => {
  const edited = polygon.edited === undefined ? "" : ` edited="${polygon.edited ? "1" : "0"}"`;
  const lineTo = polygon.lineTo
    .map((point) => `<wp:lineTo x="${intAttr(point.x)}" y="${intAttr(point.y)}"/>`)
    .join("");
  return `<wp:wrapPolygon${edited}><wp:start x="${intAttr(polygon.start.x)}" y="${intAttr(polygon.start.y)}"/>${lineTo}</wp:wrapPolygon>`;
};
