/**
 * The authored DrawingML transform a drawing node carries.
 *
 * A drawing's `a:xfrm` states a rotation and two flips, and each of them is
 * either absent or a value the author wrote. The editor used to carry all
 * three inside one CSS string (`rotate(90deg) scaleX(-1)`), which can express
 * neither an authored `rot="0"` nor an authored `flipH="0"`: both spell the
 * identity, and the identity is what an absent transform already means. Saving
 * the document then wrote the attribute away.
 *
 * So the node carries the three values as attrs, absent as `null`, and the CSS
 * string stays a projection of them for rendering. {@link authoredTransformAttrs}
 * produces both at once so the two cannot drift apart.
 */

import type { ImageTransform } from "../types/document";

/** The transform attrs a drawing node (`image`, `shape`, `textBox`) carries. */
type AuthoredTransformAttrs = {
  docxFlipH: boolean | null;
  docxFlipV: boolean | null;
  docxRotation: number | null;
  /** The rendered CSS, derived from the three above. */
  transform: string | undefined;
};

/** What a drawing node's attrs say, as they arrive from a persisted document. */
type ReadableTransformAttrs = {
  docxFlipH?: boolean | null;
  docxFlipV?: boolean | null;
  docxRotation?: number | null;
  transform?: string;
};

/**
 * The CSS the three authored values render as.
 *
 * Only what the browser has to draw: a rotation of zero and an explicit
 * "not flipped" are the identity, and emitting them would put a no-op
 * transform on every drawing that states one.
 */
const cssTransform = (
  rotation: number | null,
  flipH: boolean | null,
  flipV: boolean | null,
): string | undefined => {
  const parts: string[] = [];
  if (rotation !== null && rotation !== 0) {
    parts.push(`rotate(${rotation}deg)`);
  }
  if (flipH === true) {
    parts.push("scaleX(-1)");
  }
  if (flipV === true) {
    parts.push("scaleY(-1)");
  }
  return parts.length > 0 ? parts.join(" ") : undefined;
};

/** Project a model transform onto a drawing node's attrs. */
export const authoredTransformAttrs = (
  transform: ImageTransform | undefined,
): AuthoredTransformAttrs => {
  const docxRotation = transform?.rotation ?? null;
  const docxFlipH = transform?.flipH ?? null;
  const docxFlipV = transform?.flipV ?? null;
  return {
    docxFlipH,
    docxFlipV,
    docxRotation,
    transform: cssTransform(docxRotation, docxFlipH, docxFlipV),
  };
};

/**
 * The transform a node persisted before these attrs existed carries.
 *
 * The CSS can state neither an authored zero rotation nor an authored "not
 * flipped", which is why the attrs exist; such a node never carried them
 * either, so reading it this way loses nothing.
 */
const transformFromCss = (css: string | undefined): ImageTransform | undefined => {
  if (!css) {
    return undefined;
  }
  const transform: ImageTransform = {};
  const rotateMatch = /rotate\((?<deg>[-\d.]+)deg\)/u.exec(css);
  if (rotateMatch) {
    const rotation = Number.parseFloat(rotateMatch.groups!["deg"]!);
    if (Number.isFinite(rotation)) {
      transform.rotation = rotation;
    }
  }
  if (css.includes("scaleX(-1)")) {
    transform.flipH = true;
  }
  if (css.includes("scaleY(-1)")) {
    transform.flipV = true;
  }
  if (transform.rotation === undefined && !transform.flipH && !transform.flipV) {
    return undefined;
  }
  return transform;
};

/**
 * Read a drawing node's authored transform back.
 *
 * The attrs are the record; the CSS string answers only for a node that states
 * none of them, which is how a node persisted before they existed arrives.
 */
export const readAuthoredTransform = (
  attrs: ReadableTransformAttrs,
): ImageTransform | undefined => {
  const { docxFlipH, docxFlipV, docxRotation } = attrs;
  if (docxRotation == null && docxFlipH == null && docxFlipV == null) {
    return transformFromCss(attrs.transform);
  }
  return {
    ...(docxRotation == null ? {} : { rotation: docxRotation }),
    ...(docxFlipH == null ? {} : { flipH: docxFlipH }),
    ...(docxFlipV == null ? {} : { flipV: docxFlipV }),
  };
};
