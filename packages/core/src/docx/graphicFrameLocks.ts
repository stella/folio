/**
 * `wp:cNvGraphicFramePr > a:graphicFrameLocks` — the manipulation locks an
 * author set on a drawing's graphic frame (ECMA-376 §20.1.2.2.19).
 *
 * Parsing and serialization live together so the attribute set cannot drift:
 * a lock added to {@link ImageFrameLocks} fails the total map below until both
 * directions carry it.
 */

import type { ImageFrameLocks } from "../types/document";
import { findChild, getAttribute, parseOnOffValue } from "./xmlParser";
import type { XmlElement } from "./xmlParser";

/**
 * Model key → OOXML attribute. Insertion order is the schema's attribute
 * order, which is also the order we emit.
 */
const GRAPHIC_FRAME_LOCK_ATTRIBUTES = {
  noGrp: "noGrp",
  noDrilldown: "noDrilldown",
  noSelect: "noSelect",
  noChangeAspect: "noChangeAspect",
  noMove: "noMove",
  noResize: "noResize",
} as const satisfies Record<keyof ImageFrameLocks, string>;

const isGraphicFrameLockKey = (key: string): key is keyof typeof GRAPHIC_FRAME_LOCK_ATTRIBUTES =>
  key in GRAPHIC_FRAME_LOCK_ATTRIBUTES;

/** Every modeled lock, in schema order: the total map above keeps it exhaustive. */
export const GRAPHIC_FRAME_LOCK_KEYS = Object.keys(GRAPHIC_FRAME_LOCK_ATTRIBUTES).filter(
  isGraphicFrameLockKey,
);

const DRAWINGML_NAMESPACE = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

/**
 * Read the locks off a `wp:inline` or `wp:anchor` element.
 *
 * Returns undefined when the element is absent or carries no recognized
 * attribute: an authored `<a:graphicFrameLocks/>` with nothing on it means the
 * same as no element at all, so both stay "spec defaults" on the model.
 */
export const parseGraphicFrameLocks = (parent: XmlElement): ImageFrameLocks | undefined => {
  const locksEl = findChild(findChild(parent, "wp", "cNvGraphicFramePr"), "a", "graphicFrameLocks");
  if (!locksEl) {
    return undefined;
  }

  const locks: ImageFrameLocks = {};
  let present = false;
  for (const key of GRAPHIC_FRAME_LOCK_KEYS) {
    const value = parseOnOffValue(getAttribute(locksEl, null, GRAPHIC_FRAME_LOCK_ATTRIBUTES[key]));
    if (value !== undefined) {
      locks[key] = value;
      present = true;
    }
  }

  return present ? locks : undefined;
};

/**
 * Emit the whole `wp:cNvGraphicFramePr` element for regenerated DrawingML.
 *
 * Absent locks mean no authored frame was ever parsed (a Folio-created
 * picture), which keeps the historical `noChangeAspect="1"`.
 */
export const serializeGraphicFrameLocks = (locks: ImageFrameLocks | undefined): string => {
  const attrs = locks
    ? GRAPHIC_FRAME_LOCK_KEYS.flatMap((key) => {
        const value = locks[key];
        return value === undefined
          ? []
          : [`${GRAPHIC_FRAME_LOCK_ATTRIBUTES[key]}="${value ? "1" : "0"}"`];
      })
    : ['noChangeAspect="1"'];

  const attrList = attrs.length > 0 ? ` ${attrs.join(" ")}` : "";
  return `<wp:cNvGraphicFramePr><a:graphicFrameLocks ${DRAWINGML_NAMESPACE}${attrList}/></wp:cNvGraphicFramePr>`;
};
