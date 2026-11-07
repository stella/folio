import type { Node as ProseMirrorNode } from "prosemirror-model";

/**
 * The content-control lock rule (`@stll/docx-core/model`'s
 * `lockedControlViolation`) on a ProseMirror document: the locked controls a
 * document holds, as spans on its position scale.
 */
import {
  type ContentControlLock,
  firstLockedControlViolation,
  type LockedControlSpan,
  type LockedControlTouch,
  type LockedControlViolation,
} from "@stll/docx-core/model";

/** The node types that carry a content control's `w:lock`. */
const CONTROL_NODE_NAMES: ReadonlySet<string> = new Set(["sdt", "blockSdt"]);

const lockOf = (node: ProseMirrorNode): ContentControlLock | null => {
  if (!CONTROL_NODE_NAMES.has(node.type.name)) return null;
  const lock: unknown = node.attrs["lock"];
  switch (lock) {
    case "contentLocked":
    case "sdtContentLocked":
    case "sdtLocked":
      return lock;
    default:
      return null;
  }
};

/** Every control of `doc` whose lock forbids something, outermost first. */
export const lockedControlSpans = (
  doc: ProseMirrorNode,
): LockedControlSpan[] => {
  const spans: LockedControlSpan[] = [];
  doc.descendants((node, pos) => {
    const lock = lockOf(node);
    if (lock !== null)
      spans.push({
        lock,
        from: pos,
        to: pos + node.nodeSize,
        contentFrom: pos + 1,
        contentTo: pos + node.nodeSize - 1,
      });
    return true;
  });
  return spans;
};

/** What the locks of `doc` say about `touches`: `null` when they permit them all. */
export const lockedControlViolationIn = (
  doc: ProseMirrorNode,
  touches: readonly LockedControlTouch[],
): LockedControlViolation | null =>
  touches.length === 0
    ? null
    : firstLockedControlViolation(lockedControlSpans(doc), touches);
