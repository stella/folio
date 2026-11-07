/**
 * The content-control lock rule (`@stll/docx-core/model`'s
 * `lockedControlViolation`) on a ProseMirror document: the locked controls a
 * document holds, as spans on its position scale, and what a transform did to
 * them.
 */
import type { Node as ProseMirrorNode } from "prosemirror-model";
import type { Transform } from "prosemirror-transform";

import {
  type ContentControlChange,
  type ContentControlLock,
  contentControlLockAllows,
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
const spansByDoc = new WeakMap<ProseMirrorNode, readonly LockedControlSpan[]>();

export const lockedControlSpans = (
  doc: ProseMirrorNode,
): readonly LockedControlSpan[] => {
  const known = spansByDoc.get(doc);
  if (known !== undefined) return known;
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
  spansByDoc.set(doc, spans);
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

/** Marks that annotate content without being part of it. */
const ANNOTATION_MARKS: ReadonlySet<string> = new Set([
  "comment",
  "runIdentity",
]);
/** Attributes that identify a node without being part of its content. */
const IDENTITY_ATTRS: ReadonlySet<string> = new Set(["paraId", "textId"]);

const sameAttrs = (left: ProseMirrorNode, right: ProseMirrorNode): boolean => {
  const keys = new Set([
    ...Object.keys(left.attrs),
    ...Object.keys(right.attrs),
  ]);
  for (const key of keys) {
    if (IDENTITY_ATTRS.has(key)) continue;
    if (JSON.stringify(left.attrs[key]) !== JSON.stringify(right.attrs[key]))
      return false;
  }
  return true;
};

const contentMarks = (node: ProseMirrorNode) =>
  node.marks.filter((mark) => !ANNOTATION_MARKS.has(mark.type.name));

/** Whether two nodes hold the same content, annotations and identities aside. */
const sameLockedContent = (
  left: ProseMirrorNode,
  right: ProseMirrorNode,
): boolean => {
  if (left.type !== right.type) return false;
  if (left.isText && left.text !== right.text) return false;
  if (!sameAttrs(left, right)) return false;
  const leftMarks = contentMarks(left);
  const rightMarks = contentMarks(right);
  if (
    leftMarks.length !== rightMarks.length ||
    leftMarks.some((mark, index) => {
      const other = rightMarks[index];
      return other === undefined || !mark.eq(other);
    })
  )
    return false;
  return sameLockedChildren(left, right);
};

/** Whether two nodes' children hold the same content (their own attributes aside). */
function sameLockedChildren(
  left: ProseMirrorNode,
  right: ProseMirrorNode,
): boolean {
  if (left.childCount !== right.childCount) return false;
  for (let index = 0; index < left.childCount; index += 1)
    if (!sameLockedContent(left.child(index), right.child(index))) return false;
  return true;
}

/**
 * The node `span`'s control is in `transform`'s result, or `null` when it is
 * gone. A step may rewrite the stretch around a control with a copy of it, so
 * the control is looked for at either side of where its ends map to.
 */
const survivorOf = (
  transform: Transform,
  span: LockedControlSpan,
  before: ProseMirrorNode,
): ProseMirrorNode | null => {
  const { mapping, doc } = transform;
  for (const startSide of [1, -1] as const)
    for (const endSide of [-1, 1] as const) {
      const start = mapping.map(span.from, startSide);
      const end = mapping.map(span.to, endSide);
      const after = doc.nodeAt(start);
      if (after?.type === before.type && start + after.nodeSize === end)
        return after;
    }
  return null;
};

/**
 * What the locks of `transform`'s starting document say about its result: a
 * control whose content is locked must come through with the same content,
 * and one locked against removal must come through at all. Read from the
 * result rather than the steps, so it holds for any step a command takes.
 */
export const lockedControlViolationOfTransform = (
  transform: Transform,
): LockedControlViolation | null => {
  if (!transform.docChanged) return null;
  for (const span of lockedControlSpans(transform.before)) {
    const before = transform.before.nodeAt(span.from);
    if (before === null) continue;
    const after = survivorOf(transform, span, before);
    let change: ContentControlChange | null = null;
    if (after === null) change = "removeControl";
    else if (!sameLockedChildren(before, after)) change = "editContent";
    if (change !== null && !contentControlLockAllows(span.lock, change))
      return { lock: span.lock, change };
  }
  return null;
};
